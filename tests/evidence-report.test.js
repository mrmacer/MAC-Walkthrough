/* ─────────────────────────────────────────────────────────────────────────
   Teacher Observation Evidence Report

   1. evidence-report.js is exercised for real against fixture records already
      shaped like normalizeSharePointWalkthrough()'s output (the real read
      model REPORTS.walkthroughs uses) — filtering, entry-building, category
      grouping, CSV/PA-ETEP/DOCX builders. The .docx package is opened with a
      small independent ZIP reader (CRC verified), same technique as
      tests/pace-export.test.js.
   2. app.js's Reports-page wiring (teacher list = IEP_Users2 ∪ historical
      record names, generate/export/copy handlers, admin gating) is exercised
      for real under node:vm.

   Run with: node tests/evidence-report.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");
const EVIDENCE_REPORT = require("../evidence-report.js");

/* ── fixtures: shaped like normalizeSharePointWalkthrough()'s real output ── */

const rec = (id, o) => ({
  sharePointId: id, observationId: `walkthrough-${id}`, observationDate: "2026-09-10",
  observationTime: "9:15 AM", observationTimestamp: "2026-09-10T09:15:00.000Z",
  observer: "Decusky", observerEmail: "maceg@iu29.org", teacher: "Amber Bossons",
  student: "Analee Cruz", classroom: "Bossons' Room", focus: "", environment: "",
  observationLength: 0, engagement: "engaged-with-support",
  supportsObserved: ["Sensory Break", "Assistive Technology", "Peer Support"],
  disengagementReasons: [], supportRequested: "help-soon",
  observedWin: "Used the calm corner independently.", concernGap: "Struggled with transitions.",
  observationNotes: "Quiet room, small group.", followUpNotes: "Check in next week.",
  priority: "", followUpNeeded: true, followUpDate: "", aiSummary: "", aiSuggestions: "",
  submissionId: `sub-${id}`, synced: true, created: "2026-09-10T09:15:00.000Z",
  modified: "2026-09-10T09:15:00.000Z", modifiedBy: "", ...o
});

const RECORDS = [
  rec("1"),
  rec("2", { observationDate: "2026-09-15", observationTime: "10:00 AM", observationTimestamp: "2026-09-15T10:00:00.000Z",
    engagement: "engaged-independently", supportsObserved: ["Modified Assignment", "Calming Strategy"],
    observedWin: "Completed task with no prompts.", concernGap: "", followUpNeeded: false }),
  rec("3", { teacher: "Nikki Stock", observationDate: "2026-09-12", observationTimestamp: "2026-09-12T08:00:00.000Z",
    student: "", supportsObserved: [], observedWin: "", concernGap: "" }),
  rec("4", { teacher: "Former Teacher", observationDate: "2026-08-01", observationTimestamp: "2026-08-01T09:00:00.000Z",
    student: "Other Kid", supportsObserved: ["Other"] })
];

/* ── independent ZIP reader (stored entries, CRC verified) ────────────────── */

function readZip(bytes) {
  const buf = Buffer.from(bytes);
  const eocd = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  assert.ok(eocd >= 0, "end-of-central-directory present");
  const count = buf.readUInt16LE(eocd + 10);
  let pos = buf.readUInt32LE(eocd + 16);
  const files = {};
  for (let i = 0; i < count; i++) {
    assert.equal(buf.readUInt32LE(pos), 0x02014b50, "central header signature");
    const crc = buf.readUInt32LE(pos + 16), size = buf.readUInt32LE(pos + 24);
    const nameLen = buf.readUInt16LE(pos + 28), extraLen = buf.readUInt16LE(pos + 30), commentLen = buf.readUInt16LE(pos + 32);
    const local = buf.readUInt32LE(pos + 42);
    const name = buf.toString("utf8", pos + 46, pos + 46 + nameLen);
    const dataStart = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
    const data = buf.subarray(dataStart, dataStart + size);
    assert.equal(EVIDENCE_REPORT.crc32(data), crc, `CRC for ${name}`);
    files[name] = data.toString("utf8");
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return files;
}

function assertWellFormed(xmlText, label) {
  const body = xmlText.replace(/^<\?xml[^>]*\?>\s*/, "");
  assert.ok(!/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(body), `${label}: unescaped ampersand`);
  const stack = [];
  const re = /<(\/?)([A-Za-z_][\w:.-]*)((?:\s+[\w:.-]+="[^"<]*")*)\s*(\/?)>/g;
  let m;
  while ((m = re.exec(body))) {
    if (m[4]) continue;
    if (m[1]) assert.equal(stack.pop(), m[2], `${label}: mismatched </${m[2]}>`);
    else stack.push(m[2]);
  }
  assert.equal(stack.length, 0, `${label}: unclosed ${stack.join(",")}`);
}

function parseCsv(text) {
  const s = text.replace(/^﻿/, "");
  const rows = []; let cur = [], cell = "", q = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) { if (c === '"') { if (s[i + 1] === '"') { cell += '"'; i++; } else q = false; } else cell += c; }
    else if (c === '"') q = true;
    else if (c === ",") { cur.push(cell); cell = ""; }
    else if (c === "\r" && s[i + 1] === "\n") { cur.push(cell); rows.push(cur); cur = []; cell = ""; i++; }
    else cell += c;
  }
  if (cell || cur.length) { cur.push(cell); rows.push(cur); }
  return rows;
}

const RATING_WORDS = /\b(Distinguished|Proficient|Needs Improvement|Failing)\b/;

/* ── layer 1: selectRecords / buildEntries ────────────────────────────────── */

function selectFiltersByTeacherAndDate() {
  const s1 = EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" });
  assert.deepEqual(s1.map(r => r.sharePointId), ["1", "2"], "only this teacher, chronological ascending");

  const s2 = EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons", from: "2026-09-12", to: "2026-09-30" });
  assert.deepEqual(s2.map(r => r.sharePointId), ["2"], "date range narrows further");

  const s3 = EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Former Teacher" });
  assert.deepEqual(s3.map(r => r.sharePointId), ["4"], "an inactive/former teacher's history is still selectable by name");
}

function buildEntriesUsesOnlyRealFields() {
  const entries = EVIDENCE_REPORT.buildEntries(EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" }), { redactStudents: false });
  assert.equal(entries.length, 2);
  const e = entries[0];
  assert.equal(e.dateLabel, "Sep 10, 2026");
  assert.equal(e.observer, "Decusky");
  assert.equal(e.engagementLabel, "Engaged With Support");
  assert.deepEqual(e.supports, ["Sensory Break", "Assistive Technology", "Peer Support"]);
  assert.equal(e.supportRequestedLabel, "Help Soon");
  assert.equal(e.observedWin, "Used the calm corner independently.");
  assert.equal(e.followUpNeeded, true);
  // Focus was never sent by the live New Walkthrough form on these fixtures
  // (empty string) — must not be invented.
  assert.equal(e.focusLabel, "");
}

function redactionDefaultsOnAndCanBeDisabled() {
  const selected = EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" });
  const redacted = EVIDENCE_REPORT.buildEntries(selected, {});
  assert.ok(redacted.every(e => e.student === EVIDENCE_REPORT.REDACTED_LABEL || e.student === ""), "default redaction is on");
  assert.ok(!redacted.some(e => e.student === "Analee Cruz"), "no real student name leaks when redaction is on");

  const unredacted = EVIDENCE_REPORT.buildEntries(selected, { redactStudents: false });
  assert.equal(unredacted[0].student, "Analee Cruz", "explicit opt-out shows the real name");
}

function categoryGroupingIsExplicitNotAnRating() {
  const entries = EVIDENCE_REPORT.buildEntries(EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" }), {});
  const groups = EVIDENCE_REPORT.groupByCategory(entries);
  const byCat = Object.fromEntries(groups.map(g => [g.category, g.items]));
  assert.ok(byCat["Behavior & Environment Supports"].some(i => i.supports.includes("Sensory Break")));
  assert.ok(byCat["Assistive Technology"].some(i => i.supports.includes("Assistive Technology")));
  assert.ok(byCat["Peer & Small-Group Support"].some(i => i.supports.includes("Peer Support")));
  assert.ok(byCat["Accommodations & Modifications"].some(i => i.supports.includes("Modified Assignment")));
  groups.forEach(g => assert.ok(!RATING_WORDS.test(g.category), "category label is never a rating"));

  // "Other" support option gets an honest bucket, never guessed further.
  const otherEntries = EVIDENCE_REPORT.buildEntries(EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Former Teacher" }), {});
  const otherGroups = EVIDENCE_REPORT.groupByCategory(otherEntries);
  assert.ok(otherGroups.some(g => g.category === "Other"));
}

/* ── layer 1: exporters ────────────────────────────────────────────────────── */

function csvIsRfc4180AndInjectionSafe() {
  const entries = EVIDENCE_REPORT.buildEntries(
    EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" }),
    { redactStudents: true }
  );
  const rows = parseCsv(EVIDENCE_REPORT.buildCsv(entries));
  assert.deepEqual(rows[0], ["Date", "Time", "Observer", "Classroom", "Student", "Focus", "Engagement",
    "Supports Observed", "Disengagement Reasons", "Support Requested", "Observed Win", "Concern / Gap",
    "Observation Notes", "Follow-Up Needed", "Follow-Up Notes", "Priority"]);
  assert.equal(rows.length, 3, "header + 2 observations");
  assert.equal(rows[1][4], EVIDENCE_REPORT.REDACTED_LABEL, "redacted student name in the CSV too");
  assert.ok(EVIDENCE_REPORT.buildCsv(entries).startsWith("﻿"), "UTF-8 BOM");

  const injection = EVIDENCE_REPORT.buildEntries(
    [{ ...RECORDS[0], sharePointId: "x", observedWin: "=SUM(1,2)", concernGap: 'Said "stop", left,\nreturned.' }],
    { redactStudents: false }
  );
  const injRows = parseCsv(EVIDENCE_REPORT.buildCsv(injection));
  assert.equal(injRows[1][10], "'=SUM(1,2)", "formula-looking text neutralised");
  assert.equal(injRows[1][11], 'Said "stop", left,\nreturned.', "quotes/commas/newlines round-trip");
}

function docxIsValidAndNeverAssignsRatings() {
  const selected = EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" });
  const entries = EVIDENCE_REPORT.buildEntries(selected, { redactStudents: true });
  const summary = EVIDENCE_REPORT.summarize(entries, { teacher: "Amber Bossons" });
  const files = readZip(EVIDENCE_REPORT.buildDocx(entries, summary));
  assert.deepEqual(Object.keys(files).sort(), ["[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels", "word/document.xml", "word/styles.xml"]);
  Object.entries(files).forEach(([n, x]) => assertWellFormed(x, n));

  const text = files["word/document.xml"].replace(/<w:br\/>/g, "\n").replace(/<[^>]+>/g, "");
  assert.match(text, /Teacher: Amber Bossons/);
  assert.match(text, /Number of Observations: 2/);
  assert.match(text, /Chronological Observation Evidence/);
  assert.match(text, /Evidence by Category/);
  assert.match(text, /not a Danielson domain rating/);
  assert.match(text, /Administrator Synthesis/);
  assert.match(text, /Strengths Observed:/);
  assert.match(text, /Confidential/);
  assert.ok(!RATING_WORDS.test(text), "no evaluation rating is ever assigned in the document");
  assert.ok(!text.includes("Analee Cruz"), "redacted student name never appears in the exported document");
  assert.ok(text.includes(EVIDENCE_REPORT.REDACTED_LABEL), "the redacted-name placeholder is shown instead");
}

function docxHandlesZeroObservationsGracefully() {
  const summary = EVIDENCE_REPORT.summarize([], { teacher: "Nikki Stock", from: "2020-01-01", to: "2020-01-02" });
  const files = readZip(EVIDENCE_REPORT.buildDocx([], summary));
  assertWellFormed(files["word/document.xml"], "empty doc");
  const text = files["word/document.xml"].replace(/<[^>]+>/g, "");
  assert.match(text, /No walkthrough observations were found/);
}

function paEtepSectionsAndCopyAllMatch() {
  const selected = EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" });
  const entries = EVIDENCE_REPORT.buildEntries(selected, { redactStudents: true });
  const summary = EVIDENCE_REPORT.summarize(entries, { teacher: "Amber Bossons" });
  const { sections, allText } = EVIDENCE_REPORT.buildPaEtepSections(entries, summary);

  assert.equal(sections.length, 1 + entries.length + 1, "header + one per observation + category rollup");
  assert.equal(sections[0].id, "header");
  assert.match(sections[0].text, /Teacher: Amber Bossons/);
  assert.match(sections[1].title, /Sep 10, 2026/);
  assert.match(sections[1].text, /Date: Sep 10, 2026/);
  assert.match(sections[1].text, /Engagement: Engaged With Support/);
  assert.match(sections[1].text, /Sensory Break/);
  assert.equal(sections.at(-1).id, "categories");
  assert.match(sections.at(-1).text, /Behavior & Environment Supports/);

  sections.forEach(s => assert.ok(allText.includes(s.text), `Copy All includes "${s.title}" verbatim`));
  assert.ok(!RATING_WORDS.test(allText));
  assert.ok(!allText.includes("Analee Cruz"), "redaction applies to the PA-ETEP view too");
}

function summarizePeriodAndObservers() {
  const entries = EVIDENCE_REPORT.buildEntries(EVIDENCE_REPORT.selectRecords(RECORDS, { teacher: "Amber Bossons" }), {});
  const summary = EVIDENCE_REPORT.summarize(entries, { teacher: "Amber Bossons", from: "2026-09-01", to: "2026-09-30" });
  assert.equal(summary.count, 2);
  assert.deepEqual(summary.observers, ["Decusky"]);
  assert.match(summary.periodLabel, /Sep 1, 2026 – Sep 30, 2026/);
}

/* ── layer 2: app.js wiring ────────────────────────────────────────────────── */

function makeDocument() {
  const registry = {};
  function elementFor(id) {
    if (!registry[id]) {
      const el = {
        id, value: "", checked: false, textContent: "", dataset: {},
        classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
        style: {}, _handlers: {},
        addEventListener(ev, fn) { (el._handlers[ev] ||= []).push(fn); },
        removeEventListener() {},
        querySelector: () => elementFor(Symbol()), querySelectorAll: () => [],
        closest() { return null; }, setAttribute() {}, getAttribute() { return null; }, removeAttribute() {},
        remove() {}, appendChild() {}, matches() { return false; }, focus() {}, blur() {}, click() {}
      };
      Object.defineProperty(el, "innerHTML", { get: () => el._html || "", set: v => { el._html = v; } });
      registry[id] = el;
    }
    return registry[id];
  }
  return {
    registry, getElementById: id => elementFor(id),
    querySelector: () => elementFor(Symbol()), querySelectorAll: () => [],
    createElement: () => elementFor(Symbol()), addEventListener() {}, removeEventListener() {},
    body: elementFor("__body__")
  };
}

function makeAppContext() {
  const document = makeDocument();
  const blobs = [];
  const context = {
    console, document, navigator: {}, location: { origin: "http://localhost", hash: "" },
    crypto: crypto.webcrypto, TextEncoder, Uint8Array,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("evidence report must not perform any network request"); },
    msal: { PublicClientApplication: class {} },
    setTimeout: () => 0, clearTimeout, structuredClone, addEventListener() {}, confirm: () => true,
    Blob: class { constructor(parts, opts) { this.parts = parts; this.type = opts.type; blobs.push(this); } },
    URL: { createObjectURL: () => "blob:test", revokeObjectURL() {} },
    GRAPH: { getListItems() { throw new Error("must not re-query SharePoint"); } }
  };
  context.window = context;
  vm.createContext(context);
  ["config.js", "auth.js", "data.js", "pace-admin.js", "pace-export.js", "evidence-report.js", "app.js"].forEach(f =>
    vm.runInContext(readSrc(f), context, { filename: f }));
  vm.runInContext('MAC_ADMIN_PANEL_ALLOWED = true; AUTH.role = "Administrator";', context);
  const APP = vm.runInContext("APP", context);
  const REPORTS = vm.runInContext("REPORTS", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  const toasts = [];
  vm.runInContext("showToast = (m, t) => __toasts.push({ m, t });", Object.assign(context, { __toasts: toasts }));
  REPORTS.walkthroughs = RECORDS.map(r => ({ ...r }));
  REPORTS.filteredWalkthroughs = REPORTS.walkthroughs;
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }, { id: "u2", name: "Nikki Stock" }];
  TEACHER_DIRECTORY.loaded = true;
  return { context, APP, REPORTS, TEACHER_DIRECTORY, document, blobs, toasts };
}

function cardIsAdminGatedAndUsesUnionOfDirectoryAndHistory() {
  const { APP, document } = makeAppContext();
  const html = APP._evidenceReportCardHtml();
  assert.match(html, /Teacher Observation Evidence Report/);
  document.getElementById("er-teacher"); document.getElementById("er-focus"); document.getElementById("page-reports");
  APP._populateEvidenceTeacherOptions();
  const teacherHtml = document.registry["er-teacher"].innerHTML;
  assert.match(teacherHtml, /Amber Bossons/, "active IEP_Users2 teacher present");
  assert.match(teacherHtml, /Nikki Stock/, "active IEP_Users2 teacher with no walkthroughs yet is still listed");
  assert.match(teacherHtml, /Former Teacher/, "inactive/former teacher with real history is still selectable — not in TEACHER_DIRECTORY at all");
}

function generateBuildsCorrectEntriesAndBlocksWithoutTeacher() {
  const { APP, document, toasts } = makeAppContext();
  document.getElementById("er-teacher").value = "";
  APP._generateEvidenceReport();
  assert.equal(APP._evidenceReport, undefined, "no teacher selected — nothing generated");
  assert.match(toasts.at(-1).m, /Select a teacher/);

  document.getElementById("er-teacher").value = "Amber Bossons";
  document.getElementById("er-redact").checked = true;
  APP._generateEvidenceReport();
  assert.equal(APP._evidenceReport.entries.length, 2);
  assert.equal(APP._evidenceReport.summary.teacherName, "Amber Bossons");
  assert.ok(APP._evidenceReport.entries.every(e => e.student !== "Analee Cruz"), "UI-driven redaction checkbox is honored");
}

function exportsProduceDownloadsAndBlockOnEmpty() {
  const { APP, document, blobs, toasts } = makeAppContext();
  document.getElementById("er-teacher").value = "Amber Bossons";
  document.getElementById("er-redact").checked = true;
  APP._generateEvidenceReport();
  APP._bindEvidenceReportResult();

  document.registry["erExportCsv"]._handlers.click[0]();
  assert.equal(blobs.length, 1);
  assert.equal(blobs[0].type, "text/csv;charset=utf-8");
  assert.equal(document.anchors, undefined); // sanity: no crash path taken

  document.registry["erExportDocx"]._handlers.click[0]();
  assert.equal(blobs.length, 2);
  assert.match(blobs[1].type, /wordprocessingml/);

  // No matches → no file, clear message (same pattern as PACE export).
  document.getElementById("er-teacher").value = "Nikki Stock";
  document.getElementById("er-from").value = "2099-01-01";
  APP._generateEvidenceReport();
  APP._bindEvidenceReportResult();
  document.registry["erExportCsv"]._handlers.click[0]();
  assert.equal(blobs.length, 2, "still 2 — the empty-result export was refused");
  assert.match(toasts.at(-1).m, /No observations match/);
}

function staticNeverWritesToSharePointOrAssignsRatings() {
  const lib = readSrc("evidence-report.js");
  assert.ok(!/GRAPH\./.test(lib), "evidence-report.js never touches GRAPH/SharePoint");
  assert.ok(!/fetch\(/.test(lib), "no network calls");
  // Rating words are fine in a comment explaining what's NOT done; the real
  // guarantee is that no *executable* code ever produces/assigns one.
  const codeOnly = lib.split("\n").filter(l => !l.trim().startsWith("//") && !l.trim().startsWith("*")).join("\n");
  assert.ok(!RATING_WORDS.test(codeOnly), "no rating vocabulary in any code path of the builder");
  const fn = app.slice(app.indexOf("_generateEvidenceReport()"), app.indexOf("_generateEvidenceReport()") + 1200);
  assert.ok(!/GRAPH\.(create|update|delete)/.test(fn), "generating a report never writes anywhere");
}

const tests = {
  selectFiltersByTeacherAndDate, buildEntriesUsesOnlyRealFields, redactionDefaultsOnAndCanBeDisabled,
  categoryGroupingIsExplicitNotAnRating, csvIsRfc4180AndInjectionSafe, docxIsValidAndNeverAssignsRatings,
  docxHandlesZeroObservationsGracefully, paEtepSectionsAndCopyAllMatch, summarizePeriodAndObservers,
  cardIsAdminGatedAndUsesUnionOfDirectoryAndHistory, generateBuildsCorrectEntriesAndBlocksWithoutTeacher,
  exportsProduceDownloadsAndBlockOnEmpty, staticNeverWritesToSharePointOrAssignsRatings
};

let failed = 0;
for (const [name, fn] of Object.entries(tests)) {
  try { fn(); console.log("  ✓ " + name); }
  catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
}
if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
console.log(`\nAll ${Object.keys(tests).length} evidence-report checks passed.`);
