/* ─────────────────────────────────────────────────────────────────────────
   PATCH: New Walkthrough Student — dropdown → optional free-text input

   Students move between classrooms often, so the live New Walkthrough
   Student field is now a plain optional text input instead of a dropdown
   populated from the selected teacher's roster (STUDENT_ROSTER /
   IEP_Students_2026_27). The typed value is trimmed and saved into the
   existing SharePoint Student column (via StudentName → "Student" in
   graph.js); blank is saved as an empty string. No roster lookup, no
   validation against the roster, never required for any Focus.

   Historical Student values, normalization, Reports, the Evidence Report
   exports, Daily Pulse and PACE are all unchanged.

   Run with: node tests/walkthrough-student-text.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");

/* ── shared per-id DOM stub (same shape as walkthrough-classroom-removal) ── */

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
        remove() {}, appendChild() {}, matches() { return false; }, focus() {}, blur() {},
        click() {}, reset() {}, scrollIntoView() {}, scrollTo() {}
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

function makeContext() {
  const document = makeDocument();
  const calls = { saveWalkthrough: [], getListItems: [] };
  const context = {
    console, document, navigator: {}, location: { origin: "http://localhost", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected raw fetch — GRAPH is mocked directly"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone, addEventListener() {}, removeEventListener() {}, confirm: () => true,
    TextEncoder, Uint8Array,
    GRAPH: {
      async getListItems(listName) { calls.getListItems.push(listName); return []; },
      async getListSchema() { return {}; },
      async saveWalkthrough(payload) { calls.saveWalkthrough.push(payload); return { duplicatePrevented: false }; }
    }
  };
  context.window = context;
  vm.createContext(context);
  ["config.js", "auth.js", "data.js", "pace-admin.js", "pace-export.js", "evidence-report.js", "app.js"].forEach(f =>
    vm.runInContext(readSrc(f), context, { filename: f }));
  vm.runInContext('MAC_ADMIN_PANEL_ALLOWED = true; AUTH.role = "Administrator";', context);
  const toasts = [];
  vm.runInContext("showToast = (m, t) => __toasts.push({ m, t });", Object.assign(context, { __toasts: toasts }));
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  TEACHER_DIRECTORY.loaded = true;
  return { context, document, calls, toasts };
}

function renderedWalkthroughHtml() {
  const ctx = makeContext();
  const APP = vm.runInContext("APP", ctx.context);
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();
  return { ...ctx, html: ctx.document.registry["page-walkthrough"].innerHTML };
}

// Drives the real _submitWalkthrough() with a mocked FormData.
async function submit(fields) {
  const ctx = makeContext();
  const APP = vm.runInContext("APP", ctx.context);
  const records = [];
  vm.runInContext("DB", ctx.context).addRecord = r => { records.push(r); return { responses: r }; };
  const formData = {
    focus: "whole-class", engagementObserved: "engaged-independently", supportNeeded: "none",
    classroomStatus: "on-track", date: "2026-10-01", weekOf: "2026-W40",
    observedWin: "", concernGap: "", followUpNotes: "", supportsObserved: [],
    ...fields
  };
  ctx.context.FormData = class {
    constructor() { this.d = formData; }
    get(k) { return this.d[k] ?? null; }
    getAll(k) { const v = this.d[k]; return Array.isArray(v) ? v : (v !== undefined && v !== null ? [v] : []); }
  };
  const dateField = { value: "2026-10-01" };
  const weekOfField = { value: "2026-W40" };
  const formEl = {
    querySelector(sel) {
      if (sel === '[name="teacherId"]') return { value: "u1" };
      if (sel === '[name="date"]') return dateField;
      if (sel === '[name="weekOf"]') return weekOfField;
      return { value: "" };
    },
    querySelectorAll() { return []; },
    reset() {}
  };
  const confirmEl = ctx.document.getElementById("__confirmEl__");
  const pageEl = {
    querySelector(sel) {
      if (sel === "#walkthroughForm") return formEl;
      if (sel === "#walkthroughConfirm") return confirmEl;
      return null;
    }
  };
  await APP._submitWalkthrough(pageEl);
  return { ...ctx, payload: ctx.calls.saveWalkthrough[0], record: records[0] };
}

/* ── UI ───────────────────────────────────────────────────────────────── */

function rendersTextInputForStudent() {
  const { html } = renderedWalkthroughHtml();
  const card = html.slice(html.indexOf("section-student"), html.indexOf("section-engagement"));
  assert.match(card, /<input type="text"[^>]*name="studentName"[^>]*id="walkthroughStudent"/, "item 1: Student is a text input");
  assert.match(card, /placeholder="Enter student name\.\.\."/, "item 1: placeholder");
  assert.match(card, /Student <span class="walk-optional-tag">optional<\/span>/, "item 3: labelled optional");
  assert.ok(!/\brequired\b/.test(card.replace(/never required/, "")), "item 3: no required attribute");
  assert.ok(!/ value="[^"]/.test(card), "never auto-filled");
}

function studentDropdownIsGone() {
  const { html } = renderedWalkthroughHtml();
  const card = html.slice(html.indexOf("section-student"), html.indexOf("section-engagement"));
  assert.ok(!/<select/.test(card), "item 2: no <select> in the Student card");
  assert.ok(!/None \/ whole class|No roster available/.test(html), "item 2: no dropdown options remain");
  assert.ok(!/name="studentId"/.test(html), "item 2: the old studentId field is gone");
}

function noRosterWarning() {
  const { html } = renderedWalkthroughHtml();
  assert.ok(!/No student roster has been added for this teacher/.test(html), "item 8: warning text gone");
  assert.ok(!/walkFocusRosterHint/.test(html), "item 8: warning container gone");
  assert.ok(!/walkFocusRosterHint|No student roster has been added/.test(app), "item 8: gone from source entirely");
}

function noRosterLookupForStudent() {
  const { calls } = renderedWalkthroughHtml();
  assert.ok(!calls.getListItems.includes("IEP_Students_2026_27"), "item 7: rendering never loads the roster");
  assert.ok(!/_refreshWalkthroughStudentOptions/.test(app), "item 7: roster-population function removed (it had no other callers)");
  const walkSrc = app.slice(app.indexOf("  renderWalkthrough() {"), app.indexOf("  async _submitWalkthrough(pageEl)"));
  const submitSrc = app.slice(app.indexOf("  async _submitWalkthrough(pageEl)"), app.indexOf("DB.addRecord(responses);"));
  const code = s => s.split("\n").filter(l => !l.trim().startsWith("//")).join("\n");
  assert.ok(!/STUDENT_ROSTER/.test(code(walkSrc)), "item 7: renderWalkthrough has no STUDENT_ROSTER dependency");
  assert.ok(!/STUDENT_ROSTER/.test(code(submitSrc)), "item 7: submit has no STUDENT_ROSTER dependency");
}

/* ── submission ───────────────────────────────────────────────────────── */

async function typedNameIsTrimmedAndSaved() {
  const { payload, toasts } = await submit({ studentName: "   Jordan  Q. Smith  " });
  assert.ok(payload, "item 4: walkthrough submitted");
  assert.equal(payload.StudentName, "Jordan  Q. Smith", "item 4: saved exactly as typed, trimmed");
  assert.ok(!toasts.some(t => t.t === "error"), "no validation error");
}

async function nameNotOnAnyRosterStillSaves() {
  const { payload } = await submit({ studentName: "Visiting Student From Room 12", focus: "individual-student" });
  assert.equal(payload.StudentName, "Visiting Student From Room 12", "item 7: not validated against any roster");
}

async function blankStudentSaves() {
  for (const v of [undefined, "", "    "]) {
    const { payload, toasts } = await submit(v === undefined ? {} : { studentName: v });
    assert.ok(payload, `item 5: submitted with studentName=${JSON.stringify(v)}`);
    assert.equal(payload.StudentName, "", "item 5: blank saved as empty string");
    assert.ok(!toasts.some(t => t.t === "error"));
  }
}

async function everyFocusSavesWithOrWithoutStudent() {
  for (const focus of ["whole-class", "small-group", "individual-student"]) {
    const blank = await submit({ focus, studentName: "" });
    assert.ok(blank.payload, `item 3/6: ${focus} saves with Student blank`);
    assert.equal(blank.payload.Focus, focus);
    const named = await submit({ focus, studentName: "Alex Lee" });
    assert.equal(named.payload.StudentName, "Alex Lee", `item 6: ${focus} saves a typed Student`);
  }
  const submitFn = app.slice(app.indexOf("async _submitWalkthrough(pageEl)"), app.indexOf("const responses", app.indexOf("async _submitWalkthrough(pageEl)")));
  assert.ok(!/errors\.push\([^)]*[Ss]tudent/.test(submitFn), "item 6: no Student validation message exists");
}

async function localRecordCarriesTypedName() {
  const { record } = await submit({ studentName: " Riley " });
  assert.equal(record.studentName, "Riley", "local record stores the trimmed typed name");
  assert.equal(record.studentId, "", "no roster id is invented for a typed name");
}

/* ── data compatibility / reporting ───────────────────────────────────── */

function noSharePointSchemaChange() {
  const graph = readSrc("graph.js");
  assert.match(graph, /"Student":\s+entry\.StudentName\s+\|\| entry\.Student \|\| ""/, "StudentName still maps to the existing Student column");
}

function historicalStudentStillNormalizes() {
  const { context } = makeContext();
  const normalize = vm.runInContext("normalizeSharePointWalkthrough", context);
  const r = normalize({ id: "7", Teacher: "Amber Bossons", Student: "Historic Roster Kid", ObservationDate: "2025-03-01" });
  assert.equal(r.student, "Historic Roster Kid", "item 9: historical Student value preserved");
  const blank = normalize({ id: "8", Teacher: "Amber Bossons", ObservationDate: "2026-10-01" });
  assert.equal(blank.student || "", "", "item 9: missing Student stays blank");
  assert.match(app, /setOpts\("rf-student",\s+walkthroughs\.map\(r => r\.student\)/, "item 9: Reports Student filter unchanged");
}

function evidenceExportsUnchanged() {
  const ER = require("../evidence-report.js");
  const recs = [
    { teacher: "Amber Bossons", student: "Historic Roster Kid", observationDate: "2025-03-01", supportsObserved: [] },
    { teacher: "Amber Bossons", student: "  Typed Name ", observationDate: "2026-10-01", supportsObserved: [] },
    { teacher: "Amber Bossons", student: "", observationDate: "2026-10-02", supportsObserved: [] }
  ];
  const shown = ER.buildEntries(recs, { redactStudents: false });
  assert.deepEqual(shown.map(e => e.student), ["Historic Roster Kid", "Typed Name", ""], "item 10: stored values used as-is");
  const redacted = ER.buildEntries(recs, {});
  assert.deepEqual(redacted.map(e => e.student), [ER.REDACTED_LABEL, ER.REDACTED_LABEL, ""], "item 10: redaction default unchanged");
  assert.match(ER.buildCsv(shown).split(/\r?\n/)[0], /Classroom,Student,Focus/, "item 10: CSV columns unchanged");
}

function dailyPulseAndPaceUnaffected() {
  assert.match(app, /<select class="form-select" name="studentId" id="paceStudent">/, "item 11: PACE still uses its roster dropdown");
  assert.match(app, /if \(!studentId\)\s+\{ showToast\("Please select a student\.", "error"\); return; \}/, "item 11: PACE still requires a student");
  assert.match(app, /const rosterStudent = STUDENT_ROSTER\.find\(studentId\);/, "item 11: PACE still resolves from STUDENT_ROSTER");
  assert.match(app, /STUDENT_ROSTER\.getDailyPulseForTeacher\(/, "item 11: Daily Pulse roster untouched");
  assert.match(app, /const STUDENT_ROSTER = \{/, "STUDENT_ROSTER retained globally");
  assert.match(app, /function resolveCanonicalTeacherValue\(/, "Setup → Students helper retained");
}

const tests = {
  rendersTextInputForStudent, studentDropdownIsGone, noRosterWarning, noRosterLookupForStudent,
  typedNameIsTrimmedAndSaved, nameNotOnAnyRosterStillSaves, blankStudentSaves,
  everyFocusSavesWithOrWithoutStudent, localRecordCarriesTypedName,
  noSharePointSchemaChange, historicalStudentStillNormalizes, evidenceExportsUnchanged,
  dailyPulseAndPaceUnaffected
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} walkthrough Student text-input checks passed.`);
})();
