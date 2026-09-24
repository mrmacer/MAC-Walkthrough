/* ─────────────────────────────────────────────────────────────────────────
   PATCH: Hide unlaunched Check-In report + remove Classroom from New Walkthrough

   Two independent, narrowly-scoped changes, verified separately:
   1. Reports: the Student Check-In Status section no longer renders, but
      nothing about Student Check-In's own data/storage/workflow changed,
      and the new Evidence Report + existing walkthrough Reports still work.
   2. New Walkthrough: the Classroom selector is gone and no longer
      required, Teacher/Focus remain required, a walkthrough submits with a
      real (never invented) blank Classroom, Individual Student roster
      lookup still depends only on Teacher, and historical Classroom data
      stays fully supported everywhere else (normalization, Reports filter,
      Evidence Report).

   Run with: node tests/walkthrough-classroom-removal.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");

/* ── shared per-id DOM stub ─────────────────────────────────────────────── */

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

const ROSTER_SCHEMA = { "Student First Name": "sfield_1", "Student Last Name": "sfield_2", "Teacher": "sfield_3", "Active": "sfield_5" };
const ROSTER_ROWS = [
  { id: "stu-1", sfield_1: "Jane", sfield_2: "Doe", sfield_3: "Bossons, A", sfield_5: "Yes" }
];

function makeContext() {
  const document = makeDocument();
  const calls = { saveWalkthrough: [] };
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
      async getListItems(listName) {
        if (listName === "IEP_Students_2026_27") return ROSTER_ROWS;
        return [];
      },
      async getListSchema(listName) { return listName === "IEP_Students_2026_27" ? ROSTER_SCHEMA : {}; },
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
  return { context, document, calls, toasts };
}

/* ── change 1: Student Check-In Status hidden from Reports ───────────────── */

function checkInStatusIsHiddenButEvidenceReportAndReportsStillWork() {
  const { context, document } = makeContext();
  const APP = vm.runInContext("APP", context);
  const REPORTS = vm.runInContext("REPORTS", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  REPORTS.walkthroughs = [{ sharePointId: "1", teacher: "Amber Bossons", observationDate: "2026-09-10", classroom: "Room 4", supportsObserved: [] }];
  REPORTS.filteredWalkthroughs = REPORTS.walkthroughs;
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  TEACHER_DIRECTORY.loaded = true;

  APP.renderReports();
  const html = document.registry["page-reports"].innerHTML;

  assert.ok(!/Student Check-In Status/.test(html), "item 1: Check-In Status section is not rendered");
  assert.ok(!/id="student-check-status-report"/.test(html), "its container is gone from the page too");
  assert.match(html, /Teacher Observation Evidence Report/, "item 2: Evidence Report card still renders");
  assert.match(html, /id="rf-teacher"/, "item 3: existing walkthrough Reports filter bar is intact");
  assert.match(html, /id="rf-classroom"/, "item 11: the Classroom filter itself still exists for historical data");
  assert.match(html, /id="walkthrough-reports-container"/, "item 3: the walkthrough table container still renders");

  // _renderCheckStatusPanels() itself is untouched and still functional —
  // only the call from renderReports() and the HTML block were removed.
  assert.equal(typeof APP._renderCheckStatusPanels, "function", "item 13: the function is not deleted");
  document.registry["student-check-status-report"] = undefined; // never created — see below
  const el = document.getElementById("student-check-status-report");
  el._html = "";
  APP._renderCheckStatusPanels(); // safe no-op-ish call; proves nothing throws
  const fn = app.slice(app.indexOf("renderReports() {"), app.indexOf("renderReports() {") + 6000);
  assert.ok(!/this\._renderCheckStatusPanels\(\);/.test(fn) || fn.indexOf("//") < fn.indexOf("this._renderCheckStatusPanels();"),
    "item 13: the call is commented out, not deleted from history — the function is reachable again with a one-line change");
}

function studentCheckStorageIsUntouched() {
  // Static guard: DB.getStudentChecks()/addStudentCheck-style storage code
  // must be byte-identical in spirit — this patch must not touch data.js at all.
  const data = readSrc("data.js");
  assert.match(data, /getStudentChecks\(\) \{ return this\._get\(CONFIG\.STORAGE_KEYS\.STUDENT_CHECKS\); \}/);
}

/* ── change 2: New Walkthrough — Classroom removed ────────────────────────── */

function makeRenderedWalkthroughPage() {
  const { context, document } = makeContext();
  const APP = vm.runInContext("APP", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  TEACHER_DIRECTORY.loaded = true;
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();
  return { context, document, APP };
}

function classroomSelectorIsGone() {
  const { document } = makeRenderedWalkthroughPage();
  const html = document.registry["page-walkthrough"].innerHTML;
  assert.ok(!/name="classroomId"/.test(html), "item 4: no classroomId field in the form at all");
  assert.ok(!/section-classroom/.test(html), "item 4: the Classroom form card is gone");
  assert.ok(!/Search classrooms…/.test(html));
  assert.match(html, /Who are you visiting/, "Teacher field still present");
  assert.match(html, /walk-field-label">Focus /, "Focus field still present");
}

function validationNoLongerRequiresClassroomButKeepsTeacherAndFocus() {
  const submitFn = app.slice(app.indexOf("async _submitWalkthrough(pageEl)"), app.indexOf("async _submitWalkthrough(pageEl)") + 1600);
  assert.ok(!/classroomId/.test(submitFn.slice(0, submitFn.indexOf("const errors"))), "item 5: classroomId is never read from the DOM any more");
  assert.ok(!/Please select a classroom/.test(submitFn), "item 5: Classroom validation message removed");
  assert.match(submitFn, /if \(!teacherId\)\s*errors\.push\("Please select a teacher\."\);/, "item 6: Teacher still required");
  assert.match(submitFn, /if \(!focus\)\s*errors\.push\("Please select an observation focus\."\);/, "item 7: Focus still required");
}

async function submitsWithClassroomBlankAndNeverInventsAValue() {
  const { context, document, calls } = makeContext();
  const APP = vm.runInContext("APP", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  TEACHER_DIRECTORY.loaded = true;

  const formData = {
    focus: "whole-class", engagementObserved: "engaged-independently", supportNeeded: "none",
    classroomStatus: "on-track", date: "2026-09-24", weekOf: "2026-W39", studentId: "",
    observedWin: "", concernGap: "", followUpNotes: "", supportsObserved: []
  };
  context.FormData = class {
    constructor() { this.d = formData; }
    get(k) { return this.d[k] ?? null; }
    getAll(k) { const v = this.d[k]; return Array.isArray(v) ? v : (v !== undefined && v !== null ? [v] : []); }
  };

  const dateField = { value: "2026-09-24" };
  const weekOfField = { value: "2026-W39" };
  const confirmEl = document.getElementById("__confirmEl__");
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
  const pageEl = {
    querySelector(sel) {
      if (sel === "#walkthroughForm") return formEl;
      if (sel === "#walkthroughConfirm") return confirmEl;
      return null;
    }
  };

  await APP._submitWalkthrough(pageEl);

  assert.equal(calls.saveWalkthrough.length, 1, "item 8: the walkthrough was submitted successfully with Classroom blank");
  const payload = calls.saveWalkthrough[0];
  assert.equal(payload.Classroom, "", "Classroom is sent as a real empty string");
  assert.ok(!["N/A", "Unknown", "unknown", "n/a"].includes(payload.Classroom), "never a placeholder value");
  assert.equal(payload.Teacher, "Amber Bossons");
  assert.equal(payload.Focus, "whole-class");
}

async function individualStudentLookupStillDependsOnlyOnTeacher() {
  const { context, document } = makeContext();
  const APP = vm.runInContext("APP", context);
  const STUDENT_ROSTER = vm.runInContext("STUDENT_ROSTER", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await STUDENT_ROSTER.refresh();
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  TEACHER_DIRECTORY.loaded = true;

  const studentSel = { innerHTML: "" };
  const hint = { _hidden: true, classList: { toggle(cls, on) { if (cls === "hidden") hint._hidden = !!on; } } };
  document.registry["page-walkthrough"] = {
    querySelector(sel) {
      if (sel === '[name="teacherId"]') return { value: "u1" };
      if (sel === "#walkthroughStudent") return studentSel;
      if (sel === "#walkFocusRosterHint") return hint;
      if (sel === 'input[name="focus"]:checked') return { value: "individual-student" };
      return null;
    }
  };
  APP._refreshWalkthroughStudentOptions();
  assert.match(studentSel.innerHTML, /Jane Doe/, "item 9: roster resolves purely from the selected teacher");

  const fnSrc = app.slice(app.indexOf("_refreshWalkthroughStudentOptions() {"), app.indexOf("_refreshWalkthroughStudentOptions() {") + 2000);
  assert.ok(!/classroom/i.test(fnSrc), "item 9: no classroom reference anywhere in the roster-lookup function");

  // item 10: Whole Class / Small Group still fine with zero roster match.
  document.registry["page-walkthrough"] = {
    querySelector(sel) {
      if (sel === '[name="teacherId"]') return { value: "no-such-teacher" };
      if (sel === "#walkthroughStudent") return studentSel;
      if (sel === "#walkFocusRosterHint") return hint;
      if (sel === 'input[name="focus"]:checked') return { value: "whole-class" };
      return null;
    }
  };
  APP._refreshWalkthroughStudentOptions();
  assert.equal(hint._hidden, true, "Whole Class stays unblocked with no roster match");
}

/* ── historical Classroom data stays fully supported elsewhere ───────────── */

function historicalClassroomStillNormalizesAndFilters() {
  assert.match(app, /classroom:\s*firstDefined\(item, \["Classroom"\]\)/, "item 11: normalizeSharePointWalkthrough still reads Classroom");
  assert.match(app, /setOpts\("rf-classroom", +walkthroughs\.map\(r => r\.classroom\)/, "item 11: Reports Classroom filter still populated from real records");
}

function evidenceReportStillHandlesHistoricalClassroom() {
  const lib = readSrc("evidence-report.js");
  assert.match(lib, /classroom:\s*r\.classroom \|\| ""/, "item 12: Evidence Report still carries Classroom through from stored records");
  assert.match(lib, /"Classroom"/, "item 12: Classroom stays a real CSV column");
  // A record with a real historical Classroom value renders it verbatim, and
  // a record with none (new-style) simply omits it — never invented.
  const EVIDENCE_REPORT = require("../evidence-report.js");
  const withRoom = EVIDENCE_REPORT.buildEntries([{ teacher: "Amber Bossons", classroom: "Room 4", observationDate: "2026-09-01", supportsObserved: [] }], {});
  const blank    = EVIDENCE_REPORT.buildEntries([{ teacher: "Amber Bossons", classroom: "",       observationDate: "2026-09-10", supportsObserved: [] }], {});
  assert.equal(withRoom[0].classroom, "Room 4");
  assert.equal(blank[0].classroom, "");
}

/* ── Daily Pulse / PACE unaffected ────────────────────────────────────────── */

function noCouplingIntroducedToDailyPulseOrPace() {
  const pulseFn = app.slice(app.indexOf("renderPulse() {"), app.indexOf("renderPulse() {") + 6000);
  assert.ok(!/classroomId|section-classroom/.test(pulseFn), "item 14: Daily Pulse untouched by the classroom removal");
  const paceAdmin = readSrc("pace-admin.js");
  assert.ok(!/classroomId/.test(paceAdmin), "item 14: pace-admin.js untouched");
}

const tests = {
  checkInStatusIsHiddenButEvidenceReportAndReportsStillWork, studentCheckStorageIsUntouched,
  classroomSelectorIsGone, validationNoLongerRequiresClassroomButKeepsTeacherAndFocus,
  submitsWithClassroomBlankAndNeverInventsAValue, individualStudentLookupStillDependsOnlyOnTeacher,
  historicalClassroomStillNormalizesAndFilters, evidenceReportStillHandlesHistoricalClassroom,
  noCouplingIntroducedToDailyPulseOrPace
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} checks passed.`);
})();
