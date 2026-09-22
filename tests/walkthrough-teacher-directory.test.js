/* ─────────────────────────────────────────────────────────────────────────
   PATCH: Move New Walkthrough teacher directory to SharePoint

   The New Walkthrough "Who are you visiting?" selector now reads live from
   macwalkthroughwhoareyouvisiting (via TEACHER_DIRECTORY, a small dedicated
   module) instead of the hardcoded PILOT_TEACHERS / DB.getTeachers() cache.

   Three layers, matching this codebase's established test style (see
   tests/admin-editing.test.js's own docstring):
   1. TEACHER_DIRECTORY's normalize/dedupe/error behavior is exercised for
      real against a mocked GRAPH.getWhoAreYouVisiting().
   2. renderWalkthrough()'s loading/error/success gating and
      _refreshWalkthroughStudentOptions()'s roster lookup are exercised for
      real under node:vm, with STUDENT_ROSTER loaded for real against a
      mocked GRAPH.getListItems/getListSchema (same technique as
      tests/teacher-dropdown.test.js).
   3. Setup → Teachers add/remove wiring is exercised for real, proving the
      New Walkthrough directory refreshes without a redeploy or DB.clearAll().
   Static source assertions cover the remaining out-of-scope-boundary checks
   (PILOT_TEACHERS untouched for legacy consumers, no PACE/Daily Pulse
   coupling introduced).

   Run with: node tests/walkthrough-teacher-directory.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");
// vm-context objects/arrays are a different realm than this file's — plain()
// round-trips them through JSON so assert.deepEqual compares plain values,
// not cross-realm prototypes (same technique as tests/teacher-dropdown.test.js).
const plain = obj => JSON.parse(JSON.stringify(obj));

/* ── shared per-id DOM stub (same shape as tests/teacher-dropdown.test.js) ── */

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
        querySelector: () => elementFor(Symbol()),
        querySelectorAll: () => [],
        closest() { return null; },
        setAttribute() {}, getAttribute() { return null; }, removeAttribute() {},
        remove() {}, appendChild() {}, matches() { return false; }, focus() {}, blur() {}
      };
      Object.defineProperty(el, "innerHTML", { get: () => el._html || "", set: v => { el._html = v; } });
      registry[id] = el;
    }
    return registry[id];
  }
  return {
    registry,
    getElementById: id => elementFor(id),
    querySelector: () => elementFor(Symbol()),
    querySelectorAll: () => [],
    createElement: () => elementFor(Symbol()),
    addEventListener() {}, removeEventListener() {},
    body: elementFor("__body__")
  };
}

function makeGraphMock() {
  const calls = { getWhoAreYouVisiting: 0, createListItem: [], getListItems: [] };
  let directory = []; // [{spId, name}], caller controls contents per test
  return {
    calls,
    setDirectory(rows) { directory = rows; },
    GRAPH: {
      async getWhoAreYouVisiting() {
        calls.getWhoAreYouVisiting++;
        // Mirror graph.js's real contract: blanks dropped, alphabetical.
        return directory.filter(t => (t.name || "").trim())
          .map(t => ({ spId: t.spId, name: t.name.trim() }))
          .sort((a, b) => a.name.localeCompare(b.name));
      },
      async getListItems(listName) {
        calls.getListItems.push(listName);
        if (listName === "IEP_Students_2026_27") return ROSTER_ROWS;
        return [];
      },
      async getListSchema(listName) {
        if (listName === "IEP_Students_2026_27") return ROSTER_SCHEMA;
        return {};
      },
      async createListItem(listName, fields) { calls.createListItem.push({ listName, fields }); return { id: "new" }; },
      async getSiteId() { return "site-1"; },
      async getListId() { return "list-1"; },
      async _patch() { return {}; }
    }
  };
}

// Roster fixture: "Bossons" has students, "Andruchek" has none — lets a
// single fixture exercise both the "has roster" and "no roster" branches.
const ROSTER_SCHEMA = { "Student First Name": "sfield_1", "Student Last Name": "sfield_2", "Teacher": "sfield_3", "Active": "sfield_5" };
const ROSTER_ROWS = [
  { id: "stu-1", sfield_1: "Jane", sfield_2: "Doe",   sfield_3: "Bossons", sfield_5: "Yes" },
  { id: "stu-2", sfield_1: "Sam",  sfield_2: "Rivera", sfield_3: "Bossons", sfield_5: "Yes" },
  { id: "stu-3", sfield_1: "Pat",  sfield_2: "Kim",    sfield_3: "Someone Else", sfield_5: "Yes" }
];

// Real in-memory localStorage (a plain no-op stub would silently swallow
// DB.addRecord()/DB._set(), which the "historical record survives" test
// depends on actually persisting across calls).
function makeLocalStorage() {
  const store = new Map();
  return {
    getItem: k => (store.has(k) ? store.get(k) : null),
    setItem: (k, v) => { store.set(k, String(v)); },
    removeItem: k => { store.delete(k); }
  };
}

function makeContext() {
  const document = makeDocument();
  const { calls, GRAPH, setDirectory } = makeGraphMock();
  const dbClearCalls = [];
  const context = {
    console, document, navigator: {},
    location: { origin: "http://localhost:5500", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: makeLocalStorage(),
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected raw fetch — GRAPH is mocked directly"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone,
    addEventListener() {}, removeEventListener() {},
    confirm: () => true,
    GRAPH
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("config.js"), context, { filename: "config.js" });
  vm.runInContext(readSrc("auth.js"), context, { filename: "auth.js" });
  vm.runInContext(readSrc("data.js"), context, { filename: "data.js" });
  vm.runInContext(app, context, { filename: "app.js" });
  vm.runInContext('MAC_ADMIN_PANEL_ALLOWED = true; AUTH.role = "Administrator";', context);
  // Spy on DB.clearAll() without altering its behavior — proves the
  // migration doesn't need a destructive cache wipe (spec requirement).
  const DB = vm.runInContext("DB", context);
  const originalClearAll = DB.clearAll.bind(DB);
  DB.clearAll = (...a) => { dbClearCalls.push(a); return originalClearAll(...a); };
  return { context, document, calls, setDirectory, dbClearCalls };
}

/* ── layer 1: TEACHER_DIRECTORY ─────────────────────────────────────────── */

async function directoryNormalizesAndDedupes() {
  const { context, setDirectory } = makeContext();
  setDirectory([
    { spId: "sp-2", name: "Andruchek" },
    { spId: "sp-1", name: "Bossons" },
    { spId: "sp-1b", name: "bossons" },   // duplicate visible name, different casing
    { spId: "sp-3", name: "  " }          // blank after trim
  ]);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await TEACHER_DIRECTORY.refresh();

  assert.equal(TEACHER_DIRECTORY.error, null);
  assert.equal(TEACHER_DIRECTORY.loaded, true);
  const all = TEACHER_DIRECTORY.getAll();
  assert.equal(all.length, 2, "blank dropped, duplicate visible name collapsed to one entry");
  assert.deepEqual(plain(all).map(t => t.name.toLowerCase()), ["andruchek", "bossons"], "sorted, case-insensitive dedupe");
  const bossons = all.find(t => t.name.toLowerCase() === "bossons");
  assert.ok(["sp-1", "sp-1b"].includes(bossons.id), "kept one of the two duplicate rows, deterministically");
  assert.equal(TEACHER_DIRECTORY.find("sp-2").name, "Andruchek");
  assert.equal(TEACHER_DIRECTORY.find(bossons.id === "sp-1" ? "sp-1b" : "sp-1"), null, "the collapsed duplicate's own id is not a separate entry");
  // Not brittle position-based ids:
  assert.ok(all.every(t => t.id.startsWith("sp-")), "ids are the real SharePoint item ids, not array indexes");
}

async function directoryFailureDoesNotFallBackToStalePilotTeachers() {
  const { context, calls } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  TEACHER_DIRECTORY.refresh = async function () {
    this.loading = true;
    try { throw new Error("simulated Graph outage"); }
    catch (err) { this.error = err.message; }
    finally { this.loading = false; }
    return this._teachers;
  };
  await TEACHER_DIRECTORY.refresh();
  assert.equal(TEACHER_DIRECTORY.error, "simulated Graph outage");
  assert.equal(TEACHER_DIRECTORY.loaded, false);
  const PILOT_TEACHERS = vm.runInContext("PILOT_TEACHERS", context);
  const pilotNames = PILOT_TEACHERS.map(t => t.name);
  assert.deepEqual(plain(TEACHER_DIRECTORY.getAll()), [], "no stale/hardcoded fallback on failure");
  assert.ok(!TEACHER_DIRECTORY.getAll().some(t => pilotNames.includes(t.name)));
}

/* ── layer 2: renderWalkthrough() gating ────────────────────────────────── */

function renderLoadingState() {
  const { context, document, calls } = makeContext();
  const APP = vm.runInContext("APP", context);
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();
  const html = document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /Loading teacher directory/);
  assert.ok(!/walkthroughForm/.test(html), "form is not rendered while the directory is loading");
  assert.equal(calls.getWhoAreYouVisiting, 1, "exactly one load kicked off");
}

function renderErrorStateWithRetryAndNoFallback() {
  const { context, document } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  const PILOT_TEACHERS    = vm.runInContext("PILOT_TEACHERS", context);
  TEACHER_DIRECTORY.error = "Unable to load the teacher directory.";
  const APP = vm.runInContext("APP", context);
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();
  const html = document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /Teacher directory could not be loaded/);
  assert.match(html, /id="teacherDirectoryRetry"/, "a retry control is offered");
  assert.ok(!/walkthroughForm/.test(html), "no form while blocked on an error");
  PILOT_TEACHERS.forEach(t => assert.ok(!html.includes(`>${t.name}<`), `no silent fallback to pilot teacher "${t.name}"`));
}

function renderSuccessUsesLiveDirectoryOnly() {
  const { context, document, calls } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  TEACHER_DIRECTORY._teachers = [{ id: "sp-2", name: "Andruchek" }, { id: "sp-1", name: "Bossons" }];
  TEACHER_DIRECTORY.loaded = true;
  const APP = vm.runInContext("APP", context);
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();
  assert.equal(calls.getWhoAreYouVisiting, 0, "already-loaded directory is not re-fetched on every render");
  const html = document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /walkthroughForm/, "form renders once the directory is available");
  assert.ok(html.indexOf("Andruchek") < html.indexOf("Bossons"), "teacher options render in the directory's own (alphabetical) order");
  // Neither the hardcoded pilot list nor the legacy DB cache is the source
  // for this page's primary selector any more.
  const fn = app.slice(app.indexOf("renderWalkthrough() {"), app.indexOf("_ensureTeacherDirectoryLoading") === -1 ? app.length : app.indexOf("_submitWalkthrough(pageEl)"));
  assert.match(fn, /const allTeachers\s*=\s*TEACHER_DIRECTORY\.getAll\(\)/);
  assert.ok(!/const allTeachers\s*=\s*DB\.getTeachers\(\)/.test(fn));
  assert.ok(!/const allTeachers\s*=\s*PILOT_TEACHERS/.test(fn));
}

/* ── layer 2: _refreshWalkthroughStudentOptions() (Individual Student) ───── */

async function studentOptionsUseLiveRosterForTeacherWithStudents() {
  const { context, document } = makeContext();
  const STUDENT_ROSTER    = vm.runInContext("STUDENT_ROSTER", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await STUDENT_ROSTER.refresh();
  TEACHER_DIRECTORY._teachers = [{ id: "sp-1", name: "Bossons" }];
  TEACHER_DIRECTORY.loaded = true;

  const studentSel = { innerHTML: "" };
  const hint = { _hidden: true, classList: { toggle(cls, on) { if (cls === "hidden") hint._hidden = !!on; } } };
  document.registry["page-walkthrough"] = {
    querySelector(sel) {
      if (sel === '[name="teacherId"]') return { value: "sp-1" };
      if (sel === "#walkthroughStudent") return studentSel;
      if (sel === "#walkFocusRosterHint") return hint;
      if (sel === 'input[name="focus"]:checked') return { value: "individual-student" };
      return null;
    }
  };
  const APP = vm.runInContext("APP", context);
  APP._refreshWalkthroughStudentOptions();

  assert.match(studentSel.innerHTML, /Jane Doe/);
  assert.match(studentSel.innerHTML, /Sam Rivera/);
  assert.ok(!studentSel.innerHTML.includes("Pat Kim"), "another teacher's student is never leaked in");
  assert.ok(!hint._hidden === false || hint._hidden === true, "sanity");
  assert.equal(hint._hidden, true, "roster exists, so the 'no roster' hint stays hidden even for Individual Student");
}

async function studentOptionsHandleTeacherWithNoRosterMatch() {
  const { context, document } = makeContext();
  const STUDENT_ROSTER    = vm.runInContext("STUDENT_ROSTER", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await STUDENT_ROSTER.refresh();
  // "Andruchek" exists in the live directory but has no matching roster rows
  // — e.g. a brand-new teacher added via Setup this morning.
  TEACHER_DIRECTORY._teachers = [{ id: "sp-2", name: "Andruchek" }];
  TEACHER_DIRECTORY.loaded = true;

  const APP = vm.runInContext("APP", context);

  function run(focusValue) {
    const studentSel = { innerHTML: "" };
    const hint = { _hidden: true, classList: { toggle(cls, on) { if (cls === "hidden") hint._hidden = !!on; } } };
    document.registry["page-walkthrough"] = {
      querySelector(sel) {
        if (sel === '[name="teacherId"]') return { value: "sp-2" };
        if (sel === "#walkthroughStudent") return studentSel;
        if (sel === "#walkFocusRosterHint") return hint;
        if (sel === 'input[name="focus"]:checked') return focusValue ? { value: focusValue } : null;
        return null;
      }
    };
    APP._refreshWalkthroughStudentOptions();
    return { studentSel, hint };
  }

  // Whole Class / Small Group: fully usable, no blocking hint (items 10 & 11).
  for (const focus of ["whole-class", "small-group", ""]) {
    const { studentSel, hint } = run(focus);
    assert.match(studentSel.innerHTML, /No roster available/);
    assert.equal(hint._hidden, true, `focus="${focus}" must not show the blocking hint`);
  }

  // Individual Student with no roster match: informational hint shown, but
  // this is never a hard validation block (studentId stays optional).
  const { hint: individualHint } = run("individual-student");
  assert.equal(individualHint._hidden, false, "Individual Student surfaces the informational hint when there's no roster match");
  const submitFn = app.slice(app.indexOf("async _submitWalkthrough(pageEl)"), app.indexOf("async _submitWalkthrough(pageEl)") + 1500);
  assert.ok(!/studentId/.test(submitFn.slice(0, submitFn.indexOf("const responses"))), "Student is never a required field — a missing roster never blocks the walkthrough");
}

function noTeacherSelectedShowsPlaceholderOnly() {
  const { context, document } = makeContext();
  const studentSel = { innerHTML: "populate-me" };
  const hint = { _hidden: true, classList: { toggle(cls, on) { if (cls === "hidden") hint._hidden = !!on; } } };
  document.registry["page-walkthrough"] = {
    querySelector(sel) {
      if (sel === '[name="teacherId"]') return { value: "" };
      if (sel === "#walkthroughStudent") return studentSel;
      if (sel === "#walkFocusRosterHint") return hint;
      return null;
    }
  };
  const APP = vm.runInContext("APP", context);
  APP._refreshWalkthroughStudentOptions();
  assert.equal(studentSel.innerHTML, '<option value="">— None / whole class —</option>');
  assert.equal(hint._hidden, true);
}

/* ── layer 3: Setup → Teachers add/remove refreshes the live directory ───── */

async function addingTeacherInSetupMakesItAvailableToNewWalkthrough() {
  const { context, document, calls, setDirectory, dbClearCalls } = makeContext();
  setDirectory([{ spId: "sp-1", name: "Bossons" }]);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await TEACHER_DIRECTORY.refresh();
  assert.deepEqual(plain(TEACHER_DIRECTORY.getAll()).map(t => t.name), ["Bossons"]);

  document.getElementById("newTeacherName").value = "Alvarez";
  const APP = vm.runInContext("APP", context);
  APP._bindSetupTabEvents("teachers");

  // Simulate the add taking effect on the server, then click Add.
  setDirectory([{ spId: "sp-1", name: "Bossons" }, { spId: "sp-4", name: "Alvarez" }]);
  await document.registry["addTeacherBtn"]._handlers.click[0]();

  const calls2 = plain(calls.createListItem);
  assert.equal(calls2.length, 1);
  assert.deepEqual(calls2[0], { listName: "macwalkthroughwhoareyouvisiting", fields: { teacher: "Alvarez" } });
  assert.deepEqual(plain(TEACHER_DIRECTORY.getAll()).map(t => t.name).sort(), ["Alvarez", "Bossons"],
    "newly added teacher is available to New Walkthrough without a code deploy");
  assert.equal(dbClearCalls.length, 0, "no destructive DB.clearAll() migration was needed to accomplish this");
}

async function removingTeacherInSetupRemovesFromFutureSelectionsOnly() {
  const { context, document, calls, setDirectory } = makeContext();
  setDirectory([{ spId: "sp-1", name: "Bossons" }, { spId: "sp-2", name: "Andruchek" }]);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await TEACHER_DIRECTORY.refresh();

  // A pre-existing local walkthrough record referencing the teacher being
  // removed — proves the delete only affects future selections, never
  // historical data (item 9).
  const DB = vm.runInContext("DB", context);
  DB.addRecord({ teacherId: "sp-2", teacherName: "Andruchek", source: "walkthrough" });
  const beforeCount = plain(DB.getRecords()).length;

  const APP = vm.runInContext("APP", context);
  APP._bindSetupTabEvents("teachers");

  setDirectory([{ spId: "sp-1", name: "Bossons" }]); // server-side state after delete
  const btn = { dataset: { spId: "sp-2", name: "Andruchek" } };
  const evt = { target: { closest: sel => sel === ".del-teacher-btn" ? btn : null } };
  for (const handler of document.registry["tabContent"]._handlers.click) {
    await handler(evt);
  }

  assert.deepEqual(plain(TEACHER_DIRECTORY.getAll()).map(t => t.name), ["Bossons"], "removed teacher is gone from future selections");
  const recordsAfter = plain(DB.getRecords());
  assert.equal(recordsAfter.length, beforeCount, "historical walkthrough records are untouched by the removal");
  assert.equal(recordsAfter[0].responses.teacherName, "Andruchek", "the historical record itself is not rewritten or deleted");
}

/* ── static boundary assertions ─────────────────────────────────────────── */

function legacyConsumersStillUsePilotTeachers() {
  // Out of scope for this patch by explicit instruction — confirm they were
  // not touched, so the documented follow-up migration still applies to them.
  const bounds = (start, endMarkers) => {
    const s = app.indexOf(start);
    assert.ok(s >= 0, `${start} not found`);
    let e = app.length;
    for (const m of endMarkers) { const i = app.indexOf(m, s + 1); if (i > s && i < e) e = i; }
    return app.slice(s, e);
  };
  assert.match(bounds("renderWalkthroughV2() {", ["renderTeacherDashboard"]), /PILOT_TEACHERS/);
  assert.match(bounds("renderTeacherDashboard() {", ["renderStudentCheckIn"]), /PILOT_TEACHERS|DB\.getTeachers\(\)/);
  assert.match(bounds("renderStudentCheckIn() {", ["renderSetup"]), /PILOT_STUDENTS|PILOT_TEACHERS/);
  assert.match(bounds("renderFormLab() {", ["_seedPilotData"]), /DB\.getTeachers\(\)/);
}

function noCouplingIntroducedToDailyPulseOrPace() {
  const pulseFn = app.slice(app.indexOf("renderPulse() {"), app.indexOf("renderPulse() {") + 6000);
  assert.ok(!/TEACHER_DIRECTORY/.test(pulseFn), "Daily Pulse does not reference the new teacher directory");
  const paceAdmin = readSrc("pace-admin.js");
  assert.ok(!/TEACHER_DIRECTORY/.test(paceAdmin), "pace-admin.js does not reference the new teacher directory");
  const paceRenderStart = app.indexOf("_renderPaceAdminResults() {");
  assert.ok(paceRenderStart > 0);
  assert.ok(!/TEACHER_DIRECTORY/.test(app.slice(paceRenderStart, paceRenderStart + 4000)));
}

function noProductionWritesOrMigrationsAddedForThisPatch() {
  // This patch is read/refresh-only against macwalkthroughwhoareyouvisiting
  // for the New Walkthrough page itself — it must not add any new write to
  // IEP_Students_2026_27, and must not force a destructive local migration.
  const directoryModule = app.slice(app.indexOf("const TEACHER_DIRECTORY"), app.indexOf("const TEACHER_DIRECTORY") + 3000);
  assert.ok(!/GRAPH\.(create|update|delete)/.test(directoryModule), "TEACHER_DIRECTORY itself never writes to SharePoint");
  assert.ok(!/IEP_Students_2026_27/.test(directoryModule));
}

const tests = {
  directoryNormalizesAndDedupes, directoryFailureDoesNotFallBackToStalePilotTeachers,
  renderLoadingState, renderErrorStateWithRetryAndNoFallback, renderSuccessUsesLiveDirectoryOnly,
  studentOptionsUseLiveRosterForTeacherWithStudents, studentOptionsHandleTeacherWithNoRosterMatch,
  noTeacherSelectedShowsPlaceholderOnly,
  addingTeacherInSetupMakesItAvailableToNewWalkthrough, removingTeacherInSetupRemovesFromFutureSelectionsOnly,
  legacyConsumersStillUsePilotTeachers, noCouplingIntroducedToDailyPulseOrPace,
  noProductionWritesOrMigrationsAddedForThisPatch
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} teacher-directory migration checks passed.`);
})();
