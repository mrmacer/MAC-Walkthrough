/* ─────────────────────────────────────────────────────────────────────────
   PATCH: Use IEP_Users2 as the canonical walkthrough teacher directory

   The New Walkthrough "Who are you visiting?" selector, and the read-only
   Setup → Teachers tab, both now derive teachers from IEP_Users2 (active
   users whose Role includes Teacher) via TEACHER_DIRECTORY — replacing the
   retired dedicated macwalkthroughwhoareyouvisiting list.

   Three layers, matching this codebase's established test style (see
   tests/admin-editing.test.js's own docstring):
   1. TEACHER_DIRECTORY's IEP_Users2 sourcing/filtering/dedup is exercised
      for real against a mocked GRAPH.getListItems(), reusing the app's own
      normalizeSetupUser()/isTeacherRole() — nothing re-implemented here.
   2. renderWalkthrough()'s loading/error/success gating and
      _refreshWalkthroughStudentOptions()'s roster lookup (including the
      IEP_Users2-name → roster-canonical-name bridge via
      resolveCanonicalTeacherValue()) are exercised for real under node:vm,
      with STUDENT_ROSTER loaded for real against a mocked
      GRAPH.getListItems/getListSchema (same technique as
      tests/teacher-dropdown.test.js).
   3. Setup → Teachers' read-only rendering is exercised for real.
   Static source assertions cover the remaining boundary checks (no
   macwalkthroughwhoareyouvisiting writes anywhere, no PACE/Daily Pulse
   coupling).

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
    registry, getElementById: id => elementFor(id),
    querySelector: () => elementFor(Symbol()), querySelectorAll: () => [],
    createElement: () => elementFor(Symbol()), addEventListener() {}, removeEventListener() {},
    body: elementFor("__body__")
  };
}

// IEP_Users2 rows (field_1 Name, field_2 Role, field_3 Active, field_9 Email)
// — every case the requested test list needs in one fixture.
const USERS2_ROWS = [
  { id: "u1", field_1: "Amber Bossons",  field_2: "Teacher",                  field_3: "Yes", field_9: "bossa@iu29.org" },
  { id: "u2", field_1: "Nikki Stock",    field_2: "Teacher",                  field_3: "Yes", field_9: "stocn@iu29.org" },
  { id: "u3", field_1: "Former Teacher", field_2: "Teacher",                  field_3: "No",  field_9: "old@iu29.org" },   // inactive — must not appear
  { id: "u4", field_1: "Some Admin",     field_2: "Administrator",            field_3: "Yes", field_9: "admin@iu29.org" }, // not a teacher — must not appear
  { id: "u5", field_1: "Pat Rivera",     field_2: "Teacher, Administrator",   field_3: "Yes", field_9: "rivera@iu29.org" }, // multi-role, comma — must appear
  { id: "u6", field_1: "Sam Multi",      field_2: "Teacher; Behavior Specialist", field_3: "Yes", field_9: "sam@iu29.org" }, // multi-role, semicolon — must appear
  { id: "u7", field_1: "amber bossons",  field_2: "Teacher",                  field_3: "Yes", field_9: "dup@iu29.org" },   // duplicate visible name, different casing
  { id: "u8", field_1: "",               field_2: "Teacher",                  field_3: "Yes", field_9: "blank@iu29.org" } // blank name — must be ignored
];

// Roster: only "Bossons, A" (canonical) has a student — exercises both the
// "has roster" (Bossons) and "no roster match" (Stock, whose display name
// resolveCanonicalTeacherValue can only guess "Stock, N" for) branches.
const ROSTER_SCHEMA = { "Student First Name": "sfield_1", "Student Last Name": "sfield_2", "Teacher": "sfield_3", "Active": "sfield_5" };
const ROSTER_ROWS = [
  { id: "stu-1", sfield_1: "Jane", sfield_2: "Doe",   sfield_3: "Bossons, A", sfield_5: "Yes" },
  { id: "stu-2", sfield_1: "Sam",  sfield_2: "Rivera", sfield_3: "Bossons, A", sfield_5: "Yes" }
];

function makeGraphMock() {
  const calls = { getListItems: [], createListItem: [], getListId: [] };
  return {
    calls,
    GRAPH: {
      async getListItems(listName) {
        calls.getListItems.push(listName);
        if (listName === "IEP_Users2") return USERS2_ROWS;
        if (listName === "IEP_Students_2026_27") return ROSTER_ROWS;
        return [];
      },
      async getListSchema(listName) {
        if (listName === "IEP_Students_2026_27") return ROSTER_SCHEMA;
        return {};
      },
      async createListItem(listName, fields) { calls.createListItem.push({ listName, fields }); return { id: "new" }; },
      async getListId(listName) { calls.getListId.push(listName); return "list-id"; },
      async getSiteId() { return "site-1"; },
      async _patch() { throw new Error("must not write to any list from Setup → Teachers any more"); }
    }
  };
}

function makeContext() {
  const document = makeDocument();
  const { calls, GRAPH } = makeGraphMock();
  const context = {
    console, document, navigator: {},
    location: { origin: "http://localhost", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected raw fetch — GRAPH is mocked directly"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone, addEventListener() {}, removeEventListener() {}, confirm: () => true,
    GRAPH
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("config.js"), context, { filename: "config.js" });
  vm.runInContext(readSrc("auth.js"), context, { filename: "auth.js" });
  vm.runInContext(readSrc("data.js"), context, { filename: "data.js" });
  vm.runInContext(app, context, { filename: "app.js" });
  vm.runInContext('MAC_ADMIN_PANEL_ALLOWED = true; AUTH.role = "Administrator";', context);
  return { context, document, calls };
}

/* ── layer 1: TEACHER_DIRECTORY sourced from IEP_Users2 ───────────────────── */

async function directorySourcesFromIepUsers2() {
  const { context, calls } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await TEACHER_DIRECTORY.refresh();
  assert.deepEqual(calls.getListItems, ["IEP_Users2"], "reads IEP_Users2 via the existing generic list reader — item 1");
}

async function activeTeacherRoleFiltering() {
  const { context } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await TEACHER_DIRECTORY.refresh();
  const names = plain(TEACHER_DIRECTORY.getAll()).map(t => t.name);

  assert.ok(names.includes("Amber Bossons"), "active Teacher appears — item 2");
  assert.ok(names.includes("Nikki Stock"), "active Teacher appears — item 2");
  assert.ok(!names.includes("Former Teacher"), "inactive Teacher does not appear — item 3");
  assert.ok(!names.includes("Some Admin"), "Administrator without Teacher role does not appear — item 4");
  assert.ok(names.includes("Pat Rivera"), "multi-role 'Teacher, Administrator' still appears — item 5");
  assert.ok(names.includes("Sam Multi"), "multi-role 'Teacher; Behavior Specialist' still appears — item 5");
  assert.ok(!names.some(n => n === ""), "blank name ignored — item 6");
}

async function dedupedAndAlphabetized() {
  const { context } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await TEACHER_DIRECTORY.refresh();
  const all = plain(TEACHER_DIRECTORY.getAll());
  const bossonsRows = all.filter(t => t.name.toLowerCase() === "amber bossons");
  assert.equal(bossonsRows.length, 1, "duplicate visible name (case-insensitive) collapsed to one — item 7");
  assert.ok(["u1", "u7"].includes(bossonsRows[0].id), "kept one of the duplicate rows deterministically, by real SharePoint id");

  const names = all.map(t => t.name);
  const sorted = [...names].sort((a, b) => a.localeCompare(b));
  assert.deepEqual(names, sorted, "alphabetized — item 7");
  assert.ok(all.every(t => /^u\d$/.test(t.id)), "ids are real SharePoint item ids, never array positions");
}

/* ── layer 2: renderWalkthrough() / student options ───────────────────────── */

function renderSuccessUsesTeacherDirectoryOnly() {
  const { context, document } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }, { id: "u2", name: "Nikki Stock" }];
  TEACHER_DIRECTORY.loaded = true;
  const APP = vm.runInContext("APP", context);
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();

  const html = document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /walkthroughForm/);
  assert.match(html, /Amber Bossons/);
  assert.match(html, /Nikki Stock/);

  const fn = app.slice(app.indexOf("renderWalkthrough() {"), app.indexOf("_submitWalkthrough(pageEl)"));
  assert.match(fn, /const allTeachers\s*=\s*TEACHER_DIRECTORY\.getAll\(\)/, "New Walkthrough uses TEACHER_DIRECTORY — item 8");
  assert.ok(!/const allTeachers\s*=\s*DB\.getTeachers\(\)/.test(fn));
  assert.ok(!/const allTeachers\s*=\s*PILOT_TEACHERS/.test(fn));
}

function renderErrorNeverFallsBackToPilotTeachers() {
  const { context, document } = makeContext();
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  const PILOT_TEACHERS    = vm.runInContext("PILOT_TEACHERS", context);
  TEACHER_DIRECTORY.error = "Unable to load the teacher directory.";
  const APP = vm.runInContext("APP", context);
  APP.currentPage = "walkthrough";
  APP.renderWalkthrough();

  const html = document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /Teacher directory could not be loaded/);
  assert.match(html, /id="teacherDirectoryRetry"/, "retry offered");
  assert.ok(!/walkthroughForm/.test(html), "no form while blocked on an error");
  PILOT_TEACHERS.forEach(t => assert.ok(!html.includes(`>${t.name}<`), `no fallback to pilot teacher "${t.name}" — item 9`));
  const fn = app.slice(app.indexOf("renderWalkthrough() {"), app.indexOf("_submitWalkthrough(pageEl)"));
  const fnCode = fn.split("\n").filter(l => !l.trim().startsWith("//")).join("\n");
  assert.ok(!/PILOT_TEACHERS\.find|DB\.getTeachers\(\)/.test(fnCode), "no DB.getTeachers()/PILOT_TEACHERS fallback anywhere in renderWalkthrough's actual code — item 9");
}

async function individualStudentRosterLinkageBridgesNameFormats() {
  const { context, document } = makeContext();
  const STUDENT_ROSTER    = vm.runInContext("STUDENT_ROSTER", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await STUDENT_ROSTER.refresh();
  await TEACHER_DIRECTORY.refresh();

  const bossons = TEACHER_DIRECTORY.getAll().find(t => t.name === "Amber Bossons");
  const studentSel = { innerHTML: "" };
  const hint = { _hidden: true, classList: { toggle(cls, on) { if (cls === "hidden") hint._hidden = !!on; } } };
  document.registry["page-walkthrough"] = {
    querySelector(sel) {
      if (sel === '[name="teacherId"]') return { value: bossons.id };
      if (sel === "#walkthroughStudent") return studentSel;
      if (sel === "#walkFocusRosterHint") return hint;
      if (sel === 'input[name="focus"]:checked') return { value: "individual-student" };
      return null;
    }
  };
  const APP = vm.runInContext("APP", context);
  APP._refreshWalkthroughStudentOptions();

  // "Amber Bossons" (IEP_Users2) must resolve to "Bossons, A" (the roster's
  // own canonical value) to find these students — item 13.
  assert.match(studentSel.innerHTML, /Jane Doe/);
  assert.match(studentSel.innerHTML, /Sam Rivera/);
  assert.equal(hint._hidden, true, "roster match found, so the hint stays hidden");
}

async function wholeClassAndSmallGroupWorkWithNoRosterMatch() {
  const { context, document } = makeContext();
  const STUDENT_ROSTER    = vm.runInContext("STUDENT_ROSTER", context);
  const TEACHER_DIRECTORY = vm.runInContext("TEACHER_DIRECTORY", context);
  await STUDENT_ROSTER.refresh();
  await TEACHER_DIRECTORY.refresh();
  const stock = TEACHER_DIRECTORY.getAll().find(t => t.name === "Nikki Stock"); // no roster match, per fixture

  const APP = vm.runInContext("APP", context);
  function run(focusValue) {
    const studentSel = { innerHTML: "" };
    const hint = { _hidden: true, classList: { toggle(cls, on) { if (cls === "hidden") hint._hidden = !!on; } } };
    document.registry["page-walkthrough"] = {
      querySelector(sel) {
        if (sel === '[name="teacherId"]') return { value: stock.id };
        if (sel === "#walkthroughStudent") return studentSel;
        if (sel === "#walkFocusRosterHint") return hint;
        if (sel === 'input[name="focus"]:checked') return focusValue ? { value: focusValue } : null;
        return null;
      }
    };
    APP._refreshWalkthroughStudentOptions();
    return { studentSel, hint };
  }

  for (const focus of ["whole-class", "small-group", ""]) {
    const { studentSel, hint } = run(focus);
    assert.match(studentSel.innerHTML, /No roster available/);
    assert.equal(hint._hidden, true, `focus="${focus}" must stay fully usable — item 14`);
  }
  const { hint: individualHint } = run("individual-student");
  assert.equal(individualHint._hidden, false, "Individual Student shows the informational hint, but this is not a submit blocker");
  const submitFn = app.slice(app.indexOf("async _submitWalkthrough(pageEl)"), app.indexOf("async _submitWalkthrough(pageEl)") + 1500);
  assert.ok(!/studentId/.test(submitFn.slice(0, submitFn.indexOf("const responses"))), "Student is never a required field");
}

/* ── layer 3: Setup → Teachers, read-only ─────────────────────────────────── */

async function setupTeachersIsReadOnly() {
  const { context, document, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  await SETUP_DATA.refresh();

  const html = APP._renderTeachersTab();
  assert.ok(!/id="addTeacherBtn"/.test(html), "no Add Teacher control — item 10");
  assert.ok(!/id="newTeacherName"/.test(html), "no Add Teacher input — item 10");
  assert.ok(!/del-teacher-btn/.test(html), "no Delete Teacher control — item 10");
  assert.ok(!/grant-access-btn/.test(html), "no Grant Access control (meaningless now — everyone shown already has a login)");
  assert.match(html, /Teachers are managed through the Users directory\. Active users with the Teacher role automatically appear in walkthroughs\./,
    "explains where teachers are actually managed — item 12");
  assert.match(html, /Amber Bossons/);
  assert.match(html, /Nikki Stock/);
  assert.ok(!/Former Teacher/.test(html), "inactive teacher not shown in Setup either");

  // Binding the tab wires nothing that writes anywhere.
  APP._bindSetupTabEvents("teachers");
  assert.equal(calls.createListItem.length, 0);
  assert.equal(document.registry["addTeacherBtn"], undefined, "no such control was ever created/bound — item 10");
}

async function setupTeachersNeverTouchesTheOldList() {
  const { context, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  await SETUP_DATA.refresh();
  assert.ok(!calls.getListItems.includes("macwalkthroughwhoareyouvisiting"), "SETUP_DATA.refresh() never reads the old list — item 11");
  assert.ok(!calls.getListId.includes("macwalkthroughwhoareyouvisiting"), "never resolves the old list's id either — item 11");
  assert.ok(!("teachers" in vm.runInContext("SETUP_LISTS", context)), "SETUP_LISTS.teachers (the old list name) is gone");
}

function staticNoWritesToOldListAnywhere() {
  const codeOnly = line => !line.trim().startsWith("//");
  assert.equal(app.split("\n").filter(l => l.includes("SETUP_LISTS.teachers") && codeOnly(l)).length, 0,
    "no code reference to a removed SETUP_LISTS.teachers key (comments may still explain the retired name)");
  assert.equal(app.split("\n").filter(l => l.includes("createListItem(SETUP_LISTS") && codeOnly(l)).length, 0,
    "no createListItem call could target the retired list");
  // The list itself is only ever mentioned in explanatory/legacy comments now.
  const codeLines = app.split("\n").filter(l => l.includes("macwalkthroughwhoareyouvisiting") && codeOnly(l));
  assert.equal(codeLines.length, 0, "macwalkthroughwhoareyouvisiting appears only in comments, never in executable code");
}

/* ── static: Daily Pulse / PACE unchanged ─────────────────────────────────── */

function noCouplingIntroducedToDailyPulseOrPace() {
  const pulseFn = app.slice(app.indexOf("renderPulse() {"), app.indexOf("renderPulse() {") + 6000);
  assert.ok(!/TEACHER_DIRECTORY/.test(pulseFn), "Daily Pulse does not reference the teacher directory — item 15");
  const paceAdmin = readSrc("pace-admin.js");
  assert.ok(!/TEACHER_DIRECTORY/.test(paceAdmin), "pace-admin.js does not reference the teacher directory — item 15");
  const paceRenderStart = app.indexOf("_renderPaceAdminResults() {");
  assert.ok(paceRenderStart > 0);
  assert.ok(!/TEACHER_DIRECTORY/.test(app.slice(paceRenderStart, paceRenderStart + 4000)), "item 15");
}

const tests = {
  directorySourcesFromIepUsers2, activeTeacherRoleFiltering, dedupedAndAlphabetized,
  renderSuccessUsesTeacherDirectoryOnly, renderErrorNeverFallsBackToPilotTeachers,
  individualStudentRosterLinkageBridgesNameFormats, wholeClassAndSmallGroupWorkWithNoRosterMatch,
  setupTeachersIsReadOnly, setupTeachersNeverTouchesTheOldList, staticNoWritesToOldListAnywhere,
  noCouplingIntroducedToDailyPulseOrPace
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} IEP_Users2 teacher-directory checks passed.`);
})();
