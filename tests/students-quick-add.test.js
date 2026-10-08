/* ─────────────────────────────────────────────────────────────────────────
   Students page — quick add / edit for the production roster

   Exercises the real app.js under node:vm with a stateful, mocked
   GRAPH (IEP_Students_2026_27 + IEP_Users2 rows) so create/update payloads,
   confirm-before-success ordering, duplicate protection, defaults, and
   authorization are verified against actual behavior — not just source text.
   Existing suites (setup-student-roster, teacher-dropdown, walkthrough-*,
   pace-*) separately prove Setup → Students, PACE, and Daily Pulse still
   work; the checks here add direct no-regression assertions on top.

   Run with: node tests/students-quick-add.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");
const plain = obj => JSON.parse(JSON.stringify(obj));

/* ── DOM stub ──────────────────────────────────────────────────────────── */

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
        click() {}, scrollIntoView() {}, scrollTo() {}
      };
      Object.defineProperty(el, "innerHTML", { get: () => el._html || "", set: v => { el._html = v; } });
      registry[id] = el;
    }
    return registry[id];
  }
  return {
    registry, getElementById: id => elementFor(id),
    querySelector: () => null, querySelectorAll: () => [],
    createElement: () => elementFor(Symbol()), addEventListener() {}, removeEventListener() {},
    body: elementFor("__body__")
  };
}

/* ── stateful Graph mock ───────────────────────────────────────────────── */

const ROSTER_LIST = "IEP_Students_2026_27";
const ROSTER_SCHEMA = {
  "Student First Name": "sfield_1", "Student Last Name": "sfield_2", "Teacher": "sfield_3",
  "Classroom": "sfield_4", "Active": "sfield_5", "PACE Enabled": "sfield_6", "Daily Pulse Enabled": "sfield_7"
};
const KNOWN_COLUMNS = Object.keys(ROSTER_SCHEMA);

const USERS2_ROWS = [
  { id: "u1", field_1: "Amber Bossons", field_2: "Teacher",       field_3: "Yes", field_9: "bossa@iu29.org" },
  { id: "u2", field_1: "Nikki Stock",   field_2: "Teacher",       field_3: "Yes", field_9: "stocn@iu29.org" },
  { id: "u3", field_1: "Some Admin",    field_2: "Administrator", field_3: "Yes", field_9: "admin@iu29.org" },
  { id: "u4", field_1: "Old Teacher",   field_2: "Teacher",       field_3: "No",  field_9: "old@iu29.org" }
];

function makeGraph() {
  const state = {
    rows: [
      { id: "s1", sfield_1: "Jane", sfield_2: "Doe",  sfield_3: "Bossons, A", sfield_4: "Room 1", sfield_5: "Yes", sfield_6: "Yes", sfield_7: "No" },
      { id: "s2", sfield_1: "Old",  sfield_2: "Kid",  sfield_3: "Bossons, A", sfield_4: "",       sfield_5: "No",  sfield_6: "No",  sfield_7: "No" }
    ],
    nextId: 100, calls: { create: [], update: [], rosterReads: 0 },
    createGate: null, createError: null, updateError: null
  };
  const toInternal = fields => Object.fromEntries(Object.entries(fields).map(([k, v]) => [ROSTER_SCHEMA[k], v]));
  return {
    state,
    GRAPH: {
      async getListItems(listName) {
        if (listName === "IEP_Users2") return USERS2_ROWS;
        if (listName === ROSTER_LIST) { state.calls.rosterReads++; return state.rows.map(r => ({ ...r })); }
        return [];
      },
      async getListSchema(listName) { return listName === ROSTER_LIST ? ROSTER_SCHEMA : {}; },
      async createMappedListItem(listName, fields) {
        state.calls.create.push({ listName, fields: { ...fields } });
        if (state.createGate) await state.createGate;
        if (state.createError) throw new Error(state.createError);
        const id = String(state.nextId++);
        state.rows.push({ id, ...toInternal(fields) });
        return { id };
      },
      async updateMappedListItem(listName, itemId, fields) {
        state.calls.update.push({ listName, itemId, fields: { ...fields } });
        if (state.updateError) throw new Error(state.updateError);
        Object.assign(state.rows.find(r => r.id === itemId), toInternal(fields));
        return {};
      }
    }
  };
}

function makeContext({ role = "Administrator", panel = true } = {}) {
  const document = makeDocument();
  const { GRAPH, state } = makeGraph();
  const toasts = [];
  const context = {
    console: { ...console, log() {}, warn() {} }, document, navigator: {}, location: { origin: "http://localhost", hash: "" },
    crypto: crypto.webcrypto, TextEncoder, Uint8Array,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected raw fetch — GRAPH is mocked directly"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone, addEventListener() {}, removeEventListener() {}, confirm: () => true,
    GRAPH, __toasts: toasts
  };
  context.window = context;
  vm.createContext(context);
  ["config.js", "auth.js", "data.js", "pace-admin.js", "pace-export.js", "evidence-report.js", "app.js"].forEach(f =>
    vm.runInContext(readSrc(f), context, { filename: f }));
  vm.runInContext(`MAC_ADMIN_PANEL_ALLOWED = ${panel}; AUTH.role = ${JSON.stringify(role)}; showToast = (m, t) => __toasts.push({ m, t });`, context);
  return {
    context, document, state, toasts,
    APP: vm.runInContext("APP", context),
    STUDENT_ROSTER: vm.runInContext("STUDENT_ROSTER", context),
    TEACHER_DIRECTORY: vm.runInContext("TEACHER_DIRECTORY", context),
    SETUP_DATA: vm.runInContext("SETUP_DATA", context)
  };
}

async function readyContext(opts) {
  const c = makeContext(opts);
  await c.STUDENT_ROSTER.refresh();
  await c.TEACHER_DIRECTORY.refresh();
  c.APP.currentPage = "students";
  c.APP.renderStudents();
  return c;
}

const formHtml = c => c.document.registry["stuFormHost"].innerHTML;
const listHtml = c => c.document.registry["stuList"].innerHTML;
const fill = (c, v) => {
  const d = c.document;
  if ("first" in v)     d.getElementById("stuFirst").value = v.first;
  if ("last" in v)      d.getElementById("stuLast").value = v.last;
  if ("teacher" in v)   d.getElementById("stuTeacher").value = v.teacher;
  if ("classroom" in v) d.getElementById("stuClassroom").value = v.classroom;
  d.getElementById("stuActive").checked = "active" in v ? v.active : true;
  d.getElementById("stuPace").checked   = "pace" in v ? v.pace : false;
  d.getElementById("stuPulse").checked  = "pulse" in v ? v.pulse : false;
};

/* ── 1–2. sidebar + authorization ────────────────────────────────────────── */

async function sidebarShowsStudentsForAdminsBeforeSetupOnly() {
  const admin = makeContext();
  admin.APP.renderNav();
  const html = admin.document.registry["navList"].innerHTML;
  assert.match(html, /data-page="students"/, "item 1: Students is in the sidebar for an authorized admin");
  assert.ok(html.indexOf('data-page="students"') < html.indexOf('data-page="setup"'), "placed before Setup");
  assert.ok(html.indexOf('data-page="students"') > html.indexOf('data-page="pace"'), "near the operational tools");

  for (const opts of [{ role: "Teacher" }, { panel: false }]) {
    const other = makeContext(opts);
    other.APP.renderNav();
    assert.ok(!/data-page="students"/.test(other.document.registry["navList"].innerHTML), `hidden for ${JSON.stringify(opts)}`);
  }
}

async function unauthorizedUsersCannotReachOrUseStudentManagement() {
  const teacher = makeContext({ role: "Teacher" });
  assert.equal(vm.runInContext('canAccessRoute("students")', teacher.context), false, "item 2: teacher cannot route to it");
  const noPanel = makeContext({ panel: false });
  assert.equal(vm.runInContext('canAccessRoute("students")', noPanel.context), false, "no Admin Panel permission → no route");
  const admin = makeContext();
  assert.equal(vm.runInContext('canAccessRoute("students")', admin.context), true);

  // Even if the page function is invoked directly, nothing renders and nothing writes.
  for (const c of [teacher, noPanel]) {
    await c.STUDENT_ROSTER.refresh();
    c.APP.renderStudents();
    const html = c.document.registry["page-students"].innerHTML;
    assert.match(html, /do not have permission/);
    assert.ok(!/stuAddBtn/.test(html));
    c.APP._studentsUi.mode = "add";
    c.APP._studentsUi.draft = { first: "Mia", last: "Lopez", teacher: "Bossons, A", classroom: "", active: true, paceEnabled: false, dailyPulseEnabled: false };
    c.document.getElementById("stuFirst").value = "Mia";
    await c.APP._saveStudentForm(false);
    assert.equal(c.state.calls.create.length, 0, "no SharePoint write for an unauthorized user");
  }
}

/* ── 3–8. page, form, defaults, teacher mapping, validation ──────────────── */

async function pageLoadsRosterAndSearchFilters() {
  const c = await readyContext();
  const html = listHtml(c);
  assert.match(html, /Jane Doe/, "item 3: existing roster records listed");
  assert.match(html, /Old Kid/);
  assert.match(html, /Room 1/);
  assert.match(c.document.registry["page-students"].innerHTML, /\+ Add Student/);
  assert.equal(c.document.registry["stuCount"].textContent, "2 students");

  c.APP._studentsUi.search = "jane";
  c.APP._renderStudentsList();
  assert.match(listHtml(c), /Jane Doe/);
  assert.ok(!/Old Kid/.test(listHtml(c)), "search filters in memory");
  assert.equal(c.document.registry["stuCount"].textContent, "1 of 2");
  c.APP._studentsUi.search = "zzz";
  c.APP._renderStudentsList();
  assert.match(listHtml(c), /No students match/);
  assert.equal(c.state.calls.rosterReads, 1, "search never re-queries SharePoint");

  // Not-yet-loaded roster shows a loading state instead of an empty list.
  const cold = makeContext();
  cold.APP.currentPage = "students";
  cold.APP.renderStudents();
  assert.match(cold.document.registry["page-students"].innerHTML, /Loading student roster/);
}

async function addStudentOpensQuickFormWithCanonicalTeachersAndSafeDefaults() {
  const c = await readyContext();
  assert.equal(formHtml(c), "", "form is closed until + Add Student");
  c.APP._openStudentAdd();
  const html = formHtml(c);
  assert.match(html, /<h3>Add Student<\/h3>/, "item 4: quick-entry form opens");
  assert.match(html, /id="stuFirst"/);
  assert.match(html, /id="stuLast"/);
  assert.match(html, /Save Student/);
  assert.match(html, /Save &amp; Add Another/);

  // item 6: canonical teacher mapping — IEP_Users2 active Teachers only,
  // Amber Bossons resolves to the roster's own "Bossons, A"; no admin/inactive.
  assert.match(html, /<option value="Bossons, A">Amber Bossons<\/option>/);
  assert.match(html, /<option value="Stock, N">Nikki Stock<\/option>/);
  assert.ok(!/Some Admin|Old Teacher/.test(html), "non-teachers and inactive teachers are not offered");

  // items 7–8: only Active starts on; eligibility flags stay opt-in.
  assert.match(html, /id="stuActive" checked/);
  assert.ok(!/id="stuPace" checked/.test(html), "PACE Enabled is not blindly enabled");
  assert.ok(!/id="stuPulse" checked/.test(html), "Daily Pulse Enabled is not blindly enabled");
  assert.deepEqual(plain(vm.runInContext("QUICK_ADD_STUDENT_DEFAULTS", c.context)),
    { active: true, paceEnabled: false, dailyPulseEnabled: false });
  assert.ok(Object.isFrozen(vm.runInContext("QUICK_ADD_STUDENT_DEFAULTS", c.context)));
}

async function validationRequiresNameAndTeacher() {
  const c = await readyContext();
  c.APP._openStudentAdd();
  fill(c, { first: "", last: "Lopez", teacher: "Bossons, A" });
  await c.APP._saveStudentForm(false);
  assert.match(c.toasts.at(-1).m, /first and last name/, "item 5: name required");
  fill(c, { first: "Mia", last: "", teacher: "Bossons, A" });
  await c.APP._saveStudentForm(false);
  assert.match(c.toasts.at(-1).m, /first and last name/);
  fill(c, { first: "Mia", last: "Lopez", teacher: "" });
  await c.APP._saveStudentForm(false);
  assert.match(c.toasts.at(-1).m, /Select a teacher/);
  assert.equal(c.state.calls.create.length, 0, "nothing written for invalid input");
  assert.equal(c.APP._studentsUi.mode, "add", "form stays open");
}

/* ── 9–11. save path + confirm-before-success ───────────────────────────── */

async function saveCreatesThroughTheExistingRosterPathWithExactColumns() {
  const c = await readyContext();
  c.APP._openStudentAdd();
  fill(c, { first: "  Mia ", last: "Lopez", teacher: "Bossons, A", classroom: "Room 12", pace: true });
  await c.APP._saveStudentForm(false);

  assert.equal(c.state.calls.create.length, 1, "item 9: one create");
  const { listName, fields } = c.state.calls.create[0];
  assert.equal(listName, ROSTER_LIST, "written to IEP_Students_2026_27 via STUDENT_ROSTER._listName()");
  assert.deepEqual(plain(fields), {
    "Student First Name": "Mia", "Student Last Name": "Lopez", "Teacher": "Bossons, A", "Classroom": "Room 12",
    "Active": "Yes", "PACE Enabled": "Yes", "Daily Pulse Enabled": "No"
  });
  assert.ok(Object.keys(fields).every(k => KNOWN_COLUMNS.includes(k)), "item 20: only existing roster columns — no schema change");

  assert.match(c.toasts.find(t => /added/.test(t.m)).m, /Student "Mia Lopez" added/);
  assert.match(listHtml(c), /Mia Lopez/, "new student is on the page immediately after the refresh");
  assert.equal(c.APP._studentsUi.mode, null, "Save Student closes the form");
  assert.ok(c.SETUP_DATA.students.some(s => s.name === "Mia Lopez"), "Setup → Students' snapshot stays current");
}

async function savedStudentsFeedPaceAndDailyPulseThroughTheSharedRoster() {
  const c = await readyContext();
  c.APP._openStudentAdd();
  fill(c, { first: "Pace", last: "Kid", teacher: "Bossons, A", pace: true });
  await c.APP._saveStudentForm(false);
  c.APP._openStudentAdd();
  fill(c, { first: "Plain", last: "Kid", teacher: "Bossons, A" });
  await c.APP._saveStudentForm(false);

  const pace = c.STUDENT_ROSTER.getPaceEnabled().map(s => s.name);
  assert.ok(pace.includes("Pace Kid"), "items 16–17: PACE reads the same roster");
  assert.ok(!pace.includes("Plain Kid"), "a student saved with defaults is NOT PACE-eligible");
  const pulse = c.STUDENT_ROSTER.getDailyPulseEnabled().map(s => s.name);
  assert.ok(!pulse.includes("Plain Kid") && !pulse.includes("Pace Kid"), "Daily Pulse stays opt-in");
}

async function saveWaitsForSharePointAndFailuresNeverShowSuccess() {
  // item 10: nothing is reported until the Graph write resolves.
  const c = await readyContext();
  let release;
  c.state.createGate = new Promise(r => { release = r; });
  c.APP._openStudentAdd();
  fill(c, { first: "Slow", last: "Write", teacher: "Bossons, A" });
  const pending = c.APP._saveStudentForm(false);
  await new Promise(r => setImmediate(r));
  assert.equal(c.toasts.filter(t => /added/.test(t.m)).length, 0, "no success message while SharePoint is still working");
  assert.equal(c.APP._studentsUi.saving, true);
  assert.match(formHtml(c), /Saving…/);
  assert.equal(c.state.calls.rosterReads, 1, "roster not refreshed before confirmation");
  release();
  await pending;
  assert.equal(c.toasts.filter(t => /added/.test(t.m)).length, 1);
  assert.equal(c.state.calls.rosterReads, 2, "refreshed after confirmation");

  // item 11: a failed write is an error, never a false success, and keeps the form.
  const f = await readyContext();
  f.state.createError = "Graph 403: forbidden";
  f.APP._openStudentAdd();
  fill(f, { first: "Fail", last: "Write", teacher: "Bossons, A" });
  await f.APP._saveStudentForm(false);
  assert.equal(f.toasts.filter(t => /added/.test(t.m)).length, 0, "no false success");
  assert.match(f.toasts.at(-1).m, /Failed to add student: Graph 403/);
  assert.equal(f.toasts.at(-1).t, "error");
  assert.equal(f.APP._studentsUi.mode, "add", "form stays open with the admin's input");
  assert.equal(f.APP._studentsUi.saving, false);
  assert.match(formHtml(f), /value="Fail"/);
  assert.ok(!/Fail Write/.test(listHtml(f)), "nothing appears on the list");
  assert.equal(f.state.calls.rosterReads, 1, "no refresh after a failed write");
}

/* ── 12–13. duplicates, Save & Add Another ───────────────────────────────── */

async function duplicateProtectionWarnsAndNeverOverwrites() {
  const c = await readyContext();
  c.APP._openStudentAdd();
  fill(c, { first: "jane ", last: " DOE", teacher: "Bossons, A" });
  await c.APP._saveStudentForm(false);
  assert.equal(c.state.calls.create.length, 0, "item 12: obvious duplicate (case/space-insensitive) is not created");
  assert.equal(c.state.calls.update.length, 0, "and the existing record is never touched");
  assert.match(formHtml(c), /may already be on the roster/);
  assert.match(formHtml(c), /Jane Doe — Bossons, A · Active/);
  assert.match(formHtml(c), /Add as a separate student/);

  // Inactive matches are flagged too, with a pointer to reactivate instead.
  c.APP._studentsUi.dup = null;
  fill(c, { first: "Old", last: "Kid", teacher: "Bossons, A" });
  await c.APP._saveStudentForm(false);
  assert.equal(c.state.calls.create.length, 0);
  assert.match(formHtml(c), /Old Kid — Bossons, A · Inactive/);
  assert.match(formHtml(c), /inactive record exists/);

  // Only an explicit confirmation creates it, once.
  await c.APP._saveStudentForm(false, true);
  assert.equal(c.state.calls.create.length, 1);

  // A genuinely different name never triggers the warning.
  const d = await readyContext();
  d.APP._openStudentAdd();
  fill(d, { first: "Jane", last: "Dole", teacher: "Bossons, A" });
  await d.APP._saveStudentForm(false);
  assert.equal(d.state.calls.create.length, 1);
}

async function saveAndAddAnotherKeepsFormOpenAndResetsToDefaults() {
  const c = await readyContext();
  c.APP._openStudentAdd();
  fill(c, { first: "First", last: "One", teacher: "Bossons, A", classroom: "Room 3", pace: true, pulse: true });
  await c.APP._saveStudentForm(true);

  assert.equal(c.state.calls.create.length, 1, "item 13: saved");
  assert.equal(c.APP._studentsUi.mode, "add", "the Add Student interface stays open");
  const html = formHtml(c);
  assert.match(html, /<h3>Add Student<\/h3>/);
  assert.match(html, /id="stuFirst"[^>]*value=""/, "name cleared");
  assert.match(html, /id="stuLast"[^>]*value=""/);
  assert.match(html, /id="stuClassroom"[^>]*value=""/);
  assert.ok(!/<option value="Bossons, A" selected>/.test(html), "teacher cleared");
  assert.match(html, /id="stuActive" checked/);
  assert.ok(!/id="stuPace" checked/.test(html) && !/id="stuPulse" checked/.test(html), "flags back to defaults");
  assert.match(listHtml(c), /First One/, "the saved student is already listed");

  fill(c, { first: "Second", last: "Two", teacher: "Stock, N" });
  await c.APP._saveStudentForm(true);
  assert.equal(c.state.calls.create.length, 2, "next student entered immediately");
  assert.equal(plain(c.state.calls.create[1].fields)["Teacher"], "Stock, N");
}

/* ── 14. editing ─────────────────────────────────────────────────────────── */

async function editUsesTheExistingUpdatePathAndSendsOnlyChanges() {
  const c = await readyContext();
  c.APP._openStudentEdit("s1");
  assert.match(formHtml(c), /<h3>Edit Student<\/h3>/);
  assert.match(formHtml(c), /id="stuFirst"[^>]*value="Jane"/);
  assert.match(formHtml(c), /<option value="Bossons, A" selected>/);
  assert.ok(!/Save &amp; Add Another/.test(formHtml(c)), "no add-another on edit");

  fill(c, { first: "Jane", last: "Doe", teacher: "Bossons, A", classroom: "Room 9", active: true, pace: true, pulse: false });
  await c.APP._saveStudentForm(false);
  assert.equal(c.state.calls.create.length, 0, "item 14: editing never creates");
  assert.equal(c.state.calls.update.length, 1);
  assert.deepEqual(plain(c.state.calls.update[0]), { listName: ROSTER_LIST, itemId: "s1", fields: { "Classroom": "Room 9" } });
  assert.match(listHtml(c), /Room 9/);

  // Deactivating is an edit too — there is no delete path at all.
  c.APP._openStudentEdit("s1");
  fill(c, { first: "Jane", last: "Doe", teacher: "Bossons, A", classroom: "Room 9", active: false, pace: true });
  await c.APP._saveStudentForm(false);
  assert.deepEqual(plain(c.state.calls.update[1].fields), { "Active": "No" });

  // A rename that collides with another student is caught, but the student's own name isn't a "duplicate".
  c.APP._openStudentEdit("s1");
  fill(c, { first: "Old", last: "Kid", teacher: "Bossons, A", classroom: "Room 9", active: false, pace: true });
  await c.APP._saveStudentForm(false);
  assert.equal(c.state.calls.update.length, 2, "rename onto an existing student is held for confirmation");
  assert.match(formHtml(c), /Save anyway/);
}

async function editKeepsAFormerTeacherInsteadOfSilentlyReassigning() {
  const c = await readyContext();
  c.state.rows[0].sfield_3 = "Retired, T";
  await c.STUDENT_ROSTER.refresh();
  c.APP._openStudentEdit("s1");
  assert.match(formHtml(c), /<option value="Retired, T" selected>Retired, T \(current\)<\/option>/);
}

/* ── 15–20. nothing else changed ─────────────────────────────────────────── */

function setupStudentsAndOtherWorkflowsAreUntouched() {
  // item 15: Setup → Students still writes through the same existing path.
  assert.match(app, /await GRAPH\.createMappedListItem\(STUDENT_ROSTER\._listName\(\), \{\s*"Student First Name":\s+firstName,/);
  assert.match(app, /GRAPH\.updateMappedListItem\(STUDENT_ROSTER\._listName\(\), spId, \{ "Active": "No" \}\)/);
  assert.match(app, /_renderStudentsTab\(\) \{/);

  // items 16–17: PACE / Daily Pulse code never references the new page.
  const pulse = app.slice(app.indexOf("renderPulse() {"), app.indexOf("renderPulse() {") + 6000);
  assert.ok(!/_studentsUi|renderStudents|QUICK_ADD/.test(pulse));
  assert.ok(!/_studentsUi|QUICK_ADD|renderStudents/.test(readSrc("pace-admin.js")));
  const paceStart = app.indexOf("async renderPaceAdmin(");
  assert.ok(paceStart > 0, "PACE admin page definition found");
  assert.ok(!/_studentsUi|renderStudents/.test(app.slice(paceStart, paceStart + 4000)));
}

async function newWalkthroughStudentStaysFreeTextWithNoRosterLookup() {
  const c = makeContext();
  const TEACHER_DIRECTORY = c.TEACHER_DIRECTORY;
  TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  TEACHER_DIRECTORY.loaded = true;
  c.APP.currentPage = "walkthrough";
  c.APP.renderWalkthrough();
  const html = c.document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /<input type="text" class="form-input" name="studentName" id="walkthroughStudent"/, "item 18: free-text input");
  assert.ok(!/<select[^>]*studentId/.test(html), "no roster-driven <select>");
  assert.equal(c.state.calls.rosterReads, 0, "item 19: rendering New Walkthrough never reads the roster");

  const code = fn => fn.split("\n").filter(l => !l.trim().startsWith("//")).join("\n");
  const render = app.slice(app.indexOf("renderWalkthrough() {"), app.indexOf("renderWalkthroughV2() {"));
  assert.ok(!/STUDENT_ROSTER/.test(code(render)), "no STUDENT_ROSTER reference in the live New Walkthrough code");
}

function noSchemaOrListCreationIntroduced() {
  const graph = readSrc("graph.js");
  assert.ok(!/_post\([^)]*\/columns|_post\(`sites\/\$\{siteId\}\/lists`\s*,/.test(graph), "item 20: graph.js never creates lists or columns");
  const section = app.slice(app.indexOf("/* ── STUDENTS (quick entry)"), app.indexOf("/* ── SETUP ──"));
  assert.ok(!/GRAPH\._post|GRAPH\._patch|GRAPH\.createList\b|\/columns/.test(section), "the Students page only uses the existing mapped create/update helpers");
  assert.ok(!/removeStudent|deleteStudent|\bDELETE\b|_delete\(/.test(section), "no hard delete");
  assert.match(section, /STUDENT_ROSTER\._listName\(\)/, "list name always comes from STUDENT_ROSTER");
  assert.ok(!/["']IEP_Students_2026_27["']/.test(section), "no second hard-coded student list");
}

const tests = {
  sidebarShowsStudentsForAdminsBeforeSetupOnly, unauthorizedUsersCannotReachOrUseStudentManagement,
  pageLoadsRosterAndSearchFilters, addStudentOpensQuickFormWithCanonicalTeachersAndSafeDefaults,
  validationRequiresNameAndTeacher, saveCreatesThroughTheExistingRosterPathWithExactColumns,
  savedStudentsFeedPaceAndDailyPulseThroughTheSharedRoster, saveWaitsForSharePointAndFailuresNeverShowSuccess,
  duplicateProtectionWarnsAndNeverOverwrites, saveAndAddAnotherKeepsFormOpenAndResetsToDefaults,
  editUsesTheExistingUpdatePathAndSendsOnlyChanges, editKeepsAFormerTeacherInsteadOfSilentlyReassigning,
  setupStudentsAndOtherWorkflowsAreUntouched, newWalkthroughStudentStaysFreeTextWithNoRosterLookup,
  noSchemaOrListCreationIntroduced
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} Students-page checks passed.`);
})();
