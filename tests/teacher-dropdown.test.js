/* ─────────────────────────────────────────────────────────────────────────
   Setup → Students → Add Student: Teacher dropdown fix

   The dropdown was empty because it was sourced from
   macwalkthroughwhoareyouvisiting (an unrelated Walkthrough picklist) via
   SETUP_DATA.teachers, not IEP_Users2. This suite runs the REAL
   _renderStudentsTab()/_bindSetupTabEvents("students") code under node:vm
   against a mocked GRAPH (recording every call), proving the dropdown is
   now built from active Teacher rows in IEP_Users2, and that the VALUE
   saved is the canonical IEP_Students_2026_27 Teacher format — not the
   IEP_Users2 display name verbatim — so Daily Pulse/PACE's exact-string
   teacher matching keeps working.

   Run with: node tests/teacher-dropdown.test.js
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
    addEventListener() {},
    body: elementFor("__body__")
  };
}

// IEP_Users2 rows: an active Teacher already assigned students in the
// roster (Amber Bossons → existing roster value "Bossons, A"), an active
// multi-role Teacher with NO existing roster record (must fall back to a
// constructed value), an inactive Teacher, and an active non-teacher.
const USERS2_SCHEMA = { "Name": "field_1", "Role": "field_2", "Active": "field_3", "Email": "field_9" };
const USERS2_ROWS = [
  { id: "u1", field_1: "Amber Bossons", field_2: "Teacher", field_3: "Yes", field_9: "bossa@iu29.org" },
  { id: "u2", field_1: "Pat Rivera", field_2: "Teacher; Behavior Specialist", field_3: "Yes", field_9: "rivera@iu29.org" },
  { id: "u3", field_1: "Former Teacher", field_2: "Teacher", field_3: "No", field_9: "old@iu29.org" },
  { id: "u4", field_1: "Some Admin", field_2: "Administrator", field_3: "Yes", field_9: "admin@iu29.org" }
];

const ROSTER_SCHEMA = {
  "Student First Name": "sfield_1", "Student Last Name": "sfield_2", "Teacher": "sfield_3",
  "Classroom": "sfield_4", "Active": "sfield_5", "PACE Enabled": "sfield_6", "Daily Pulse Enabled": "sfield_7"
};
const ROSTER_ROWS = [
  { id: "s1", sfield_1: "Jane", sfield_2: "Doe", sfield_3: "Bossons, A", sfield_4: "Room 1", sfield_5: "Yes", sfield_6: "Yes", sfield_7: "No" }
];

function makeGraphMock() {
  const calls = { getListItems: [], createMappedListItem: [], updateMappedListItem: [], getWhoAreYouVisiting: 0 };
  return {
    calls,
    GRAPH: {
      async getListItems(listName) {
        calls.getListItems.push(listName);
        if (listName === "IEP_Users2") return USERS2_ROWS;
        if (listName === "IEP_Students_2026_27") return ROSTER_ROWS;
        if (listName === "macwalkthroughwhoareyouvisiting") throw new Error("must not be read directly as items — getWhoAreYouVisiting() is the only reader, and it must not be used for this dropdown");
        return [];
      },
      async getListSchema(listName) {
        if (listName === "IEP_Students_2026_27") return ROSTER_SCHEMA;
        return {};
      },
      async getWhoAreYouVisiting() { calls.getWhoAreYouVisiting++; return [{ spId: "t1", name: "Should Not Be Used" }]; },
      async createMappedListItem(listName, fields) { calls.createMappedListItem.push({ listName, fields }); return { id: "new-item" }; },
      async updateMappedListItem(listName, itemId, fields) { calls.updateMappedListItem.push({ listName, itemId, fields }); return {}; }
    }
  };
}

function makeContext() {
  const document = makeDocument();
  const { calls, GRAPH } = makeGraphMock();
  const context = {
    console, document, navigator: {},
    location: { origin: "http://localhost:5500", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected raw fetch — GRAPH is mocked directly"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone,
    addEventListener() {},
    confirm: () => true,
    GRAPH
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("config.js"), context, { filename: "config.js" });
  vm.runInContext(readSrc("auth.js"), context, { filename: "auth.js" });
  vm.runInContext(readSrc("data.js"), context, { filename: "data.js" });
  vm.runInContext(app, context, { filename: "app.js" });
  vm.runInContext("MAC_ADMIN_PANEL_ALLOWED = true;", context);
  return { context, document, calls };
}

async function verifyDropdownSourcedFromUsers2() {
  const { context, document, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();

  assert.ok(calls.getListItems.includes("IEP_Users2"), "teacher options must be read from IEP_Users2");

  const html = APP._renderStudentsTab();

  // Active teachers present, by their IEP_Users2 display name (human-readable label)
  assert.match(html, />Amber Bossons</, "an active Teacher must appear in the dropdown");
  assert.match(html, />Pat Rivera</, "an active multi-role user whose roles include Teacher must appear");

  // Excluded
  assert.doesNotMatch(html, />Former Teacher</, "an inactive Teacher must not appear");
  assert.doesNotMatch(html, />Some Admin</, "a non-teacher must not appear");
  assert.doesNotMatch(html, /Should Not Be Used/, "macwalkthroughwhoareyouvisiting names must never appear in this dropdown");
}

async function verifyCanonicalValuesNotDisplayNames() {
  const { context } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();
  const html = APP._renderStudentsTab();

  // Amber Bossons already has an existing roster student under "Bossons, A"
  // — the option's saved VALUE must be that exact existing string, with the
  // human-friendly name only as the visible label.
  assert.match(html, /<option value="Bossons, A">Amber Bossons<\/option>/,
    "the saved value must be the canonical roster key already in use, not the IEP_Users2 display name");

  // Pat Rivera has no existing roster record — falls back to a constructed
  // "Last, F" guess (IEP_Users2's own "Last, First" convention, abbreviated).
  assert.match(html, /<option value="Rivera, P">Pat Rivera<\/option>/,
    "with no existing roster record, a constructed Last, F fallback must be used, not the plain display name");
}

async function verifyAddStudentWritesCanonicalTeacherValue() {
  const { context, document, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();

  document.getElementById("newStudentFirstName").value = "New";
  document.getElementById("newStudentLastName").value  = "Student";
  document.getElementById("newStudentTeacher").value   = "Bossons, A"; // what the <select> would actually submit
  document.getElementById("newStudentActive").checked  = true;

  await document.getElementById("addStudentBtn")._handlers.click[0]();

  assert.equal(calls.createMappedListItem.length, 1);
  const write = calls.createMappedListItem[0];
  assert.equal(write.listName, "IEP_Students_2026_27");
  assert.equal(plain(write.fields)["Teacher"], "Bossons, A",
    "Add Student must write the canonical roster value, matching Daily Pulse/PACE's existing Teacher = \"Bossons, A\" students");
}

async function verifyDailyPulseMatchingSemanticsPreserved() {
  // STUDENT_ROSTER._matchesTeacher() (unmodified by this patch) is exact,
  // case/whitespace-insensitive string equality — confirms a student saved
  // with the resolved canonical value is actually findable by a teacher
  // signed in under that same roster name, the same way Daily Pulse's
  // getDailyPulseForTeacher() already relies on.
  const { context } = makeContext();
  const STUDENT_ROSTER = vm.runInContext("STUDENT_ROSTER", context);
  await STUDENT_ROSTER.refresh();
  const forBossons = STUDENT_ROSTER.getForTeacher("Bossons, A");
  assert.equal(forBossons.length, 1);
  assert.equal(forBossons[0].name, "Jane Doe");
  assert.equal(STUDENT_ROSTER.getForTeacher("Amber Bossons").length, 0,
    "matching is against the canonical roster value, not the IEP_Users2 display name");
}

function verifyNoUsers2OrWalkthroughTeacherWrites() {
  const studentsBlock = app.match(/if \(tab === "students"\) \{[^]*?\n    \}/)[0];
  assert.doesNotMatch(studentsBlock, /createMappedListItem\(SETUP_LISTS\.users/, "must never write to IEP_Users2");
  assert.doesNotMatch(studentsBlock, /updateMappedListItem\(SETUP_LISTS\.users/, "must never write to IEP_Users2");
  assert.doesNotMatch(studentsBlock, /macwalkthroughwhoareyouvisiting/i, "must never write to the Walkthrough picklist either");
}

async function verifyNoWritesOccurJustFromRendering() {
  const { context, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();
  APP._renderStudentsTab();
  assert.equal(calls.createMappedListItem.length, 0, "loading the page must never create a record");
  assert.equal(calls.updateMappedListItem.length, 0, "loading the page must never patch a record");
}

function verifyNoEditStudentTeacherPathExists() {
  // Setup currently has no "edit an existing student's teacher" capability
  // at all — only Add Student and the Active/Inactive toggle. Confirmed
  // directly rather than assumed, so this stays accurate if that changes.
  assert.doesNotMatch(app, /edit-student-btn|editStudent\(/i,
    "no student-edit capability exists yet — if one is added later it must reuse this same IEP_Users2 + canonical-value logic, not a new source");
}

Promise.all([
  verifyDropdownSourcedFromUsers2(),
  verifyCanonicalValuesNotDisplayNames(),
  verifyAddStudentWritesCanonicalTeacherValue(),
  verifyDailyPulseMatchingSemanticsPreserved(),
  verifyNoWritesOccurJustFromRendering()
])
  .then(() => {
    verifyNoUsers2OrWalkthroughTeacherWrites();
    verifyNoEditStudentTeacherPathExists();
    console.log("Teacher dropdown (IEP_Users2 → canonical roster value) tests passed.");
  })
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
