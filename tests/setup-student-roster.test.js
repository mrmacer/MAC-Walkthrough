/* ─────────────────────────────────────────────────────────────────────────
   Setup → Students: redirect to the production roster (IEP_Students_2026_27)

   Runs the REAL Setup → Students render/bind/write code paths under
   node:vm against a mocked GRAPH (recording every call), exercising
   SETUP_DATA.refresh(), APP._renderStudentsTab(), and
   APP._bindSetupTabEvents("students")'s Add/Deactivate/Reactivate
   handlers — not just source-regex matching, for the same reason
   tests/admin-editing.test.js runs graph.js for real: this is a write
   path, and "does it actually call the right list/fields" needs to be
   proven, not just implied by the source text.

   Run with: node tests/setup-student-roster.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");

// Objects built by code running inside a vm.Context (e.g. mapFields()'s
// `{}` literal in graph.js) have that context's own Object.prototype, not
// this test file's — assert.deepEqual (aliased to deepStrictEqual under
// node:assert/strict) treats that as unequal even with identical own
// properties. Round-tripping through JSON strips it to a plain object in
// this realm before comparing.
const plain = obj => JSON.parse(JSON.stringify(obj));

/* ── DOM stub with a stable per-id element registry, so a value set on an
   element before a handler runs is the same object the handler reads. ── */
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

function makeGraphMock() {
  const calls = { getListItems: [], createMappedListItem: [], updateMappedListItem: [], getWhoAreYouVisiting: 0 };
  const ROSTER_SCHEMA = {
    "Student First Name": "field_1", "Student Last Name": "field_2", "Teacher": "field_3",
    "Classroom": "field_4", "Active": "field_5", "PACE Enabled": "field_6", "Daily Pulse Enabled": "field_7"
  };
  const ROSTER_ROWS = [
    { id: "roster-1", field_1: "Jane", field_2: "Doe", field_3: "Stock", field_4: "Room 1", field_5: "Yes", field_6: "Yes", field_7: "No" },
    { id: "roster-2", field_1: "Sam", field_2: "Lee", field_3: "Bossons", field_4: "Room 2", field_5: "No", field_6: "No", field_7: "Yes" }
  ];
  return {
    calls,
    GRAPH: {
      async getListItems(listName) {
        calls.getListItems.push(listName);
        if (listName === "IEP_Students_2026_27") return ROSTER_ROWS;
        if (listName === "IEP_Skook_Pilot_Students") throw new Error("must not be read for Setup Students");
        return [];
      },
      async getListSchema(listName) {
        if (listName === "IEP_Students_2026_27") return ROSTER_SCHEMA;
        return {};
      },
      async getWhoAreYouVisiting() { calls.getWhoAreYouVisiting++; return [{ spId: "t1", name: "Stock" }]; },
      async createMappedListItem(listName, fields) { calls.createMappedListItem.push({ listName, fields }); return { id: "new-item" }; },
      async updateMappedListItem(listName, itemId, fields) { calls.updateMappedListItem.push({ listName, itemId, fields }); return {}; }
    }
  };
}

function makeContext({ adminPanelAllowed = true } = {}) {
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
  if (adminPanelAllowed) {
    vm.runInContext("MAC_ADMIN_PANEL_ALLOWED = true;", context);
  }
  return { context, document, calls };
}

async function verifyRefreshReadsProductionRosterOnly() {
  const { context, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  await SETUP_DATA.refresh();

  assert.ok(calls.getListItems.includes("IEP_Students_2026_27"), "Setup must read IEP_Students_2026_27");
  assert.ok(!calls.getListItems.includes("IEP_Skook_Pilot_Students"), "Setup must NOT read the old pilot list");
  assert.equal(SETUP_DATA.students.length, 2, "existing production students must render");
  assert.deepEqual(SETUP_DATA.students.map(s => s.name), ["Jane Doe", "Sam Lee"],
    "first/last names must combine correctly for display while staying separate fields internally");
  assert.equal(SETUP_DATA.students[0].firstName, "Jane");
  assert.equal(SETUP_DATA.students[0].lastName, "Doe");
}

async function verifySourceLabelReflectsProductionRoster() {
  const { context } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  await SETUP_DATA.refresh();
  const APP = vm.runInContext("APP", context);
  const html = APP._renderStudentsTab();
  assert.match(html, /Source: IEP_STUDENTS_2026_27/, "source label must reflect the production roster");
  assert.doesNotMatch(html, /IEP_SKOOK_PILOT_STUDENTS/i, "the old pilot list must not appear as the source");
  assert.match(html, /Jane Doe/);
  assert.match(html, /Sam Lee/);
  assert.match(html, />Teacher</); // column header, not the old "Pilot ID" column
  assert.doesNotMatch(html, /Pilot ID/i, "the old Pilot ID field/column must be gone");
}

async function verifyAddStudentWritesProductionFieldsWithDefaults() {
  const { context, document, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  // Matches renderSetup()'s own real flow: the active tab must be set before
  // refresh() runs, because refresh() triggers _refreshSetupUI() internally,
  // which renders/binds whatever tab is currently active.
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();

  document.getElementById("newStudentFirstName").value = "Alex";
  document.getElementById("newStudentLastName").value  = "Kim";
  document.getElementById("newStudentTeacher").value   = "Stock";
  document.getElementById("newStudentClassroom").value = "Room 3";
  document.getElementById("newStudentActive").checked  = true;
  document.getElementById("newStudentPace").checked    = false;
  document.getElementById("newStudentPulse").checked   = false;

  await document.getElementById("addStudentBtn")._handlers.click[0]();

  assert.equal(calls.createMappedListItem.length, 1, "Add Student must issue exactly one write");
  const write = calls.createMappedListItem[0];
  assert.equal(write.listName, "IEP_Students_2026_27", "must write to the production roster");
  assert.notEqual(write.listName, "IEP_Skook_Pilot_Students");
  assert.deepEqual(plain(write.fields), {
    "Student First Name": "Alex",
    "Student Last Name":  "Kim",
    "Teacher":             "Stock",
    "Classroom":           "Room 3",
    "Active":              "Yes",
    "PACE Enabled":        "No",
    "Daily Pulse Enabled": "No"
  });
  assert.equal("Title" in write.fields, false, "no fabricated Pilot ID field");
  assert.equal("Student Name" in write.fields, false, "must not combine into the old pilot-style single name field");
}

function verifyDefaultCheckboxMarkup() {
  // The stub DOM can't read the `checked` attribute out of a raw HTML
  // string, so the actual default state is asserted directly against the
  // rendered markup for the three checkboxes.
  const formFn = app.match(/_renderStudentsTab\(\) \{[^]*?\n  \},/)[0];
  assert.match(formFn, /id="newStudentActive"[^>]*checked>/, "Active must default to checked (Yes)");
  assert.doesNotMatch(formFn, /id="newStudentPace"[^>]*checked>/, "PACE Enabled must default to unchecked (No)");
  assert.doesNotMatch(formFn, /id="newStudentPulse"[^>]*checked>/, "Daily Pulse Enabled must default to unchecked (No)");
}

async function verifyToggleTogglePatchesExistingItemNoDuplicate() {
  const { context, document, calls } = makeContext();
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();

  const tc = document.registry["tabContent"];
  const fakeEvent = (selector, dataset) => ({ target: { closest: sel => (sel === selector ? { dataset } : null) } });

  await tc._handlers.click[0](fakeEvent(".deactivate-student-btn", { spId: "roster-1", name: "Jane Doe" }));

  assert.equal(calls.createMappedListItem.length, 0, "deactivating must never create a new record");
  assert.equal(calls.updateMappedListItem.length, 1, "deactivating must issue exactly one PATCH");
  const patch = calls.updateMappedListItem[0];
  assert.equal(patch.listName, "IEP_Students_2026_27");
  assert.equal(patch.itemId, "roster-1", "must PATCH the existing item's own SharePoint id");
  assert.deepEqual(plain(patch.fields), { "Active": "No" });

  await tc._handlers.click[0](fakeEvent(".reactivate-student-btn", { spId: "roster-2", name: "Sam Lee" }));
  assert.equal(calls.createMappedListItem.length, 0, "reactivating must never create a new record either");
  assert.equal(calls.updateMappedListItem.length, 2);
  assert.equal(calls.updateMappedListItem[1].itemId, "roster-2");
  assert.deepEqual(plain(calls.updateMappedListItem[1].fields), { "Active": "Yes" });
}

async function verifyAdminPanelRequiredForWrites() {
  const { context, document, calls } = makeContext({ adminPanelAllowed: false });
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  const APP = vm.runInContext("APP", context);
  APP._setupActiveTab = "students";
  await SETUP_DATA.refresh();

  document.getElementById("newStudentFirstName").value = "Blocked";
  document.getElementById("newStudentLastName").value  = "User";
  document.getElementById("newStudentTeacher").value   = "Stock";

  await document.getElementById("addStudentBtn")._handlers.click[0]();
  assert.equal(calls.createMappedListItem.length, 0, "Add Student must be blocked when MAC_ADMIN_PANEL_ALLOWED is false");

  const tc = document.registry["tabContent"];
  const fakeEvent = (selector, dataset) => ({ target: { closest: sel => (sel === selector ? { dataset } : null) } });
  await tc._handlers.click[0](fakeEvent(".deactivate-student-btn", { spId: "roster-1", name: "Jane Doe" }));
  assert.equal(calls.updateMappedListItem.length, 0, "deactivate must also be blocked when MAC_ADMIN_PANEL_ALLOWED is false");
}

function verifyRequireSetupAdminUsesAdminPanelBoundary() {
  assert.match(app, /function requireSetupAdmin\(\) \{\s*if \(!MAC_ADMIN_PANEL_ALLOWED\) throw new Error/,
    "requireSetupAdmin() must use the same MAC_ADMIN_PANEL_ALLOWED boundary as the rest of MAC");
  assert.doesNotMatch(app, /function requireSetupAdmin\(\) \{\s*if \(!AUTH\.isAdmin\)/,
    "the old, weaker AUTH.isAdmin-only check must be gone");
}

function verifyNoMigrationOrDeletionOfPilotList() {
  // The old list must still be defined (not deleted from source) and must
  // never be written to as part of this patch.
  assert.match(app, /students:\s*"IEP_Skook_Pilot_Students"/, "the old pilot list constant must remain defined");
  assert.doesNotMatch(app, /createMappedListItem\(SETUP_LISTS\.students/, "nothing may write new records into the old pilot list");
  assert.doesNotMatch(app, /updateMappedListItem\(SETUP_LISTS\.students/, "nothing may patch records in the old pilot list");
  assert.doesNotMatch(app, /deleteListItem/, "no delete capability was introduced against any list");
  // Its only remaining reference must be STUDENT_ROSTER's own defensive fallback.
  const nonCommentRefs = app.split("\n").filter(l => l.includes("SETUP_LISTS.students") && !l.trim().startsWith("//"));
  assert.equal(nonCommentRefs.length, 1, "SETUP_LISTS.students must have exactly one live reference left: STUDENT_ROSTER's fallback");
  assert.match(nonCommentRefs[0], /return SETUP_LISTS\.students;/);
}

Promise.all([
  verifyRefreshReadsProductionRosterOnly(),
  verifySourceLabelReflectsProductionRoster(),
  verifyAddStudentWritesProductionFieldsWithDefaults(),
  verifyToggleTogglePatchesExistingItemNoDuplicate(),
  verifyAdminPanelRequiredForWrites()
])
  .then(() => {
    verifyDefaultCheckboxMarkup();
    verifyRequireSetupAdminUsesAdminPanelBoundary();
    verifyNoMigrationOrDeletionOfPilotList();
    console.log("Setup Students → production roster tests passed.");
  })
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
