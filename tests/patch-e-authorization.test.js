/* ─────────────────────────────────────────────────────────────────────────
   PATCH E — MAC-Walkthrough Admin Panel gate

   Covers:
   1. canAccessRoute()'s fixed default (no longer "anything but teacher gets
      full access") via a full APP.init() run against a mocked GRAPH/AUTH —
      app.js loads cleanly under node:vm with a minimal DOM/root stub.
   2. Protected data (any SharePoint list besides IEP_Users2/IEP_App_Users)
      must never be requested for a denied user, and the router itself
      must never even start listening for hash changes.

   Run with: node tests/patch-e-authorization.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");

function readSrc(file) { return fs.readFileSync(path.join(ROOT, file), "utf8"); }

// Self-referential: every querySelector/querySelectorAll on a stub element
// returns further stub elements (never null), since renderWalkthrough() and
// friends chain arbitrarily deep DOM traversal that a bare object graph
// can't otherwise satisfy. This app's rendering pipeline is out of scope
// for PATCH E — only the gate ahead of it is under test here.
function stubElement() {
  const node = {
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    addEventListener() {}, removeEventListener() {}, textContent: "", value: "", style: {}, dataset: {},
    querySelector() { return stubElement(); }, querySelectorAll() { return []; }, focus() {}, blur() {},
    innerHTML: "", appendChild() {}, removeAttribute() {}, setAttribute() {}, getAttribute() { return null; },
    closest() { return null; }, remove() {}, click() {}, scrollIntoView() {}, matches() { return false; },
    insertAdjacentHTML() {}, before() {}, after() {}, replaceWith() {}, cloneNode() { return stubElement(); }
  };
  return node;
}

function makeContext({ iepUsersRow = null, appUsersRows = [], appUsersError = null } = {}) {
  const protectedListsRequested = [];
  const listeners = {};
  const document = {
    getElementById: () => stubElement(),
    querySelectorAll: () => [],
    querySelector: () => stubElement(),
    addEventListener() {},
    createElement: () => stubElement(),
    body: stubElement()
  };
  const context = {
    console, document, navigator: {},
    location: { origin: "http://localhost:5500", hash: "", protocol: "http:", reload() {} },
    crypto: crypto.webcrypto,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected network request"); },
    msal: { PublicClientApplication: class {} },
    requestAnimationFrame: fn => fn(),
    setTimeout, clearTimeout, structuredClone,
    protectedListsRequested,
    addEventListener(event, handler) {
      listeners[event] = handler;
      if (event === "DOMContentLoaded") context._init = handler;
    },
    _listeners: listeners,
    GRAPH: {
      async findUserByEmail() { return iepUsersRow; },
      async getListItemsByDisplayName(listName) {
        if (listName === "IEP_App_Users") {
          if (appUsersError) throw new Error(appUsersError);
          return appUsersRows;
        }
        protectedListsRequested.push(listName);
        return [];
      },
      async getListItems(listName) { protectedListsRequested.push(listName); return []; },
      async getAllListItems(listName) { protectedListsRequested.push(listName); return []; }
    }
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("config.js"), context, { filename: "config.js" });
  vm.runInContext(readSrc("auth.js"), context, { filename: "auth.js" });
  vm.runInContext(readSrc("data.js"), context, { filename: "data.js" });
  vm.runInContext(readSrc("iep-app-users.js"), context, { filename: "iep-app-users.js" });
  vm.runInContext(readSrc("pace-admin.js"), context, { filename: "pace-admin.js" });
  vm.runInContext(readSrc("app.js"), context, { filename: "app.js" });
  return context;
}

function iepUsersRow(email, role) {
  return { Title: "u1", field_1: "Test User", field_2: role, field_3: true, Email: email };
}

async function verifyAdministratorWithAdminPanelYesIsAllowed() {
  const context = makeContext({
    iepUsersRow: iepUsersRow("admin@iu29.org", "Administrator"),
    appUsersRows: [{ id: "1", Email: "admin@iu29.org", "Admin Panel": "Yes" }]
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "admin@iu29.org", name: "Admin User" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  // init() must complete without throwing and reach the real render
  // pipeline (renderWalkthrough — the default landing page's initial paint
  // doesn't itself call GRAPH, it's backed by local PILOT_TEACHERS/DB data
  // until an actual submit; MAC_ADMIN_PANEL_ALLOWED === true is the direct
  // proof the gate passed and execution proceeded past the early return).
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), true);
}

async function verifyAdministratorWithAdminPanelNoIsDenied() {
  const context = makeContext({
    iepUsersRow: iepUsersRow("admin@iu29.org", "Administrator"),
    appUsersRows: [{ id: "1", Email: "admin@iu29.org", "Admin Panel": "No" }]
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "admin@iu29.org", name: "Admin User" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), false);
  assert.deepEqual(context.protectedListsRequested, [], "no protected list may be requested for a denied user — navigate() must never run");
  assert.equal("hashchange" in context._listeners, false, "the router must never start listening for hash changes for a denied user");
}

async function verifyTeacherRoleIsDenied() {
  // The explicit PATCH E test case: a "teacher" role with no IEP_App_Users
  // row must be denied entirely — not routed to #pulse the way this app
  // used to. Teachers belong in the standalone Daily Pulse app.
  const context = makeContext({
    iepUsersRow: iepUsersRow("teacher@iu29.org", "Teacher"),
    appUsersRows: []
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "teacher@iu29.org", name: "A Teacher" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), false);
  assert.deepEqual(context.protectedListsRequested, []);
}

async function verifyBehaviorSpecialistRoleIsDenied() {
  const context = makeContext({
    iepUsersRow: iepUsersRow("specialist@iu29.org", "Behavior Specialist"),
    appUsersRows: []
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "specialist@iu29.org", name: "A Specialist" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), false, "FIXED: a non-teacher role must no longer default-allow");
  assert.deepEqual(context.protectedListsRequested, []);
}

async function verifyUnknownRoleWithoutAppRowIsDenied() {
  const context = makeContext({
    iepUsersRow: iepUsersRow("mystery@iu29.org", "Some New Role"),
    appUsersRows: []
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "mystery@iu29.org", name: "Mystery Person" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), false);
}

async function verifyAdministratorWithoutRowGetsMigrationFallback() {
  const context = makeContext({
    iepUsersRow: iepUsersRow("admin@iu29.org", "Administrator"),
    appUsersRows: [] // no IEP_App_Users row yet
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "admin@iu29.org", name: "Admin User" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), true, "an active Administrator must still get in during migration");
}

async function verifyLookupFailureDeniesEvenAnAdministrator() {
  const context = makeContext({
    iepUsersRow: iepUsersRow("admin@iu29.org", "Administrator"),
    appUsersError: "Graph 403: insufficient privileges"
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "admin@iu29.org", name: "Admin User" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  assert.equal(vm.runInContext("MAC_ADMIN_PANEL_ALLOWED", context), false, "a lookup failure must deny even an active Administrator");
  assert.deepEqual(context.protectedListsRequested, []);
}

async function verifyDirectHashCannotBypassTheGate() {
  // Even if this app's router were somehow reachable, canAccessRoute()
  // itself must independently refuse a protected route for a denied
  // session — defense in depth beyond "the listener never registers."
  const context = makeContext({
    iepUsersRow: iepUsersRow("specialist@iu29.org", "Behavior Specialist"),
    appUsersRows: []
  });
  const auth = vm.runInContext("AUTH", context);
  auth.account = { username: "specialist@iu29.org", name: "A Specialist" };
  auth.init = async function () { await this.loadPilotUserFromSharePoint(); };
  await context._init();
  const canAccessRoute = vm.runInContext("canAccessRoute", context);
  for (const route of ["pace", "reports", "dashboard", "setup", "storage"]) {
    assert.equal(canAccessRoute(route), false, `#${route} must be denied for a Behavior Specialist with no Admin Panel grant`);
  }
}

Promise.all([
  verifyAdministratorWithAdminPanelYesIsAllowed(),
  verifyAdministratorWithAdminPanelNoIsDenied(),
  verifyTeacherRoleIsDenied(),
  verifyBehaviorSpecialistRoleIsDenied(),
  verifyUnknownRoleWithoutAppRowIsDenied(),
  verifyAdministratorWithoutRowGetsMigrationFallback(),
  verifyLookupFailureDeniesEvenAnAdministrator(),
  verifyDirectHashCannotBypassTheGate()
])
  .then(() => console.log("PATCH E (MAC-Walkthrough Admin Panel gate) tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
