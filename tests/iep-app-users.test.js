/* ─────────────────────────────────────────────────────────────────────────
   PATCH A — IEP_App_Users reader coverage

   Plain-script style, matching this project's other tests (no test
   framework — assertions throw, Promise.all + process.exitCode surfaces
   failures). Run with: node tests/iep-app-users.test.js

   Covers only iep-app-users.js in isolation. No app.js/auth.js behavior
   changes yet in PATCH A (including the canAccessRoute() default-allow gap
   from the Phase 1 audit) — that arrives with the Admin Panel gate patch.
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");

function loadModule(context) {
  vm.runInContext(fs.readFileSync(path.join(ROOT, "iep-app-users.js"), "utf8"), context, { filename: "iep-app-users.js" });
}

function makeContext(overrides = {}) {
  return vm.createContext({
    console,
    CONFIG: { APP_USERS_LIST: "IEP_App_Users" },
    ...overrides
  });
}

async function verifyBooleanNormalization() {
  const context = makeContext();
  loadModule(context);
  const normalize = vm.runInContext("normalizeAppUsersBoolean", context);
  for (const truthy of ["Yes", "yes", " YES ", "true", "1", true, 1]) {
    assert.equal(normalize(truthy), true);
  }
  for (const falsy of ["No", "no", "false", "0", false, 0]) {
    assert.equal(normalize(falsy), false);
  }
  assert.equal(normalize(undefined, true), true);
  assert.equal(normalize("", false), false);
  assert.equal(normalize(null), false);
}

async function verifyDefaultIsMigrationMode() {
  const context = makeContext();
  loadModule(context);
  assert.equal(vm.runInContext("ACCESS_MODEL_CONFIG", context).ENFORCE_APP_USERS, false);
}

async function verifyResolveFindsMatchByEmailCaseInsensitive() {
  let calledArgs = null;
  const context = makeContext({
    GRAPH: {
      async getListItemsByDisplayName(...args) {
        calledArgs = args;
        return [
          { id: "1", "Display Name": "Admin User", "Role": "Administrator", "Email": "Admin@example.org", "Admin Panel": "Yes", "PACE": "Yes", "Walkthrough": "Yes", "Daily Pulse": "Yes", "Teacher/Classroom": "" }
        ];
      }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("  ADMIN@EXAMPLE.ORG  ");
  assert.deepEqual(calledArgs, ["IEP_App_Users", true]);
  assert.ok(row);
  assert.equal(appUsers.found, true);
  assert.equal(appUsers.lookupError, null);
  assert.deepEqual(JSON.parse(JSON.stringify(appUsers.permissions)), { dailyPulse: true, pace: true, walkthrough: true, adminPanel: true });
}

async function verifyPartialAdminPermissionsAreExplicit() {
  // Explicit permissions stay authoritative even for an Administrator role
  // — Admin Panel = Yes with PACE = No must not be "upgraded" to full access.
  const context = makeContext({
    GRAPH: {
      async getListItemsByDisplayName() {
        return [{ id: "1", "Role": "Administrator", "Email": "partial-admin@example.org", "Admin Panel": "Yes", "PACE": "No", "Walkthrough": "No", "Daily Pulse": "No" }];
      }
    }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  await appUsers.resolve("partial-admin@example.org");
  assert.deepEqual(JSON.parse(JSON.stringify(appUsers.permissions)), { dailyPulse: false, pace: false, walkthrough: false, adminPanel: true });
}

async function verifyResolveNoMatchReturnsNullCleanly() {
  const context = makeContext({
    GRAPH: { async getListItemsByDisplayName() { return [{ id: "1", Email: "someone-else@example.org" }]; } }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("nobody@example.org");
  assert.equal(row, null);
  assert.equal(appUsers.found, false);
  assert.equal(appUsers.permissions, null);
  assert.equal(appUsers.lookupError, null, "no match is not an error");
}

async function verifyResolveFailsClosedOnGraphError() {
  const context = makeContext({
    GRAPH: { async getListItemsByDisplayName() { throw new Error("Graph 403: insufficient privileges"); } }
  });
  loadModule(context);
  const appUsers = vm.runInContext("APP_USERS", context);
  const row = await appUsers.resolve("admin@example.org");
  assert.equal(row, null, "a failed lookup must never resolve to a match");
  assert.match(appUsers.lookupError, /insufficient privileges/);
}

Promise.all([
  verifyBooleanNormalization(),
  verifyDefaultIsMigrationMode(),
  verifyResolveFindsMatchByEmailCaseInsensitive(),
  verifyPartialAdminPermissionsAreExplicit(),
  verifyResolveNoMatchReturnsNullCleanly(),
  verifyResolveFailsClosedOnGraphError()
])
  .then(() => console.log("IEP_App_Users reader (PATCH A) tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
