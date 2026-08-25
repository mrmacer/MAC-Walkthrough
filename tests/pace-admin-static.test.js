const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const read = file => fs.readFileSync(path.join(root, file), "utf8");
const app = read("app.js");
const dashboard = read("dashboard-sync.js");
const paceAdmin = read("pace-admin.js");
const index = read("index.html");

assert.match(app, /const TEACHER_ALLOWED_ROUTES = new Set\(\["pulse"\]\)/);
assert.match(app, /pace:\s+\(\) => this\.renderPaceAdmin\(\)/);
assert.match(app, /if \(!AUTH\.isAdmin \|\| USER_CONTEXT\.isViewingAsTeacher\)/);
assert.match(app, /Administrator access is required to view PACE analytics and history/);
assert.match(app, /id="pf-from"/);
assert.match(app, /id="pf-student"/);
assert.match(app, /id="pf-search"/);
assert.match(app, /id="pace-history-modal"/);

assert.doesNotMatch(dashboard, /DB\.getPaceLogs/);
assert.match(dashboard, /PACE_ADMIN\.load\(\)/);
assert.doesNotMatch(paceAdmin, /\.(?:createListItem|createMappedListItem|updateMappedListItem|savePaceVisit|_post|_patch)\s*\(/);

const paceAdminIndex = index.indexOf('<script src="pace-admin.js"></script>');
const appIndex = index.indexOf('<script src="app.js"></script>');
const syncIndex = index.indexOf('<script src="dashboard-sync.js"></script>');
assert.ok(paceAdminIndex >= 0 && paceAdminIndex < appIndex && appIndex < syncIndex, "script load order must expose PACE_ADMIN before app/dashboard use");

for (const workflow of ["renderWalkthrough", "renderPulse", "renderStudentCheckIn", "renderReports"]) {
  assert.match(app, new RegExp(`\\b${workflow}\\(`), `${workflow} must remain present`);
}

console.log("PACE admin route and read-only integration checks passed.");
