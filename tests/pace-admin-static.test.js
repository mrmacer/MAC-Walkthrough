const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "..");
const read = file => fs.readFileSync(path.join(root, file), "utf8");
const app = read("app.js");
const dashboard = read("dashboard-sync.js");
const paceAdmin = read("pace-admin.js");
const index = read("index.html");
const styles = read("styles.css");

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

// Student PACE History visit-detail UI (Notes priority patch)
assert.match(app, /_paceVisitDetailHtml\(visit, availability\)/);
assert.match(app, /class="pace-visit-notes-text"/, "full Notes text must render, not stay hidden behind a summary-only view");
assert.match(app, /No notes recorded\./);
assert.match(app, /openVisits/, "student history must surface open/in-progress visits, not just completed ones");
assert.match(app, /Currently in PACE/);
assert.doesNotMatch(app, /pace-visit-notes-text[^]{0,400}substring\(|pace-visit-notes-text[^]{0,400}\.slice\(/,
  "the rendered Notes block must not be truncated with substring/slice");

// Notes must preserve line breaks wherever rendered — both the detail panel
// and the main #pace table's existing Notes disclosure.
assert.match(styles, /\.pace-visit-notes-text\s*\{[^}]*white-space:\s*pre-wrap/,
  "detail Notes block must preserve line breaks (white-space: pre-wrap)");
assert.match(styles, /\.pace-notes p \{[^}]*white-space:\s*pre-wrap/,
  "main table's Notes disclosure must also preserve line breaks, not collapse them");

// Shared normalizer only — the detail UI must consume PACE_ADMIN's own
// specialists/roomInfo, not re-derive/re-parse visit fields itself.
assert.match(app, /PACE_ADMIN\.roomInfo\(/);
assert.match(paceAdmin, /specialists:\s*list\(firstValue\(row, FIELD_ALIASES\.specialist\)\)/);
assert.match(paceAdmin, /function roomInfo\(/);

console.log("PACE admin route and read-only integration checks passed.");
