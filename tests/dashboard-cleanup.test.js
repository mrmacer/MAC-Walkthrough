/* ─────────────────────────────────────────────────────────────────────────
   Dashboard cleanup — administrative correction patch

   Runs the REAL APP.renderDashboard() under node:vm against seeded
   localStorage data and captures what it actually writes into
   #page-dashboard's innerHTML, then asserts on that captured HTML — not
   just source-regex matching — so "removed sections absent" / "active
   metrics still calculate correctly" are genuinely exercised. AUTH stays
   unauthenticated (no account), so the async SharePoint-refresh half of
   renderDashboard() never runs; this test only covers the synchronous
   local-data render, which is where the removed sections/cards and the
   Walkthrough Support Requests relabel/recount all live.

   Run with: node tests/dashboard-cleanup.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");

function stubElement() {
  const node = {
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    addEventListener() {}, removeEventListener() {}, textContent: "", value: "", style: {}, dataset: {},
    querySelector() { return stubElement(); }, querySelectorAll() { return []; }, focus() {}, blur() {},
    innerHTML: "", appendChild() {}, removeAttribute() {}, setAttribute() {}, getAttribute() { return null; },
    closest() { return null; }, remove() {}, click() {}, scrollIntoView() {}, matches() { return false; }
  };
  return node;
}

function makeCapturingElement() {
  const el = stubElement();
  let html = "";
  Object.defineProperty(el, "innerHTML", { get: () => html, set: v => { html = v; } });
  return el;
}

function makeContext() {
  const store = {};
  const dashboardEl = makeCapturingElement();
  const document = {
    getElementById: id => (id === "page-dashboard" ? dashboardEl : stubElement()),
    querySelector: () => stubElement(),
    querySelectorAll: () => [],
    createElement: () => stubElement(),
    addEventListener() {},
    body: stubElement()
  };
  const context = {
    console, document, navigator: {},
    location: { origin: "http://localhost:5500", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: {
      getItem: k => (k in store ? store[k] : null),
      setItem: (k, v) => { store[k] = v; },
      removeItem: k => { delete store[k]; }
    },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected network request — renderDashboard must not go live when AUTH.isAuthenticated is false"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone,
    addEventListener() {} // app.js's own DOMContentLoaded auto-boot at file scope — never fired in this test
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("config.js"), context, { filename: "config.js" });
  vm.runInContext(readSrc("auth.js"), context, { filename: "auth.js" });
  vm.runInContext(readSrc("data.js"), context, { filename: "data.js" });
  vm.runInContext(readSrc("app.js"), context, { filename: "app.js" });
  return { context, store, dashboardEl };
}

function seed(store, key, value) { store[key] = JSON.stringify(value); }

async function verifyDashboardRender() {
  const { context, store, dashboardEl } = makeContext();
  const AUTH = vm.runInContext("AUTH", context);
  AUTH.role = "administrator"; // getter-backed AUTH.isAdmin needs a role, not a direct assignment

  // Two walkthrough support requests (one help-soon, one urgent) and one
  // "none" — Walkthrough Support Requests must count only the two real ones.
  seed(store, "skookWalkthrough_records", [
    { id: "r1", submittedAt: new Date().toISOString(), responses: { teacherId: "t1", classroomId: "c1", supportNeeded: "help-soon", classroomStatus: "monitor", supportsObserved: ["Visual Schedule"] } },
    { id: "r2", submittedAt: new Date().toISOString(), responses: { teacherId: "t2", classroomId: "c2", supportNeeded: "urgent", classroomStatus: "needs-support", supportsObserved: [] } },
    { id: "r3", submittedAt: new Date().toISOString(), responses: { teacherId: "t1", classroomId: "c1", supportNeeded: "none", classroomStatus: "on-track", supportsObserved: [] } }
  ]);
  seed(store, "skookWalkthrough_dailyPulses", [
    { id: "p1", timestamp: new Date().toISOString(), student: "Student A", pulseStatus: "concern", categories: ["Behavior"], supportLevel: "" },
    { id: "p2", timestamp: new Date().toISOString(), student: "Student B", pulseStatus: "great", categories: ["Behavior"], supportLevel: "need-help-now" }
  ]);
  seed(store, "skookWalkthrough_teachers", [{ id: "t1", name: "Teacher One" }, { id: "t2", name: "Teacher Two" }]);
  seed(store, "skookWalkthrough_classrooms", [{ id: "c1", name: "Room 1", teacherId: "t1" }, { id: "c2", name: "Room 2", teacherId: "t2" }]);

  const APP = vm.runInContext("APP", context);
  await APP.renderDashboard();
  const html = dashboardEl.innerHTML;

  // ── Removed sections absent ──────────────────────────────────────────
  for (const removedTitle of ["Student Pulse Trends", "Convergence Engine", "Emerging Patterns", "Student Voice Check-Ins", "Help Requests"]) {
    assert.doesNotMatch(html, new RegExp(`card-title">${removedTitle}\\b`), `"${removedTitle}" section must be removed from the dashboard`);
  }
  // PACE Activity is a different, retained card — make sure the removal
  // regexes above didn't accidentally also eat it.
  assert.match(html, /PACE Activity/, "PACE Activity card must remain");

  // ── Removed metric cards absent ──────────────────────────────────────
  for (const removedCard of ["Classrooms Visited", "Student Check-Ins Today", "Check-Outs Today", "Hard Mornings Today", "Rough Endings Today"]) {
    assert.doesNotMatch(html, new RegExp(`stat-label">${removedCard}<`), `"${removedCard}" metric card must be removed`);
  }

  // ── Active metric cards remain, correctly labeled ─────────────────────
  for (const keptCard of ["Total Walkthroughs", "This Week", "Teachers Visited", "Follow-Up Opportunities", "Help Soon", "Urgent",
    "Daily Pulse Entries", "Students Flagged Today", "Most Common Concern", "PACE Visits Today", "Students Seen Today", "Average PACE Duration"]) {
    assert.match(html, new RegExp(`stat-label"[^>]*>${keptCard}<`), `"${keptCard}" metric card must remain`);
  }

  // ── Support Requests relabeled and walkthrough-only ───────────────────
  assert.match(html, /stat-label">Walkthrough Support Requests</, "the card must be relabeled");
  assert.doesNotMatch(html, /stat-label">Support Requests</, "the old unqualified label must be gone");
  assert.match(html, /id="d-pulse-support">2</, "must count only the 2 real walkthrough support requests (help-soon + urgent), not the dead Daily Pulse contribution");

  // ── Other real metrics computed correctly from the seeded data ────────
  assert.match(html, /id="d-walk-total">3</);
  assert.match(html, /id="d-walk-helpsoon">1</);
  assert.match(html, /id="d-walk-urgent">1</);
  assert.match(html, /id="d-walk-followups">2</, "Follow-Up Opportunities counts any non-\"none\" support flag (help-soon + urgent)");
  assert.match(html, /id="d-pulse-total">2</);
  assert.match(html, /id="d-pulse-flagged">1</, "only the \"concern\" pulse counts as flagged today");
  assert.match(html, /id="d-pulse-concern">Behavior</, "Most Common Concern must still read from live Category data");
}

verifyDashboardRender()
  .then(() => console.log("Dashboard cleanup tests passed."))
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
