/* ─────────────────────────────────────────────────────────────────────────
   PERMANENT FIX: getWhoAreYouVisiting() list-ID resolution

   Root cause recap: GRAPH.getWhoAreYouVisiting() hardcoded the SharePoint
   list's internal name directly into the Graph items URL
   (/lists/macwalkthroughwhoareyouvisiting/items), instead of resolving it
   to the list's real GUID via GRAPH.getListId() first — the pattern every
   other list reader in graph.js already uses. Graph's /lists/{id} endpoint
   requires the GUID, so this always 404'd (itemNotFound).

   Two layers, matching this codebase's established test style (see
   tests/admin-editing.test.js's own docstring):
   1. graph.js's getWhoAreYouVisiting() runs for REAL under node:vm against
      a mocked fetch, so the exact request sequence (site lookup → list
      discovery → items-by-real-GUID) is genuinely verified, not just
      pattern-matched from source.
   2. app.js's SETUP_DATA/Setup → Teachers error-surfacing is exercised for
      real under node:vm with GRAPH mocked directly (same technique as
      tests/teacher-dropdown.test.js).
   Static source assertions cover the "temporary diagnostic fully removed"
   and "no hardcoded GUID" requirements.

   Run with: node tests/teacher-directory-list-resolution.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app   = readSrc("app.js");
const graph = readSrc("graph.js");
const plain = obj => JSON.parse(JSON.stringify(obj));

const LIVE_LIST_ID = "baa21a52-3f47-415e-8803-28c565ea2578"; // the real tenant GUID, from live discovery — never hardcoded in source

/* ── layer 1: graph.js, real code, mocked fetch ───────────────────────────── */

function makeGraphContext({ listExists = true } = {}) {
  const calls = [];
  const SITE_ID = "site-1";
  const rows = [
    { id: "item-1", fields: { teacher: "Andruchek" } },
    { id: "item-2", fields: { teacher: "Bossons" } },
    { id: "item-3", fields: { teacher: "  " } } // blank — must be dropped
  ];

  const fetchMock = async (url) => {
    calls.push({ url });
    if (url.includes("siu29.sharepoint.com")) {
      return { ok: true, json: async () => ({ id: SITE_ID }) };
    }
    if (url.includes(`/sites/${SITE_ID}/lists?`)) {
      const value = listExists
        ? [
            { id: "list-users",    name: "IEP_Users2",     displayName: "IEP_Users2" },
            { id: LIVE_LIST_ID,    name: "macwalkthroughwhoareyouvisiting", displayName: "mac-walkthrough-who-are-you-visiting" }
          ]
        : [{ id: "list-users", name: "IEP_Users2", displayName: "IEP_Users2" }];
      return { ok: true, json: async () => ({ value }) };
    }
    if (url.includes(`/sites/${SITE_ID}/lists/${LIVE_LIST_ID}/items`)) {
      return { ok: true, json: async () => ({ value: rows }) };
    }
    if (url.includes("/lists/macwalkthroughwhoareyouvisiting/items")) {
      // The exact pre-fix (broken) request shape — must never be hit again.
      return { ok: false, status: 404, text: async () => JSON.stringify({ error: { code: "itemNotFound", message: "The specified list was not found" } }) };
    }
    throw new Error("Unexpected fetch in test: " + url);
  };

  const context = { console, window: {}, AUTH: { acquireGraphToken: async () => "fake-token" }, fetch: fetchMock };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(graph, context, { filename: "graph.js" });
  return { calls, GRAPH: vm.runInContext("GRAPH", context) };
}

async function resolvesListNameThroughGetListId() {
  const { calls, GRAPH } = makeGraphContext();
  const result = await GRAPH.getWhoAreYouVisiting();

  // Item 1 & 3: the resolution actually happened — a list-discovery request
  // was made, and the real GUID (not the internal name) was used for items.
  const listDiscovery = calls.filter(c => /\/lists\?/.test(c.url));
  assert.equal(listDiscovery.length, 1, "exactly one list-discovery request (getListId's own GET)");
  const itemsRequests = calls.filter(c => /\/items\?/.test(c.url));
  assert.equal(itemsRequests.length, 1, "exactly one items request");
  assert.match(itemsRequests[0].url, new RegExp(`/lists/${LIVE_LIST_ID}/items`), "items request uses the real resolved GUID");

  // Item 2: the raw internal name must never appear as the /lists/{id} path
  // segment — that's exactly the bug being fixed.
  calls.forEach(c => assert.ok(
    !/\/lists\/macwalkthroughwhoareyouvisiting\//.test(c.url),
    `internal name used directly as a list-id path segment: ${c.url}`
  ));

  // Item 4: teacher normalization unchanged — blank dropped, sorted, spId/name shape.
  assert.deepEqual(plain(result), [
    { spId: "item-1", name: "Andruchek" },
    { spId: "item-2", name: "Bossons" }
  ]);
}

async function itemsRequestUsesTheResolvedGuidNotTheName() {
  const { GRAPH } = makeGraphContext();
  // getListId() caches by name — confirm a second call reuses the cache
  // (no second list-discovery round trip) yet still resolves to the GUID.
  await GRAPH.getWhoAreYouVisiting();
  const { calls: calls2 } = { calls: [] };
  const before = GRAPH._listIdCache["macwalkthroughwhoareyouvisiting"];
  assert.equal(before, LIVE_LIST_ID, "resolved id is cached under the internal name");
}

async function missingListStillFailsClearlyNotSilently() {
  const { GRAPH } = makeGraphContext({ listExists: false });
  await assert.rejects(
    () => GRAPH.getWhoAreYouVisiting(),
    /SharePoint list not found: macwalkthroughwhoareyouvisiting/,
    "a genuinely missing list still fails, with getListId's clear message — not a raw, confusing Graph 404 body"
  );
}

function noHardcodedGuidOrRawNamePath() {
  const fn = graph.slice(graph.indexOf("async getWhoAreYouVisiting()"), graph.indexOf("async getWhoAreYouVisiting()") + 900);
  assert.ok(!fn.includes(LIVE_LIST_ID), "the live tenant GUID is never hardcoded in source");
  assert.ok(!/\/lists\/macwalkthroughwhoareyouvisiting\//.test(fn), "the internal name is never used directly as a /lists/{id} path segment");
  assert.match(fn, /getListId\(\s*["']macwalkthroughwhoareyouvisiting["']\s*\)/, "resolves through getListId(), the same pattern every other reader uses");
}

/* ── layer 2: SETUP_DATA / Setup → Teachers error surfacing ──────────────── */

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

function makeAppContext({ teachersFail = false } = {}) {
  const document = makeDocument();
  const context = {
    console, document, navigator: {},
    location: { origin: "http://localhost", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected raw fetch — GRAPH is mocked directly"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone, addEventListener() {}, removeEventListener() {}, confirm: () => true,
    GRAPH: {
      async getListItems(listName) { return listName === "IEP_Users2" ? [] : []; },
      async getListSchema() { return {}; },
      async getWhoAreYouVisiting() {
        if (teachersFail) throw new Error("Graph 404: list not found");
        return [{ spId: "sp-1", name: "Andruchek" }, { spId: "sp-2", name: "Bossons" }];
      }
    }
  };
  context.window = context;
  vm.createContext(context);
  ["config.js", "auth.js", "data.js", "pace-admin.js", "pace-export.js", "app.js"].forEach(f =>
    vm.runInContext(readSrc(f), context, { filename: f }));
  vm.runInContext('MAC_ADMIN_PANEL_ALLOWED = true; AUTH.role = "Administrator";', context);
  const APP = vm.runInContext("APP", context);
  const SETUP_DATA = vm.runInContext("SETUP_DATA", context);
  return { context, document, APP, SETUP_DATA };
}

async function teachersLoadFailureIsSurfacedNotSwallowed() {
  const { SETUP_DATA, APP } = makeAppContext({ teachersFail: true });
  await SETUP_DATA.refresh();
  assert.equal(SETUP_DATA.teachers.length, 0);
  assert.match(SETUP_DATA.teachersError, /Graph 404/, "the real rejection reason is kept, not discarded");

  const html = APP._renderTeachersTab();
  assert.match(html, /could not be loaded/i, "Setup → Teachers shows a load-failure banner");
  assert.ok(!/No teachers found/.test(html), "a failed load is never rendered as an indistinguishable empty state");
}

async function successfulTeachersLoadStillRendersNormally() {
  const { SETUP_DATA, APP } = makeAppContext({ teachersFail: false });
  await SETUP_DATA.refresh();
  assert.equal(SETUP_DATA.teachersError, null);
  assert.equal(SETUP_DATA.teachers.length, 2);

  const html = APP._renderTeachersTab();
  assert.ok(!/could not be loaded/i.test(html), "no error banner on a successful load");
  assert.match(html, /Andruchek/);
  assert.match(html, /Bossons/);
  assert.match(html, /Teachers \(2\)/);
}

function usersLoadFailureStillSurfacesSeparately() {
  // Regression guard: fixing the Teachers error path must not disturb the
  // pre-existing Users error path (SETUP_DATA.error), which the Users tab's
  // own banner is written for.
  const src = app.slice(app.indexOf("const SETUP_DATA = {"), app.indexOf("const SETUP_DATA = {") + 2000);
  assert.match(src, /usersR\.status === "rejected"\)\s*this\.error\s*=/);
  assert.match(src, /teachersR\.status === "rejected"/);
  assert.ok(!/this\.error\s*=\s*teachersR/.test(src), "the two error states stay independent (teachersError, not error)");
}

/* ── static: temporary diagnostic fully removed ───────────────────────────── */

function temporaryDiagnosticCodeIsGone() {
  [app, graph].forEach(src => {
    ["listSiteLists", "spDiagRunBtn", "_renderDiagTab", "_spDiagResult", "SP Diagnostic", "TEMPORARY DIAGNOSTIC"]
      .forEach(marker => assert.ok(!src.includes(marker), `leftover diagnostic marker "${marker}" still present`));
  });
  assert.ok(!app.includes('"diag"'), 'the "diag" Setup tab is gone');
  const tabsLine = app.match(/const tabs = \[[^\]]*\];/)[0];
  assert.deepEqual(JSON.parse(tabsLine.replace("const tabs = ", "").replace(";", "")),
    ["users", "teachers", "students", "classrooms", "data"], "Setup tabs are back to the normal five");
}

/* ── static: New Walkthrough still on TEACHER_DIRECTORY, no PILOT_TEACHERS fallback ── */

function newWalkthroughStillUsesTeacherDirectoryOnly() {
  const fn = app.slice(app.indexOf("renderWalkthrough() {"), app.indexOf("_submitWalkthrough(pageEl)"));
  assert.match(fn, /const allTeachers\s*=\s*TEACHER_DIRECTORY\.getAll\(\)/);
  assert.ok(!/const allTeachers\s*=\s*DB\.getTeachers\(\)/.test(fn));
  assert.ok(!/const allTeachers\s*=\s*PILOT_TEACHERS/.test(fn));
  assert.ok(!/PILOT_TEACHERS\.find/.test(fn), "no PILOT_TEACHERS fallback anywhere in renderWalkthrough()");
}

const tests = {
  resolvesListNameThroughGetListId, itemsRequestUsesTheResolvedGuidNotTheName,
  missingListStillFailsClearlyNotSilently, noHardcodedGuidOrRawNamePath,
  teachersLoadFailureIsSurfacedNotSwallowed, successfulTeachersLoadStillRendersNormally,
  usersLoadFailureStillSurfacesSeparately, temporaryDiagnosticCodeIsGone,
  newWalkthroughStillUsesTeacherDirectoryOnly
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} list-resolution fix checks passed.`);
})();
