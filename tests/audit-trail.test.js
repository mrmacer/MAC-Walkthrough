/* ─────────────────────────────────────────────────────────────────────────
   Audit trail — Modified / Editor (Part 5)

   Exercises the REAL normalizeSharePointWalkthrough()/personDisplayName()
   (app.js) and PACE_ADMIN.normalizeVisit() (pace-admin.js) against the
   several shapes SharePoint's "Editor" Person field can plausibly come back
   as from Graph's $expand=fields, plus the case where it's absent entirely
   — the read must never throw and must never fabricate a name.

   Run with: node tests/audit-trail.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");

function stubElement() {
  const node = {
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    addEventListener() {}, removeEventListener() {}, textContent: "", value: "", style: {}, dataset: {},
    querySelector() { return stubElement(); }, querySelectorAll() { return []; },
    innerHTML: "", appendChild() {}, setAttribute() {}, getAttribute() { return null; },
    closest() { return null; }, remove() {}, matches() { return false; }
  };
  return node;
}

function makeContext() {
  const context = {
    console,
    document: { getElementById: () => stubElement(), querySelector: () => stubElement(), querySelectorAll: () => [], createElement: () => stubElement(), addEventListener() {}, body: stubElement() },
    navigator: {}, location: { origin: "http://localhost:5500", hash: "" },
    crypto: crypto.webcrypto,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: async () => { throw new Error("unexpected network request"); },
    msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone,
    addEventListener() {}
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("config.js"), context, { filename: "config.js" });
  vm.runInContext(readSrc("auth.js"), context, { filename: "auth.js" });
  vm.runInContext(readSrc("data.js"), context, { filename: "data.js" });
  vm.runInContext(app, context, { filename: "app.js" });
  return context;
}

function verifyPersonDisplayNameShapes() {
  const context = makeContext();
  const personDisplayName = vm.runInContext("personDisplayName", context);

  assert.equal(personDisplayName({ DisplayName: "Jane Smith" }), "Jane Smith");
  assert.equal(personDisplayName({ displayName: "Jane Smith" }), "Jane Smith");
  assert.equal(personDisplayName({ LookupValue: "Jane Smith", LookupId: 6 }), "Jane Smith");
  assert.equal(personDisplayName({ Title: "Jane Smith" }), "Jane Smith");
  assert.equal(personDisplayName({ LookupId: 6 }), "", "an object with none of the known name fields must yield \"\", never a guess");
  assert.equal(personDisplayName(undefined), "");
  assert.equal(personDisplayName(null), "");
  assert.equal(personDisplayName(""), "");
  assert.equal(personDisplayName("Plain String Editor"), "Plain String Editor", "a plain string value must pass through unchanged");
}

function verifyWalkthroughModifiedByExtraction() {
  const context = makeContext();
  const normalize = vm.runInContext("normalizeSharePointWalkthrough", context);

  const withEditor = normalize({ id: "1", Modified: "2026-09-10T12:00:00Z", Editor: { DisplayName: "Pat Rivera" } });
  assert.equal(withEditor.modified, "2026-09-10T12:00:00Z");
  assert.equal(withEditor.modifiedBy, "Pat Rivera");

  const withoutEditor = normalize({ id: "2", Modified: "2026-09-10T12:00:00Z" });
  assert.equal(withoutEditor.modified, "2026-09-10T12:00:00Z", "Modified must still be preserved even when Editor is absent");
  assert.equal(withoutEditor.modifiedBy, "", "missing Editor must yield \"\", not a fabricated name or a crash");

  const withUnusableEditor = normalize({ id: "3", Modified: "2026-09-10T12:00:00Z", Editor: { LookupId: 9 } });
  assert.equal(withUnusableEditor.modifiedBy, "", "an Editor object with no usable name field must yield \"\", never invent one");

  // Must never throw regardless of shape.
  assert.doesNotThrow(() => normalize({ id: "4" }));
}

function verifyPaceModifiedByExtraction() {
  const context = makeContext();
  vm.runInContext(readSrc("pace-admin.js"), context, { filename: "pace-admin.js" });
  const PACE_ADMIN = vm.runInContext("PACE_ADMIN", context);

  const withEditor = PACE_ADMIN.normalizeVisit({ Student: "A", Date: "2026-09-10", Modified: "2026-09-10T12:00:00Z", Editor: { displayName: "Kelly Marchetti" } });
  assert.equal(withEditor.modified, "2026-09-10T12:00:00Z");
  assert.equal(withEditor.modifiedBy, "Kelly Marchetti");

  const withoutEditor = PACE_ADMIN.normalizeVisit({ Student: "A", Date: "2026-09-10", Modified: "2026-09-10T12:00:00Z" });
  assert.equal(withoutEditor.modified, "2026-09-10T12:00:00Z");
  assert.equal(withoutEditor.modifiedBy, "", "missing Editor must yield \"\", not a crash or a guess");

  assert.doesNotThrow(() => PACE_ADMIN.normalizeVisit({ Student: "A", Date: "2026-09-10" }));
}

function verifyRenderingNeverAssumesEditorIsPresent() {
  // Static safety net: the rendered "Last modified by" text must always be
  // built with a fallback, never a bare interpolation that would print
  // "undefined"/blow up on an empty modifiedBy.
  assert.match(app, /record\.modifiedBy \|\| "Not available from the current SharePoint read"/,
    "the walkthrough detail modal must fall back rather than assume Editor is present");
  assert.match(app, /visit\.modified \? fmtDateTime\(visit\.modified\) : "Not available"/,
    "the PACE visit detail panel must fall back rather than assume Modified is present");
  assert.match(app, /visit\.modifiedBy \? " by " \+ escHtml\(visit\.modifiedBy\) : ""/,
    "the PACE visit detail panel must omit \"by ...\" entirely when Editor is unavailable, not print \"by \"");
}

try {
  verifyPersonDisplayNameShapes();
  verifyWalkthroughModifiedByExtraction();
  verifyPaceModifiedByExtraction();
  verifyRenderingNeverAssumesEditorIsPresent();
  console.log("Audit trail (Modified/Editor) tests passed.");
} catch (error) {
  console.error(error);
  process.exitCode = 1;
}
