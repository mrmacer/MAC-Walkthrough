/* ─────────────────────────────────────────────────────────────────────────
   Admin editing (Walkthrough + PACE) — administrative correction patch

   Two layers, matching this project's existing test styles:
   1. graph.js's updateWalkthrough()/updatePaceVisit() run for REAL under
      node:vm against a mocked fetch/AUTH, so PATCH-vs-POST, the exact URL
      (list + item id), and the exact field payload are genuinely verified —
      not just pattern-matched from source (same technique as
      tests/patch-e-authorization.test.js).
   2. app.js's UI wiring (authorization gate, Classroom omission, stale
      field-name absence, Duration-not-directly-editable, Time Out
      conditional inclusion, Cancel performs no write) is verified with
      static source assertions, matching tests/pace-admin-static.test.js's
      established approach for this codebase's DOM-heavy areas.

   Run with: node tests/admin-editing.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const app = readSrc("app.js");

/* ── Layer 1: graph.js updateWalkthrough()/updatePaceVisit() for real ────── */

function makeGraphContext() {
  const calls = [];
  const SITE_ID = "site-1";
  const LISTS = {
    IEP_Walkthrough_Observations: {
      id: "list-walk",
      columns: {
        "Teacher": "Teacher", "Student": "Student", "Focus": "Focus",
        "Environment": "Environment", "Engagement": "Engagement",
        "SupportObserved": "SupportObserved", "DisengagementReasons": "DisengagementReasons",
        "SupportRequested": "SupportRequested", "Observed Win": "ObservedWin",
        "Concern / Gap": "ConcernGap", "Observation Notes": "ObservationNotes",
        "Follow-Up Notes": "FollowUpNotes", "Follow-Up Date": "FollowUpDate",
        "Priority": "Priority", "Follow-Up Needed": "FollowUpNeeded",
        "Classroom": "Classroom", "Submission ID": "SubmissionID",
        "Observation ID": "ObservationID", "Modified": "Modified", "Editor": "Editor"
      }
    },
    IEP_Pace_Visits: {
      id: "list-pace",
      columns: {
        "Student": "Student", "Room": "Room", "Time In": "TimeIn", "Time Out": "TimeOut",
        "Duration": "Duration", "Reason": "Reason", "Intervention Used": "InterventionUsed",
        "Behavior Specialist": "BehaviorSpecialist", "Teacher Came From": "TeacherCameFrom",
        "SCM Used": "SCMUsed", "Notes": "Notes", "Modified": "Modified", "Editor": "Editor"
      }
    }
  };

  const fetchMock = async (url, options = {}) => {
    const method = options.method || "GET";
    calls.push({ method, url, body: options.body ? JSON.parse(options.body) : null });

    if (url.includes("siu29.sharepoint.com")) {
      return { ok: true, json: async () => ({ id: SITE_ID }) };
    }
    if (url.includes(`/sites/${SITE_ID}/lists?`)) {
      return { ok: true, json: async () => ({ value: Object.entries(LISTS).map(([name, def]) => ({ id: def.id, name, displayName: name })) }) };
    }
    const columnsMatch = url.match(/\/sites\/[^/]+\/lists\/([^/]+)\/columns/);
    if (columnsMatch) {
      const listDef = Object.values(LISTS).find(l => l.id === columnsMatch[1]);
      return { ok: true, json: async () => ({ value: Object.entries(listDef.columns).map(([displayName, name]) => ({ displayName, name, text: {} })) }) };
    }
    const patchMatch = url.match(/\/sites\/[^/]+\/lists\/([^/]+)\/items\/([^/]+)\/fields/);
    if (patchMatch && method === "PATCH") {
      return { ok: true, status: 200, json: async () => ({ id: patchMatch[2] }) };
    }
    throw new Error("Unexpected fetch in test: " + method + " " + url);
  };

  const context = {
    console,
    window: {},
    AUTH: { acquireGraphToken: async () => "fake-token" },
    fetch: fetchMock
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(readSrc("graph.js"), context, { filename: "graph.js" });
  return { context, calls, GRAPH: vm.runInContext("GRAPH", context) };
}

async function verifyWalkthroughEditPatchesNotPosts() {
  const { calls, GRAPH } = makeGraphContext();
  await GRAPH.updateWalkthrough("wt-item-42", {
    "Teacher": "Ms. Rivera", "Focus": "small-group", "Follow-Up Needed": true
  });

  const writes = calls.filter(c => c.method !== "GET");
  assert.equal(writes.length, 1, "editing must issue exactly one write call");
  assert.equal(writes[0].method, "PATCH", "editing an existing walkthrough must PATCH, never POST");
  assert.match(writes[0].url, /\/items\/wt-item-42\/fields$/, "PATCH must target the exact existing SharePoint item id");
  assert.match(writes[0].url, /\/lists\/list-walk\//, "PATCH must target IEP_Walkthrough_Observations, not V2 or any other list");
  assert.deepEqual(writes[0].body, { Teacher: "Ms. Rivera", Focus: "small-group", FollowUpNeeded: true },
    "only the fields actually passed must be written — nothing else touched");
}

async function verifyWalkthroughEditNeverSendsClassroomWhenOmitted() {
  const { calls, GRAPH } = makeGraphContext();
  await GRAPH.updateWalkthrough("wt-item-7", { "Teacher": "Mr. Owusu" });
  const patch = calls.find(c => c.method === "PATCH");
  assert.equal("Classroom" in patch.body, false, "Classroom must never be sent unless the caller explicitly includes it");
}

async function verifyWalkthroughEditNeverSendsProtectedFields() {
  const { calls, GRAPH } = makeGraphContext();
  // Simulates what would happen if a caller mistakenly tried to resend a
  // protected field — mapFields()'s existing blocklist (unmodified by this
  // patch) must still strip it before the PATCH is sent.
  await GRAPH.updateWalkthrough("wt-item-9", { "Teacher": "X", "Modified": "2020-01-01", "Editor": "Someone" });
  const patch = calls.find(c => c.method === "PATCH");
  assert.equal("Modified" in patch.body, false, "Modified must never be writable");
  assert.equal("Editor" in patch.body, false, "Editor must never be writable");
}

async function verifyWalkthroughUpdateRequiresItemId() {
  const { GRAPH } = makeGraphContext();
  await assert.rejects(() => GRAPH.updateWalkthrough(null, { Teacher: "X" }), /item id is required/);
  await assert.rejects(() => GRAPH.updateWalkthrough("", { Teacher: "X" }), /item id is required/);
}

async function verifyPaceEditPatchesNotPosts() {
  const { calls, GRAPH } = makeGraphContext();
  await GRAPH.updatePaceVisit("pace-item-3", { "Notes": "Corrected note text", "SCM Used": true });

  const writes = calls.filter(c => c.method !== "GET");
  assert.equal(writes.length, 1, "editing must issue exactly one write call");
  assert.equal(writes[0].method, "PATCH", "editing an existing PACE visit must PATCH, never POST");
  assert.match(writes[0].url, /\/items\/pace-item-3\/fields$/);
  assert.match(writes[0].url, /\/lists\/list-pace\//, "PATCH must target IEP_Pace_Visits");
  assert.deepEqual(writes[0].body, { Notes: "Corrected note text", SCMUsed: true });
}

async function verifyPaceEditUsesModernFieldNamesNotStaleOnes() {
  const { calls, GRAPH } = makeGraphContext();
  // The exact confirmed-live display names the audit established — using
  // the stale "Behavior"/"Interventions" names here would silently drop the
  // field (mapFields warns and skips unmapped display names).
  await GRAPH.updatePaceVisit("pace-item-5", {
    "Room": "PACE Room 1", "Reason": "Needs a Break", "Intervention Used": "Sensory Break",
    "Behavior Specialist": "Kelly Marchetti", "Teacher Came From": "Sickle"
  });
  const patch = calls.find(c => c.method === "PATCH");
  assert.deepEqual(patch.body, {
    Room: "PACE Room 1", Reason: "Needs a Break", InterventionUsed: "Sensory Break",
    BehaviorSpecialist: "Kelly Marchetti", TeacherCameFrom: "Sickle"
  }, "modern field names must map through to their live internal columns");
}

async function verifyPaceEditStaleFieldNamesAreDroppedNotSent() {
  const { calls, GRAPH } = makeGraphContext();
  // If something did send the old names, they must not silently reappear as
  // some other column — the live schema simply has no "Behavior"/
  // "Interventions" column, so mapFields() drops them.
  await GRAPH.updatePaceVisit("pace-item-6", { "Behavior": "Peer conflict", "Interventions": "Calm space", "Notes": "kept" });
  const patch = calls.find(c => c.method === "PATCH");
  assert.deepEqual(patch.body, { Notes: "kept" }, "stale Behavior/Interventions display names must not map to any live column");
}

async function verifyPaceEditRequiresItemId() {
  const { GRAPH } = makeGraphContext();
  await assert.rejects(() => GRAPH.updatePaceVisit(null, { Notes: "x" }), /item id is required/);
}

async function verifyPaceEditOmittingTimeOutTouchesNothing() {
  const { calls, GRAPH } = makeGraphContext();
  // Editing only Notes on an open visit — Time Out/Duration must never
  // appear in the payload, so the item's open state cannot change.
  await GRAPH.updatePaceVisit("pace-item-open", { "Notes": "Still working with student" });
  const patch = calls.find(c => c.method === "PATCH");
  assert.equal("TimeOut" in patch.body, false);
  assert.equal("Duration" in patch.body, false);
}

/* ── Layer 2: app.js UI wiring — static assertions ────────────────────────
   Matching tests/pace-admin-static.test.js's established approach for this
   codebase's DOM-heavy render methods. */

function verifyWalkthroughEditGatedOnAdminPanel() {
  assert.match(app, /MAC_ADMIN_PANEL_ALLOWED \? `\s*<div class="detail-section detail-actions">\s*<button type="button" class="btn btn-secondary" id="walkthroughEditBtn">Edit<\/button>/,
    "the walkthrough Edit button must only render when MAC_ADMIN_PANEL_ALLOWED is true");
  assert.match(app, /function renderWalkthroughEditForm\(record\) \{\s*if \(!MAC_ADMIN_PANEL_ALLOWED\) return;/,
    "the edit form entry point must also independently re-check the gate (defense in depth)");
  assert.match(app, /async function saveWalkthroughEdit\(record, formEl\) \{[^]*?if \(!MAC_ADMIN_PANEL_ALLOWED\) return;/,
    "saving must independently re-check the gate too");
}

function verifyPaceEditGatedOnAdminPanel() {
  assert.match(app, /MAC_ADMIN_PANEL_ALLOWED \? `\s*<div class="detail-actions"[^]*?pace-visit-edit-btn/,
    "the PACE Edit button must only render when MAC_ADMIN_PANEL_ALLOWED is true");
  assert.match(app, /_savePaceVisitEdit\(visit, formEl\) \{\s*if \(this\._paceEditSaving\) return;\s*if \(!MAC_ADMIN_PANEL_ALLOWED\) return;/,
    "saving a PACE edit must independently re-check the gate");
}

function verifyWalkthroughCancelPerformsNoWrite() {
  const cancelBlock = app.match(/document\.getElementById\("walkthroughEditCancel"\)\.addEventListener\("click", \(\) => \{[^}]*\}\);/);
  assert.ok(cancelBlock, "walkthrough Cancel handler must exist");
  assert.doesNotMatch(cancelBlock[0], /GRAPH\./, "Cancel must never call GRAPH — zero writes");
  assert.match(cancelBlock[0], /renderReportDetailModal\(record\)/, "Cancel must just re-render the unmodified record");
}

function verifyPaceCancelPerformsNoWrite() {
  const cancelBlock = app.match(/pace-visit-edit-cancel[^]*?btn\.addEventListener\("click", e => \{[^}]*\}\);/);
  assert.ok(cancelBlock, "PACE Cancel handler must exist");
  assert.doesNotMatch(cancelBlock[0], /GRAPH\./, "Cancel must never call GRAPH — zero writes");
}

function verifyWalkthroughProtectedFieldsNeverBuilt() {
  const saveFn = app.match(/async function saveWalkthroughEdit[^]*?\n}\n/)[0];
  for (const protectedField of ['"Observation ID"', '"Session ID"', '"Submission ID"', '"Created"', '"Modified"', '"AI Summary"', '"AI Suggestions"', '"Classroom"']) {
    assert.doesNotMatch(saveFn, new RegExp(protectedField + "\\s*:"),
      `${protectedField} must never be assembled into the walkthrough edit payload`);
  }
}

function verifyPaceDurationNotDirectlyEditable() {
  const formFn = app.match(/_paceVisitEditFormHtml\(visit, availability\) \{[^]*?\n  \},/)[0];
  assert.doesNotMatch(formFn, /name="duration"/i, "Duration must not be a direct form input");
  assert.match(formFn, /Duration is calculated automatically/, "the form must explain Duration is derived, not editable");
}

function verifyPaceTimeOutConditionalInclusion() {
  const saveFn = app.match(/_savePaceVisitEdit\(visit, formEl\) \{[^]*?\n  \},/)[0];
  assert.match(saveFn, /timeOutChanged\s*=\s*newTimeOut\s*!==\s*\(visit\.timeOut/,
    "Time Out must be compared against the original value before deciding whether to send it");
  assert.match(saveFn, /if \(timeOutChanged\) displayFields\["Time Out"\] = newTimeOut;/,
    "Time Out must only be included in the payload when it actually changed");
  assert.match(saveFn, /PACE_ADMIN\.calculateDuration\(/,
    "Duration must be recalculated via the existing shared calculateDuration(), not a second algorithm");
}

function verifyStaleSavePaceVisitNotReused() {
  // The legacy GRAPH.savePaceVisit() (unrelated, unreachable write path)
  // must not have been copied into the new edit payload construction.
  const saveFn = app.match(/_savePaceVisitEdit\(visit, formEl\) \{[^]*?\n  \},/)[0];
  assert.doesNotMatch(saveFn, /"Behavior":/, "must not use the stale legacy \"Behavior\" display name");
  assert.doesNotMatch(saveFn, /"Interventions":/, "must not use the stale legacy \"Interventions\" display name");
}

function verifyV2WalkthroughUntouched() {
  assert.doesNotMatch(app.match(/async function saveWalkthroughEdit[^]*?\n}\n/)[0], /_V2/,
    "editing must never target IEP_Walkthrough_Observations_V2");
  // updateWalkthrough() itself (graph.js) is covered by Layer 1's URL assertions.
}

Promise.all([
  verifyWalkthroughEditPatchesNotPosts(),
  verifyWalkthroughEditNeverSendsClassroomWhenOmitted(),
  verifyWalkthroughEditNeverSendsProtectedFields(),
  verifyWalkthroughUpdateRequiresItemId(),
  verifyPaceEditPatchesNotPosts(),
  verifyPaceEditUsesModernFieldNamesNotStaleOnes(),
  verifyPaceEditStaleFieldNamesAreDroppedNotSent(),
  verifyPaceEditRequiresItemId(),
  verifyPaceEditOmittingTimeOutTouchesNothing()
])
  .then(() => {
    verifyWalkthroughEditGatedOnAdminPanel();
    verifyPaceEditGatedOnAdminPanel();
    verifyWalkthroughCancelPerformsNoWrite();
    verifyPaceCancelPerformsNoWrite();
    verifyWalkthroughProtectedFieldsNeverBuilt();
    verifyPaceDurationNotDirectlyEditable();
    verifyPaceTimeOutConditionalInclusion();
    verifyStaleSavePaceVisitNotReused();
    verifyV2WalkthroughUntouched();
    console.log("Admin editing (Walkthrough + PACE) tests passed.");
  })
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
