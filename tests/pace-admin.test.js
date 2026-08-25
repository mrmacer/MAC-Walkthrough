const assert = require("node:assert/strict");
const PACE_ADMIN = require("../pace-admin.js");

const completedRow = {
  id: "17",
  Student: "Sample Student",
  Date: "2026-08-24T00:00:00Z",
  "Time In": "09:05",
  "Time Out": "09:35",
  Duration: 30,
  Reason: "Needs a Break, Emotional Dysregulation",
  "Intervention Used": "Calm space",
  Notes: "Simulated test record"
};

const visit = PACE_ADMIN.normalizeVisit(completedRow);
assert.equal(visit.date, "2026-08-24");
assert.equal(visit.durationMinutes, 30);
assert.deepEqual(visit.reasons, ["Needs a Break", "Emotional Dysregulation"]);
assert.equal(visit.isCompleted, true);
assert.equal(visit.scmUsed, null, "missing SCM must stay unknown, not become false");

const calculated = PACE_ADMIN.normalizeVisit({
  Student: "Older Record", Date: "2026-08-23", "Time In": "10:00", "Time Out": "10:45",
  Behavior: "Peer conflict", Interventions: "Restorative conversation"
});
assert.equal(calculated.durationMinutes, 45, "duration should derive safely from valid same-day times");
assert.deepEqual(calculated.reasons, ["Peer conflict"]);
assert.deepEqual(calculated.supports, ["Restorative conversation"]);

const incomplete = PACE_ADMIN.normalizeVisit({ Student: "Older Record", Date: "2026-08-22", "Time In": "13:00" });
assert.equal(incomplete.isCompleted, false);
assert.equal(incomplete.durationMinutes, null);

const confirmedSchema = {
  Student: "Student", Date: "Date", "Time In": "Time_x0020_In", "Time Out": "Time_x0020_Out",
  Duration: "Duration", Reason: "Reason", "Intervention Used": "Intervention_x0020_Used",
  Notes: "Notes", "Staff Member": "StaffMember", "Return Status": "Return_x0020_Status"
};
const availability = PACE_ADMIN.getAvailability(confirmedSchema);
assert.equal(availability.duration, true);
assert.equal(availability.specialist, true);
assert.equal(availability.paceRoom, false);
assert.equal(availability.scm, false);
assert.equal(availability.teacherCameFrom, false);

const visits = [visit, calculated, incomplete];
const metrics = PACE_ADMIN.dashboardMetrics(visits, "2026-08-24");
assert.equal(metrics.visitsToday, 1);
assert.equal(metrics.studentsSeenToday, 1);
assert.equal(metrics.scmEvents, null, "SCM metric must be unavailable when the field is absent");
assert.equal(metrics.mostCommonReason[0], "Emotional Dysregulation");

const filtered = PACE_ADMIN.filterVisits(visits, { reason: "Peer conflict", duration: "45-plus" });
assert.equal(filtered.length, 1);
assert.equal(filtered[0].student, "Older Record");

const history = PACE_ADMIN.studentHistory(visits, "Older Record", "2026-08-24");
assert.equal(history.visits, 1, "incomplete historical rows remain visible but do not count as completed visits");
assert.equal(history.totalMinutes, 45);

// Provider contract: reads all pages + schema and never calls a write method.
let readCalls = 0;
let writeCalls = 0;
global.GRAPH = {
  async getListItemsByDisplayName(name, allPages) {
    readCalls++;
    assert.equal(name, "IEP_Pace_Visits");
    assert.equal(allPages, true);
    return [completedRow];
  },
  async getListSchema(name) {
    readCalls++;
    assert.equal(name, "IEP_Pace_Visits");
    return confirmedSchema;
  },
  async createListItem() { writeCalls++; },
  async updateMappedListItem() { writeCalls++; }
};

(async () => {
  PACE_ADMIN.invalidate();
  const loaded = await PACE_ADMIN.load(true);
  assert.equal(loaded.visits.length, 1);
  assert.equal(readCalls, 2);
  assert.equal(writeCalls, 0);
  console.log("PACE admin compatibility tests passed.");
})().catch(err => {
  console.error(err);
  process.exitCode = 1;
});
