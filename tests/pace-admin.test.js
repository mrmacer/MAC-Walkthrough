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

// ── PATCH: multiple Behavior Specialists, PACE Room identity, open-visit
//    history, and long/multi-paragraph Notes ─────────────────────────────

const multiSpecialistVisit = PACE_ADMIN.normalizeVisit({
  Student: "Izen Sosa", Date: "2026-08-27", "Time In": "11:21", "Time Out": "11:52",
  "PACE Room": "pace-room-1", "Behavior Specialist": "Kelly Marchetti, Sharon Morgan",
  "Teacher Came From": "Sickle", Reason: "Emotional Dysregulation, Needs a Break",
  "Intervention Used": "Sensory Break, Verbal Processing", "SCM Used": "No",
  Notes: "Line one of the note.\nLine two of the note.\n\nA new paragraph after a blank line."
});
assert.deepEqual(multiSpecialistVisit.specialists, ["Kelly Marchetti", "Sharon Morgan"],
  "comma-separated specialists must be parsed into a list, not kept as one string");
assert.equal(multiSpecialistVisit.teacherCameFrom, "Sickle");
assert.equal(multiSpecialistVisit.scmUsed, false);
assert.equal(multiSpecialistVisit.notes.includes("\n\n"), true,
  "multi-paragraph notes must round-trip with line breaks intact, not be flattened");
assert.equal(multiSpecialistVisit.notes.length > 60, true);

const semicolonSpecialistVisit = PACE_ADMIN.normalizeVisit({
  Student: "Semicolon Test", Date: "2026-08-27", "Behavior Specialist": "Kelly Marchetti; Sharon Morgan"
});
assert.deepEqual(semicolonSpecialistVisit.specialists, ["Kelly Marchetti", "Sharon Morgan"],
  "semicolon-separated specialists must also be parsed into a list");

const singleSpecialistVisit = PACE_ADMIN.normalizeVisit({
  Student: "Older Single-Specialist Record", Date: "2026-08-01", "Staff Member": "Nikki Stock"
});
assert.deepEqual(singleSpecialistVisit.specialists, ["Nikki Stock"],
  "an older plain single-name record must still yield a one-element list, not break");

assert.deepEqual(PACE_ADMIN.roomInfo("pace-room-1"), { label: "PACE Room 1", hallway: "Yellow Hall", raw: "pace-room-1" });
assert.deepEqual(PACE_ADMIN.roomInfo("pace-room-2"), { label: "PACE Room 2", hallway: "Green Hall", raw: "pace-room-2" });
assert.equal(PACE_ADMIN.roomInfo(""), null, "a blank room value must not be rendered as a fabricated room");
assert.equal(PACE_ADMIN.roomInfo("PACE Room 3").label, "PACE Room 3",
  "an unrecognized future room value must display as-is rather than guessing a hallway");

// A currently-open visit (today, no Time Out) must surface separately from
// completed history, and must not corrupt average/total duration metrics.
const openToday = PACE_ADMIN.normalizeVisit({ Student: "Open Case", Date: "2026-08-24", "Time In": "13:10" });
const oldDangling = PACE_ADMIN.normalizeVisit({ Student: "Open Case", Date: "2026-07-01", "Time In": "09:00" });
const openCompleted = PACE_ADMIN.normalizeVisit({
  Student: "Open Case", Date: "2026-08-20", "Time In": "09:00", "Time Out": "09:20", Duration: 20
});
const openCaseHistory = PACE_ADMIN.studentHistory([openToday, oldDangling, openCompleted], "Open Case", "2026-08-24");
assert.equal(openCaseHistory.openVisits.length, 2, "both the today-open and old-dangling visits are surfaced as open");
assert.equal(openCaseHistory.recentVisits.length, 1, "only the completed visit appears in recentVisits");
assert.equal(openCaseHistory.averageDuration, 20, "the open visits' missing duration must not affect the average");
assert.equal(openToday.scmUsed, null, "an open visit legitimately has no SCM answer yet");

// ── PATCH: PACE Room + SCM surfaced throughout admin reporting ───────────

// Room normalization end-to-end: normalizeVisit()'s raw paceRoom value
// chained through roomInfo(), not a literal string re-typed in the test.
assert.equal(multiSpecialistVisit.paceRoom, "pace-room-1");
const room1 = PACE_ADMIN.roomInfo(multiSpecialistVisit.paceRoom);
assert.equal(room1.label, "PACE Room 1");
assert.equal(room1.hallway, "Yellow Hall");

const room2Visit = PACE_ADMIN.normalizeVisit({ Student: "Room 2 Test", Date: "2026-08-27", "PACE Room": "pace-room-2" });
const room2 = PACE_ADMIN.roomInfo(room2Visit.paceRoom);
assert.equal(room2.label, "PACE Room 2");
assert.equal(room2.hallway, "Green Hall");

const missingRoomVisit = PACE_ADMIN.normalizeVisit({ Student: "No Room Recorded", Date: "2026-08-27" });
assert.equal(missingRoomVisit.paceRoom, "", "a record with no PACE Room column value must not fabricate one");
assert.equal(PACE_ADMIN.roomInfo(missingRoomVisit.paceRoom), null);

const scmYesVisit = PACE_ADMIN.normalizeVisit({ Student: "SCM Yes Test", Date: "2026-08-27", "SCM Used": "Yes" });
assert.equal(scmYesVisit.scmUsed, true);
const scmNoVisit = PACE_ADMIN.normalizeVisit({ Student: "SCM No Test", Date: "2026-08-27", "SCM Used": "No" });
assert.equal(scmNoVisit.scmUsed, false);
const scmMissingVisit = PACE_ADMIN.normalizeVisit({ Student: "SCM Missing Test", Date: "2026-08-27" });
assert.equal(scmMissingVisit.scmUsed, null, "no SCM Used value must stay unknown, never become false");

// Room / SCM filters
const roomFiltered = PACE_ADMIN.filterVisits([multiSpecialistVisit, room2Visit], { paceRoom: "pace-room-1" });
assert.deepEqual(roomFiltered.map(v => v.student), ["Izen Sosa"], "room filter must return only matching-room visits");

const scmFilterVisits = [scmYesVisit, scmNoVisit, scmMissingVisit];
assert.deepEqual(PACE_ADMIN.filterVisits(scmFilterVisits, { scm: "yes" }).map(v => v.student), ["SCM Yes Test"]);
assert.deepEqual(PACE_ADMIN.filterVisits(scmFilterVisits, { scm: "no" }).map(v => v.student), ["SCM No Test"]);
assert.equal(PACE_ADMIN.filterVisits(scmFilterVisits, {}).length, 3, "no SCM filter applied returns every visit, missing included");

// SCM Events metric: only explicit Yes counts, blank/null never counted as No
const scmMetricVisits = [
  PACE_ADMIN.normalizeVisit({ Student: "A", Date: "2026-08-24", "Time In": "09:00", "Time Out": "09:20", "SCM Used": "Yes" }),
  PACE_ADMIN.normalizeVisit({ Student: "B", Date: "2026-08-24", "Time In": "10:00", "Time Out": "10:20", "SCM Used": "Yes" }),
  PACE_ADMIN.normalizeVisit({ Student: "C", Date: "2026-08-24", "Time In": "11:00", "Time Out": "11:20", "SCM Used": "No" }),
  PACE_ADMIN.normalizeVisit({ Student: "D", Date: "2026-08-24", "Time In": "12:00", "Time Out": "12:20" }), // completed, SCM never recorded
  PACE_ADMIN.normalizeVisit({ Student: "E", Date: "2026-08-24", "Time In": "13:00" }) // open, SCM blank until completion
];
const scmSummary = PACE_ADMIN.summarize(scmMetricVisits);
assert.equal(scmSummary.scmEvents, 2, "SCM Events must count only the 2 explicit Yes visits, not the blank/open ones");

const scmHistory = PACE_ADMIN.studentHistory(
  [{ ...scmMetricVisits[0], student: "SCM Metric Student" }, { ...scmMetricVisits[2], student: "SCM Metric Student" }],
  "SCM Metric Student", "2026-08-24"
);
assert.equal(scmHistory.scmEvents, 1, "student-history SCM Events must also count only explicit Yes");

// Full student-history detail: Room + SCM appear alongside every other field
const fullDetailVisit = PACE_ADMIN.normalizeVisit({
  Student: "Full Detail Student", Date: "2026-08-27", "Time In": "11:21", "Time Out": "11:52",
  "PACE Room": "pace-room-1", "Behavior Specialist": "Kelly Marchetti, Sharon Morgan",
  "Teacher Came From": "Sickle", Reason: "Emotional Dysregulation, Needs a Break",
  "Intervention Used": "Sensory Break, Verbal Processing", "SCM Used": "No",
  Notes: "Full detail check."
});
assert.equal(fullDetailVisit.paceRoom, "pace-room-1");
assert.equal(fullDetailVisit.scmUsed, false);
assert.equal(fullDetailVisit.teacherCameFrom, "Sickle");
assert.deepEqual(fullDetailVisit.specialists, ["Kelly Marchetti", "Sharon Morgan"]);
assert.deepEqual(fullDetailVisit.reasons, ["Emotional Dysregulation", "Needs a Break"]);
assert.deepEqual(fullDetailVisit.supports, ["Sensory Break", "Verbal Processing"]);
assert.equal(fullDetailVisit.notes, "Full detail check.");

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
