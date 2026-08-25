/* PACE administrator read model.
 *
 * IEP_Pace_Visits remains the source of truth. This module converts evolving
 * SharePoint display-name fields into one defensive shape used by the
 * dashboard, explorer, and student history. It never writes records.
 */

(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.PACE_ADMIN = api;
})(typeof globalThis !== "undefined" ? globalThis : window, function () {
  const LIST_NAME = "IEP_Pace_Visits";
  const CACHE_MS  = 90_000;

  const FIELD_ALIASES = {
    student:         ["Student", "Student Name", "StudentName", "studentName"],
    date:            ["Date", "Visit Date", "date"],
    timeIn:          ["Time In", "TimeIn", "timeIn"],
    timeOut:         ["Time Out", "TimeOut", "timeOut"],
    duration:        ["Duration", "Duration Minutes", "DurationMinutes", "durationMinutes"],
    paceRoom:        ["PACE Room", "Pace Room", "paceRoom"],
    specialist:      ["Behavior Specialist", "Staff Member", "Submitted By", "submittedByName"],
    teacherCameFrom: ["Teacher Came From", "Teacher", "teacherCameFrom"],
    reason:          ["Reason", "Behavior", "Behaviors", "behaviors"],
    support:         ["Intervention Used", "Interventions", "Support", "Supports", "interventions"],
    scm:             ["SCM Used", "SCM", "scmUsed"],
    notes:           ["Notes", "Visit Notes", "notes"],
    submittedAt:     ["Submitted At", "Created", "createdAt", "timestamp"],
    returnStatus:    ["Return Status", "returnStatus"]
  };

  let _cache = null;
  let _cacheAt = 0;

  function firstValue(row, aliases) {
    for (const key of aliases) {
      if (!Object.prototype.hasOwnProperty.call(row || {}, key)) continue;
      const value = row[key];
      if (value !== undefined && value !== null && value !== "") return value;
    }
    return null;
  }

  function hasField(schema, aliases) {
    return aliases.some(name => Object.prototype.hasOwnProperty.call(schema || {}, name));
  }

  function text(value) {
    if (value === undefined || value === null) return "";
    if (typeof value === "object") {
      return String(value.displayName || value.DisplayName || value.LookupValue || value.title || value.name || "").trim();
    }
    return String(value).trim();
  }

  function list(value) {
    if (Array.isArray(value)) return value.map(text).filter(Boolean);
    return String(value || "").split(/[,;]\s*/).map(v => v.trim()).filter(Boolean);
  }

  function booleanOrNull(value) {
    if (value === undefined || value === null || value === "") return null;
    if (typeof value === "boolean") return value;
    if (typeof value === "number") return value === 1;
    const normalized = String(value).trim().toLowerCase();
    if (["yes", "true", "1"].includes(normalized)) return true;
    if (["no", "false", "0"].includes(normalized)) return false;
    return null;
  }

  function dateOnly(value) {
    const raw = text(value);
    const match = raw.match(/^(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : "";
  }

  function calculateDuration(timeIn, timeOut) {
    const match = value => String(value || "").match(/^(\d{1,2}):(\d{2})/);
    const start = match(timeIn);
    const end   = match(timeOut);
    if (!start || !end) return null;
    const minutes = (Number(end[1]) * 60 + Number(end[2])) - (Number(start[1]) * 60 + Number(start[2]));
    // The completed-visit tracker requires same-day visits. Do not infer an
    // overnight duration from malformed or legacy records.
    return Number.isFinite(minutes) && minutes >= 0 ? minutes : null;
  }

  function duration(value, timeIn, timeOut) {
    if (value !== undefined && value !== null && value !== "") {
      const n = Number(value);
      if (Number.isFinite(n) && n >= 0) return Math.round(n);
    }
    return calculateDuration(timeIn, timeOut);
  }

  function normalizeVisit(row) {
    const timeIn  = text(firstValue(row, FIELD_ALIASES.timeIn));
    const timeOut = text(firstValue(row, FIELD_ALIASES.timeOut));
    const visitDuration = duration(firstValue(row, FIELD_ALIASES.duration), timeIn, timeOut);
    return {
      id:              text(row?.id || row?.ID || ""),
      student:         text(firstValue(row, FIELD_ALIASES.student)),
      date:            dateOnly(firstValue(row, FIELD_ALIASES.date)),
      timeIn,
      timeOut,
      durationMinutes: visitDuration,
      paceRoom:        text(firstValue(row, FIELD_ALIASES.paceRoom)),
      specialist:      text(firstValue(row, FIELD_ALIASES.specialist)),
      teacherCameFrom: text(firstValue(row, FIELD_ALIASES.teacherCameFrom)),
      reasons:         list(firstValue(row, FIELD_ALIASES.reason)),
      supports:        list(firstValue(row, FIELD_ALIASES.support)),
      scmUsed:         booleanOrNull(firstValue(row, FIELD_ALIASES.scm)),
      notes:           text(firstValue(row, FIELD_ALIASES.notes)),
      submittedAt:     text(firstValue(row, FIELD_ALIASES.submittedAt)),
      returnStatus:    text(firstValue(row, FIELD_ALIASES.returnStatus)),
      isCompleted:     Boolean(timeOut)
    };
  }

  function getAvailability(schema) {
    return {
      duration:        hasField(schema, FIELD_ALIASES.duration),
      paceRoom:        hasField(schema, FIELD_ALIASES.paceRoom),
      specialist:      hasField(schema, FIELD_ALIASES.specialist),
      teacherCameFrom: hasField(schema, ["Teacher Came From"]),
      reason:          hasField(schema, FIELD_ALIASES.reason),
      support:         hasField(schema, FIELD_ALIASES.support),
      scm:             hasField(schema, FIELD_ALIASES.scm),
      notes:           hasField(schema, FIELD_ALIASES.notes),
      submittedAt:     hasField(schema, FIELD_ALIASES.submittedAt),
      returnStatus:    hasField(schema, FIELD_ALIASES.returnStatus)
    };
  }

  function inRange(visit, from, to) {
    if (from && (!visit.date || visit.date < from)) return false;
    if (to && (!visit.date || visit.date > to)) return false;
    return true;
  }

  function filterVisits(visits, filters) {
    const f = filters || {};
    const search = String(f.search || "").trim().toLowerCase();
    return (visits || []).filter(visit => {
      if (!inRange(visit, f.from, f.to)) return false;
      if (f.student && visit.student !== f.student) return false;
      if (f.specialist && visit.specialist !== f.specialist) return false;
      if (f.teacherCameFrom && visit.teacherCameFrom !== f.teacherCameFrom) return false;
      if (f.paceRoom && visit.paceRoom !== f.paceRoom) return false;
      if (f.reason && !visit.reasons.includes(f.reason)) return false;
      if (f.scm === "yes" && visit.scmUsed !== true) return false;
      if (f.scm === "no" && visit.scmUsed !== false) return false;
      if (f.duration === "45-plus" && !(visit.durationMinutes >= 45)) return false;
      if (f.duration === "under-45" && !(visit.durationMinutes !== null && visit.durationMinutes < 45)) return false;
      if (search) {
        const haystack = [visit.student, visit.specialist, visit.teacherCameFrom, visit.paceRoom,
          ...visit.reasons, ...visit.supports, visit.notes].join(" ").toLowerCase();
        if (!haystack.includes(search)) return false;
      }
      return true;
    });
  }

  function counts(values) {
    const result = {};
    values.filter(Boolean).forEach(value => { result[value] = (result[value] || 0) + 1; });
    return result;
  }

  function topValue(values) {
    return Object.entries(counts(values)).sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0] || null;
  }

  function addDays(dateString, amount) {
    const d = new Date(`${dateString}T12:00:00`);
    if (Number.isNaN(d.getTime())) return "";
    d.setDate(d.getDate() + amount);
    return d.toISOString().slice(0, 10);
  }

  function summarize(visits, options) {
    const opts = options || {};
    const selected  = (visits || []).filter(v => inRange(v, opts.from, opts.to));
    const completed = selected.filter(v => v.isCompleted);
    const withDuration = completed.filter(v => v.durationMinutes !== null);
    const withScm      = completed.filter(v => v.scmUsed !== null);
    const reasonValues = completed.flatMap(v => v.reasons);
    return {
      visits:           completed.length,
      uniqueStudents:   new Set(completed.map(v => v.student).filter(Boolean)).size,
      totalMinutes:     withDuration.length ? withDuration.reduce((sum, v) => sum + v.durationMinutes, 0) : null,
      averageDuration:  withDuration.length ? Math.round(withDuration.reduce((sum, v) => sum + v.durationMinutes, 0) / withDuration.length) : null,
      longestDuration:  withDuration.length ? Math.max(...withDuration.map(v => v.durationMinutes)) : null,
      scmEvents:        withScm.length ? withScm.filter(v => v.scmUsed).length : null,
      needsBreak:       reasonValues.length ? reasonValues.filter(r => r.trim().toLowerCase() === "needs a break").length : null,
      mostCommonReason: topValue(reasonValues),
      mostCommonTeacher: topValue(completed.map(v => v.teacherCameFrom).filter(Boolean)),
      completed,
      records: selected
    };
  }

  function detectDashboardFlags(visits, today) {
    const from = addDays(today, -29);
    const recent = (visits || []).filter(v => v.isCompleted && inRange(v, from, today));
    const perStudentDay = {};
    recent.forEach(v => {
      if (!v.student || !v.date) return;
      const key = `${v.date}\u0000${v.student}`;
      perStudentDay[key] = (perStudentDay[key] || 0) + 1;
    });
    return {
      extendedVisits: recent.filter(v => v.durationMinutes !== null && v.durationMinutes >= 45).length,
      repeatVisitDays: Object.values(perStudentDay).filter(n => n >= 3).length,
      scmEvents: recent.some(v => v.scmUsed !== null) ? recent.filter(v => v.scmUsed).length : null
    };
  }

  function dashboardMetrics(visits, today) {
    const periodFrom = addDays(today, -29);
    const todaySummary = summarize(visits, { from: today, to: today });
    const period = summarize(visits, { from: periodFrom, to: today });
    return {
      visitsToday:       todaySummary.visits,
      studentsSeenToday: todaySummary.uniqueStudents,
      averageDuration:   period.averageDuration,
      scmEvents:         period.scmEvents,
      mostCommonReason:  period.mostCommonReason,
      longestVisit:      period.longestDuration,
      recentVisits:      [...period.completed].sort(sortNewest).slice(0, 8),
      flags:             detectDashboardFlags(visits, today),
      periodFrom,
      periodTo: today
    };
  }

  function sortNewest(a, b) {
    const key = v => `${v.date || ""}T${v.timeIn || "00:00"}|${v.submittedAt || ""}`;
    return key(b).localeCompare(key(a));
  }

  function studentHistory(visits, student, today) {
    const from = addDays(today, -29);
    const matching = (visits || []).filter(v => v.student === student);
    const summary = summarize(matching, { from, to: today });
    return {
      student,
      from,
      to: today,
      ...summary,
      recentVisits: [...matching].filter(v => v.isCompleted).sort(sortNewest).slice(0, 10)
    };
  }

  async function load(force) {
    const now = Date.now();
    if (!force && _cache && now - _cacheAt < CACHE_MS) return _cache;
    if (typeof GRAPH === "undefined") throw new Error("Microsoft Graph data provider is unavailable.");
    // getListItemsByDisplayName already resolves/caches the schema. Read it
    // afterward so the normal Graph cache avoids a duplicate columns request.
    const rows = await GRAPH.getListItemsByDisplayName(LIST_NAME, true);
    const schema = await GRAPH.getListSchema(LIST_NAME);
    _cache = {
      visits: rows.map(normalizeVisit),
      schema,
      availability: getAvailability(schema)
    };
    _cacheAt = now;
    return _cache;
  }

  function invalidate() {
    _cache = null;
    _cacheAt = 0;
  }

  return {
    LIST_NAME,
    FIELD_ALIASES,
    calculateDuration,
    normalizeVisit,
    getAvailability,
    filterVisits,
    summarize,
    dashboardMetrics,
    detectDashboardFlags,
    studentHistory,
    sortNewest,
    load,
    invalidate
  };
});
