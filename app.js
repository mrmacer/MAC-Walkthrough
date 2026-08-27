/* ── Utilities ───────────────────────────────────────────────────────────────── */

function escHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fmtDate(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

function fmtDateShort(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric" });
}

function fmtDateTime(iso) {
  if (!iso) return "—";
  const d = new Date(iso);
  return d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" }) +
    " " + d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
}

// Formats a stored "HH:MM" (24-hour) time string as a 12-hour clock label,
// e.g. "11:21 AM". Returns "" for anything that doesn't parse.
function fmtClockTime(hhmm) {
  const m = String(hhmm || "").match(/^(\d{1,2}):(\d{2})/);
  if (!m) return "";
  const d = new Date();
  d.setHours(Number(m[1]), Number(m[2]), 0, 0);
  return d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
}

function getWeekOf(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  if (isNaN(d)) return "";
  const day = d.getDay();
  const monday = new Date(d);
  monday.setDate(d.getDate() - (day === 0 ? 6 : day - 1));
  return monday.toISOString().slice(0, 10);
}

function showToast(msg, type = "success") {
  const existing = document.querySelector(".toast");
  if (existing) existing.remove();
  const el = document.createElement("div");
  el.className = "toast toast-" + type;
  el.innerHTML = (type === "success" ? "✓ " : "⚠ ") + escHtml(msg);
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 3200);
}

function showSyncError(message) {
  const existing = document.getElementById("syncErrorBox");
  const box = existing || document.createElement("div");
  box.id        = "syncErrorBox";
  box.className = "sync-error-box";
  box.innerHTML = `
    <div class="sync-error-title">⚠ SharePoint sync failed — data saved locally</div>
    <pre class="sync-error-body">${escHtml(String(message || "Unknown error"))}</pre>
    <button class="sync-error-dismiss" onclick="this.closest('#syncErrorBox').remove()">Dismiss</button>`;
  if (!existing) {
    const anchor =
      document.querySelector("#pulseConfirm") ||
      document.querySelector(".content-area") ||
      document.body;
    anchor.parentNode.insertBefore(box, anchor.nextSibling);
  }
  box.scrollIntoView({ behavior: "smooth", block: "start" });
}

function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

function statusBadgeHtml(status) {
  const emoji = CONFIG.STATUS_EMOJI[status] || "";
  const label = CONFIG.CLASSROOM_STATUS_OPTIONS.find(o => o.value === status)?.label || status;
  const cls   = { "on-track":"badge-green","monitor":"badge-yellow","needs-support":"badge-orange","immediate-follow-up":"badge-red" }[status] || "badge-slate";
  return `<span class="badge ${cls}">${emoji} ${escHtml(label)}</span>`;
}

function supportNeededBadgeHtml(val) {
  const label = CONFIG.SUPPORT_NEEDED_OPTIONS.find(o => o.value === val)?.label || val;
  const cls   = { "none":"badge-slate","talk-at-weekly":"badge-blue","help-soon":"badge-orange","urgent":"badge-red" }[val] || "badge-slate";
  return `<span class="badge ${cls}">${escHtml(label)}</span>`;
}

/* ── Pilot Roster ─────────────────────────────────────────────────────────── */
// LEGACY: Admin/prototype roster dependency (New Walkthrough, V2 Walkthrough,
// Teacher Dashboard, Student Check-In). Daily Pulse and PACE no longer use
// this source — see STUDENT_ROSTER below, which loads live from SharePoint.
// Migrate the remaining consumers listed above to SharePoint in a future
// privacy-hardening patch. Do not add any more student names here.

const DATA_VERSION     = "pilot-roster-v4-student-checkin";
const DATA_VERSION_KEY = "iepSkookDataVersion";

const PILOT_ADMIN = {
  id:   "admin-decusky",
  name: "Decusky",
  role: "administrator"
};

const PILOT_TEACHERS = [
  {
    id:   "teacher-stock",
    name: "Stock",
    role: "teacher",
    students: [
      { id: "antonio-h",   name: "Antonio H."   },
      { id: "david-m",     name: "David M."     },
      { id: "bryanna-c",   name: "Bryanna C."   },
      { id: "xavier-o",    name: "Xavier O."    },
      { id: "siah-l",      name: "Siah L."      },
      { id: "cornelius-d", name: "Cornelius D." }
    ]
  },
  {
    id:   "teacher-bossons",
    name: "Bossons",
    role: "teacher",
    students: [
      { id: "christian-e",  name: "Christian E."  },
      { id: "william-m",    name: "William M."    },
      { id: "ruby-m",       name: "Ruby M."       },
      { id: "leonay-r",     name: "Leonay R."     },
      { id: "yosmauri-s",   name: "Yosmauri S."   },
      { id: "maximillan-s", name: "Maximillan S." }
    ]
  },
  {
    id:   "teacher-andruchek",
    name: "Andruchek",
    role: "teacher",
    students: [
      { id: "savior-l", name: "Savior L." },
      { id: "heaven-g", name: "Heaven G." },
      { id: "briant-u", name: "Briant U." },
      { id: "ariel-l",  name: "Ariel L."  },
      { id: "chase-w",  name: "Chase W."  },
      { id: "dekari-s", name: "Dekari S." }
    ]
  },
  {
    id:   "teacher-hughes",
    name: "Hughes",
    role: "teacher",
    students: [
      { id: "jayden-c",    name: "Jayden C."    },
      { id: "jeremy-h",    name: "Jeremy H."    },
      { id: "thomas-c",    name: "Thomas C."    },
      { id: "markenna-h",  name: "Markenna H."  },
      { id: "landyn-k",    name: "Landyn K."    },
      { id: "austin-w",    name: "Austin W."    },
      { id: "andrew-b",    name: "Andrew B."    },
      { id: "emma-f",      name: "Emma F."      }
    ]
  },
  {
    id:   "teacher-kenny",
    name: "Kenny",
    role: "teacher",
    students: [
      { id: "miya-m",      name: "Miya M."      },
      { id: "ayden-k",     name: "Ayden K."     },
      { id: "julianna-t",  name: "Julianna T."  },
      { id: "alhaji-f",    name: "Alhaji F."    },
      { id: "layla-g",     name: "Layla G."     }
    ]
  },
  {
    id:       "teacher-mamrosh",
    name:     "Mamrosh",
    role:     "teacher",
    students: []
  }
];

const PILOT_STUDENTS = PILOT_TEACHERS.flatMap(teacher =>
  teacher.students.map(student => ({
    ...student,
    teacherId:   teacher.id,
    teacherName: teacher.name
  }))
);

/* ── Student Roster Provider (2026–27) ───────────────────────────────────────── */
// Shared, live SharePoint-backed roster for Daily Pulse + PACE. See
// CONFIG.STUDENT_ROSTER_LIST in config.js for the single place the real
// 2026–27 list name gets inserted. Do not add student data to source files —
// this provider only ever knows HOW to load the roster, never WHO is on it
// until it fetches from SharePoint at runtime.

const ROSTER_PLACEHOLDER = "REPLACE_WITH_2026_27_STUDENT_LIST";

function slugifyTeacherId(name) {
  const slug = String(name || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  return slug ? `teacher-${slug}` : "";
}

// Normalizes a SharePoint Yes/No column to a boolean. Graph may already
// return an actual boolean for some column types — trust that directly
// rather than string-comparing it. `defaultValue` applies only when the
// column is absent/blank (e.g. the legacy fallback list, which has no
// PACE Enabled / Daily Pulse Enabled columns at all).
function normalizeRosterBoolean(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value === 1;
  const s = String(value).trim().toLowerCase();
  return s === "yes" || s === "true" || s === "1";
}

const STUDENT_ROSTER = {
  _students:      [],
  loading:        false,
  loaded:         false,
  error:          null,
  _usingFallback: false,

  _listName() {
    const configured = String(CONFIG.STUDENT_ROSTER_LIST || "").trim();
    if (configured && configured !== ROSTER_PLACEHOLDER) {
      this._usingFallback = false;
      return configured;
    }
    // CONFIG.STUDENT_ROSTER_LIST has been reset to the placeholder. Fall
    // back to the old pilot roster list so Daily Pulse/PACE don't hard-fail
    // during development. Under normal production configuration (a real
    // list name in CONFIG.STUDENT_ROSTER_LIST) this branch never runs.
    this._usingFallback = true;
    return SETUP_LISTS.students;
  },

  // Production schema (IEP_Students_2026_27) has separate first/last name
  // columns and no "Student ID" column — see the schema comment in
  // config.js. The legacyName fallback only matters for the old pilot list.
  _normalize(row) {
    const firstName = String(row["Student First Name"] || "").trim();
    const lastName  = String(row["Student Last Name"]  || "").trim();
    const legacyName = String(row["Student Name"] || row["StudentName"] || row["Name"] || "").trim();
    const name = (firstName || lastName) ? `${firstName} ${lastName}`.trim() : legacyName;
    if (!name) return null;

    const teacherName = String(row["Teacher"] || row["TeacherName"] || "").trim();

    return {
      // The SharePoint list item's own stable item id (surfaced as `row.id`
      // below, from GRAPH.getListItems) — never derived from the student's
      // name, which can change or collide across students.
      id:                row.id != null ? String(row.id) : name,
      firstName,
      lastName,
      name,
      teacherName,
      teacherId:         slugifyTeacherId(teacherName),
      classroom:         String(row["Classroom"] || "").trim(),
      // Fail-safe: a blank/missing PACE Enabled or Daily Pulse Enabled cell
      // defaults to NOT eligible, unlike Active (which keeps its existing
      // default-true behavior for compatibility).
      active:            normalizeRosterBoolean(row["Active"],             true),
      paceEnabled:       normalizeRosterBoolean(row["PACE Enabled"],       false),
      dailyPulseEnabled: normalizeRosterBoolean(row["Daily Pulse Enabled"], false)
    };
  },

  async refresh() {
    this.loading = true;
    this.error   = null;
    const listName = this._listName();
    if (this._usingFallback) {
      // Generic, non-identifying notice only — never log roster contents.
      console.warn(`STUDENT_ROSTER: CONFIG.STUDENT_ROSTER_LIST is still the placeholder — using fallback list "${listName}".`);
    }
    try {
      const [raw, schema] = await Promise.all([
        GRAPH.getListItems(listName),
        GRAPH.getListSchema(listName)
      ]);
      const inv = {};
      Object.entries(schema).forEach(([display, internal]) => { inv[internal] = display; });
      this._students = raw
        .map(item => {
          const row = {};
          Object.entries(item).forEach(([k, v]) => { row[inv[k] || k] = v; });
          row.id = item.id; // stable SharePoint list item id — not subject to display-name guessing
          return this._normalize(row);
        })
        .filter(Boolean);
      this.loaded = true;
      // Generic diagnostic only — a count, never the records themselves.
      console.log(`Student roster loaded: ${this._students.length} record(s).`);
    } catch (err) {
      console.error("Student roster load failed:", err.message || String(err));
      this.error = err.message || "Unable to load the student roster.";
    } finally {
      this.loading = false;
    }
    return this._students;
  },

  getAll()    { return this._students; },
  getActive() { return this._students.filter(s => s.active); },
  find(id)    { return this._students.find(s => s.id === id) || null; },
  findByName(name) { return this._students.find(s => s.name === name) || null; },

  // Shared teacher-name match used by every teacher-scoped getter below —
  // case/whitespace-insensitive, but never guesses: a record with no
  // Teacher value never matches any teacher (it can still surface through
  // the unscoped admin-facing getters).
  _matchesTeacher(student, teacherName) {
    const target = String(teacherName || "").trim().toLowerCase();
    if (!target || !student.teacherName) return false;
    return student.teacherName.trim().toLowerCase() === target;
  },

  // General active roster for a given teacher (matched by teacher NAME —
  // the only linkage the SharePoint roster offers).
  getForTeacher(teacherName) {
    return this.getActive().filter(s => this._matchesTeacher(s, teacherName));
  },

  // PACE: active + PACE-enabled. Intentionally NOT restricted by teacher —
  // PACE is a shared intervention workflow, not a per-classroom one; a
  // teacher may need to log a visit for a student outside their own
  // classroom. Same result for admin and teacher.
  getPaceEnabled() {
    return this.getActive().filter(s => s.paceEnabled);
  },

  // Daily Pulse (administrator view): active + Daily-Pulse-enabled, no
  // teacher restriction — preserves the existing admin UX of seeing every
  // eligible student at once.
  getDailyPulseEnabled() {
    return this.getActive().filter(s => s.dailyPulseEnabled);
  },

  // Daily Pulse (teacher view): active + Daily-Pulse-enabled + assigned to
  // this specific teacher.
  getDailyPulseForTeacher(teacherName) {
    return this.getDailyPulseEnabled().filter(s => this._matchesTeacher(s, teacherName));
  }
};

/* ── Teacher Context ─────────────────────────────────────────────────────────── */

const TEACHER_IDENTITIES = {
  "teacher-stock": {
    userId:      "teacher-stock",
    teacherName: "Stock",
    displayName: "Nikki Stock",
    email:       "stocn@iu29.org",
    classroom:   "Stock's Room"
  },
  "teacher-bossons": {
    userId:      "teacher-bossons",
    teacherName: "Bossons",
    displayName: "Amber Bossons",
    email:       "bossa@iu29.org",
    classroom:   "Bossons' Room"
  }
};

const AVAILABLE_TEACHERS = PILOT_TEACHERS
  .filter(t => TEACHER_IDENTITIES[t.id])
  .map(t => TEACHER_IDENTITIES[t.id]);

const TEACHER_ALLOWED_ROUTES = new Set(["pulse"]);

const USER_CONTEXT = {
  actualUser:    null,
  viewedTeacher: null,

  get isViewingAsTeacher() { return AUTH.isAdmin && this.viewedTeacher !== null; },
  get effectiveRole() {
    return this.isViewingAsTeacher ? "teacher" : normalizeUserRole(AUTH.role);
  },
  get effectiveTeacher() {
    if (this.isViewingAsTeacher) return this.viewedTeacher;
    if (AUTH.isTeacher) {
      return TEACHER_IDENTITIES[AUTH.pilotUserId] || {
        userId:      AUTH.pilotUserId,
        teacherName: AUTH.pilotUser?.Name || "",
        displayName: AUTH.pilotUser?.Name || AUTH.displayName || "",
        email:       AUTH.account?.username || "",
        classroom:   ""
      };
    }
    return null;
  }
};

function canAccessRoute(route) {
  if (USER_CONTEXT.effectiveRole === "teacher") return TEACHER_ALLOWED_ROUTES.has(route);
  return true;
}

function recordBelongsToTeacher(record, teacher, roster) {
  const teacherName   = String(teacher?.teacherName || "").trim().toLowerCase();
  const recordTeacher = String(record?.teacher || record?.teacherName || "").trim().toLowerCase();
  if (recordTeacher && recordTeacher === teacherName) return true;
  const rosterNames   = new Set(roster.map(s => String(s.name || s.studentName || "").trim().toLowerCase()));
  const recordStudent = String(record?.student || "").trim().toLowerCase();
  return recordStudent ? rosterNames.has(recordStudent) : false;
}

/* ── Setup / SharePoint Admin ────────────────────────────────────────────────── */

const SETUP_LISTS = {
  users:      "IEP_Users2",
  teachers:   "macwalkthroughwhoareyouvisiting",
  students:   "IEP_Skook_Pilot_Students",
  classrooms: "IEP_School_Settings"
};

const APP_CACHE = {
  _data: {},
  set(key, value) { this._data[key] = value; },
  get(key)        { return this._data[key]; },
  clear(key)      { if (key) { delete this._data[key]; } else { this._data = {}; } }
};

function requireSetupAdmin() {
  if (!AUTH.isAdmin) throw new Error("Admin access required.");
}

function normalizeSetupUser(item) {
  return {
    spId:    item.id || item.spId,
    userId:  item.Title || item.UserID || item.User_x0020_ID || "",
    name:    item.field_1 || item.Name || item.Title || "",
    role:    item.field_2 || item.Role || "",
    email:   item.Email || item.email || item.EMail || "",
    active:  item.field_3 === true || item.field_3 === "Yes" || item.Active === true || item.Active === "Yes",
    raw:     item
  };
}

function normalizeSetupTeacher(item) {
  return {
    spId:      item.id || item.spId,
    name:      (item.fields?.teacher || item.teacher || item.Title || "").trim(),
    hasAccess: PILOT_TEACHERS.some(p => p.name.toLowerCase() === (item.fields?.teacher || item.teacher || "").trim().toLowerCase()),
    raw:       item
  };
}

function normalizeSetupStudent(item) {
  return {
    spId:      item.id || item.spId,
    pilotId:   item.PilotID || item.Pilot_x0020_ID || item.Title || "",
    name:      item.StudentName || item.Student_x0020_Name || item.Name || "",
    teacher:   item.Teacher || item.TeacherName || "",
    classroom: item.Classroom || "",
    active:    item.Active === true || item.Active === "Yes" || item.Active === undefined,
    raw:       item
  };
}

function normalizeSetupClassroom(item) {
  return {
    spId:    item.id || item.spId,
    name:    item.Title || item.ClassroomName || item.Classroom || "",
    teacher: item.Teacher || item.TeacherName || "",
    school:  item.School || "",
    raw:     item
  };
}

function createSetupUserId(role, name) {
  const slug = String(name || "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "");
  const prefix = String(role || "user").toLowerCase().split(" ")[0];
  return `${prefix}-${slug}-${Date.now().toString(36)}`;
}

function getNextPilotId(students) {
  const nums = students
    .map(s => parseInt(String(s.pilotId || "").replace(/^P0*/, ""), 10))
    .filter(n => !isNaN(n));
  const next = nums.length ? Math.max(...nums) + 1 : 1;
  return "P" + String(next).padStart(3, "0");
}

const SETUP_DATA = {
  users:      [],
  teachers:   [],
  students:   [],
  classrooms: [],
  loading:    false,
  error:      null,

  async refresh() {
    this.loading = true;
    this.error   = null;
    const [usersR, teachersR, studentsR, classroomsR] = await Promise.allSettled([
      GRAPH.getListItems(SETUP_LISTS.users),
      GRAPH.getWhoAreYouVisiting(),
      GRAPH.getListItems(SETUP_LISTS.students).catch(() => []),
      GRAPH.getListItems(SETUP_LISTS.classrooms).catch(() => [])
    ]);
    this.users      = usersR.status      === "fulfilled" ? usersR.value.map(normalizeSetupUser)      : [];
    this.teachers   = teachersR.status   === "fulfilled" ? teachersR.value.map(normalizeSetupTeacher) : [];
    this.students   = studentsR.status   === "fulfilled" ? studentsR.value.map(normalizeSetupStudent) : [];
    this.classrooms = classroomsR.status === "fulfilled" ? classroomsR.value.map(normalizeSetupClassroom) : [];
    if (usersR.status === "rejected") this.error = usersR.reason?.message || "Failed to load users.";
    this.loading = false;
    if (typeof APP !== "undefined" && APP._refreshSetupUI) APP._refreshSetupUI();
  }
};

/* ── Reports Module ─────────────────────────────────────────────────────────── */

function firstDefined(obj, keys, fallback = "") {
  for (const key of keys) {
    const v = obj?.[key];
    if (v !== undefined && v !== null && v !== "") return v;
  }
  return fallback;
}

function normalizeBoolean(value) {
  return value === true || value === 1 || value === "1" ||
    String(value).toLowerCase() === "true" || String(value).toLowerCase() === "yes";
}

function normalizeWalkthroughArray(value) {
  if (Array.isArray(value)) return value.map(String).map(v => v.trim()).filter(Boolean);
  return String(value || "").split(",").map(v => v.trim()).filter(Boolean);
}

function normalizePaceScmUsed(item) {
  const v = item.scmUsed !== undefined ? item.scmUsed
    : item.SCMUsed !== undefined ? item.SCMUsed
    : item["SCM Used"] !== undefined ? item["SCM Used"]
    : item.SCM_x0020_Used;
  if (v === null || v === undefined) return null;
  return v === true || v === 1 || String(v).toLowerCase() === "true";
}

/* ── Walkthrough submission guards ──────────────────────────────────────────── */

let _walkthroughInFlight = false;
const _walkthroughDraft  = { submissionId: null };

function ensureWalkthroughSubmissionId() {
  if (!_walkthroughDraft.submissionId) {
    _walkthroughDraft.submissionId = crypto.randomUUID();
  }
  return _walkthroughDraft.submissionId;
}

function _handleWalkthroughSubmit(e) {
  e.preventDefault();
  e.stopPropagation();
  console.count("handleWalkthroughSubmit called");
  APP._submitWalkthrough(document.getElementById("page-walkthrough"));
}

function deduplicateWalkthroughReports(records) {
  const seen = new Map();
  records.forEach(record => {
    const key = record.submissionId || record.observationId || record.sharePointId;
    if (key && !seen.has(key)) seen.set(key, record);
    else if (!key) seen.set(Symbol(), record);
  });
  return [...seen.values()];
}

function normalizeSharePointWalkthrough(item) {
  const observationTimestamp = firstDefined(item, [
    "ObservationTimestamp", "Observation_x0020_Timestamp", "SubmittedAt", "Submitted_x0020_At", "Created"
  ]);
  const observationDate = firstDefined(item, ["ObservationDate", "Observation_x0020_Date", "Date"]);
  return {
    sharePointId:         item.id || "",
    observationId:        firstDefined(item, ["ObservationID", "Observation_x0020_ID", "Title"]),
    sessionId:            firstDefined(item, ["SessionID", "Session_x0020_ID"]),
    observationDate,
    observationTime:      firstDefined(item, ["ObservationTime", "Observation_x0020_Time"]),
    observationTimestamp,
    observer:             firstDefined(item, ["Observer", "field_1"]),
    observerEmail:        firstDefined(item, ["ObserverEmail", "Observer_x0020_Email"]),
    teacher:              firstDefined(item, ["Teacher"]),
    student:              firstDefined(item, ["Student"]),
    classroom:            firstDefined(item, ["Classroom"]),
    focus:                firstDefined(item, ["Focus"]),
    environment:          firstDefined(item, ["Environment"]),
    observationLength:    Number(firstDefined(item, ["ObservationLength", "Observation_x0020_Length"], 0)) || 0,
    engagement:           firstDefined(item, ["Engagement"]),
    supportsObserved:     normalizeWalkthroughArray(firstDefined(item, ["SupportObserved", "SupportsObserved", "Supports_x0020_Observed"])),
    disengagementReasons: normalizeWalkthroughArray(firstDefined(item, ["DisengagementReasons", "Disengagement_x0020_Reasons"])),
    supportRequested:     firstDefined(item, ["SupportRequested", "SupportsRequested", "Support_x0020_Requested"]),
    observedWin:          firstDefined(item, ["ObservedWin", "Observed_x0020_Win"]),
    concernGap:           firstDefined(item, ["ConcernGap", "Concern_x0020__x002f__x0020_Gap", "Concern_x0020_Gap"]),
    observationNotes:     firstDefined(item, ["ObservationNotes", "Observation_x0020_Notes", "Notes"]),
    followUpNotes:        firstDefined(item, ["FollowUpNotes", "Follow_x002d_Up_x0020_Notes"]),
    priority:             firstDefined(item, ["Priority"]),
    followUpNeeded:       normalizeBoolean(firstDefined(item, ["FollowUpNeeded", "Follow_x002d_Up_x0020_Needed"])),
    followUpDate:         firstDefined(item, ["FollowUpDate", "Follow_x002d_Up_x0020_Date"]),
    aiSummary:            firstDefined(item, ["AISummary", "AI_x0020_Summary"]),
    aiSuggestions:        firstDefined(item, ["AISuggestions", "AI_x0020_Suggestions"]),
    submissionId:         firstDefined(item, ["SubmissionID", "Submission_x0020_ID"]),
    synced:               normalizeBoolean(firstDefined(item, ["Synced"])),
    created:              firstDefined(item, ["Created"]),
    modified:             firstDefined(item, ["Modified"])
  };
}

function getWalkthroughSortDate(record) {
  if (record.observationTimestamp) {
    const p = new Date(record.observationTimestamp);
    if (!isNaN(p)) return p;
  }
  if (record.observationDate) {
    const combined = record.observationTime
      ? `${record.observationDate} ${record.observationTime}`
      : record.observationDate;
    const p = new Date(combined);
    if (!isNaN(p)) return p;
  }
  const c = new Date(record.created || 0);
  return isNaN(c) ? new Date(0) : c;
}

function uniqueSorted(values) {
  return [...new Set(values.map(v => String(v || "").trim()).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));
}

function formatReportDateTime(record) {
  if (record.observationDate) {
    const d = new Date(record.observationDate + "T12:00:00");
    if (!isNaN(d)) {
      const ds = d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
      return record.observationTime ? `${ds} · ${record.observationTime}` : ds;
    }
  }
  const d = getWalkthroughSortDate(record);
  return d.getTime() === 0 ? "—"
    : d.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

const REPORTS = {
  rawWalkthroughs:      [],
  walkthroughs:         [],
  filteredWalkthroughs: [],
  loading: false,
  error:   null,
  async load() {
    this.loading = true;
    this.error   = null;
    renderReportsLoadingState();
    try {
      const items = await GRAPH.getListItems("IEP_Walkthrough_Observations");
      this.rawWalkthroughs      = items || [];
      const normalized          = this.rawWalkthroughs.map(normalizeSharePointWalkthrough);
      this.walkthroughs         = deduplicateWalkthroughReports(normalized);
      this.filteredWalkthroughs = [...this.walkthroughs];
      populateReportsFilterOptions(this.walkthroughs);
      applyReportsFilters();
    } catch (err) {
      console.error("Failed to load walkthrough reports:", err);
      this.error = err.message || String(err);
      renderReportsErrorState(this.error);
    } finally {
      this.loading = false;
    }
  }
};

function renderReportsLoadingState() {
  const s = document.getElementById("reports-status");
  if (s) s.innerHTML = `<div class="reports-loading">Loading walkthrough records from SharePoint…</div>`;
  const c = document.getElementById("walkthrough-reports-container");
  if (c) c.innerHTML = "";
  const sm = document.getElementById("reports-summary");
  if (sm) sm.innerHTML = "";
}

function renderReportsErrorState(msg) {
  const s = document.getElementById("reports-status");
  if (!s) return;
  s.innerHTML = `
    <div class="reports-error">
      <p><strong>Reports could not be loaded from SharePoint.</strong></p>
      <details><summary>Technical details</summary><pre>${escHtml(msg)}</pre></details>
    </div>`;
}

function populateReportsFilterOptions(walkthroughs) {
  function setOpts(id, values, allLabel) {
    const sel = document.getElementById(id);
    if (!sel) return;
    const current = sel.value;
    sel.innerHTML = `<option value="">${allLabel}</option>` +
      uniqueSorted(values).map(v =>
        `<option value="${escHtml(v)}"${current === v ? " selected" : ""}>${escHtml(v)}</option>`
      ).join("");
  }
  setOpts("rf-teacher",    walkthroughs.map(r => r.teacher),                         "All Teachers");
  setOpts("rf-classroom",  walkthroughs.map(r => r.classroom),                        "All Classrooms");
  setOpts("rf-student",    walkthroughs.map(r => r.student).filter(Boolean),          "All Students");
  setOpts("rf-observer",   walkthroughs.map(r => r.observer),                         "All Observers");
  setOpts("rf-engagement", walkthroughs.map(r => r.engagement).filter(Boolean),       "All");
  setOpts("rf-support",    walkthroughs.map(r => r.supportRequested).filter(Boolean), "All");
  setOpts("rf-priority",   walkthroughs.map(r => r.priority).filter(Boolean),         "All");
}

function getReportsFilterValues() {
  const g = id => (document.getElementById(id)?.value || "").trim();
  return {
    fromDate:        g("rf-from"),
    toDate:          g("rf-to"),
    teacher:         g("rf-teacher"),
    classroom:       g("rf-classroom"),
    student:         g("rf-student"),
    observer:        g("rf-observer"),
    engagement:      g("rf-engagement"),
    supportRequested:g("rf-support"),
    priority:        g("rf-priority"),
    followUpNeeded:  g("rf-followup"),
    searchText:      g("rf-search")
  };
}

function applyReportsFilters() {
  const f = getReportsFilterValues();
  const filtered = REPORTS.walkthroughs.filter(r => {
    const d = getWalkthroughSortDate(r);
    if (f.fromDate && d < new Date(`${f.fromDate}T00:00:00`)) return false;
    if (f.toDate   && d > new Date(`${f.toDate}T23:59:59`))   return false;
    if (f.teacher          && r.teacher          !== f.teacher)          return false;
    if (f.classroom        && r.classroom        !== f.classroom)        return false;
    if (f.student          && r.student          !== f.student)          return false;
    if (f.observer         && r.observer         !== f.observer)         return false;
    if (f.engagement       && r.engagement       !== f.engagement)       return false;
    if (f.supportRequested && r.supportRequested !== f.supportRequested) return false;
    if (f.priority         && r.priority         !== f.priority)         return false;
    if (f.followUpNeeded === "yes" && !r.followUpNeeded) return false;
    if (f.followUpNeeded === "no"  &&  r.followUpNeeded) return false;
    if (f.searchText) {
      const blob = [
        r.observationId, r.sessionId, r.observer, r.teacher, r.student, r.classroom,
        r.engagement, r.supportRequested, r.observedWin, r.concernGap,
        r.observationNotes, r.followUpNotes,
        r.supportsObserved.join(" "), r.disengagementReasons.join(" ")
      ].join(" ").toLowerCase();
      if (!blob.includes(f.searchText.toLowerCase())) return false;
    }
    return true;
  });
  REPORTS.filteredWalkthroughs = filtered.sort((a, b) =>
    getWalkthroughSortDate(b) - getWalkthroughSortDate(a)
  );
  const status = document.getElementById("reports-status");
  if (status) status.innerHTML = "";
  renderReportsSummary(REPORTS.filteredWalkthroughs);
  renderWalkthroughReportsTable(REPORTS.filteredWalkthroughs);
}

function calculateReportsSummary(records) {
  const supportCounts = { none: 0, weekly: 0, soon: 0, urgent: 0 };
  records.forEach(r => {
    const s = String(r.supportRequested || "").toLowerCase();
    if (!s || s === "none") supportCounts.none++;
    else if (s.includes("weekly")) supportCounts.weekly++;
    else if (s.includes("soon"))   supportCounts.soon++;
    else if (s.includes("urgent")) supportCounts.urgent++;
  });
  return {
    total:          records.length,
    teacherCount:   new Set(records.map(r => r.teacher).filter(Boolean)).size,
    classroomCount: new Set(records.map(r => r.classroom).filter(Boolean)).size,
    studentCount:   new Set(records.map(r => r.student).filter(Boolean)).size,
    followUpCount:  records.filter(r => r.followUpNeeded).length,
    supportCounts
  };
}

function renderReportsSummary(records) {
  const el = document.getElementById("reports-summary");
  if (!el) return;
  const s = calculateReportsSummary(records);
  const supCounts = {};
  records.forEach(r => r.supportsObserved.forEach(sup => { supCounts[sup] = (supCounts[sup] || 0) + 1; }));
  const topSup = Object.entries(supCounts).sort((a, b) => b[1] - a[1]);
  const maxSup = topSup[0]?.[1] || 1;

  el.innerHTML = `
    <div class="summary-grid">
      <div class="card">
        <div class="card-title">This Report</div>
        <ul class="status-list">
          <li class="status-row"><span class="status-label-text">Walkthroughs</span><span class="status-count">${s.total}</span></li>
          <li class="status-row"><span class="status-label-text">Teachers visited</span><span class="status-count">${s.teacherCount}</span></li>
          <li class="status-row"><span class="status-label-text">Classrooms</span><span class="status-count">${s.classroomCount}</span></li>
          <li class="status-row"><span class="status-label-text">Students named</span><span class="status-count">${s.studentCount}</span></li>
          <li class="status-row"><span class="status-label-text">Follow-Up Needed</span><span class="status-count">${s.followUpCount}</span></li>
        </ul>
      </div>
      <div class="card">
        <div class="card-title">Support Requested</div>
        <ul class="status-list">
          <li class="status-row"><span class="status-label-text">No action needed</span><span class="status-count">${s.supportCounts.none}</span></li>
          <li class="status-row"><span class="status-label-text">Discuss at weekly meeting</span><span class="status-count">${s.supportCounts.weekly}</span></li>
          <li class="status-row"><span class="status-label-text">Help soon</span><span class="status-count">${s.supportCounts.soon}</span></li>
          <li class="status-row"><span class="status-label-text">Urgent</span><span class="status-count">${s.supportCounts.urgent}</span></li>
        </ul>
      </div>
    </div>
    ${topSup.length > 0 ? `
    <div class="card mb-20">
      <div class="card-title">Supports Observed (${records.length} walkthroughs)</div>
      <ul class="support-bar-list">${topSup.map(([sup, n]) => `
        <li class="support-bar-row">
          <div class="support-bar-meta"><span>${escHtml(sup)}</span><span>${n}</span></div>
          <div class="support-bar-track"><div class="support-bar-fill" style="width:${Math.round(n / maxSup * 100)}%"></div></div>
        </li>`).join("")}
      </ul>
    </div>` : ""}`;
}

function renderWalkthroughReportsTable(records) {
  const container = document.getElementById("walkthrough-reports-container");
  if (!container) return;
  if (!records.length) {
    container.innerHTML = `<div class="empty-state"><div class="empty-icon">📋</div><p>No walkthrough records match the current filters.</p></div>`;
    return;
  }
  container.innerHTML = `
    <div class="card">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:8px">
        <div class="card-title" style="margin-bottom:0">Walkthrough Records (${records.length})</div>
        <div class="btn-group">
          <button class="btn btn-secondary btn-sm" id="exportReportsCsvBtn">Export CSV</button>
          <button class="btn btn-secondary btn-sm" id="exportReportsJsonBtn">Export JSON</button>
        </div>
      </div>
      <div class="table-wrap">
        <table class="reports-table">
          <thead>
            <tr>
              <th>Date / Time</th><th>Teacher</th><th>Classroom</th><th>Student</th>
              <th>Observer</th><th>Engagement</th><th>Supports</th>
              <th>Support Requested</th><th>Follow-Up</th><th></th>
            </tr>
          </thead>
          <tbody>
            ${records.map(r => `
              <tr>
                <td style="white-space:nowrap">${formatReportDateTime(r)}</td>
                <td>${escHtml(r.teacher || "—")}</td>
                <td>${escHtml(r.classroom || "—")}</td>
                <td>${escHtml(r.student || "Whole class")}</td>
                <td>${escHtml(r.observer || "—")}</td>
                <td>${escHtml(r.engagement || "—")}</td>
                <td style="max-width:160px;font-size:12px">${escHtml(r.supportsObserved.join(", ") || "—")}</td>
                <td>${escHtml(r.supportRequested || "None")}</td>
                <td>${r.followUpNeeded ? "Yes" : "No"}</td>
                <td><button class="report-view-btn" data-report-id="${escHtml(r.sharePointId)}">View</button></td>
              </tr>`).join("")}
          </tbody>
        </table>
      </div>
    </div>`;

  document.getElementById("exportReportsCsvBtn").addEventListener("click", exportReportsCsv);
  document.getElementById("exportReportsJsonBtn").addEventListener("click", exportReportsJson);
  container.querySelectorAll(".report-view-btn").forEach(btn =>
    btn.addEventListener("click", () => openWalkthroughReportDetails(btn.dataset.reportId))
  );
}

function openWalkthroughReportDetails(sharePointId) {
  const record = REPORTS.walkthroughs.find(r => String(r.sharePointId) === String(sharePointId));
  if (!record) return;
  renderReportDetailModal(record);
}

function renderReportDetailModal(record) {
  const body = document.getElementById("report-modal-body");
  if (!body) return;

  function row(label, value) {
    if (value === "" || value === null || value === undefined) return "";
    return `<div class="detail-row"><span class="detail-label">${escHtml(label)}</span><span class="detail-value">${escHtml(String(value))}</span></div>`;
  }
  function boolRow(label, value) {
    return `<div class="detail-row"><span class="detail-label">${escHtml(label)}</span><span class="detail-value">${value ? "Yes" : "No"}</span></div>`;
  }
  function listRow(label, arr) {
    if (!arr || !arr.length) return "";
    return `<div class="detail-row"><span class="detail-label">${escHtml(label)}</span><span class="detail-value">${escHtml(arr.join(", "))}</span></div>`;
  }
  function tsRow(label, iso) {
    if (!iso) return "";
    const d = new Date(iso);
    return isNaN(d) ? "" : `<div class="detail-row"><span class="detail-label">${escHtml(label)}</span><span class="detail-value">${d.toLocaleString("en-US", { month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit" })}</span></div>`;
  }

  body.innerHTML = `
    <div class="detail-section">
      <div class="detail-section-title">Observation</div>
      ${row("Observation ID", record.observationId)}
      ${row("Session ID", record.sessionId)}
      ${row("Observation Date", record.observationDate)}
      ${row("Observation Time", record.observationTime)}
      ${tsRow("Submitted", record.observationTimestamp || record.created)}
      ${row("Observer", record.observer)}
      ${row("Observer Email", record.observerEmail)}
    </div>
    <div class="detail-section">
      <div class="detail-section-title">Location</div>
      ${row("Teacher", record.teacher)}
      ${row("Classroom", record.classroom)}
      ${row("Student", record.student || "Whole class")}
      ${row("Focus", record.focus)}
      ${row("Environment", record.environment)}
      ${record.observationLength ? row("Observation Length", record.observationLength + " min") : ""}
    </div>
    <div class="detail-section">
      <div class="detail-section-title">Observation Data</div>
      ${row("Engagement", record.engagement)}
      ${listRow("Supports Observed", record.supportsObserved)}
      ${listRow("Disengagement Reasons", record.disengagementReasons)}
      ${row("Support Requested", record.supportRequested)}
      ${boolRow("Follow-Up Needed", record.followUpNeeded)}
    </div>
    <div class="detail-section">
      <div class="detail-section-title">Narrative</div>
      ${row("Observed Win", record.observedWin)}
      ${row("Concern / Gap", record.concernGap)}
      ${row("Observation Notes", record.observationNotes)}
      ${row("Follow-Up Notes", record.followUpNotes)}
      ${row("Follow-Up Date", record.followUpDate)}
      ${row("Priority", record.priority)}
    </div>
    ${record.aiSummary || record.aiSuggestions ? `
    <div class="detail-section">
      <div class="detail-section-title">AI Insights</div>
      ${row("AI Summary", record.aiSummary)}
      ${row("AI Suggestions", record.aiSuggestions)}
    </div>` : ""}
    <div class="detail-section">
      <div class="detail-section-title">System</div>
      ${row("Submission ID", record.submissionId)}
      ${row("SharePoint ID", record.sharePointId)}
      ${tsRow("Created", record.created)}
      ${tsRow("Modified", record.modified)}
    </div>`;

  const modal = document.getElementById("report-detail-modal");
  if (modal) {
    modal.classList.remove("hidden");
    document.body.style.overflow = "hidden";
  }
}

function closeReportDetailModal() {
  const modal = document.getElementById("report-detail-modal");
  if (modal) modal.classList.add("hidden");
  document.body.style.overflow = "";
}

function exportReportsCsv() {
  const records = REPORTS.filteredWalkthroughs;
  const headers = [
    "Observation ID","Observation Date","Observation Time","Submission Timestamp",
    "Observer","Teacher","Classroom","Student","Focus","Engagement",
    "Supports Observed","Disengagement Reasons","Support Requested",
    "Observed Win","Concern / Gap","Observation Notes",
    "Follow-Up Needed","Follow-Up Notes","Priority"
  ];
  const cell = v => `"${String(v || "").replace(/"/g, '""')}"`;
  const rows = records.map(r => [
    cell(r.observationId), cell(r.observationDate), cell(r.observationTime),
    cell(r.observationTimestamp || r.created), cell(r.observer), cell(r.teacher),
    cell(r.classroom), cell(r.student), cell(r.focus), cell(r.engagement),
    cell(r.supportsObserved.join("; ")), cell(r.disengagementReasons.join("; ")),
    cell(r.supportRequested), cell(r.observedWin), cell(r.concernGap),
    cell(r.observationNotes), cell(r.followUpNeeded ? "Yes" : "No"),
    cell(r.followUpNotes), cell(r.priority)
  ].join(","));
  const ts = new Date().toISOString().slice(0, 10);
  downloadFile(`walkthrough-report-${ts}.csv`, [headers.join(","), ...rows].join("\n"), "text/csv");
  showToast("CSV exported.");
}

function exportReportsJson() {
  const records = REPORTS.filteredWalkthroughs;
  const ts = new Date().toISOString().slice(0, 10);
  downloadFile(`walkthrough-report-${ts}.json`,
    JSON.stringify({ exportedAt: new Date().toISOString(), count: records.length, records }, null, 2),
    "application/json");
  showToast("JSON exported.");
}

/* ── Daily Pulse Scales ───────────────────────────────────────────────────────── */

const DAILY_PULSE_SCALES = {
  academicProgress: {
    label:  "Academic Progress",
    prompt: "How did the student participate or demonstrate skills in reading and/or math?",
    choices: [
      { value: "not-addressed",       label: "Not addressed today",                          tone: "neutral"  },
      { value: "significant-support", label: "Participated but required significant support", tone: "negative" },
      { value: "some-support",        label: "Demonstrated skills with some support",         tone: "positive" },
      { value: "independent",         label: "Demonstrated skills independently",             tone: "positive" },
      { value: "noticeable-progress", label: "Made noticeable progress",                      tone: "positive" }
    ],
    detailFields: [
      { key: "skillArea", label: "Skill, activity, or area of progress", type: "text", optional: true }
    ]
  },
  behavior: {
    label:  "Behavior",
    prompt: "Which statement best describes the student's behavior today?",
    choices: [
      { value: "excellent", label: "Excellent day — no behavioral concerns",                              tone: "positive" },
      { value: "minor",     label: "Minor concerns — redirection was effective",                          tone: "negative" },
      { value: "moderate",  label: "Moderate concerns — repeated behaviors or additional support needed", tone: "negative" },
      { value: "major",     label: "Major concerns — unsafe behavior occurred",                           tone: "negative" }
    ],
    multiSelectFields: [
      { key: "behaviorsObserved", label: "Behaviors observed", options: [
        "Hitting", "Biting", "Scratching", "Kicking", "Elopement / attempted elopement",
        "Property destruction", "Verbal aggression", "Refusal / noncompliance", "Other"
      ]}
    ],
    detailFields: [
      { key: "behaviorContext", label: "What occurred, possible trigger, and staff response", type: "textarea", optional: true },
      { key: "otherBehavior",  label: "Other behavior", type: "text", optional: true, showWhenOption: "Other" }
    ]
  },
  communication: {
    label:  "Communication",
    prompt: "Consider verbal communication, gestures, signs, picture symbols, or an AAC device.",
    choices: [
      { value: "did-not-communicate",   label: "Did not communicate needs today",             tone: "negative" },
      { value: "significant-prompting", label: "Communicated with significant prompting",      tone: "negative" },
      { value: "occasional-prompting",  label: "Communicated with occasional prompting",       tone: "positive" },
      { value: "independent",           label: "Communicated independently and appropriately", tone: "positive" },
      { value: "not-observed",          label: "Not observed / not applicable",                tone: "neutral"  }
    ],
    multiSelectFields: [
      { key: "communicationMethods", label: "Communication method", options: [
        "Verbal", "Gestures / sign language", "Picture symbols", "AAC / communication device", "Other"
      ]}
    ],
    detailFields: [
      { key: "communicationComments",    label: "Comments or observations",   type: "textarea", optional: true },
      { key: "otherCommunicationMethod", label: "Other communication method", type: "text", optional: true, showWhenOption: "Other" }
    ]
  },
  socialSkills: {
    label:  "Social Skills",
    prompt: "How did the student interact with peers or adults?",
    choices: [
      { value: "avoided",             label: "Avoided or did not interact with others",        tone: "negative" },
      { value: "significant-support", label: "Interacted only with significant adult support", tone: "negative" },
      { value: "some-support",        label: "Had some appropriate interactions with support", tone: "positive" },
      { value: "independent",         label: "Interacted appropriately and independently",     tone: "positive" },
      { value: "not-observed",        label: "Not observed / not applicable",                  tone: "neutral"  }
    ],
    detailFields: [
      { key: "interactionNotes", label: "Interaction with peers or adults", type: "textarea", optional: true }
    ]
  },
  selfRegulation: {
    label:  "Self-Regulation",
    prompt: "How well did the student manage emotions, sensory needs, frustration, and changes?",
    choices: [
      { value: "regulated",           label: "Remained calm and regulated throughout the day",         tone: "positive" },
      { value: "recovered-minimal",   label: "Became dysregulated but recovered with minimal support", tone: "positive" },
      { value: "repeated-support",    label: "Required repeated adult support or coping strategies",   tone: "negative" },
      { value: "significant-dysreg",  label: "Significant dysregulation affected learning or safety",  tone: "negative" },
      { value: "not-observed",        label: "Not observed / not applicable",                          tone: "neutral"  }
    ],
    multiSelectFields: [
      { key: "regulationStrategies", label: "Strategies used", options: [
        "Break", "Quiet / calming area", "Sensory tool", "Movement",
        "Breathing / calming strategy", "Adult reassurance", "Visual schedule", "Other"
      ]}
    ],
    detailFields: [
      { key: "regulationComments",      label: "Comments or observations", type: "textarea", optional: true },
      { key: "otherRegulationStrategy", label: "Other strategy",           type: "text",     optional: true, showWhenOption: "Other" }
    ]
  },
  transitionDifficulty: {
    label:  "Transition Difficulty",
    prompt: "How did the student manage transitions today?",
    choices: [
      { value: "none",         label: "No difficulty — transitioned independently",                                  tone: "positive" },
      { value: "minor",        label: "Minor difficulty — needed one or two prompts",                                tone: "negative" },
      { value: "some",         label: "Some difficulty — needed repeated prompting or support",                      tone: "negative" },
      { value: "major",        label: "Major difficulty — refusal, aggression, elopement, or significant distress", tone: "negative" },
      { value: "not-observed", label: "Not observed / not applicable",                                              tone: "neutral"  }
    ],
    detailFields: [
      { key: "difficultTransition", label: "Location or transition that was difficult", type: "textarea", optional: true, showForValues: ["minor", "some", "major"] }
    ]
  },
  iepGoalProgress: {
    label:  "IEP Goal Progress",
    prompt: "How did the student progress toward an IEP goal today?",
    choices: [
      { value: "not-addressed",        label: "Goal was not addressed today",             tone: "neutral"  },
      { value: "no-progress",          label: "Goal addressed — no progress observed",    tone: "negative" },
      { value: "some-progress",        label: "Goal addressed — some progress observed",  tone: "positive" },
      { value: "significant-progress", label: "Goal addressed — significant progress",    tone: "positive" },
      { value: "independent",          label: "Goal completed independently",             tone: "positive" }
    ],
    multiSelectFields: [
      { key: "goalAreas", label: "Goal area addressed", options: [
        "Academic", "Behavior", "Communication", "Social skills",
        "Functional / life skills", "Related service", "Other"
      ]}
    ],
    detailFields: [
      { key: "otherGoalArea", label: "Other goal area", type: "text", optional: true, showWhenOption: "Other" }
    ]
  }
};

/* ── APP Object ──────────────────────────────────────────────────────────────── */
const APP = {
  currentPage: "walkthrough",
  _v2ExpandedStudents: {},

  async init() {
    const loginScreen = document.getElementById("loginScreen");
    const loginError  = document.getElementById("loginError");
    const loginBtn    = document.getElementById("loginBtn");

    try {
      await AUTH.init();
    } catch (err) {
      console.error("AUTH.init failed:", err);
      loginScreen.classList.remove("hidden");
      loginError.textContent = "Authentication service failed to load. Check your connection and try again.";
      loginError.classList.remove("hidden");
      return;
    }

    if (!AUTH.account) {
      loginScreen.classList.remove("hidden");
      loginBtn.addEventListener("click", () => AUTH.login());
      return;
    }

    if (AUTH.isUnauthorized) {
      loginScreen.classList.remove("hidden");
      loginError.textContent = AUTH.lookupError ||
        "You signed in successfully, but your account has not been added to IEP Skook yet. Please contact an administrator.";
      loginError.classList.remove("hidden");
      const debugPanel = document.getElementById("authDebugPanel");
      if (debugPanel && window.IEP_AUTH_DEBUG) {
        debugPanel.textContent = JSON.stringify(window.IEP_AUTH_DEBUG, null, 2);
        debugPanel.classList.remove("hidden");
      }
      loginBtn.textContent = "Sign in with a different account";
      loginBtn.addEventListener("click", () => AUTH.logout());
      return;
    }

    this.currentUserId = AUTH.pilotUserId;

    USER_CONTEXT.actualUser = {
      userId: AUTH.pilotUserId,
      name:   AUTH.pilotUser?.Name || AUTH.displayName,
      email:  AUTH.account?.username || "",
      role:   AUTH.role
    };

    if (AUTH.isAdmin) {
      try {
        const saved = sessionStorage.getItem("iepViewAsTeacher");
        if (saved) {
          const parsed = JSON.parse(saved);
          USER_CONTEXT.viewedTeacher = AVAILABLE_TEACHERS.find(t => t.userId === parsed?.userId) || null;
          if (!USER_CONTEXT.viewedTeacher) sessionStorage.removeItem("iepViewAsTeacher");
        }
      } catch { sessionStorage.removeItem("iepViewAsTeacher"); }
    }

    this._runMigration();
    this.bindNav();
    this.bindSidebar();
    this._seedPilotData();
    this._ensureMamrosh();
    this._bindUserSelector();
    document.getElementById("fabBtn").addEventListener("click", () => this.goTo("walkthrough"));

    const initHash = (location.hash || "").replace("#", "") || "";
    let landing;
    if (AUTH.isTeacher || USER_CONTEXT.isViewingAsTeacher) {
      // Daily Pulse is the teacher default landing page; still honor a valid
      // deep link (e.g. a reload while on #pace) rather than always bouncing.
      landing = (initHash && canAccessRoute(initHash)) ? initHash : "pulse";
    } else {
      landing = (initHash && canAccessRoute(initHash)) ? initHash : "walkthrough";
    }
    this.navigate(landing);

    window.addEventListener("hashchange", () => {
      const p = (location.hash || "").replace("#", "") || "walkthrough";
      this.navigate(p);
    });
  },

  bindNav() {
    document.getElementById("navList").addEventListener("click", e => {
      const link = e.target.closest(".nav-link");
      if (!link) return;
      e.preventDefault();
      this.goTo(link.dataset.page);
    });
  },

  goTo(page) {
    const nextHash = `#${page}`;
    // Assigning the current hash does not fire hashchange. Render directly when
    // role/context changes have left the URL pointing at the requested route.
    if (location.hash === nextHash) {
      this.navigate(page);
    } else {
      location.hash = page;
    }
  },

  bindSidebar() {
    const btn     = document.getElementById("menuBtn");
    const sidebar = document.getElementById("sidebar");
    const overlay = document.getElementById("sidebarOverlay");
    const close   = () => { sidebar.classList.remove("open"); overlay.classList.add("hidden"); };
    btn.addEventListener("click", () => {
      const open = sidebar.classList.toggle("open");
      overlay.classList.toggle("hidden", !open);
    });
    overlay.addEventListener("click", close);
  },

  // Identity is derived from the authenticated AUTH.role, never from
  // membership in the hardcoded PILOT_TEACHERS list — a real teacher who
  // isn't one of the original pilot teachers must still resolve as a
  // teacher (isAdmin: false), not silently fall back to administrator.
  getCurrentUser() {
    if (USER_CONTEXT.isViewingAsTeacher) {
      const t = USER_CONTEXT.viewedTeacher;
      return { id: t.userId, name: t.displayName, teacherName: t.teacherName, role: "teacher", isAdmin: false };
    }
    if (AUTH.isTeacher) {
      const t  = USER_CONTEXT.effectiveTeacher || {};
      const id = t.userId || slugifyTeacherId(t.teacherName || t.displayName) || AUTH.pilotUserId || "";
      return {
        id,
        name:        t.displayName || t.teacherName || AUTH.pilotUser?.Name || AUTH.displayName || "",
        teacherName: t.teacherName || t.displayName || "",
        role:        "teacher",
        isAdmin:     false
      };
    }
    // Administrator (or any other authenticated/authorized role) — preserves
    // the previous single "admin-decusky" identity used throughout the app.
    return { ...PILOT_ADMIN, name: USER_CONTEXT.actualUser?.name || AUTH.pilotUser?.Name || PILOT_ADMIN.name, isAdmin: true };
  },

  setCurrentUser(userId) {
    this.currentUserId = userId;
    localStorage.setItem("iepSkookCurrentUser", userId);
    const indicator = document.getElementById("topBarUser");
    if (indicator) {
      const u = this.getCurrentUser();
      indicator.textContent = u.name + (u.isAdmin ? " (Admin)" : "");
    }
    this.navigate(this.currentPage);
  },

  _bindUserSelector() {
    const nameEl = document.getElementById("authDisplayName");
    if (nameEl) nameEl.textContent = AUTH.pilotUser?.Name || AUTH.displayName;

    const signOutBtn = document.getElementById("signOutBtn");
    if (signOutBtn) signOutBtn.addEventListener("click", () => AUTH.logout());

    this._renderTopBarControls();
  },

  renderNav() {
    const navList = document.getElementById("navList");
    if (!navList) return;

    const isTeacherView = USER_CONTEXT.effectiveRole === "teacher";
    const cp = this.currentPage;

    const icon = {
      walkthrough:      `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path d="M13.586 3.586a2 2 0 112.828 2.828l-.793.793-2.828-2.828.793-.793zM11.379 5.793L3 14.172V17h2.828l8.38-8.379-2.83-2.828z"/></svg>`,
      pace:             `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h9"/><path d="M10 12h11"/><path d="M18 9l3 3-3 3"/></svg>`,
      dashboard:        `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path d="M3 4a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1H4a1 1 0 01-1-1V4zm0 8a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1H4a1 1 0 01-1-1v-4zm8-8a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 01-1-1V4zm0 8a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 01-1-1v-4z"/></svg>`,
      teacherDashboard: `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path d="M3 4a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1H4a1 1 0 01-1-1V4zm0 8a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1H4a1 1 0 01-1-1v-4zm8-8a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 01-1-1V4zm0 8a1 1 0 011-1h4a1 1 0 011 1v4a1 1 0 01-1 1h-4a1 1 0 01-1-1v-4z"/></svg>`,
      reports:          `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path d="M2 11a1 1 0 011-1h2a1 1 0 011 1v5a1 1 0 01-1 1H3a1 1 0 01-1-1v-5zM8 7a1 1 0 011-1h2a1 1 0 011 1v9a1 1 0 01-1 1H9a1 1 0 01-1-1V7zm6-4a1 1 0 011-1h2a1 1 0 011 1v13a1 1 0 01-1 1h-2a1 1 0 01-1-1V3z"/></svg>`,
      pulse:            `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>`,
      checkin:          `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="8" r="4"/><path d="M6 20v-1a6 6 0 0 1 12 0v1"/><path d="M16 11l2 2 4-4"/></svg>`,
      formlab:          `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M9 2a1 1 0 000 2h2a1 1 0 100-2H9z"/><path fill-rule="evenodd" d="M4 5a2 2 0 012-2 3 3 0 003 3h2a3 3 0 003-3 2 2 0 012 2v11a2 2 0 01-2 2H6a2 2 0 01-2-2V5zm3 4a1 1 0 000 2h.01a1 1 0 100-2H7zm3 0a1 1 0 000 2h3a1 1 0 100-2h-3zm-3 4a1 1 0 100 2h.01a1 1 0 100-2H7zm3 0a1 1 0 100 2h3a1 1 0 100-2h-3z" clip-rule="evenodd"/></svg>`,
      setup:            `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.061 2.042-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z" clip-rule="evenodd"/></svg>`,
      storage:          `<svg class="nav-icon" viewBox="0 0 20 20" fill="currentColor"><path d="M4 4a2 2 0 00-2 2v1h16V6a2 2 0 00-2-2H4z"/><path fill-rule="evenodd" d="M18 9H2v5a2 2 0 002 2h12a2 2 0 002-2V9zM4 13a1 1 0 011-1h1a1 1 0 110 2H5a1 1 0 01-1-1zm5-1a1 1 0 100 2h1a1 1 0 100-2H9z" clip-rule="evenodd"/></svg>`,
      v2walkthrough:    `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/></svg>`
    };

    const items = isTeacherView
      ? [
          { page: "pulse", label: "Daily Pulse", cls: "nav-link" },
          { page: "pace",  label: "PACE Log",     cls: "nav-link" }
        ]
      : [
          { page: "walkthrough",   label: "New Walkthrough",  cls: "nav-link nav-primary-action", aria: "New Walkthrough" },
          { page: "pace",          label: "PACE Log",         cls: "nav-link" },
          { page: "dashboard",     label: "Dashboard",        cls: "nav-link" },
          { page: "reports",       label: "Reports",          cls: "nav-link" },
          { page: "pulse",         label: "Daily Pulse",      cls: "nav-link" },
          { page: "checkin",       label: "Student Check-In", cls: "nav-link" },
          { page: "formlab",       label: "Form Lab",         cls: "nav-link" },
          { page: "setup",         label: "Setup",            cls: "nav-link" },
          { page: "storage",       label: "Storage Settings", cls: "nav-link" },
          { page: "v2walkthrough", label: "V2 Walkthrough",   cls: "nav-link nav-prototype-link", aria: "V2 Walkthrough Prototype", badge: "Prototype" }
        ];

    navList.innerHTML = items.map(item => `
      <li><a class="${item.cls}${cp === item.page ? " active" : ""}"
             data-page="${escHtml(item.page)}" href="#${escHtml(item.page)}"
             ${item.aria ? `aria-label="${escHtml(item.aria)}"` : ""}>
        ${icon[item.page] || ""}
        <span>${escHtml(item.label)}</span>
        ${item.badge ? `<span class="nav-prototype-badge">${escHtml(item.badge)}</span>` : ""}
      </a></li>`).join("");
  },

  _renderTopBarControls() {
    const actualName = USER_CONTEXT.actualUser?.name || AUTH.pilotUser?.Name || AUTH.displayName || "";
    const indicator  = document.getElementById("topBarUser");
    if (indicator) {
      indicator.textContent = AUTH.isAdmin ? `${actualName} (Admin)` : actualName;
    }

    const viewAsControl = document.getElementById("viewAsControl");
    if (viewAsControl) {
      if (AUTH.isAdmin) {
        const selected = USER_CONTEXT.viewedTeacher?.userId || "";
        viewAsControl.innerHTML = `
          <label class="view-as-label" for="viewAsSelect">View As</label>
          <select class="view-as-select" id="viewAsSelect" aria-label="View as teacher">
            <option value="">Administrator View</option>
            ${AVAILABLE_TEACHERS.map(t =>
              `<option value="${escHtml(t.userId)}"${t.userId === selected ? " selected" : ""}>${escHtml(t.displayName)}</option>`
            ).join("")}
          </select>`;
        document.getElementById("viewAsSelect").addEventListener("change", e => {
          this.setViewedTeacher(e.target.value || null);
        });
      } else {
        viewAsControl.innerHTML = "";
      }
    }

    const banner = document.getElementById("viewAsBanner");
    if (banner) {
      if (USER_CONTEXT.isViewingAsTeacher) {
        const t = USER_CONTEXT.viewedTeacher;
        banner.innerHTML = `
          <div class="view-as-banner">
            <div class="view-as-banner-text">
              <strong>Viewing as ${escHtml(t.displayName)}</strong>
              <span>You are still signed in as ${escHtml(actualName)}.</span>
            </div>
            <button class="view-as-exit-btn" id="exitTeacherView">Exit Teacher View</button>
          </div>`;
        banner.classList.remove("hidden");
        document.getElementById("exitTeacherView").addEventListener("click", () => this.setViewedTeacher(null));
      } else {
        banner.innerHTML = "";
        banner.classList.add("hidden");
      }
    }
  },

  setViewedTeacher(userId) {
    if (!AUTH.isAdmin) return;
    if (!userId) {
      USER_CONTEXT.viewedTeacher = null;
      sessionStorage.removeItem("iepViewAsTeacher");
    } else {
      const teacher = AVAILABLE_TEACHERS.find(t => t.userId === userId);
      if (!teacher) return;
      USER_CONTEXT.viewedTeacher = teacher;
      sessionStorage.setItem("iepViewAsTeacher", JSON.stringify(teacher));
    }
    // A draft belongs to the roster it was started under. Never carry an admin
    // or another teacher's selected student across a context switch.
    this._pulseDraft = null;
    this._renderTopBarControls();
    this.goTo(userId ? "pulse" : "dashboard");
  },

  _runMigration() {
    const stored = localStorage.getItem(DATA_VERSION_KEY);
    if (stored !== DATA_VERSION) {
      DB.clearAll();
      localStorage.setItem(DATA_VERSION_KEY, DATA_VERSION);
    }
  },

  navigate(page) {
    const pages = ["walkthrough", "v2walkthrough", "pulse", "pace", "checkin", "dashboard", "reports", "formlab", "setup", "storage", "teacherDashboard"];

    if (!canAccessRoute(page)) {
      page = USER_CONTEXT.effectiveRole === "teacher" ? "pulse" : "dashboard";
    } else if (!pages.includes(page)) {
      page = USER_CONTEXT.effectiveRole === "teacher" ? "pulse" : "walkthrough";
    }

    this.currentPage = page;

    pages.forEach(p => {
      document.getElementById("page-" + p)?.classList.toggle("hidden", p !== page);
    });

    this.renderNav();

    const titles = {
      walkthrough:      "New Walkthrough",
      v2walkthrough:    "V2 Walkthrough",
      pulse:            "Daily Pulse",
      pace:             "PACE Log",
      checkin:          "Student Check-In",
      dashboard:        "Dashboard",
      reports:          "Reports",
      formlab:          "Form Lab",
      setup:            "Setup",
      storage:          "Storage Settings",
      teacherDashboard: USER_CONTEXT.isViewingAsTeacher
        ? `Teacher View — ${USER_CONTEXT.viewedTeacher.displayName}`
        : "My Classroom Dashboard"
    };
    document.getElementById("pageTitle").textContent = titles[page] || page;

    document.getElementById("sidebar").classList.remove("open");
    document.getElementById("sidebarOverlay").classList.add("hidden");

    const noFabPages = ["walkthrough", "v2walkthrough", "pulse", "pace", "checkin", "teacherDashboard"];
    const fab = document.getElementById("fabBtn");
    if (fab) fab.classList.toggle("hidden", noFabPages.includes(page));

    const render = {
      walkthrough:      () => this.renderWalkthrough(),
      v2walkthrough:    () => this.renderWalkthroughV2(),
      pulse:            () => this.renderPulse(),
      pace:             () => this.renderPaceAdmin(),
      checkin:          () => this.renderStudentCheckIn(),
      dashboard:        () => this.renderDashboard(),
      reports:          () => this.renderReports(),
      formlab:          () => this.renderFormLab(),
      setup:            () => this.renderSetup(),
      storage:          () => this.renderStorage(),
      teacherDashboard: () => this.renderTeacherDashboard()
    };
    render[page]?.();
  },

  /* ── SEARCHABLE SELECT COMPONENT ────────────────────────────────────────────── */

  _buildSearchSelect(options, placeholder, name, required) {
    const opts = options.map(o =>
      `<div class="ss-option" data-value="${escHtml(o.value)}">${escHtml(o.label)}</div>`
    ).join("");
    const reqAttr  = required ? ' data-required="true"' : '';
    const autoShow = options.length < 30 ? ' data-auto-show="true"' : '';
    return `
      <div class="ss-wrap" data-name="${escHtml(name)}"${reqAttr}${autoShow}>
        <div style="position:relative">
          <input type="text" class="form-input ss-input" placeholder="${escHtml(placeholder)}"
                 autocomplete="off" autocorrect="off" spellcheck="false"
                 aria-label="${escHtml(placeholder)}" aria-autocomplete="list" aria-haspopup="listbox">
          <input type="hidden" name="${escHtml(name)}" class="ss-value">
          <svg class="ss-chevron" viewBox="0 0 20 20" fill="currentColor" width="16" height="16">
            <path fill-rule="evenodd" d="M5.293 7.293a1 1 0 011.414 0L10 10.586l3.293-3.293a1 1 0 111.414 1.414l-4 4a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414z" clip-rule="evenodd"/>
          </svg>
        </div>
        <div class="ss-dropdown" role="listbox">
          ${opts || '<div class="ss-no-results">No options available</div>'}
        </div>
      </div>`;
  },

  _bindSearchSelects(container) {
    container.querySelectorAll(".ss-wrap").forEach(wrap => {
      const input    = wrap.querySelector(".ss-input");
      const hidden   = wrap.querySelector(".ss-value");
      const dropdown = wrap.querySelector(".ss-dropdown");
      const options  = wrap.querySelectorAll(".ss-option");
      const autoShow = wrap.dataset.autoShow === "true";
      let selectedValue = "";
      let selectedLabel = "";

      const openDrop = () => {
        wrap.classList.add("open");
        // When list is short enough, always show everything on open so
        // the user can tap-to-browse without typing first.
        filterOpts(autoShow ? "" : input.value);
      };

      const closeDrop = () => {
        wrap.classList.remove("open");
        if (!selectedValue) { input.value = ""; }
        else { input.value = selectedLabel; }
      };

      const selectOpt = (value, label) => {
        selectedValue = value;
        selectedLabel = label;
        hidden.value  = value;
        input.value   = label;
        options.forEach(o => o.classList.toggle("ss-active", o.dataset.value === value));
        closeDrop();
        wrap.dispatchEvent(new CustomEvent("ss:change", { bubbles: true, detail: { value, label } }));
      };

      const filterOpts = (q) => {
        const lq = q.toLowerCase().trim();
        let visible = 0;
        options.forEach(o => {
          const match = !lq || o.textContent.toLowerCase().includes(lq);
          o.classList.toggle("ss-hidden", !match);
          if (match) visible++;
        });
        let noResults = dropdown.querySelector(".ss-no-results-msg");
        if (visible === 0 && lq) {
          if (!noResults) {
            noResults = document.createElement("div");
            noResults.className = "ss-no-results ss-no-results-msg";
            noResults.textContent = "No matches";
            dropdown.appendChild(noResults);
          }
        } else if (noResults) {
          noResults.remove();
        }
      };

      input.addEventListener("focus", openDrop);
      input.addEventListener("input", e => {
        openDrop();
        filterOpts(e.target.value);
        if (selectedValue && e.target.value !== selectedLabel) {
          selectedValue = "";
          hidden.value  = "";
          options.forEach(o => o.classList.remove("ss-active"));
        }
      });

      dropdown.addEventListener("mousedown", e => e.preventDefault());
      dropdown.addEventListener("click", e => {
        const opt = e.target.closest(".ss-option");
        if (opt && !opt.classList.contains("ss-hidden")) {
          selectOpt(opt.dataset.value, opt.textContent.trim());
        }
      });

      document.addEventListener("click", e => {
        if (!wrap.contains(e.target)) closeDrop();
      });
    });
  },

  /* ── OPTION CARD BINDING ─────────────────────────────────────────────────────── */

  _bindOptionCards(container) {
    container.querySelectorAll(".option-cards").forEach(group => {
      group.querySelectorAll(".option-card").forEach(card => {
        const input = card.querySelector("input");
        if (!input) return;
        const isCheckbox = input.type === "checkbox";

        card.addEventListener("click", () => {
          if (isCheckbox) {
            card.classList.toggle("selected", input.checked);
          } else {
            group.querySelectorAll(".option-card").forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
          }
        });

        input.addEventListener("change", () => {
          if (isCheckbox) {
            card.classList.toggle("selected", input.checked);
          } else {
            group.querySelectorAll(".option-card").forEach(c => c.classList.remove("selected"));
            card.classList.add("selected");
          }
        });
      });
    });

    container.querySelectorAll(".chip").forEach(chip => {
      const input = chip.querySelector("input[type='checkbox']");
      if (!input) return;
      chip.addEventListener("click", () => {
        chip.classList.toggle("selected", input.checked);
      });
      input.addEventListener("change", () => {
        chip.classList.toggle("selected", input.checked);
      });
    });
  },

  /* ── WALKTHROUGH PAGE ────────────────────────────────────────────────────────── */

  renderWalkthrough() {
    const el         = document.getElementById("page-walkthrough");
    const user       = this.getCurrentUser();
    const allTeachers   = DB.getTeachers();
    const allClassrooms = DB.getClassrooms();
    const teachers   = user.isAdmin ? allTeachers   : allTeachers.filter(t => t.id === user.id);
    const classrooms = user.isAdmin ? allClassrooms : allClassrooms.filter(c => c.teacherId === user.id);
    const availStudents = user.isAdmin ? PILOT_STUDENTS : PILOT_STUDENTS.filter(s => s.teacherId === user.id);
    const today      = new Date().toISOString().slice(0, 10);
    const weekOf     = getWeekOf(today);

    const noClassrooms = classrooms.length === 0;

    el.innerHTML = `
      <div class="walk-page-header">
        <h2 class="walk-page-title">New Walkthrough</h2>
        <p class="walk-page-sub">Capture classroom support observations quickly and consistently.</p>
      </div>

      ${noClassrooms ? `
        <div class="warning-banner" style="margin-bottom:16px">
          <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18" style="flex-shrink:0"><path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"/></svg>
          <div>Add at least one classroom in <a href="#setup" style="color:var(--orange);font-weight:700">Setup</a> before recording walkthroughs.</div>
        </div>` : ""}

      <div id="walkthroughConfirm" class="hidden"></div>

      <form id="walkthroughForm" novalidate>

        <!-- Auto-filled date bar -->
        <div class="walk-date-bar">
          <label class="walkthrough-date-picker">
            <span class="walkthrough-date-icon">📅</span>
            <input id="walkthrough-date" type="date" name="date" value="${today}" aria-label="Walkthrough date">
          </label>
          <span id="walkthrough-week-label" class="walkthrough-week-label"></span>
          <input type="hidden" id="wWeekOf" name="weekOf" value="${weekOf}">
        </div>

        <!-- Field 1: Who are you visiting? -->
        <div class="form-card walk-card section-visitor">
          <div class="walk-field-label">Who are you visiting? <span class="req">*</span></div>
          ${this._buildSearchSelect(
            teachers.map(t => ({ value: t.id, label: t.name })),
            "Search by name…", "teacherId", true
          )}
        </div>

        <!-- Field 2: Classroom -->
        <div class="form-card walk-card section-classroom">
          <div class="walk-field-label">Classroom <span class="req">*</span></div>
          ${this._buildSearchSelect(
            classrooms.map(c => ({ value: c.id, label: c.name + (c.roomNumber ? "  ·  Room " + c.roomNumber : "") })),
            "Search classrooms…", "classroomId", true
          )}
        </div>

        <!-- Field 3: Focus -->
        <div class="form-card walk-card section-focus">
          <div class="walk-field-label">Focus <span class="req">*</span></div>
          <div class="option-cards">
            ${CONFIG.FOCUS_OPTIONS.map(o => `
              <label class="option-card">
                <input type="radio" name="focus" value="${o.value}">
                <span>${escHtml(o.label)}</span>
              </label>`).join("")}
          </div>
          <div id="walkFocusRosterHint" class="hidden" style="margin-top:10px;padding:8px 12px;background:#fef9c3;border:1px solid #fde68a;border-radius:8px;font-size:13px;color:#78350f">
            No student roster has been added for this teacher. Select Whole Class or Small Group, or continue without naming a student.
          </div>
        </div>

        <!-- Field 4: Student (optional) -->
        <div class="form-card walk-card section-student">
          <div class="walk-field-label">Student <span class="walk-optional-tag">optional</span></div>
          <select class="form-select" name="studentId" id="walkthroughStudent">
            <option value="">— None / whole class —</option>
            ${availStudents.map(s => `<option value="${escHtml(s.id)}">${escHtml(s.name)}${user.isAdmin ? "  —  " + escHtml(s.teacherName) : ""}</option>`).join("")}
          </select>
        </div>

        <!-- Field 5: Engagement Observed -->
        <div class="form-card walk-card section-engagement">
          <div class="walk-field-label">Engagement Observed <span class="req">*</span></div>
          <div class="option-cards">
            ${CONFIG.ENGAGEMENT_OPTIONS.map(o => `
              <label class="option-card engagement-option" data-value="${o.value}">
                <input type="radio" name="engagementObserved" value="${o.value}">
                <span>${escHtml(o.label)}</span>
              </label>`).join("")}
          </div>
        </div>

        <!-- Field 6: Supports Observed -->
        <div class="form-card walk-card section-supports">
          <div class="walk-field-label">Supports Observed <span class="walk-optional-tag">optional</span></div>
          <div class="chip-grid">
            ${CONFIG.SUPPORTS_OPTIONS.map(s => `
              <label class="chip">
                <input type="checkbox" name="supportsObserved" value="${escHtml(s)}">
                <span>${escHtml(s)}</span>
              </label>`).join("")}
          </div>
        </div>

        <!-- Field 7: Observed Win -->
        <div class="form-card walk-card section-win">
          <div class="walk-field-label">One Observed Win <span class="walk-optional-tag">optional</span></div>
          <textarea class="form-textarea" name="observedWin"
            placeholder="What went well in this classroom?" rows="2"></textarea>
        </div>

        <!-- Field 8: Concern or Gap -->
        <div class="form-card walk-card section-concern">
          <div class="walk-field-label">One Concern or Gap <span class="walk-optional-tag">optional</span></div>
          <textarea class="form-textarea" name="concernGap"
            placeholder="Any patterns or gaps worth noting?" rows="2"></textarea>
        </div>

        <!-- Field 9: Support Needed -->
        <div class="form-card walk-card section-support-needed">
          <div class="walk-field-label">Support Needed <span class="req">*</span></div>
          <div class="option-cards option-cards-col">
            <label class="option-card option-card-row support-needed-option">
              <input type="radio" name="supportNeeded" value="none">
              <span>🌿 None</span>
            </label>
            <label class="option-card option-card-row support-needed-option" data-urgency="talk">
              <input type="radio" name="supportNeeded" value="talk-at-weekly">
              <span>💬 Discuss at Weekly Meeting</span>
            </label>
            <label class="option-card option-card-row support-needed-option" data-urgency="help-soon">
              <input type="radio" name="supportNeeded" value="help-soon">
              <span>🛠️ Help Soon</span>
            </label>
            <label class="option-card option-card-row support-needed-option" data-urgency="urgent">
              <input type="radio" name="supportNeeded" value="urgent">
              <span>⚡ Urgent</span>
            </label>
          </div>
        </div>

        <!-- Field 10: Classroom Status -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Classroom Status <span class="req">*</span></div>
          <div class="option-cards">
            <label class="option-card" data-status="on-track">
              <input type="radio" name="classroomStatus" value="on-track">
              <span>🟢 On Track</span>
            </label>
            <label class="option-card" data-status="monitor">
              <input type="radio" name="classroomStatus" value="monitor">
              <span>🟡 Monitor</span>
            </label>
            <label class="option-card" data-status="needs-support">
              <input type="radio" name="classroomStatus" value="needs-support">
              <span>🟠 Needs Support</span>
            </label>
            <label class="option-card" data-status="immediate-follow-up">
              <input type="radio" name="classroomStatus" value="immediate-follow-up">
              <span>🔴 Follow-Up</span>
            </label>
          </div>
        </div>

        <!-- Field 11: Follow-Up Notes -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Follow-Up Notes <span class="walk-optional-tag">optional</span></div>
          <textarea class="form-textarea" name="followUpNotes"
            placeholder="Any specific follow-up actions or reminders." rows="2"></textarea>
        </div>

        <!-- Sticky save bar -->
        <div class="save-bar">
          <button type="submit" id="save-walkthrough-button" class="btn btn-primary btn-lg btn-full save-btn">Save Walkthrough</button>
        </div>

      </form>
    `;

    // Bind components
    this._bindSearchSelects(el);
    this._bindOptionCards(el);

    // Date picker: initialize and keep week label in sync
    function formatWalkthroughWeekLabel(dateValue) {
      if (!dateValue) return "";
      const date = new Date(`${dateValue}T12:00:00`);
      return `Week of ${date.toLocaleDateString("en-US", { month: "short", day: "numeric" })}`;
    }
    function updateWalkthroughWeekLabel() {
      const input = document.getElementById("walkthrough-date");
      const label = document.getElementById("walkthrough-week-label");
      if (!input || !label) return;
      label.textContent = formatWalkthroughWeekLabel(input.value);
    }
    const walkthroughDateInput = document.getElementById("walkthrough-date");
    if (walkthroughDateInput && !walkthroughDateInput.value) {
      walkthroughDateInput.value = new Date().toISOString().slice(0, 10);
    }
    updateWalkthroughWeekLabel();
    walkthroughDateInput?.addEventListener("change", () => {
      const wo = getWeekOf(walkthroughDateInput.value);
      document.getElementById("wWeekOf").value = wo;
      updateWalkthroughWeekLabel();
    });

    // Show/hide the "no roster" focus hint based on teacher + focus selection
    function updateFocusRosterHint() {
      const hint      = document.getElementById("walkFocusRosterHint");
      if (!hint) return;
      const teacherId = el.querySelector('[name="teacherId"]')?.value || "";
      const ptTeacher = PILOT_TEACHERS.find(t => t.id === teacherId);
      const hasRoster = (ptTeacher?.students || []).length > 0;
      const focus     = el.querySelector('input[name="focus"]:checked')?.value || "";
      hint.classList.toggle("hidden", !(focus === "individual-student" && !hasRoster));
    }

    // Auto-suggest classroom + filter students when teacher is selected
    el.querySelector('[data-name="teacherId"]').addEventListener("ss:change", e => {
      const teacherId = e.detail.value;
      const classroomsForTeacher = DB.getClassrooms().filter(c => c.teacherId === teacherId);
      if (classroomsForTeacher.length === 1) {
        const classroomWrap = el.querySelector('[data-name="classroomId"]');
        const hiddenInput   = classroomWrap.querySelector(".ss-value");
        if (!hiddenInput.value) {
          const opt = classroomWrap.querySelector(`.ss-option[data-value="${classroomsForTeacher[0].id}"]`);
          if (opt) opt.click();
        }
      }
      const studentSel = el.querySelector("#walkthroughStudent");
      if (studentSel) {
        const ptTeacher  = PILOT_TEACHERS.find(t => t.id === teacherId);
        const ptStudents = ptTeacher?.students || [];
        if (ptStudents.length > 0) {
          studentSel.innerHTML = '<option value="">— None / whole class —</option>' +
            ptStudents.map(s => `<option value="${escHtml(s.id)}">${escHtml(s.name)}</option>`).join("");
        } else {
          studentSel.innerHTML = '<option value="">— No roster available —</option>';
        }
      }
      updateFocusRosterHint();
    });

    // Update hint when focus changes
    el.querySelectorAll('input[name="focus"]').forEach(r =>
      r.addEventListener("change", updateFocusRosterHint)
    );

    // Form submit — named function prevents duplicate bindings across re-renders
    ensureWalkthroughSubmissionId();
    const _wForm = el.querySelector("#walkthroughForm");
    _wForm.removeEventListener("submit", _handleWalkthroughSubmit);
    _wForm.addEventListener("submit", _handleWalkthroughSubmit);
  },

  async _submitWalkthrough(pageEl) {
    if (_walkthroughInFlight) {
      console.warn("Duplicate walkthrough submission blocked.");
      return;
    }

    const form = pageEl.querySelector("#walkthroughForm");
    const fd   = new FormData(form);

    const teacherId   = form.querySelector('[name="teacherId"]').value;
    const classroomId = form.querySelector('[name="classroomId"]').value;
    const focus       = fd.get("focus");
    const engagement  = fd.get("engagementObserved");
    const support     = fd.get("supportNeeded");
    const status      = fd.get("classroomStatus");

    const errors = [];
    if (!teacherId)   errors.push("Please select a teacher.");
    if (!classroomId) errors.push("Please select a classroom.");
    if (!focus)       errors.push("Please select an observation focus.");
    if (!engagement)  errors.push("Please select engagement observed.");
    if (!support)     errors.push("Please select support needed.");
    if (!status)      errors.push("Please select classroom status.");

    if (errors.length) {
      showToast(errors[0], "error");
      const firstRequired = form.querySelector(
        '.ss-wrap[data-required="true"] .ss-value:placeholder-shown, ' +
        'input[type="radio"]:required'
      );
      if (firstRequired) firstRequired.closest(".walk-card")?.scrollIntoView({ behavior: "smooth", block: "center" });
      return;
    }

    _walkthroughInFlight = true;
    const saveButton = document.getElementById("save-walkthrough-button");
    if (saveButton) { saveButton.disabled = true; saveButton.textContent = "Saving…"; }

    try {
      const user          = this.getCurrentUser();
      const pilotTeacher  = PILOT_TEACHERS.find(t => t.id === teacherId);
      const rawStudentId  = fd.get("studentId") || "";
      const pilotStudent  = PILOT_STUDENTS.find(s => s.id === rawStudentId);
      const submissionId  = ensureWalkthroughSubmissionId();

      const responses = {
        date:               document.getElementById("walkthrough-date")?.value || fd.get("date") || new Date().toISOString().slice(0, 10),
        weekOf:             fd.get("weekOf"),
        teacherId,
        teacherName:        pilotTeacher?.name || "",
        classroomId,
        focus,
        studentId:          rawStudentId,
        studentName:        pilotStudent?.name || "",
        engagementObserved: engagement,
        supportsObserved:   fd.getAll("supportsObserved"),
        observedWin:        (fd.get("observedWin") || "").trim(),
        concernGap:         (fd.get("concernGap") || "").trim(),
        supportNeeded:      support,
        classroomStatus:    status,
        followUpNotes:      (fd.get("followUpNotes") || "").trim(),
        submittedById:      user.id,
        submittedByName:    user.name,
        submissionId,
        source:             "walkthrough"
      };

      DB.addRecord(responses);

      const classroomName = DB.getClassrooms().find(x => x.id === classroomId)?.name || "";
      const _now = new Date();
      const spPayload = {
        SubmissionID:         submissionId,
        ObservationID:        `walkthrough-${Date.now()}`,
        ObservationDate:      responses.date                || _now.toLocaleDateString(),
        ObservationTime:      _now.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }),
        ObservationTimestamp: _now.toISOString(),
        Observer:             responses.submittedByName     || "",
        Teacher:              responses.teacherName         || "",
        StudentName:          responses.studentName         || "",
        Classroom:            classroomName,
        Focus:                responses.focus               || "",
        ClassroomStatus:      responses.classroomStatus     || "",
        Engagement:           responses.engagementObserved  || "",
        SupportObserved:      responses.supportsObserved    || [],
        DisengagementReasons: [],
        SupportRequested:     responses.supportNeeded       || "",
        FollowUpNeeded:       responses.supportNeeded !== "none",
        ObservedWin:          responses.observedWin         || "",
        ConcernGap:           responses.concernGap          || "",
        FollowUpNotes:        responses.followUpNotes       || "",
        Environment:          "",
        Duration:             "",
        Notes:                "",
        Priority:             "",
        FollowUpDate:         "",
        AISummary:            "",
        AISuggestions:        ""
      };

      console.log("Submitting walkthrough once:", submissionId);

      let spSynced = false;
      let spError  = null;
      try {
        console.count("GRAPH.saveWalkthrough called");
        const result = await GRAPH.saveWalkthrough(spPayload);
        if (result?.duplicatePrevented) {
          console.warn("SharePoint duplicate prevented for:", submissionId);
        }
        spSynced = true;
        _walkthroughDraft.submissionId = null;
        REPORTS.rawWalkthroughs = [];
        REPORTS.walkthroughs    = [];
        if (APP.currentPage === "reports") REPORTS.load();
      } catch (err) {
        spError = err;
        console.error("Walkthrough SharePoint sync failed:", err);
      }

      const teachers    = DB.getTeachers();
      const classrooms  = DB.getClassrooms();
      const t           = teachers.find(x => x.id === teacherId);
      const c           = classrooms.find(x => x.id === classroomId);
      const statusLabel = CONFIG.CLASSROOM_STATUS_OPTIONS.find(o => o.value === status)?.label || status;

      const syncLine = spSynced
        ? `<p class="confirm-sync-ok">✅ Walkthrough saved locally &nbsp;·&nbsp; ✅ Synced to SharePoint</p>`
        : `<p class="confirm-sync-ok">⚠️ Saved locally</p>
           <p class="confirm-sync-err">❌ SharePoint sync failed: ${escHtml(spError?.message || String(spError))}</p>`;

      const confirmEl = pageEl.querySelector("#walkthroughConfirm");
      confirmEl.className = "save-confirm";
      confirmEl.innerHTML = `
        <div class="save-confirm-check">✓</div>
        <h3>Walkthrough Saved</h3>
        <p><strong>${t ? escHtml(t.name) : "—"}</strong> · ${c ? escHtml(c.name) : "—"}</p>
        <p>${fmtDate(responses.date + "T12:00:00")} &nbsp;·&nbsp; ${CONFIG.STATUS_EMOJI[status] || ""} ${escHtml(statusLabel)}</p>
        ${syncLine}
        <div class="confirm-actions">
          <button class="btn btn-primary btn-lg btn-full" id="anotherBtn">
            Record Another Walkthrough
          </button>
          <a href="#dashboard" class="btn btn-secondary btn-full">View Dashboard</a>
        </div>`;

      confirmEl.scrollIntoView({ behavior: "smooth", block: "start" });

      const savedDate   = form.querySelector('[name="date"]').value;
      const savedWeekOf = form.querySelector('[name="weekOf"]').value;
      form.reset();
      form.querySelector('[name="date"]').value   = savedDate;
      form.querySelector('[name="weekOf"]').value = savedWeekOf;
      form.querySelectorAll(".option-card").forEach(c => c.classList.remove("selected"));
      form.querySelectorAll(".chip").forEach(c => c.classList.remove("selected"));
      form.querySelectorAll(".ss-input").forEach(i => { i.value = ""; });
      form.querySelectorAll(".ss-value").forEach(i => { i.value = ""; });
      form.querySelectorAll(".ss-option").forEach(o => o.classList.remove("ss-active"));
      form.querySelectorAll(".ss-wrap").forEach(w => w.classList.remove("open"));

      showToast("Walkthrough saved!");

      document.getElementById("anotherBtn").addEventListener("click", () => {
        confirmEl.className = "hidden";
        confirmEl.innerHTML = "";
        ensureWalkthroughSubmissionId();
        document.getElementById("contentArea").scrollTo({ top: 0, behavior: "smooth" });
      });
    } finally {
      _walkthroughInFlight = false;
      if (saveButton) { saveButton.disabled = false; saveButton.textContent = "Save Walkthrough"; }
    }
  },

  /* ── V2 WALKTHROUGH ─────────────────────────────────────────────────────────── */

  renderWalkthroughV2() {
    const el   = document.getElementById("page-v2walkthrough");
    const user = this.getCurrentUser();

    const teachers = user.isAdmin ? PILOT_TEACHERS : PILOT_TEACHERS.filter(t => t.id === user.id);

    const draftKey  = "iepWalkthroughV2Draft";
    let draft = {};
    try { draft = JSON.parse(localStorage.getItem(draftKey) || "{}"); } catch (_) {}
    const today = new Date().toISOString().slice(0, 10);
    if (draft.date !== today) draft = { date: today, teacherId: null, classroomId: null, students: {} };

    const selectedTeacherId = draft.teacherId || (teachers.length === 1 ? teachers[0].id : null);

    const teacherOptions = teachers.map(t =>
      `<option value="${t.id}" ${t.id === selectedTeacherId ? "selected" : ""}>${t.name}</option>`
    ).join("");

    const classrooms = DB.getClassrooms();
    const classroomOptions = classrooms.map(c =>
      `<option value="${c.id}" ${c.id === draft.classroomId ? "selected" : ""}>${c.name}</option>`
    ).join("");

    const engagementOpts = ["", "high", "medium", "low"].map(v =>
      `<option value="${v}">${v ? (v.charAt(0).toUpperCase() + v.slice(1)) : "— select —"}</option>`
    ).join("");

    const supportOpts = [
      "Visual Supports", "Sensory Break", "Behavioral Support", "Peer Support",
      "Adult Proximity", "Reduced Demands", "Communication Device", "Fidget/Tool",
      "Scheduled Break", "Quiet Space"
    ];

    const supportReqOpts = [
      { v: "none",         l: "None" },
      { v: "talk-weekly",  l: "Discuss at Weekly Meeting" },
      { v: "help-soon",    l: "Help Soon" },
      { v: "urgent",       l: "Urgent" }
    ];

    const statusColors = { green: "#22c55e", yellow: "#eab308", red: "#ef4444" };
    const expandedStudents = this._v2ExpandedStudents;

    function buildStudentCard(student) {
      const sd = (draft.students || {})[student.id] || {};
      const chosen = sd.status || "";

      const statusBtns = ["green", "yellow", "red"].map(s => {
        const active = chosen === s;
        const emoji  = s === "green" ? "🟢" : s === "yellow" ? "🟡" : "🔴";
        return `<button type="button"
          class="v2-status-btn ${active ? "v2-status-active" : ""}"
          data-sid="${student.id}" data-status="${s}"
          style="--v2-color:${statusColors[s]}"
          aria-label="${s}">${emoji}</button>`;
      }).join("");

      const supChecks = supportOpts.map(opt => {
        const checked = (sd.supportsObserved || []).includes(opt) ? "checked" : "";
        return `<label class="v2-check-label">
          <input type="checkbox" class="v2-support-check" data-sid="${student.id}" value="${opt}" ${checked}> ${opt}
        </label>`;
      }).join("");

      const reqSel = supportReqOpts.map(o =>
        `<option value="${o.v}" ${(sd.supportRequested || "none") === o.v ? "selected" : ""}>${o.l}</option>`
      ).join("");

      const engSel = ["", "high", "medium", "low"].map(v =>
        `<option value="${v}" ${(sd.engagement || "") === v ? "selected" : ""}>${v ? (v.charAt(0).toUpperCase() + v.slice(1)) : "— select —"}</option>`
      ).join("");

      const expanded = !!expandedStudents[student.id];

      const detailsPanel = expanded ? `
          <div class="v2-card-details" id="v2details-${student.id}">
            <div class="v2-field">
              <label class="v2-label">Engagement</label>
              <select class="v2-engagement" data-sid="${student.id}">${engSel}</select>
            </div>
            <div class="v2-field">
              <label class="v2-label">Supports Observed</label>
              <div class="v2-check-grid">${supChecks}</div>
            </div>
            <div class="v2-field">
              <label class="v2-label">Support Requested</label>
              <select class="v2-support-req" data-sid="${student.id}">${reqSel}</select>
            </div>
            <div class="v2-field">
              <label class="v2-label">Observed Win</label>
              <textarea class="v2-textarea" data-sid="${student.id}" data-field="observedWin" rows="2">${sd.observedWin || ""}</textarea>
            </div>
            <div class="v2-field">
              <label class="v2-label">Concern / Gap</label>
              <textarea class="v2-textarea" data-sid="${student.id}" data-field="concernGap" rows="2">${sd.concernGap || ""}</textarea>
            </div>
            <div class="v2-field">
              <label class="v2-label">Notes</label>
              <textarea class="v2-textarea" data-sid="${student.id}" data-field="notes" rows="2">${sd.notes || ""}</textarea>
            </div>
            <div class="v2-field">
              <label class="v2-label">Follow-Up Notes</label>
              <textarea class="v2-textarea" data-sid="${student.id}" data-field="followUpNotes" rows="2">${sd.followUpNotes || ""}</textarea>
            </div>
            <div class="v2-field v2-field-inline">
              <label class="v2-check-label">
                <input type="checkbox" class="v2-followup-needed" data-sid="${student.id}" ${sd.followUpNeeded ? "checked" : ""}> Follow-Up Needed
              </label>
            </div>
          </div>` : "";

      return `
        <div class="v2-student-card ${chosen ? "v2-card-has-status v2-card-" + chosen : ""}" id="v2card-${student.id}">
          <div class="v2-card-header">
            <span class="v2-student-name">${student.name}</span>
            <div class="v2-status-row">
              ${statusBtns}
              <button type="button" class="v2-more-btn" data-sid="${student.id}"
                aria-expanded="${expanded}">${expanded ? "Less ▴" : "More ▾"}</button>
            </div>
          </div>
          ${detailsPanel}
        </div>`;
    }

    const selectedTeacher = teachers.find(t => t.id === selectedTeacherId);
    const students = selectedTeacher ? selectedTeacher.students : [];
    const isClassroomLevel = !!(selectedTeacher && students.length === 0);
    const studentCards = students.map(buildStudentCard).join("");
    const v2ClassroomName = classrooms.find(c => c.id === draft.classroomId)?.name
      || (selectedTeacher ? selectedTeacher.name + "'s Room" : "");
    const classroomCard = isClassroomLevel
      ? buildStudentCard({ id: "__classroom__", name: v2ClassroomName })
      : "";

    el.innerHTML = `
      <div class="walk-page-header">
        <h2 class="walk-page-title">V2 Walkthrough</h2>
        <p class="walk-page-sub">Tap a student's status, expand for details, then sync all to SharePoint.</p>
      </div>

      <div id="v2Confirm" class="hidden"></div>

      <div class="v2-meta-row">
        <div class="v2-meta-field">
          <label class="v2-label">Teacher</label>
          <select id="v2TeacherSel" class="v2-select">
            <option value="">— select teacher —</option>
            ${teacherOptions}
          </select>
        </div>
        <div class="v2-meta-field">
          <label class="v2-label">Classroom</label>
          <select id="v2ClassroomSel" class="v2-select">
            <option value="">— select classroom —</option>
            ${classroomOptions}
          </select>
        </div>
        <div class="v2-meta-field">
          <label class="v2-label">Date</label>
          <input type="date" id="v2DateInput" class="v2-select" value="${draft.date || today}">
        </div>
      </div>

      <div id="v2StudentGrid" class="v2-student-grid">
        ${!selectedTeacher
          ? '<p class="v2-empty">Select a teacher to load students.</p>'
          : isClassroomLevel
            ? classroomCard
            : studentCards}
      </div>

      <div class="v2-footer">
        <button type="button" id="v2SyncBtn" class="btn-primary" ${!selectedTeacher ? "disabled" : ""}>
          Sync to SharePoint
        </button>
        <button type="button" id="v2ClearBtn" class="btn-ghost">Clear Draft</button>
      </div>
    `;

    // ── wire up interactions ──────────────────────────────────────────

    const saveDraft = () => localStorage.setItem(draftKey, JSON.stringify(draft));

    // Teacher select
    document.getElementById("v2TeacherSel").addEventListener("change", e => {
      draft.teacherId = e.target.value || null;
      draft.students  = {};
      saveDraft();
      this.renderWalkthroughV2();
    });

    // Classroom select
    document.getElementById("v2ClassroomSel").addEventListener("change", e => {
      draft.classroomId = e.target.value || null;
      saveDraft();
    });

    // Date input
    document.getElementById("v2DateInput").addEventListener("change", e => {
      draft.date = e.target.value || today;
      saveDraft();
    });

    const grid = document.getElementById("v2StudentGrid");

    // Status buttons
    grid.addEventListener("click", e => {
      const btn = e.target.closest(".v2-status-btn");
      if (!btn) return;
      const sid    = btn.dataset.sid;
      const status = btn.dataset.status;
      if (!draft.students[sid]) draft.students[sid] = {};
      draft.students[sid].status = status;
      saveDraft();

      // Update card appearance
      const card = document.getElementById("v2card-" + sid);
      card.classList.remove("v2-card-green", "v2-card-yellow", "v2-card-red");
      card.classList.add("v2-card-has-status", "v2-card-" + status);
      card.querySelectorAll(".v2-status-btn").forEach(b =>
        b.classList.toggle("v2-status-active", b.dataset.status === status)
      );
    });

    // More/less toggle — re-render so details are only in DOM when expanded
    grid.addEventListener("click", e => {
      const btn = e.target.closest(".v2-more-btn");
      if (!btn) return;
      const sid = btn.dataset.sid;
      expandedStudents[sid] = !expandedStudents[sid];
      this.renderWalkthroughV2();
    });

    // Engagement select
    grid.addEventListener("change", e => {
      const sel = e.target.closest(".v2-engagement");
      if (!sel) return;
      const sid = sel.dataset.sid;
      if (!draft.students[sid]) draft.students[sid] = {};
      draft.students[sid].engagement = sel.value;
      saveDraft();
    });

    // Support checkboxes
    grid.addEventListener("change", e => {
      const cb = e.target.closest(".v2-support-check");
      if (!cb) return;
      const sid = cb.dataset.sid;
      if (!draft.students[sid]) draft.students[sid] = {};
      const prev = draft.students[sid].supportsObserved || [];
      draft.students[sid].supportsObserved = cb.checked
        ? [...prev, cb.value]
        : prev.filter(v => v !== cb.value);
      saveDraft();
    });

    // Support requested select
    grid.addEventListener("change", e => {
      const sel = e.target.closest(".v2-support-req");
      if (!sel) return;
      const sid = sel.dataset.sid;
      if (!draft.students[sid]) draft.students[sid] = {};
      draft.students[sid].supportRequested = sel.value;
      saveDraft();
    });

    // Textareas
    grid.addEventListener("input", e => {
      const ta = e.target.closest(".v2-textarea");
      if (!ta) return;
      const sid   = ta.dataset.sid;
      const field = ta.dataset.field;
      if (!draft.students[sid]) draft.students[sid] = {};
      draft.students[sid][field] = ta.value;
      saveDraft();
    });

    // Follow-up needed checkboxes
    grid.addEventListener("change", e => {
      const cb = e.target.closest(".v2-followup-needed");
      if (!cb) return;
      const sid = cb.dataset.sid;
      if (!draft.students[sid]) draft.students[sid] = {};
      draft.students[sid].followUpNeeded = cb.checked;
      saveDraft();
    });

    // Clear draft
    document.getElementById("v2ClearBtn").addEventListener("click", () => {
      if (!confirm("Clear all V2 Walkthrough draft data for today?")) return;
      localStorage.removeItem(draftKey);
      this.renderWalkthroughV2();
    });

    // Sync button
    document.getElementById("v2SyncBtn").addEventListener("click", async () => {
      const syncBtn   = document.getElementById("v2SyncBtn");
      const confirmEl = document.getElementById("v2Confirm");
      if (!AUTH.isAuthenticated) {
        confirmEl.className = "confirm-card";
        confirmEl.innerHTML = `<p class="confirm-sync-err">⚠️ Not signed in to SharePoint. Sign in first.</p>`;
        return;
      }

      const teacher   = teachers.find(t => t.id === draft.teacherId);
      const classroom = classrooms.find(c => c.id === draft.classroomId)?.name
        || (teacher ? teacher.name + "'s Room" : "");
      const date      = draft.date || today;

      syncBtn.disabled    = true;
      syncBtn.textContent = "Syncing…";

      let results;

      if (isClassroomLevel) {
        // Classroom-level observation (no student roster)
        const cd = (draft.students || {})["__classroom__"] || {};
        if (!cd.status) {
          syncBtn.disabled    = false;
          syncBtn.textContent = "Sync to SharePoint";
          confirmEl.className = "confirm-card";
          confirmEl.innerHTML = `<p class="confirm-sync-err">Select a status (🟢🟡🔴) for ${classroom} first.</p>`;
          confirmEl.classList.remove("hidden");
          return;
        }
        results = await Promise.allSettled([
          GRAPH.saveWalkthroughV2({
            ObservationID:    `v2-${draft.teacherId}-${Date.now()}`,
            ObservationDate:  date,
            Observer:         AUTH.pilotUser?.Name || AUTH.displayName || "",
            Teacher:          teacher?.name || "",
            Student:          "",
            Classroom:        classroom,
            Focus:            "Whole Class",
            Status:           cd.status           || "",
            Engagement:       cd.engagement       || "",
            SupportObserved:  cd.supportsObserved  || [],
            SupportRequested: cd.supportRequested  || "none",
            ObservedWin:      cd.observedWin       || "",
            ConcernGap:       cd.concernGap        || "",
            Notes:            cd.notes             || "",
            FollowUpNotes:    cd.followUpNotes     || "",
            FollowUpNeeded:   !!cd.followUpNeeded,
            SubmissionID:     crypto.randomUUID()
          })
        ]);
      } else {
        const toSync = students.filter(s => (draft.students[s.id] || {}).status);
        if (toSync.length === 0) {
          syncBtn.disabled    = false;
          syncBtn.textContent = "Sync to SharePoint";
          confirmEl.className = "confirm-card";
          confirmEl.innerHTML = `<p class="confirm-sync-err">No students have a status set. Tap 🟢🟡🔴 first.</p>`;
          confirmEl.classList.remove("hidden");
          return;
        }
        results = await Promise.allSettled(
          toSync.map(s => {
            const sd = draft.students[s.id] || {};
            return GRAPH.saveWalkthroughV2({
              ObservationID:    `v2-${s.id}-${Date.now()}`,
              ObservationDate:  date,
              Observer:         AUTH.pilotUser?.Name || AUTH.displayName || "",
              Teacher:          teacher?.name || "",
              Student:          s.name,
              Classroom:        classroom,
              Status:           sd.status          || "",
              Engagement:       sd.engagement      || "",
              SupportObserved:  sd.supportsObserved || [],
              SupportRequested: sd.supportRequested || "none",
              ObservedWin:      sd.observedWin     || "",
              ConcernGap:       sd.concernGap      || "",
              Notes:            sd.notes           || "",
              FollowUpNotes:    sd.followUpNotes   || "",
              FollowUpNeeded:   !!sd.followUpNeeded,
              SubmissionID:     crypto.randomUUID()
            });
          })
        );
      }

      const ok   = results.filter(r => r.status === "fulfilled").length;
      const fail = results.filter(r => r.status === "rejected");

      syncBtn.disabled    = false;
      syncBtn.textContent = "Sync to SharePoint";

      confirmEl.className = "confirm-card";
      confirmEl.classList.remove("hidden");
      if (fail.length === 0) {
        REPORTS.rawWalkthroughs = [];
        REPORTS.walkthroughs    = [];
        if (APP.currentPage === "reports") REPORTS.load();
        const label = isClassroomLevel ? "classroom observation" : `student${ok !== 1 ? "s" : ""}`;
        confirmEl.innerHTML = `<p class="confirm-sync-ok">✅ ${ok} ${label} synced to SharePoint.</p>`;
      } else {
        const errMsgs = fail.map(r => escHtml(r.reason?.message || String(r.reason))).join("<br>");
        confirmEl.innerHTML = `
          <p class="confirm-sync-ok">✅ ${ok} synced</p>
          <p class="confirm-sync-err">❌ ${fail.length} failed:<br>${errMsgs}</p>`;
      }
      confirmEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
    });
  },

  /* ── DAILY PULSE V2 ─────────────────────────────────────────────────────────── */

  _pulseDraft: null,

  _getPulseRoster() {
    const user = this.getCurrentUser();
    return user.isAdmin
      ? STUDENT_ROSTER.getDailyPulseEnabled()
      : STUDENT_ROSTER.getDailyPulseForTeacher(user.teacherName || user.name);
  },

  // Shared by renderPulse()/renderPaceLog(): kicks off a roster load if one
  // isn't already in flight, and re-renders the given page once it settles.
  _ensureRosterLoading(pageName) {
    if (STUDENT_ROSTER.loading) return;
    STUDENT_ROSTER.refresh().then(() => {
      if (this.currentPage === pageName) this.navigate(pageName);
    });
  },

  _rosterStatusHtml(title) {
    if (STUDENT_ROSTER.error) {
      return `
        <div class="walk-page-header"><h2 class="walk-page-title">${escHtml(title)}</h2></div>
        <div class="warning-banner">
          <div>
            <strong>Student roster could not be loaded.</strong><br>
            Please refresh or contact your administrator.
          </div>
        </div>`;
    }
    return `
      <div class="walk-page-header"><h2 class="walk-page-title">${escHtml(title)}</h2></div>
      <div class="card"><p class="empty-state-text">Loading student roster…</p></div>`;
  },

  _initPulseDraft(studentName) {
    this._pulseDraft = {
      studentName:      studentName || "",
      attendanceStatus: null,
      overallStatus:    "",
      selectedCategories: [],
      categoryResponses: {
        academicProgress:     null,
        behavior:             null,
        communication:        null,
        socialSkills:         null,
        selfRegulation:       null,
        transitionDifficulty: null,
        iepGoalProgress:      null
      },
      quickNote: "",
      date:      new Date().toISOString().slice(0, 10),
      recordId:  `pulse-${crypto.randomUUID()}`
    };
  },

  renderPulse() {
    const el = document.getElementById("page-pulse");

    if (!STUDENT_ROSTER.loaded && !STUDENT_ROSTER.error) {
      el.innerHTML = this._rosterStatusHtml("Daily Pulse");
      this._ensureRosterLoading("pulse");
      return;
    }
    if (STUDENT_ROSTER.error) {
      el.innerHTML = this._rosterStatusHtml("Daily Pulse");
      return;
    }

    const roster = this._getPulseRoster();
    const user   = this.getCurrentUser();

    if (!this._pulseDraft) this._initPulseDraft("");
    if (this._pulseDraft.studentName && !roster.some(s => s.name === this._pulseDraft.studentName)) {
      this._initPulseDraft("");
    }

    if (roster.length === 0) {
      const emptyMsg = user.isAdmin
        ? "No Daily Pulse-enabled students are currently available."
        : "No Daily Pulse-enabled students are assigned to you.";
      el.innerHTML = `
        <div class="walk-page-header"><h2 class="walk-page-title">Daily Pulse</h2></div>
        <div class="card"><p class="empty-state-text">${escHtml(emptyMsg)}</p></div>`;
      return;
    }

    const draft     = this._pulseDraft;
    const today     = draft.date;
    const allPulses = DB.getDailyPulses();
    const completedToday = new Set(
      allPulses
        .filter(p => (p.date || (p.timestamp || "").slice(0, 10)) === today)
        .map(p => p.student)
    );

    const total      = roster.length;
    const completed  = roster.filter(s => completedToday.has(s.name)).length;
    const remaining  = total - completed;
    const currentIdx = draft.studentName
      ? roster.findIndex(s => s.name === draft.studentName)
      : -1;
    const studentNum = currentIdx >= 0 ? currentIdx + 1 : 0;

    const alreadyDoneToday  = !!(draft.studentName && completedToday.has(draft.studentName));
    const showAttendance    = !!draft.studentName;
    const showPresent       = draft.attendanceStatus === "present";
    const showAbsentConfirm = draft.attendanceStatus === "absent";

    el.innerHTML = `
      <div class="walk-page-header">
        <h2 class="walk-page-title">Daily Pulse</h2>
        ${total > 0 ? `<div class="pulse-progress-bar">
          <span class="pulse-progress-text">${studentNum > 0 ? `Student ${studentNum} of ${total}` : `${total} students`}</span>
          <span class="pulse-progress-detail">${completed} completed &bull; ${remaining} remaining</span>
        </div>` : ""}
      </div>

      <!-- Student -->
      <div class="form-card walk-card">
        <div class="walk-field-label">Student <span class="req">*</span></div>
        <select class="form-select" id="pulseStudentSelect">
          <option value="">— Choose student —</option>
          ${roster.map(s => {
            const done = completedToday.has(s.name) ? " ✓" : "";
            return `<option value="${escHtml(s.name)}"${draft.studentName === s.name ? " selected" : ""}>${escHtml(s.name)}${user.isAdmin ? " — " + escHtml(s.teacherName) : ""}${done}</option>`;
          }).join("")}
        </select>
        ${alreadyDoneToday ? `<p class="pulse-already-done-notice">Already completed today — submitting again will create a second entry.</p>` : ""}
      </div>

      <!-- Attendance -->
      <div class="form-card walk-card${showAttendance ? "" : " pulse-section-hidden"}" id="pulseAttendanceCard">
        <div class="walk-field-label">Attendance <span class="req">*</span></div>
        <p class="pulse-section-help">Was the student present today?</p>
        <div class="attendance-options" role="radiogroup" aria-label="Student attendance">
          <button type="button" class="attendance-option attendance-present${draft.attendanceStatus === "present" ? " is-selected" : ""}"
            data-attendance="present" aria-pressed="${draft.attendanceStatus === "present"}">Present</button>
          <button type="button" class="attendance-option attendance-absent${draft.attendanceStatus === "absent" ? " is-selected" : ""}"
            data-attendance="absent" aria-pressed="${draft.attendanceStatus === "absent"}">Absent</button>
        </div>
        <div id="pulseAbsentConfirm" class="${showAbsentConfirm ? "pulse-absent-confirm" : "pulse-section-hidden"}">
          <p>Mark <strong>${escHtml(draft.studentName)}</strong> absent and continue?</p>
          <div class="pulse-absent-confirm-btns">
            <button type="button" class="btn btn-danger" id="pulseConfirmAbsentBtn">Confirm Absent</button>
            <button type="button" class="btn btn-secondary" id="pulseCancelAbsentBtn">Cancel</button>
          </div>
        </div>
      </div>

      <!-- Present-only sections -->
      <div id="pulsePresenceSections"${showPresent ? "" : ` class="pulse-section-hidden"`}>

        <!-- Overall Day -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Overall Day <span class="req">*</span></div>
          <div class="option-cards pulse-status-cards">
            <label class="option-card pulse-status-card${draft.overallStatus === "great" ? " selected" : ""}" data-pulse="great">
              <input type="radio" name="pulseStatus" value="great"${draft.overallStatus === "great" ? " checked" : ""}>
              <span class="pulse-status-emoji">😊</span>
              <span>Great Day</span>
            </label>
            <label class="option-card pulse-status-card${draft.overallStatus === "struggles" ? " selected" : ""}" data-pulse="struggles">
              <input type="radio" name="pulseStatus" value="struggles"${draft.overallStatus === "struggles" ? " checked" : ""}>
              <span class="pulse-status-emoji">😐</span>
              <span>Some Struggles</span>
            </label>
            <label class="option-card pulse-status-card${draft.overallStatus === "concern" ? " selected" : ""}" data-pulse="concern">
              <input type="radio" name="pulseStatus" value="concern"${draft.overallStatus === "concern" ? " checked" : ""}>
              <span class="pulse-status-emoji">☹️</span>
              <span>Significant Concern</span>
            </label>
          </div>
        </div>

        <!-- Focus Areas -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Focus Areas <span class="walk-optional-tag">optional</span></div>
          <p class="pulse-section-help">Select one or more to add detail.</p>
          <div class="pulse-category-grid">
            ${Object.entries(DAILY_PULSE_SCALES).map(([key, scale]) => {
              const isSel  = draft.selectedCategories.includes(key);
              const isDone = !!(draft.categoryResponses[key]?.rating);
              return `<button type="button"
                class="pulse-category-btn${isSel ? " is-selected" : ""}${isDone ? " is-complete" : ""}"
                data-category-key="${key}" aria-pressed="${isSel}">
                ${escHtml(scale.label)}${isDone ? ' <span class="pulse-cat-check">&#10003;</span>' : ""}
              </button>`;
            }).join("")}
          </div>
          <div id="pulseScalePanels">
            ${draft.selectedCategories.map(key => this._buildScalePanelHtml(key)).join("")}
          </div>
        </div>

        <!-- Quick Note -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Quick Note <span class="walk-optional-tag">optional</span></div>
          <textarea class="form-textarea" id="pulseQuickNote"
            placeholder="Any additional observations or context." rows="2">${escHtml(draft.quickNote)}</textarea>
        </div>

        <div id="pulseFormError" class="pulse-section-hidden pulse-form-error"></div>

        <div class="save-bar">
          <button type="button" class="btn btn-primary btn-lg btn-full save-btn" id="pulseSubmitBtn">
            Submit Daily Pulse
          </button>
        </div>

      </div>`;

    this._bindPulseEvents(el);
    this._bindOptionCards(el);
  },

  _buildScalePanelHtml(categoryKey) {
    const scale    = DAILY_PULSE_SCALES[categoryKey];
    if (!scale) return "";
    const response  = this._pulseDraft.categoryResponses[categoryKey] || { rating: "", multiSelect: {}, details: {} };
    const hasRating = !!response.rating;

    const choicesHtml = scale.choices.map(c => {
      const isSel = response.rating === c.value;
      return `<button type="button"
        class="scale-choice-btn scale-choice-${c.tone}${isSel ? " is-selected has-check" : ""}"
        data-category="${categoryKey}" data-value="${escHtml(c.value)}"
        aria-pressed="${isSel}">${escHtml(c.label)}</button>`;
    }).join("");

    return `<section class="pulse-scale-panel" data-scale-panel="${categoryKey}">
      <header class="pulse-scale-header">
        <div>
          <h3>${escHtml(scale.label)}</h3>
          <p>${escHtml(scale.prompt)}</p>
        </div>
        <span class="scale-completion-badge${hasRating ? " visible" : ""}">Complete</span>
      </header>
      <div class="scale-choice-grid">${choicesHtml}</div>
      <div class="scale-detail-sections-wrapper">${this._buildScaleDetailHtml(scale, response, categoryKey)}</div>
    </section>`;
  },

  _buildScaleDetailHtml(scale, response, categoryKey) {
    let html = "";
    if (scale.multiSelectFields) {
      scale.multiSelectFields.forEach(field => {
        const selected = response.multiSelect?.[field.key] || [];
        html += `<div class="scale-detail-section">
          <div class="scale-detail-label">${escHtml(field.label)}</div>
          <div class="scale-chip-grid">
            ${field.options.map(opt => {
              const isSel = selected.includes(opt);
              return `<button type="button"
                class="pulse-detail-chip${isSel ? " is-selected" : ""}"
                data-category="${categoryKey}" data-field="${escHtml(field.key)}" data-option="${escHtml(opt)}"
                aria-pressed="${isSel}">${escHtml(opt)}</button>`;
            }).join("")}
          </div>
        </div>`;
      });
    }
    if (scale.detailFields) {
      scale.detailFields.forEach(field => {
        if (field.showForValues && !field.showForValues.includes(response.rating)) return;
        if (field.showWhenOption) {
          const anySelected = Object.values(response.multiSelect || {}).some(arr =>
            Array.isArray(arr) && arr.includes(field.showWhenOption)
          );
          if (!anySelected) return;
        }
        const val = response.details?.[field.key] || "";
        if (field.type === "textarea") {
          html += `<div class="scale-detail-section">
            <div class="scale-detail-label">${escHtml(field.label)} <span class="scale-optional-hint">(optional)</span></div>
            <textarea class="scale-text-field form-textarea"
              data-category="${categoryKey}" data-field="${escHtml(field.key)}"
              rows="2">${escHtml(val)}</textarea>
          </div>`;
        } else {
          html += `<div class="scale-detail-section">
            <div class="scale-detail-label">${escHtml(field.label)} <span class="scale-optional-hint">(optional)</span></div>
            <input type="text" class="scale-text-field form-input"
              data-category="${categoryKey}" data-field="${escHtml(field.key)}"
              value="${escHtml(val)}">
          </div>`;
        }
      });
    }
    return html;
  },

  _bindPulseEvents(el) {
    // Student selection
    el.querySelector("#pulseStudentSelect")?.addEventListener("change", e => {
      this._initPulseDraft(e.target.value);
      this.renderPulse();
      document.getElementById("page-pulse")
        .querySelector("#pulseAttendanceCard")
        ?.scrollIntoView({ behavior: "smooth", block: "nearest" });
    });

    // Attendance buttons
    el.querySelectorAll(".attendance-option").forEach(btn => {
      btn.addEventListener("click", () => {
        if (!this._pulseDraft.studentName) { showToast("Please select a student first.", "error"); return; }
        const status = btn.dataset.attendance;
        this._pulseDraft.attendanceStatus = status;

        el.querySelectorAll(".attendance-option").forEach(b => {
          const isSel = b.dataset.attendance === status;
          b.classList.toggle("is-selected", isSel);
          b.setAttribute("aria-pressed", String(isSel));
        });

        const absentConfirm    = el.querySelector("#pulseAbsentConfirm");
        const presenceSections = el.querySelector("#pulsePresenceSections");
        if (status === "absent") {
          absentConfirm.className    = "pulse-absent-confirm";
          presenceSections.className = "pulse-section-hidden";
        } else {
          absentConfirm.className    = "pulse-section-hidden";
          presenceSections.className = "";
        }
      });
    });

    // Absent confirm / cancel
    el.querySelector("#pulseConfirmAbsentBtn")?.addEventListener("click", () => this._submitPulseAbsent());
    el.querySelector("#pulseCancelAbsentBtn")?.addEventListener("click", () => {
      this._pulseDraft.attendanceStatus = null;
      el.querySelectorAll(".attendance-option").forEach(b => {
        b.classList.remove("is-selected");
        b.setAttribute("aria-pressed", "false");
      });
      el.querySelector("#pulseAbsentConfirm").className = "pulse-section-hidden";
    });

    // Overall day radio
    el.querySelectorAll('input[name="pulseStatus"]').forEach(radio => {
      radio.addEventListener("change", () => {
        this._pulseDraft.overallStatus = radio.value;
      });
    });

    // Category toggle buttons
    el.querySelectorAll(".pulse-category-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const key        = btn.dataset.categoryKey;
        const panelsEl   = el.querySelector("#pulseScalePanels");
        const isSelected = this._pulseDraft.selectedCategories.includes(key);

        if (isSelected) {
          this._pulseDraft.selectedCategories = this._pulseDraft.selectedCategories.filter(k => k !== key);
          btn.classList.remove("is-selected");
          btn.setAttribute("aria-pressed", "false");
          panelsEl.querySelector(`[data-scale-panel="${key}"]`)?.remove();
        } else {
          this._pulseDraft.selectedCategories.push(key);
          btn.classList.add("is-selected");
          btn.setAttribute("aria-pressed", "true");
          panelsEl.insertAdjacentHTML("beforeend", this._buildScalePanelHtml(key));
          const newPanel = panelsEl.querySelector(`[data-scale-panel="${key}"]`);
          if (newPanel) {
            this._bindScalePanelEvents(newPanel);
            newPanel.scrollIntoView({ behavior: "smooth", block: "nearest" });
          }
        }
      });
    });

    // Scale panels already rendered (from draft state)
    el.querySelectorAll(".pulse-scale-panel").forEach(panel => this._bindScalePanelEvents(panel));

    // Quick note
    el.querySelector("#pulseQuickNote")?.addEventListener("input", e => {
      this._pulseDraft.quickNote = e.target.value;
    });

    // Submit
    el.querySelector("#pulseSubmitBtn")?.addEventListener("click", () => this._submitPulsePresent(el));
  },

  _bindScalePanelEvents(panelEl) {
    panelEl.querySelectorAll(".scale-choice-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const categoryKey = btn.dataset.category;
        const value       = btn.dataset.value;
        const response    = this._pulseDraft.categoryResponses[categoryKey] ||
          { rating: "", multiSelect: {}, details: {} };
        response.rating = value;
        this._pulseDraft.categoryResponses[categoryKey] = response;

        panelEl.querySelectorAll(".scale-choice-btn").forEach(b => {
          const isSel = b.dataset.value === value;
          b.classList.toggle("is-selected", isSel);
          b.classList.toggle("has-check", isSel);
          b.setAttribute("aria-pressed", String(isSel));
        });

        panelEl.querySelector(".scale-completion-badge")?.classList.add("visible");

        const catBtn = document.querySelector(`.pulse-category-btn[data-category-key="${categoryKey}"]`);
        if (catBtn) {
          catBtn.classList.add("is-complete");
          if (!catBtn.querySelector(".pulse-cat-check")) {
            catBtn.insertAdjacentHTML("beforeend", ' <span class="pulse-cat-check">&#10003;</span>');
          }
        }

        const wrapper = panelEl.querySelector(".scale-detail-sections-wrapper");
        if (wrapper) {
          wrapper.innerHTML = this._buildScaleDetailHtml(DAILY_PULSE_SCALES[categoryKey], response, categoryKey);
          this._bindScalePanelDetailEvents(wrapper);
        }
      });
    });

    const wrapper = panelEl.querySelector(".scale-detail-sections-wrapper");
    if (wrapper) this._bindScalePanelDetailEvents(wrapper);
  },

  _bindScalePanelDetailEvents(wrapper) {
    wrapper.querySelectorAll(".pulse-detail-chip").forEach(chip => {
      chip.addEventListener("click", () => {
        const categoryKey = chip.dataset.category;
        const fieldKey    = chip.dataset.field;
        const optionVal   = chip.dataset.option;
        const response    = this._pulseDraft.categoryResponses[categoryKey] ||
          { rating: "", multiSelect: {}, details: {} };
        if (!response.multiSelect) response.multiSelect = {};
        if (!response.multiSelect[fieldKey]) response.multiSelect[fieldKey] = [];

        const arr = response.multiSelect[fieldKey];
        const idx = arr.indexOf(optionVal);
        if (idx >= 0) {
          arr.splice(idx, 1);
          chip.classList.remove("is-selected");
          chip.setAttribute("aria-pressed", "false");
        } else {
          arr.push(optionVal);
          chip.classList.add("is-selected");
          chip.setAttribute("aria-pressed", "true");
        }
        this._pulseDraft.categoryResponses[categoryKey] = response;

        // Re-render detail section to show/hide conditional fields
        const panelEl = chip.closest(".pulse-scale-panel");
        if (panelEl) {
          const w = panelEl.querySelector(".scale-detail-sections-wrapper");
          if (w) {
            w.innerHTML = this._buildScaleDetailHtml(DAILY_PULSE_SCALES[categoryKey], response, categoryKey);
            this._bindScalePanelDetailEvents(w);
          }
        }
      });
    });

    wrapper.querySelectorAll(".scale-text-field").forEach(field => {
      field.addEventListener("input", () => {
        const categoryKey = field.dataset.category;
        const fieldKey    = field.dataset.field;
        const response    = this._pulseDraft.categoryResponses[categoryKey] ||
          { rating: "", multiSelect: {}, details: {} };
        if (!response.details) response.details = {};
        response.details[fieldKey] = field.value;
        this._pulseDraft.categoryResponses[categoryKey] = response;
      });
    });
  },

  _pulseInFlight: false,

  async _submitPulseAbsent() {
    if (this._pulseInFlight) {
      console.warn("Duplicate Daily Pulse submission blocked.");
      return;
    }
    const draft         = this._pulseDraft;
    const user          = this.getCurrentUser();
    const rosterStudent = STUDENT_ROSTER.findByName(draft.studentName);
    const teacherName   = rosterStudent?.teacherName || user.teacherName || user.name || "";
    const teacherId      = rosterStudent?.teacherId   || user.id   || "";
    const studentId      = rosterStudent?.id          || "";

    const submittedByName  = USER_CONTEXT.isViewingAsTeacher
      ? (USER_CONTEXT.actualUser?.name || user.name)
      : user.name;
    const submittedByEmail = USER_CONTEXT.isViewingAsTeacher
      ? (USER_CONTEXT.actualUser?.email || AUTH.account?.username || "")
      : (AUTH.account?.username || "");

    this._pulseInFlight = true;
    const confirmBtn = document.getElementById("pulseConfirmAbsentBtn");
    if (confirmBtn) { confirmBtn.disabled = true; confirmBtn.textContent = "Saving…"; }

    try {
      const entry = DB.addDailyPulse({
        id:                draft.recordId,
        student:           draft.studentName,
        studentId,
        date:              draft.date,
        attendanceStatus:  "absent",
        pulseStatus:       "",
        categories:        [],
        categoryResponses: {},
        note:              "",
        supportLevel:      "",
        teacherId,
        teacherName,
        submittedById:    user.id,
        submittedByName,
        source:           "daily-pulse"
      });

      try {
        await GRAPH.saveDailyPulse({
          RecordID:          entry.id,
          Date:              draft.date,
          Teacher:           teacherName,
          Student:           draft.studentName,
          AttendanceStatus:  "absent",
          Present:           false,
          Status:            "",
          Categories:        [],
          CategoryResponses: {},
          QuickNote:         "",
          SubmittedBy:       submittedByName,
          SubmittedByEmail:  submittedByEmail,
          SubmittedAt:       entry.timestamp
        });
        showToast(`Marked ${draft.studentName} absent.`);
        this._advancePulseToNextStudent();
      } catch (err) {
        // Local copy is saved (DB.addDailyPulse above), but per production
        // reliability requirements do not mark this student complete or
        // advance until SharePoint confirms — draft.recordId, draft.studentName,
        // and draft.attendanceStatus ("absent") are all left untouched so a
        // retry reuses the same stable id (duplicate-safe) and the same
        // Absent selection.
        console.error("Daily Pulse SharePoint sync failed:", err.message || String(err));
        showSyncError(err.message || String(err));
      }
    } finally {
      this._pulseInFlight = false;
      if (confirmBtn) { confirmBtn.disabled = false; confirmBtn.textContent = "Confirm Absent"; }
    }
  },

  async _submitPulsePresent(el) {
    if (this._pulseInFlight) {
      console.warn("Duplicate Daily Pulse submission blocked.");
      return;
    }

    const draft = this._pulseDraft;

    if (!draft.studentName)   { this._showPulseError(el, "Please select a student."); return; }
    if (!draft.overallStatus) { this._showPulseError(el, "Please select how today went."); return; }

    for (const catKey of draft.selectedCategories) {
      const resp = draft.categoryResponses[catKey];
      if (!resp?.rating) {
        this._showPulseError(el, `Please select a response for ${DAILY_PULSE_SCALES[catKey]?.label || catKey}.`);
        el.querySelector(`[data-scale-panel="${catKey}"]`)?.scrollIntoView({ behavior: "smooth", block: "start" });
        return;
      }
    }

    const user          = this.getCurrentUser();
    const rosterStudent = STUDENT_ROSTER.findByName(draft.studentName);
    const teacherName    = rosterStudent?.teacherName || user.teacherName || user.name || "";
    const teacherId      = rosterStudent?.teacherId   || user.id   || "";
    const studentId      = rosterStudent?.id          || "";

    const submittedByName  = USER_CONTEXT.isViewingAsTeacher
      ? (USER_CONTEXT.actualUser?.name || user.name)
      : user.name;
    const submittedByEmail = USER_CONTEXT.isViewingAsTeacher
      ? (USER_CONTEXT.actualUser?.email || AUTH.account?.username || "")
      : (AUTH.account?.username || "");

    const noteEl = el.querySelector("#pulseQuickNote");
    if (noteEl) draft.quickNote = noteEl.value;

    const responses = structuredClone(draft.categoryResponses);
    const statusLabel = { great: "Great Day", struggles: "Some Struggles", concern: "Significant Concern" }[draft.overallStatus] || draft.overallStatus;

    this._pulseInFlight = true;
    const submitBtn = el.querySelector("#pulseSubmitBtn");
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = "Saving…"; }

    try {
      const entry = DB.addDailyPulse({
        id:                draft.recordId,
        student:           draft.studentName,
        studentId,
        date:              draft.date,
        attendanceStatus:  "present",
        pulseStatus:       draft.overallStatus,
        categories:        [...draft.selectedCategories],
        categoryResponses: responses,
        note:              draft.quickNote || "",
        supportLevel:      "",
        teacherId,
        teacherName,
        submittedById:    user.id,
        submittedByName,
        source:           "daily-pulse"
      });

      try {
        await GRAPH.saveDailyPulse({
          RecordID:          entry.id,
          Date:              draft.date,
          Teacher:           teacherName,
          Student:           draft.studentName,
          AttendanceStatus:  "present",
          Present:           true,
          Status:            statusLabel,
          Categories:        [...draft.selectedCategories],
          CategoryResponses: responses,
          QuickNote:         draft.quickNote || "",
          SubmittedBy:       submittedByName,
          SubmittedByEmail:  submittedByEmail,
          SubmittedAt:       entry.timestamp
        });
        showToast("Daily pulse saved and synced to SharePoint!");
        this._advancePulseToNextStudent();
      } catch (err) {
        // Local copy is saved (DB.addDailyPulse above), but per production
        // reliability requirements do not clear/advance this student's form
        // until SharePoint confirms — let the teacher see the failure and retry.
        console.error("Daily Pulse SharePoint sync failed:", err.message || String(err));
        showSyncError(err.message || String(err));
      }
    } finally {
      this._pulseInFlight = false;
      if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = "Submit Daily Pulse"; }
    }
  },

  _showPulseError(el, msg) {
    const errEl = el?.querySelector("#pulseFormError");
    if (errEl) {
      errEl.textContent = msg;
      errEl.className   = "pulse-form-error";
      errEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
    } else {
      showToast(msg, "error");
    }
  },

  _advancePulseToNextStudent() {
    const roster     = this._getPulseRoster();
    const currentIdx = roster.findIndex(s => s.name === this._pulseDraft.studentName);
    const next       = roster[currentIdx + 1] || null;

    if (next) {
      this._initPulseDraft(next.name);
      this.renderPulse();
      showToast(`Ready for ${next.name}.`);
      document.getElementById("page-pulse").scrollIntoView({ behavior: "smooth", block: "start" });
    } else {
      this._pulseDraft = null;
      const el = document.getElementById("page-pulse");
      // Teachers no longer have a dashboard route — only show that link to admins.
      const dashboardLink = USER_CONTEXT.effectiveRole === "teacher"
        ? ""
        : `<a href="#teacherDashboard" class="btn btn-primary btn-lg">View Dashboard</a>`;
      el.innerHTML = `
        <div class="walk-page-header">
          <h2 class="walk-page-title">Daily Pulse</h2>
        </div>
        <div class="pulse-completion-state">
          <div class="pulse-completion-icon">&#10003;</div>
          <h3 class="pulse-completion-title">Daily Pulse Complete</h3>
          <p class="pulse-completion-sub">All students have been recorded for today.</p>
          <div class="confirm-actions">
            <button class="btn btn-secondary btn-lg" id="pulseStartAgainBtn">Start Again</button>
            ${dashboardLink}
          </div>
        </div>`;
      el.querySelector("#pulseStartAgainBtn")?.addEventListener("click", () => {
        this._initPulseDraft("");
        this.renderPulse();
      });
    }
  },

  /* ── PACE LOG ───────────────────────────────────────────────────────────────── */

  _paceScmUsed: null,

  async renderPaceAdmin(forceRefresh = false) {
    const el = document.getElementById("page-pace");

    // Navigation authorization already blocks teacher deep links. Keep a
    // second check at the renderer boundary so PACE student history never
    // becomes accessible through a direct method call or stale hash state.
    if (!AUTH.isAdmin || USER_CONTEXT.isViewingAsTeacher) {
      el.innerHTML = `<div class="page-header"><h2>PACE Log</h2></div>
        <div class="warning-banner">Administrator access is required to view PACE analytics and history.</div>`;
      return;
    }

    el.innerHTML = `<div class="walk-page-header">
        <h2 class="walk-page-title">PACE Log</h2>
        <p class="walk-page-sub">Completed visit history from IEP_Pace_Visits.</p>
      </div>
      <div class="card pace-admin-loading"><div class="reports-loading"><span class="spinner"></span> Loading PACE visits from SharePoint…</div></div>`;

    try {
      if (forceRefresh) PACE_ADMIN.invalidate();
      const data = await PACE_ADMIN.load(forceRefresh);
      if (this.currentPage !== "pace") return;
      this._paceAdminData = data;

      const values = key => [...new Set(data.visits.flatMap(v => Array.isArray(v[key]) ? v[key] : [v[key]]).filter(Boolean))]
        .sort((a, b) => a.localeCompare(b));
      const students    = values("student");
      const specialists = values("specialists");
      const teachers    = values("teacherCameFrom");
      const rooms       = values("paceRoom");
      const reasons     = values("reasons");
      const optionHtml = (items, label) => `<option value="">${escHtml(label)}</option>${items.map(item =>
        `<option value="${escHtml(item)}">${escHtml(item)}</option>`).join("")}`;

      const today = new Date().toISOString().slice(0, 10);
      const start = new Date(`${today}T12:00:00`);
      start.setDate(start.getDate() - 29);
      const from = start.toISOString().slice(0, 10);
      const missing = [
        ["PACE Room", data.availability.paceRoom],
        ["SCM Used", data.availability.scm],
        ["Teacher Came From", data.availability.teacherCameFrom]
      ].filter(([, available]) => !available).map(([name]) => name);

      el.innerHTML = `
        <div class="walk-page-header pace-admin-header">
          <div>
            <h2 class="walk-page-title">PACE Log</h2>
            <p class="walk-page-sub">Investigate completed PACE visits from the authoritative SharePoint list.</p>
          </div>
          <button class="btn btn-secondary btn-sm" id="paceAdminRefresh">↻ Refresh from SharePoint</button>
        </div>

        <div class="info-banner" style="margin-bottom:16px">
          <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18" style="flex-shrink:0"><path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd"/></svg>
          <div>PACE Room Tracker records visits; this administrator view is read-only. Select a student name to review individual history.</div>
        </div>

        ${missing.length ? `<div class="pace-schema-note"><strong>Schema-aware view:</strong> ${escHtml(missing.join(", "))} ${missing.length === 1 ? "is" : "are"} not available in the current list, so related filters and statistics are hidden.</div>` : ""}

        <div class="filter-bar pace-filter-bar">
          <div class="pace-filter-heading"><div class="card-title">Filters</div><button class="btn btn-secondary btn-sm" id="paceClearFilters">Clear</button></div>
          <div class="filter-row pace-filter-grid">
            <div class="form-field"><label class="form-label" for="pf-from">From</label><input class="form-input" type="date" id="pf-from" value="${from}"></div>
            <div class="form-field"><label class="form-label" for="pf-to">To</label><input class="form-input" type="date" id="pf-to" value="${today}"></div>
            <div class="form-field"><label class="form-label" for="pf-student">Student</label><select class="form-select" id="pf-student">${optionHtml(students, "All Students")}</select></div>
            ${specialists.length ? `<div class="form-field"><label class="form-label" for="pf-specialist">Behavior Specialist</label><select class="form-select" id="pf-specialist">${optionHtml(specialists, "All Specialists")}</select></div>` : ""}
            ${teachers.length && data.availability.teacherCameFrom ? `<div class="form-field"><label class="form-label" for="pf-teacher">Teacher Came From</label><select class="form-select" id="pf-teacher">${optionHtml(teachers, "All Teachers")}</select></div>` : ""}
            ${rooms.length && data.availability.paceRoom ? `<div class="form-field"><label class="form-label" for="pf-room">PACE Room</label><select class="form-select" id="pf-room">${optionHtml(rooms, "All Rooms")}</select></div>` : ""}
            ${reasons.length ? `<div class="form-field"><label class="form-label" for="pf-reason">Reason</label><select class="form-select" id="pf-reason">${optionHtml(reasons, "All Reasons")}</select></div>` : ""}
            ${data.availability.scm ? `<div class="form-field"><label class="form-label" for="pf-scm">SCM</label><select class="form-select" id="pf-scm"><option value="">All</option><option value="yes">Used</option><option value="no">Not Used</option></select></div>` : ""}
            ${data.visits.some(v => v.durationMinutes !== null) ? `<div class="form-field"><label class="form-label" for="pf-duration">Duration</label><select class="form-select" id="pf-duration"><option value="">Any Duration</option><option value="under-45">Under 45 minutes</option><option value="45-plus">45+ minutes</option></select></div>` : ""}
            <div class="form-field pace-search-field"><label class="form-label" for="pf-search">Search</label><input class="form-input" type="search" id="pf-search" placeholder="Search student, reason, support, or notes…"></div>
          </div>
        </div>

        <div id="paceAdminResults"></div>

        <div id="pace-history-modal" class="report-modal hidden" role="dialog" aria-modal="true" aria-labelledby="pace-history-title">
          <div class="report-modal-backdrop" id="pace-history-backdrop"></div>
          <div class="report-modal-panel pace-history-panel">
            <div class="report-modal-header"><h3 class="report-modal-title" id="pace-history-title">Student PACE History</h3><button class="report-modal-close" id="pace-history-close" aria-label="Close">✕</button></div>
            <div class="report-modal-body" id="pace-history-body"></div>
          </div>
        </div>`;

      const apply = () => this._renderPaceAdminResults();
      el.querySelectorAll("#pf-from,#pf-to,#pf-student,#pf-specialist,#pf-teacher,#pf-room,#pf-reason,#pf-scm,#pf-duration")
        .forEach(input => input.addEventListener("change", apply));
      el.querySelector("#pf-search")?.addEventListener("input", apply);
      el.querySelector("#paceClearFilters").addEventListener("click", () => {
        el.querySelectorAll(".pace-filter-bar input,.pace-filter-bar select").forEach(input => { input.value = ""; });
        apply();
      });
      el.querySelector("#paceAdminRefresh").addEventListener("click", () => this.renderPaceAdmin(true));
      const closeHistory = () => el.querySelector("#pace-history-modal")?.classList.add("hidden");
      el.querySelector("#pace-history-close").addEventListener("click", closeHistory);
      el.querySelector("#pace-history-backdrop").addEventListener("click", closeHistory);
      apply();
    } catch (err) {
      console.error("PACE administrator read failed:", err.message || String(err));
      el.innerHTML = `<div class="walk-page-header"><h2 class="walk-page-title">PACE Log</h2><p class="walk-page-sub">Completed visit history from IEP_Pace_Visits.</p></div>
        <div class="reports-error"><strong>PACE visits could not be loaded from SharePoint.</strong><p>${escHtml(err.message || String(err))}</p><button class="btn btn-secondary btn-sm" id="paceRetryLoad">Try Again</button></div>`;
      el.querySelector("#paceRetryLoad")?.addEventListener("click", () => this.renderPaceAdmin(true));
    }
  },

  _getPaceAdminFilters() {
    const value = id => document.getElementById(id)?.value || "";
    return {
      from: value("pf-from"), to: value("pf-to"), student: value("pf-student"),
      specialist: value("pf-specialist"), teacherCameFrom: value("pf-teacher"),
      paceRoom: value("pf-room"), reason: value("pf-reason"), scm: value("pf-scm"),
      duration: value("pf-duration"), search: value("pf-search")
    };
  },

  _renderPaceAdminResults() {
    const container = document.getElementById("paceAdminResults");
    if (!container || !this._paceAdminData) return;
    const filtered = PACE_ADMIN.filterVisits(this._paceAdminData.visits, this._getPaceAdminFilters())
      .sort(PACE_ADMIN.sortNewest);
    const completed = filtered.filter(v => v.isCompleted);
    const summary = PACE_ADMIN.summarize(filtered);
    const shown = filtered.slice(0, 200);
    const roomLabel = room => {
      const info = PACE_ADMIN.roomInfo(room);
      return info ? (info.hallway ? `${info.label} · ${info.hallway}` : info.label) : "—";
    };
    const showSpecialist = this._paceAdminData.availability.specialist && this._paceAdminData.visits.some(v => v.specialists.length);
    const showTeacher    = this._paceAdminData.availability.teacherCameFrom && this._paceAdminData.visits.some(v => v.teacherCameFrom);
    const showRoom       = this._paceAdminData.availability.paceRoom && this._paceAdminData.visits.some(v => v.paceRoom);
    const showScm        = this._paceAdminData.availability.scm && this._paceAdminData.visits.some(v => v.scmUsed !== null);

    container.innerHTML = `
      <div class="pace-results-summary">
        <div><strong>${filtered.length}</strong> record${filtered.length !== 1 ? "s" : ""}</div>
        <div><strong>${summary.uniqueStudents}</strong> student${summary.uniqueStudents !== 1 ? "s" : ""}</div>
        ${summary.averageDuration !== null ? `<div><strong>${summary.averageDuration} min</strong> average</div>` : ""}
        ${filtered.length !== completed.length ? `<div class="pace-legacy-count">${filtered.length - completed.length} older/incomplete</div>` : ""}
      </div>
      ${shown.length === 0 ? `<div class="card"><div class="empty-state"><div class="empty-icon">🔎</div><p>No PACE visits match these filters.</p></div></div>` : `
        <div class="card pace-results-card">
          <div class="table-wrap"><table class="reports-table pace-admin-table"><thead><tr>
            <th>Date / Time</th><th>Student</th><th>Duration</th><th>Reason / Support</th>
            ${showSpecialist ? "<th>Specialist</th>" : ""}
            ${showTeacher ? "<th>Teacher Came From</th>" : ""}
            ${showRoom ? "<th>Room</th>" : ""}
            ${showScm ? "<th>SCM</th>" : ""}
          </tr></thead><tbody>${shown.map(visit => `<tr>
            <td><strong>${visit.date ? escHtml(fmtDate(visit.date + "T12:00:00")) : "—"}</strong><div class="pace-cell-sub">${escHtml(visit.timeIn || "—")} → ${escHtml(visit.timeOut || "—")}</div></td>
            <td>${visit.student ? `<button class="pace-student-link" data-student="${escHtml(visit.student)}">${escHtml(visit.student)}</button>` : "—"}${!visit.isCompleted ? `<div><span class="pace-record-badge">Older / incomplete</span></div>` : ""}</td>
            <td>${visit.durationMinutes !== null ? `${visit.durationMinutes} min` : "—"}</td>
            <td><div class="pace-tag-row">${visit.reasons.map(reason => `<span class="pace-tag">${escHtml(reason)}</span>`).join("")}</div>${visit.supports.length ? `<div class="pace-cell-sub">Support: ${escHtml(visit.supports.join(", "))}</div>` : ""}${visit.notes ? `<details class="pace-notes"><summary>Notes</summary><p>${escHtml(visit.notes)}</p></details>` : ""}</td>
            ${showSpecialist ? `<td>${visit.specialists.length ? escHtml(visit.specialists.join(", ")) : "—"}</td>` : ""}
            ${showTeacher ? `<td>${escHtml(visit.teacherCameFrom || "—")}</td>` : ""}
            ${showRoom ? `<td>${escHtml(roomLabel(visit.paceRoom))}</td>` : ""}
            ${showScm ? `<td>${visit.scmUsed === null ? "—" : visit.scmUsed ? `<span class="badge badge-red">Yes</span>` : "No"}</td>` : ""}
          </tr>`).join("")}</tbody></table></div>
          ${filtered.length > shown.length ? `<p class="pace-result-limit">Showing the newest ${shown.length} of ${filtered.length} matching records. Narrow the filters to investigate further.</p>` : ""}
        </div>`}`;

    container.querySelectorAll(".pace-student-link").forEach(button =>
      button.addEventListener("click", () => this._openPaceStudentHistory(button.dataset.student))
    );
  },

  _openPaceStudentHistory(student) {
    if (!AUTH.isAdmin || USER_CONTEXT.isViewingAsTeacher || !this._paceAdminData) return;
    const modal = document.getElementById("pace-history-modal");
    const body  = document.getElementById("pace-history-body");
    if (!modal || !body) return;
    const today = new Date().toISOString().slice(0, 10);
    const history = PACE_ADMIN.studentHistory(this._paceAdminData.visits, student, today);
    const availability = this._paceAdminData.availability;
    const a = {
      ...availability,
      specialist: availability.specialist && [...history.recentVisits, ...history.openVisits].some(v => v.specialists.length)
    };
    const metric = (label, value) => `<div class="stat-card"><div class="stat-label">${escHtml(label)}</div><div class="stat-value">${escHtml(value)}</div></div>`;

    body.innerHTML = `
      <div class="pace-history-heading"><h3>${escHtml(student)}</h3><p>Last 30 days · ${escHtml(fmtDate(history.from + "T12:00:00"))}–${escHtml(fmtDate(history.to + "T12:00:00"))}</p></div>
      <div class="stat-grid pace-history-stats">
        ${metric("PACE Visits", history.visits)}
        ${history.totalMinutes !== null ? metric("Total PACE Time", `${history.totalMinutes} min`) : ""}
        ${history.averageDuration !== null ? metric("Average Visit", `${history.averageDuration} min`) : ""}
        ${a.scm && history.scmEvents !== null ? metric("SCM Events", history.scmEvents) : ""}
        ${a.reason && history.needsBreak !== null ? metric("Needs a Break", history.needsBreak) : ""}
      </div>
      <div class="pace-history-highlights">
        ${history.mostCommonReason ? `<div><span>Most Common Reason</span><strong>${escHtml(history.mostCommonReason[0])}</strong></div>` : ""}
        ${a.teacherCameFrom && history.mostCommonTeacher ? `<div><span>Most Common Teacher Came From</span><strong>${escHtml(history.mostCommonTeacher[0])}</strong></div>` : ""}
      </div>
      ${history.openVisits.length ? `
      <h4 class="pace-history-recent-title">Open Visit${history.openVisits.length !== 1 ? "s" : ""}</h4>
      <div class="pace-visit-list pace-visit-list-open">
        ${history.openVisits.map(visit => this._paceVisitRowHtml(visit, a, today)).join("")}
      </div>` : ""}
      <h4 class="pace-history-recent-title">Recent Visits</h4>
      ${history.recentVisits.length
        ? `<div class="pace-visit-list">${history.recentVisits.map(visit => this._paceVisitRowHtml(visit, a, today)).join("")}</div>`
        : `<div class="empty-state"><p>No completed visits in this period.</p></div>`}`;
    document.getElementById("pace-history-title").textContent = `${student} — PACE History`;
    modal.classList.remove("hidden");
  },

  // One expandable "Recent Visits" row: a compact summary line (the
  // <summary>) plus the full visit-detail panel beneath it (Part 3),
  // rendered inline via native <details> — no second modal, no JS toggle
  // wiring needed. `today` distinguishes a currently-open visit from an
  // old record that simply never received a Time Out.
  _paceVisitRowHtml(visit, availability, today) {
    const dateLabel = visit.date ? fmtDate(visit.date + "T12:00:00") : "—";
    const reasonSummary = visit.reasons.join(", ") || "—";
    const specialistSummary = availability.specialist && visit.specialists.length
      ? escHtml(visit.specialists.join(", ")) : "";
    const statusBadge = !visit.isCompleted
      ? `<span class="pace-record-badge${visit.date === today ? " pace-badge-open-now" : ""}">${visit.date === today ? "Currently in PACE" : "Open / incomplete"}</span>`
      : "";

    return `
      <details class="pace-visit-row">
        <summary class="pace-visit-summary">
          <span class="pace-visit-summary-date">${escHtml(dateLabel)}</span>
          <span class="pace-visit-summary-duration">${visit.durationMinutes !== null ? `${visit.durationMinutes} min` : (visit.isCompleted ? "—" : "In progress")}</span>
          <span class="pace-visit-summary-reason">${escHtml(reasonSummary)}</span>
          ${specialistSummary ? `<span class="pace-visit-summary-specialist">${specialistSummary}</span>` : ""}
          ${statusBadge}
          <span class="pace-visit-summary-toggle" aria-hidden="true">
            <span class="pace-visit-summary-toggle-closed">View Details</span>
            <span class="pace-visit-summary-toggle-open">Hide Details</span>
          </span>
        </summary>
        ${this._paceVisitDetailHtml(visit, availability)}
      </details>`;
  },

  // Part 3 — the full visit-detail panel. Every field is gated by the live
  // schema's `availability` flags (never fabricated) and every per-row gap
  // renders a plain placeholder rather than being silently omitted.
  _paceVisitDetailHtml(visit, availability) {
    const room = availability.paceRoom ? PACE_ADMIN.roomInfo(visit.paceRoom) : null;
    const field = (label, value) => `
      <div class="pace-detail-field"><span>${escHtml(label)}</span><strong>${value}</strong></div>`;
    const tagField = (label, items) => `
      <div class="pace-detail-field pace-detail-field-tags"><span>${escHtml(label)}</span>
        ${items.length ? `<div class="pace-tag-row">${items.map(i => `<span class="pace-tag">${escHtml(i)}</span>`).join("")}</div>` : "<strong>—</strong>"}
      </div>`;

    const timeRange = (visit.timeIn || visit.timeOut)
      ? `${visit.timeIn ? escHtml(fmtClockTime(visit.timeIn)) : "—"} → ${visit.timeOut ? escHtml(fmtClockTime(visit.timeOut)) : (visit.isCompleted ? "—" : "In progress")}`
      : "Not recorded";

    const scmValue = !availability.scm ? null
      : visit.scmUsed === null ? "Not recorded" : (visit.scmUsed ? "Yes" : "No");

    return `
      <div class="pace-visit-details">
        <div class="pace-detail-grid">
          ${field("Date", escHtml(visit.date ? fmtDate(visit.date + "T12:00:00") : "Not recorded"))}
          ${availability.paceRoom ? field("PACE Room", escHtml(room ? (room.hallway ? `${room.label} · ${room.hallway}` : room.label) : "Not recorded")) : ""}
          ${field("Time", timeRange)}
          ${field("Duration", visit.durationMinutes !== null ? `${visit.durationMinutes} minutes` : (visit.isCompleted ? "Not recorded" : "Visit is still open"))}
          ${availability.specialist ? field(
            `Behavior Interventionist${visit.specialists.length > 1 ? "s" : ""}`,
            visit.specialists.length ? visit.specialists.map(escHtml).join("<br>") : "Not recorded"
          ) : ""}
          ${availability.teacherCameFrom ? field("Teacher Came From", escHtml(visit.teacherCameFrom || "Not recorded")) : ""}
          ${scmValue !== null ? field("SCM", escHtml(scmValue)) : ""}
        </div>
        ${tagField("Reason", visit.reasons)}
        ${tagField("Support / Intervention", visit.supports)}
        <div class="pace-visit-notes-block">
          <span class="pace-detail-field-label">Notes</span>
          ${visit.notes
            ? `<div class="pace-visit-notes-text">${escHtml(visit.notes)}</div>`
            : `<p class="pace-visit-notes-empty">No notes recorded.</p>`}
        </div>
      </div>`;
  },

  // Legacy operational form retained as an internal rollback path only. The
  // PACE route now renders renderPaceAdmin(), and teachers cannot reach it.
  _renderLegacyPaceLog() {
    const el = document.getElementById("page-pace");

    if (!STUDENT_ROSTER.loaded && !STUDENT_ROSTER.error) {
      el.innerHTML = this._rosterStatusHtml("PACE Log");
      this._ensureRosterLoading("pace");
      return;
    }
    if (STUDENT_ROSTER.error) {
      el.innerHTML = this._rosterStatusHtml("PACE Log");
      return;
    }

    this._paceScmUsed     = null;
    this._paceSubmissionId = null; // fresh form → fresh stable submission id, generated on first submit attempt
    const user = this.getCurrentUser();
    const paceLogs = DB.getPaceLogs();

    // PACE is intentionally NOT restricted to the signed-in teacher's own
    // classroom — it's a shared intervention workflow, so both admins and
    // teachers see every active + PACE-enabled student, school-wide.
    // Grouping by teacher below is presentation only; it does not remove
    // any otherwise-eligible student.
    const paceRoster = STUDENT_ROSTER.getPaceEnabled();

    if (paceRoster.length === 0) {
      el.innerHTML = `
        <div class="walk-page-header"><h2 class="walk-page-title">PACE Log</h2></div>
        <div class="card"><p class="empty-state-text">No PACE-enabled students are currently available.</p></div>`;
      return;
    }

    // Open logs are likewise shown to everyone — a teacher may need to close
    // out a log for a student outside their own classroom.
    const openLogs = paceLogs.filter(p => p.isOpen);

    const byTeacher = new Map();
    paceRoster.forEach(s => {
      const key = s.teacherName || "Unassigned";
      if (!byTeacher.has(key)) byTeacher.set(key, []);
      byTeacher.get(key).push(s);
    });
    const studentOptionsHtml = [...byTeacher.entries()].map(([teacherName, students]) => `
      <optgroup label="${escHtml(teacherName)}">
        ${students.map(s => `<option value="${escHtml(s.id)}">${escHtml(s.name)} — ${escHtml(teacherName)}</option>`).join("")}
      </optgroup>`).join("");

    const now     = new Date();
    const today   = now.toISOString().slice(0, 10);
    const timeNow = now.toTimeString().slice(0, 5);

    el.innerHTML = `
      <div class="walk-page-header">
        <h2 class="walk-page-title">PACE Log</h2>
        <p class="walk-page-sub">Track PACE room support — behavior, intervention, and return status.</p>
      </div>

      <div class="info-banner" style="margin-bottom:20px">
        <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18" style="flex-shrink:0"><path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd"/></svg>
        <div>PACE entries may include protected student information. Keep entries brief, factual, and support-focused.</div>
      </div>

      ${openLogs.length > 0 ? `
      <div class="card" style="margin-bottom:24px">
        <div class="card-title">Open PACE Logs <span class="pace-open-count">${openLogs.length}</span></div>
        <div id="openPaceLogs">
          ${openLogs.map(log => {
            const elapsedMin = Math.round((Date.now() - new Date(log.timestamp).getTime()) / 60000);
            const roomLabel  = log.paceRoom === "pace-room-1" ? "PACE Room 1" : "PACE Room 2";
            return `<div class="pace-open-card">
              <div class="pace-open-header">
                <div>
                  <div class="pace-open-student">${escHtml(log.studentName)}</div>
                  <div class="pace-open-room">${escHtml(roomLabel)}</div>
                </div>
                <div class="pace-open-time-info">
                  <div class="pace-open-in">In: ${escHtml(log.timeIn)}</div>
                  <div class="pace-still-open">still open · ${elapsedMin}m ago</div>
                </div>
              </div>
              <button class="btn btn-primary btn-sm pace-close-btn" data-log-id="${escHtml(log.id)}">Close Log</button>
            </div>`;
          }).join("")}
        </div>
      </div>` : ""}

      <div class="walk-page-header" style="margin-top:0;padding-bottom:4px">
        <h3 style="font-size:17px;font-weight:700;color:var(--text-primary)">New PACE Entry</h3>
      </div>

      <div id="paceConfirm" class="hidden"></div>

      <form id="paceForm" novalidate>

        <!-- PACE Room -->
        <div class="form-card walk-card">
          <div class="walk-field-label">PACE Room <span class="req">*</span></div>
          <div class="option-cards">
            <label class="option-card">
              <input type="radio" name="paceRoom" value="pace-room-1">
              <span>PACE Room 1</span>
            </label>
            <label class="option-card">
              <input type="radio" name="paceRoom" value="pace-room-2">
              <span>PACE Room 2</span>
            </label>
          </div>
        </div>

        <!-- Student -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Student <span class="req">*</span></div>
          <select class="form-select" name="studentId" id="paceStudent">
            <option value="">— Choose student —</option>
            ${studentOptionsHtml}
          </select>
        </div>

        <!-- Date -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Date <span class="req">*</span></div>
          <input type="date" class="form-input" name="date" value="${today}">
        </div>

        <!-- Time In -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Time In <span class="req">*</span></div>
          <input type="time" class="form-input" name="timeIn" id="paceTimeIn" value="${timeNow}">
        </div>

        <!-- Time Out -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Time Out <span class="walk-optional-tag">optional — fill when student leaves</span></div>
          <input type="time" class="form-input" name="timeOut" id="paceTimeOut">
          <div class="pace-duration" id="paceDuration">Duration: still open</div>
        </div>

        <!-- Behavior -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Behavior that led to PACE support <span class="req">*</span></div>
          <div class="chip-grid">
            ${CONFIG.PACE_BEHAVIOR_OPTIONS.filter(b => b !== "Other").map(b => `
              <label class="chip"><input type="checkbox" name="behaviors" value="${escHtml(b)}"><span>${escHtml(b)}</span></label>`).join("")}
            <label class="chip"><input type="checkbox" name="behaviors" value="Other" id="behaviorOtherChip"><span>Other</span></label>
          </div>
          <div id="behaviorOtherWrap" style="display:none;margin-top:10px">
            <input type="text" class="form-input" name="behaviorOther" placeholder="Describe other behavior…" maxlength="200">
          </div>
        </div>

        <!-- Interventions -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Intervention(s) utilized during time held <span class="req">*</span></div>
          <div class="chip-grid">
            ${CONFIG.PACE_INTERVENTION_OPTIONS.filter(i => i !== "Other").map(i => `
              <label class="chip"><input type="checkbox" name="interventions" value="${escHtml(i)}"><span>${escHtml(i)}</span></label>`).join("")}
            <label class="chip"><input type="checkbox" name="interventions" value="Other" id="interventionOtherChip"><span>Other</span></label>
          </div>
          <div id="interventionOtherWrap" style="display:none;margin-top:10px">
            <input type="text" class="form-input" name="interventionOther" placeholder="Describe other intervention…" maxlength="200">
          </div>
        </div>

        <!-- Return Status -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Student return status <span class="req">*</span></div>
          <div class="option-cards option-cards-col">
            ${CONFIG.PACE_RETURN_OPTIONS.filter(o => o.value !== "other").map(o => `
              <label class="option-card option-card-row">
                <input type="radio" name="returnStatus" value="${escHtml(o.value)}">
                <span>${escHtml(o.label)}</span>
              </label>`).join("")}
            <label class="option-card option-card-row">
              <input type="radio" name="returnStatus" value="other" id="returnStatusOther">
              <span>Other</span>
            </label>
          </div>
          <div id="returnOtherWrap" style="display:none;margin-top:10px">
            <input type="text" class="form-input" name="returnOther" placeholder="Describe return outcome…" maxlength="200">
          </div>
        </div>

        <!-- SCM Used -->
        <div class="form-card walk-card pace-scm-section">
          <div class="walk-field-label">SCM Used? <span class="req">*</span></div>
          <p class="form-help">Was Safe Crisis Management used during this PACE visit?</p>
          <div class="pace-scm-options" role="radiogroup" aria-label="SCM used">
            <button type="button" class="pace-scm-option${this._paceScmUsed === false ? " is-selected" : ""}"
              data-scm-value="false" aria-pressed="${this._paceScmUsed === false}">No</button>
            <button type="button" class="pace-scm-option${this._paceScmUsed === true ? " is-selected" : ""}"
              data-scm-value="true" aria-pressed="${this._paceScmUsed === true}">Yes</button>
          </div>
        </div>

        <!-- Notes -->
        <div class="form-card walk-card">
          <div class="walk-field-label">Notes <span class="walk-optional-tag">optional</span></div>
          <textarea class="form-textarea" name="notes" rows="3"
            placeholder="Briefly note what happened, what helped, or what should be followed up on."></textarea>
        </div>

        <div class="save-bar">
          <button type="submit" class="btn btn-primary btn-lg btn-full save-btn" id="paceSubmitBtn">Save PACE Entry</button>
        </div>

      </form>`;

    this._bindOptionCards(el);

    // SCM button wiring
    el.querySelectorAll(".pace-scm-option").forEach(btn => {
      btn.addEventListener("click", () => {
        this._paceScmUsed = btn.dataset.scmValue === "true";
        el.querySelectorAll(".pace-scm-option").forEach(opt => {
          const selected = opt.dataset.scmValue === String(this._paceScmUsed);
          opt.classList.toggle("is-selected", selected);
          opt.setAttribute("aria-pressed", String(selected));
        });
      });
    });

    // Duration calculator
    const timeInEl   = el.querySelector("#paceTimeIn");
    const timeOutEl  = el.querySelector("#paceTimeOut");
    const durationEl = el.querySelector("#paceDuration");
    const calcDuration = () => {
      const tin  = timeInEl?.value;
      const tout = timeOutEl?.value;
      if (!tin || !tout) { durationEl.textContent = "Duration: still open"; return; }
      const [ih, im] = tin.split(":").map(Number);
      const [oh, om] = tout.split(":").map(Number);
      let mins = (oh * 60 + om) - (ih * 60 + im);
      if (mins < 0) mins += 24 * 60;
      durationEl.textContent = `Duration: ${mins} minute${mins !== 1 ? "s" : ""}`;
    };
    timeInEl?.addEventListener("change", calcDuration);
    timeOutEl?.addEventListener("change", calcDuration);

    // Conditional "Other" text inputs
    el.querySelector("#behaviorOtherChip")?.addEventListener("change", e => {
      el.querySelector("#behaviorOtherWrap").style.display = e.target.checked ? "" : "none";
    });
    el.querySelector("#interventionOtherChip")?.addEventListener("change", e => {
      el.querySelector("#interventionOtherWrap").style.display = e.target.checked ? "" : "none";
    });
    el.querySelectorAll('[name="returnStatus"]').forEach(r => {
      r.addEventListener("change", () => {
        el.querySelector("#returnOtherWrap").style.display = (r.value === "other" && r.checked) ? "" : "none";
      });
    });

    // Close log buttons
    el.querySelectorAll(".pace-close-btn").forEach(btn => {
      btn.addEventListener("click", () => this._closePaceLog(btn.dataset.logId));
    });

    el.querySelector("#paceForm").addEventListener("submit", e => {
      e.preventDefault();
      this._submitPaceLog(el);
    });
  },

  _paceInFlight: false,

  async _submitPaceLog(pageEl) {
    if (this._paceInFlight) {
      console.warn("Duplicate PACE submission blocked.");
      return;
    }

    const form  = pageEl.querySelector("#paceForm");
    const fd    = new FormData(form);
    const user  = this.getCurrentUser();

    const studentId     = fd.get("studentId")   || "";
    const paceRoom      = fd.get("paceRoom")     || "";
    const date          = fd.get("date")         || "";
    const timeIn        = fd.get("timeIn")       || "";
    const behaviors     = fd.getAll("behaviors");
    const interventions = fd.getAll("interventions");
    const returnStatus  = fd.get("returnStatus") || "";

    if (!paceRoom)               { showToast("Please select a PACE room.", "error"); return; }
    if (!studentId)              { showToast("Please select a student.", "error"); return; }
    if (!date)                   { showToast("Please enter a date.", "error"); return; }
    if (!timeIn)                 { showToast("Please enter time in.", "error"); return; }
    if (behaviors.length === 0)    { showToast("Please select at least one behavior.", "error"); return; }
    if (interventions.length === 0) { showToast("Please select at least one intervention.", "error"); return; }
    if (!returnStatus)           { showToast("Please select student return status.", "error"); return; }
    if (this._paceScmUsed === null) { showToast("Please indicate whether SCM was used.", "error"); return; }

    // Stable per-form submission id — generated once, reused across a retry
    // after a failed sync, so a re-submit of the same form is recognized as
    // the same PACE visit rather than creating a duplicate SharePoint row.
    if (!this._paceSubmissionId) {
      this._paceSubmissionId = (typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : `pace-${Date.now()}`;
    }
    const submissionId = this._paceSubmissionId;

    const rosterStudent = STUDENT_ROSTER.find(studentId);

    this._paceInFlight = true;
    const submitBtn = pageEl.querySelector("#paceSubmitBtn");
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = "Saving…"; }

    try {
      const entry = DB.addPaceLog({
        id:                submissionId,
        date,
        studentId,
        studentName:       rosterStudent?.name        || "",
        teacherId:         rosterStudent?.teacherId   || "",
        teacherName:       rosterStudent?.teacherName || "",
        paceRoom,
        timeIn,
        timeOut:           fd.get("timeOut")           || "",
        behaviors,
        behaviorOther:     (fd.get("behaviorOther")    || "").trim(),
        interventions,
        interventionOther: (fd.get("interventionOther") || "").trim(),
        returnStatus,
        returnOther:       (fd.get("returnOther")       || "").trim(),
        scmUsed:           this._paceScmUsed,
        notes:             (fd.get("notes")             || "").trim(),
        submittedById:     user.id,
        submittedByName:   user.name
      });

      const paceRoomLabel = paceRoom === "pace-room-1" ? "PACE Room 1" : "PACE Room 2";
      const returnLabel   = CONFIG.PACE_RETURN_OPTIONS.find(o => o.value === returnStatus)?.label || returnStatus;

      let spError = null;
      try {
        await GRAPH.savePaceVisit(entry);
      } catch (err) {
        spError = err;
        console.error("PACE SharePoint sync failed:", err.message || String(err));
      }

      if (spError) {
        // Do not clear the form until the SharePoint save succeeds — the
        // local copy is safe (DB.addPaceLog above), but keep the entered
        // data on screen so the teacher can retry without re-typing.
        showSyncError(spError.message || String(spError));
        return;
      }

      const confirmEl = pageEl.querySelector("#paceConfirm");
      // Teachers no longer have a dashboard route — only show that link to admins.
      const dashboardLink = USER_CONTEXT.effectiveRole === "teacher"
        ? ""
        : `<a href="#dashboard" class="btn btn-secondary btn-full">View Dashboard</a>`;
      confirmEl.className = "save-confirm";
      confirmEl.innerHTML = `
        <div class="save-confirm-check">✓</div>
        <h3>PACE Entry Saved</h3>
        <p><strong>${escHtml(rosterStudent?.name || studentId)}</strong> · ${escHtml(paceRoomLabel)}</p>
        <p>${escHtml(fmtDate(date + "T12:00:00"))} · ${escHtml(returnLabel)}</p>
        ${entry.isOpen ? `<p style="color:var(--orange);font-size:13px;margin-top:4px">Log is open — time out not yet recorded.</p>` : ""}
        <p class="confirm-sync-ok">✅ Saved locally &nbsp;·&nbsp; ✅ Synced to SharePoint</p>
        <div class="confirm-actions">
          <button class="btn btn-primary btn-lg btn-full" id="anotherPaceBtn">Record Another PACE Entry</button>
          ${dashboardLink}
        </div>`;

      confirmEl.scrollIntoView({ behavior: "smooth", block: "start" });
      this._paceScmUsed      = null;
      this._paceSubmissionId = null;
      form.reset();
      form.querySelectorAll(".option-card").forEach(c => c.classList.remove("selected"));
      form.querySelectorAll(".chip").forEach(c => c.classList.remove("selected"));
      form.querySelectorAll(".pace-scm-option").forEach(o => {
        o.classList.remove("is-selected");
        o.setAttribute("aria-pressed", "false");
      });
      pageEl.querySelector("#behaviorOtherWrap").style.display    = "none";
      pageEl.querySelector("#interventionOtherWrap").style.display = "none";
      pageEl.querySelector("#returnOtherWrap").style.display       = "none";
      pageEl.querySelector("#paceDuration").textContent            = "Duration: still open";
      showToast("PACE entry saved and synced to SharePoint!");

      document.getElementById("anotherPaceBtn").addEventListener("click", () => {
        confirmEl.className = "hidden";
        confirmEl.innerHTML = "";
        document.getElementById("contentArea").scrollTo({ top: 0, behavior: "smooth" });
        this._renderLegacyPaceLog();
      });
    } finally {
      this._paceInFlight = false;
      if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = "Save PACE Entry"; }
    }
  },

  _closePaceLog(logId) {
    const timeOut = new Date().toTimeString().slice(0, 5);
    const updated = DB.updatePaceLog(logId, { timeOut });
    if (!updated) { showToast("Log not found.", "error"); return; }
    const durText = updated.durationMinutes !== null ? `${updated.durationMinutes} min` : "unknown";
    showToast(`Log closed. Duration: ${durText}`);
    this._renderLegacyPaceLog();
  },

  /* ── TEACHER DASHBOARD ───────────────────────────────────────────────────────── */

  renderTeacherDashboard() {
    const el = document.getElementById("page-teacherDashboard");
    const teacher = USER_CONTEXT.effectiveTeacher;

    if (!teacher) {
      el.innerHTML = `<div class="walk-page-header"><h2 class="walk-page-title">My Classroom Dashboard</h2></div>
        <div class="info-banner">No teacher context is available.</div>`;
      return;
    }

    el.innerHTML = `<div class="walk-page-header">
      <h2 class="walk-page-title">My Classroom Dashboard</h2>
      <p class="walk-page-sub">Loading your classroom data…</p>
    </div>`;

    const pt = PILOT_TEACHERS.find(t => t.id === teacher.userId);
    const roster = (pt?.students || []).map(s => ({ ...s, teacherName: teacher.teacherName }));

    if (roster.length === 0) {
      el.innerHTML = `<div class="walk-page-header">
        <h2 class="walk-page-title">My Classroom Dashboard</h2>
      </div>
      <div class="card"><p class="empty-state-text">No active students are currently assigned to your classroom.</p></div>`;
      return;
    }

    const allPulses     = DB.getDailyPulses();
    const teacherPulses = allPulses.filter(p => recordBelongsToTeacher(p, teacher, roster));

    const today      = new Date().toISOString().slice(0, 10);
    const weekStart  = getWeekOf(today);
    const pulseDate   = p => p.date || (p.timestamp || "").slice(0, 10);
    const todayPulses = teacherPulses.filter(p => pulseDate(p) === today);
    const weekPulses  = teacherPulses.filter(p => getWeekOf(pulseDate(p)) === weekStart);

    const greenToday  = todayPulses.filter(p => p.pulseStatus === "great").length;
    const yellowToday = todayPulses.filter(p => p.pulseStatus === "struggles").length;
    const redToday    = todayPulses.filter(p => p.pulseStatus === "concern").length;

    const studentsWithEntry = new Set(todayPulses.map(p => (p.student || "").trim().toLowerCase()));
    const noEntryToday = roster.filter(s => !studentsWithEntry.has(s.name.trim().toLowerCase()));

    const statusEmoji = { great: "😊", struggles: "😐", concern: "☹️" };
    const statusLabel = { great: "Great Day", struggles: "Some Struggles", concern: "Significant Concern" };

    const studentCards = roster.map(student => {
      const sName = student.name.trim().toLowerCase();
      const history = teacherPulses
        .filter(p => (p.student || "").trim().toLowerCase() === sName)
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
      const latest   = history[0];
      const weekCount = history.filter(p => getWeekOf(pulseDate(p)) === weekStart).length;
      const emoji     = latest ? (statusEmoji[latest.pulseStatus] || "—") : "—";
      const label     = latest ? (statusLabel[latest.pulseStatus] || latest.pulseStatus) : "No entries yet";
      const lastDate  = latest ? fmtDateShort(latest.timestamp) : "—";
      return `<div class="teacher-student-card">
        <div class="teacher-student-name">${escHtml(student.name)}</div>
        <div class="teacher-student-status">${emoji} ${escHtml(label)}</div>
        <div class="teacher-student-meta">Last entry: ${escHtml(lastDate)} · ${weekCount} this week</div>
      </div>`;
    }).join("");

    const viewAsHeader = USER_CONTEXT.isViewingAsTeacher
      ? `<div class="teacher-view-label">Teacher View: ${escHtml(teacher.displayName)}</div>`
      : "";

    el.innerHTML = `
      <div class="walk-page-header">
        ${viewAsHeader}
        <h2 class="walk-page-title">My Classroom Dashboard</h2>
        <p class="walk-page-sub">${escHtml(teacher.classroom || `${teacher.teacherName}'s Classroom`)}</p>
      </div>

      <div class="stat-grid">
        <div class="stat-card"><div class="stat-value">${roster.length}</div><div class="stat-label">My Students</div></div>
        <div class="stat-card"><div class="stat-value">${todayPulses.length}</div><div class="stat-label">Pulse Entries Today</div></div>
        <div class="stat-card stat-green"><div class="stat-value">${greenToday}</div><div class="stat-label">On Track Today</div></div>
        <div class="stat-card stat-yellow"><div class="stat-value">${yellowToday}</div><div class="stat-label">Monitor Today</div></div>
        <div class="stat-card stat-red"><div class="stat-value">${redToday}</div><div class="stat-label">Needs Support Today</div></div>
        <div class="stat-card"><div class="stat-value">${weekPulses.length}</div><div class="stat-label">This Week's Entries</div></div>
      </div>

      <div class="card" style="margin-top:24px">
        <div class="card-title">My Students</div>
        <div class="teacher-student-grid">${studentCards}</div>
      </div>

      ${noEntryToday.length > 0 ? `
      <div class="card" style="margin-top:16px">
        <div class="card-title">No Entry Today</div>
        <div class="teacher-no-entry-row">
          ${noEntryToday.map(s => `<span class="teacher-no-entry-chip">${escHtml(s.name)}</span>`).join("")}
        </div>
      </div>` : ""}

      ${teacherPulses.length === 0 ? `
      <div class="card" style="margin-top:16px">
        <p class="empty-state-text">No Daily Pulse entries have been recorded yet.</p>
      </div>` : ""}
    `;
  },

  /* ── STUDENT CHECK-IN ────────────────────────────────────────────────────────── */

  renderStudentCheckIn() {
    const el   = document.getElementById("page-checkin");
    const user = this.getCurrentUser();
    const today = new Date().toISOString().slice(0, 10);

    let studentOptionsHtml;
    if (user.isAdmin) {
      studentOptionsHtml = PILOT_TEACHERS.map(t => `
        <optgroup label="${escHtml(t.name)}">
          ${t.students.map(s => `<option value="${escHtml(s.id)}">${escHtml(s.name)} — ${escHtml(t.name)}</option>`).join("")}
        </optgroup>`).join("");
    } else {
      studentOptionsHtml = PILOT_STUDENTS.filter(s => s.teacherId === user.id)
        .map(s => `<option value="${escHtml(s.id)}">${escHtml(s.name)}</option>`).join("");
    }

    const checkinBtns = CONFIG.CHECKIN_STATUS_OPTIONS.map(o => `
      <button type="button" class="checkin-btn checkin-btn-${o.value === "ready" ? "green" : o.value === "not-sure" ? "yellow" : "red"}"
              data-value="${escHtml(o.value)}" data-group="morning">
        <span class="checkin-btn-emoji">${o.emoji}</span>
        <span class="checkin-btn-label">${escHtml(o.label)}</span>
      </button>`).join("");

    const checkoutBtns = CONFIG.CHECKOUT_STATUS_OPTIONS.map(o => `
      <button type="button" class="checkin-btn checkin-btn-${o.value === "better-good" ? "green" : o.value === "same" ? "yellow" : "red"}"
              data-value="${escHtml(o.value)}" data-group="checkout">
        <span class="checkin-btn-emoji">${o.emoji}</span>
        <span class="checkin-btn-label">${escHtml(o.label)}</span>
      </button>`).join("");

    el.innerHTML = `
      <div class="walk-page-header">
        <h2 class="walk-page-title">Student Check-In</h2>
        <p class="walk-page-sub">A quick check — how are you doing today?</p>
      </div>

      <div class="checkin-privacy-note">
        Student check-ins are private support information. They help adults notice patterns and offer support.
      </div>

      <!-- Student selector -->
      <div class="form-card walk-card">
        <div class="walk-field-label">Student</div>
        <select class="form-select" id="checkinStudent">
          <option value="">— Who is checking in? —</option>
          ${studentOptionsHtml}
        </select>
      </div>

      <!-- Today Status Card -->
      <div id="checkinTodayStatus" class="hidden"></div>

      <div id="checkinConfirm" class="hidden"></div>

      <!-- Morning Check-In -->
      <div class="checkin-section">
        <div class="checkin-section-header">
          <span class="checkin-section-icon">🌅</span>
          <div>
            <div class="checkin-section-title">Morning Check-In</div>
            <div class="checkin-section-sub">How are you coming into school today?</div>
          </div>
        </div>
        <div class="checkin-big-buttons">
          ${checkinBtns}
        </div>
        <input type="hidden" id="morningStatus" value="">
        <div class="checkin-note-wrap">
          <div class="checkin-note-prompt">Want to say anything about it?</div>
          <textarea class="form-textarea" id="morningNote" rows="2"
            placeholder="You can write a few words if you want."></textarea>
        </div>
        <button type="button" class="btn btn-primary btn-lg btn-full" id="submitCheckinBtn">
          Submit Check-In
        </button>
      </div>

      <!-- End-of-Day Check-Out -->
      <div class="checkin-section" style="margin-top:20px">
        <div class="checkin-section-header">
          <span class="checkin-section-icon">🌇</span>
          <div>
            <div class="checkin-section-title">End-of-Day Check-Out</div>
            <div class="checkin-section-sub">How are you leaving school today?</div>
          </div>
        </div>
        <div class="checkin-big-buttons">
          ${checkoutBtns}
        </div>
        <input type="hidden" id="checkoutStatus" value="">
        <div class="checkin-note-wrap">
          <div class="checkin-note-prompt">Anything you want adults to know before tomorrow?</div>
          <textarea class="form-textarea" id="checkoutNote" rows="2"
            placeholder="Optional — one sentence is enough."></textarea>
        </div>
        <button type="button" class="btn btn-primary btn-lg btn-full" id="submitCheckoutBtn">
          Submit Check-Out
        </button>
      </div>`;

    // Student select → refresh today status card
    const studentSel = el.querySelector("#checkinStudent");
    const todayCard  = el.querySelector("#checkinTodayStatus");

    const refreshTodayCard = () => {
      const sid = studentSel.value;
      if (!sid) { todayCard.className = "hidden"; todayCard.innerHTML = ""; return; }
      const allChecks = DB.getStudentChecks();
      const todayChecks  = allChecks.filter(c => c.studentId === sid && c.date === today);
      const morningEntry = todayChecks.find(c => c.type === "check-in");
      const eveningEntry = todayChecks.find(c => c.type === "check-out");
      const ps = PILOT_STUDENTS.find(s => s.id === sid);

      const statusHtml = (entry) => {
        if (!entry) return '<span class="checkin-status-pending">Not yet completed</span>';
        const colorMap = {
          "ready":"var(--green)", "not-sure":"var(--yellow)", "hard-morning":"var(--red)",
          "better-good":"var(--green)", "same":"var(--yellow)", "rough-day":"var(--red)"
        };
        const color = colorMap[entry.status] || "var(--text-muted)";
        return `<span style="color:${color};font-weight:700">${escHtml(entry.statusLabel)}</span>`;
      };

      todayCard.className = "checkin-today-card";
      todayCard.innerHTML = `
        <div class="checkin-today-name">${escHtml(ps?.name || sid)}</div>
        <div class="checkin-today-row">
          <span class="checkin-today-label">Morning:</span>
          ${statusHtml(morningEntry)}
          ${morningEntry?.note ? `<span class="checkin-today-note">"${escHtml(morningEntry.note)}"</span>` : ""}
        </div>
        <div class="checkin-today-row">
          <span class="checkin-today-label">End of Day:</span>
          ${statusHtml(eveningEntry)}
          ${eveningEntry?.note ? `<span class="checkin-today-note">"${escHtml(eveningEntry.note)}"</span>` : ""}
        </div>`;
    };

    studentSel.addEventListener("change", refreshTodayCard);

    // Big button selection (single-select per group)
    el.querySelectorAll(".checkin-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const group = btn.dataset.group;
        el.querySelectorAll(`.checkin-btn[data-group="${group}"]`).forEach(b => b.classList.remove("checkin-btn-selected"));
        btn.classList.add("checkin-btn-selected");
        el.querySelector(group === "morning" ? "#morningStatus" : "#checkoutStatus").value = btn.dataset.value;
      });
    });

    el.querySelector("#submitCheckinBtn").addEventListener("click",  () => this._submitStudentCheck(el, "check-in",  refreshTodayCard));
    el.querySelector("#submitCheckoutBtn").addEventListener("click", () => this._submitStudentCheck(el, "check-out", refreshTodayCard));
  },

  _submitStudentCheck(pageEl, type, refreshCallback) {
    const user     = this.getCurrentUser();
    const sid      = pageEl.querySelector("#checkinStudent").value;
    const isCheckin = type === "check-in";
    const status    = pageEl.querySelector(isCheckin ? "#morningStatus" : "#checkoutStatus").value;
    const note      = (pageEl.querySelector(isCheckin ? "#morningNote" : "#checkoutNote").value || "").trim();

    if (!sid)    { showToast("Please select a student.", "error"); return; }
    if (!status) { showToast(`Please select a ${isCheckin ? "morning" : "end-of-day"} status.`, "error"); return; }

    const ps          = PILOT_STUDENTS.find(s => s.id === sid);
    const today       = new Date().toISOString().slice(0, 10);
    const allOpts     = isCheckin ? CONFIG.CHECKIN_STATUS_OPTIONS : CONFIG.CHECKOUT_STATUS_OPTIONS;
    const statusLabel = allOpts.find(o => o.value === status)?.label || status;
    const statusEmoji = allOpts.find(o => o.value === status)?.emoji || "";
    const typeLabel   = isCheckin ? "Morning Check-In" : "End-of-Day Check-Out";

    DB.addStudentCheck({
      type,
      date:            today,
      studentId:       sid,
      studentName:     ps?.name        || "",
      teacherId:       ps?.teacherId   || "",
      teacherName:     ps?.teacherName || "",
      status,
      statusLabel,
      note,
      submittedById:   user.id,
      submittedByName: user.name
    });

    const confirmEl = pageEl.querySelector("#checkinConfirm");
    confirmEl.className = "save-confirm";
    confirmEl.innerHTML = `
      <div class="save-confirm-check">✓</div>
      <h3>${escHtml(typeLabel)} Saved</h3>
      <p><strong>${escHtml(ps?.name || sid)}</strong></p>
      <p>${statusEmoji} ${escHtml(statusLabel)}</p>
      ${note ? `<p style="font-size:13px;color:var(--text-secondary);margin-top:4px">"${escHtml(note)}"</p>` : ""}
      <div class="confirm-actions">
        <button class="btn btn-secondary btn-full" id="checkinDoneBtn">Done</button>
        <a href="#dashboard" class="btn btn-secondary btn-full">View Dashboard</a>
      </div>`;

    confirmEl.scrollIntoView({ behavior: "smooth", block: "start" });

    // Reset this section's controls only
    pageEl.querySelector(isCheckin ? "#morningStatus" : "#checkoutStatus").value = "";
    pageEl.querySelector(isCheckin ? "#morningNote"   : "#checkoutNote").value   = "";
    pageEl.querySelectorAll(`.checkin-btn[data-group="${isCheckin ? "morning" : "checkout"}"]`)
          .forEach(b => b.classList.remove("checkin-btn-selected"));

    if (refreshCallback) refreshCallback();
    showToast(`${typeLabel} saved!`);

    document.getElementById("checkinDoneBtn").addEventListener("click", () => {
      confirmEl.className = "hidden";
      confirmEl.innerHTML = "";
      document.getElementById("contentArea").scrollTo({ top: 0, behavior: "smooth" });
    });
  },

  /* ── DASHBOARD ─────────────────────────────────────────────────────────────── */

  async renderDashboard() {
    const el           = document.getElementById("page-dashboard");
    if (!AUTH.isAdmin || USER_CONTEXT.isViewingAsTeacher) {
      el.innerHTML = `<div class="page-header"><h2>Dashboard</h2></div>
        <div class="warning-banner">Administrator access is required to view dashboard analytics.</div>`;
      return;
    }
    const records      = DB.getRecords();
    const pulses       = DB.getDailyPulses();
    const studentChecks = DB.getStudentChecks();
    const teachers     = DB.getTeachers();
    const classrooms   = DB.getClassrooms();
    const weekRecords  = DB.getRecordsThisWeek();

    // ── User context ──
    const user = this.getCurrentUser();
    const currentStudents = user.isAdmin
      ? PILOT_STUDENTS
      : PILOT_STUDENTS.filter(s => s.teacherId === user.id);
    const filteredPulses  = user.isAdmin ? pulses : pulses.filter(p => {
      const ps = PILOT_STUDENTS.find(s => s.name === p.student);
      return ps?.teacherId === user.id;
    });
    const filteredRecords = user.isAdmin ? records : records.filter(r => r.responses.teacherId === user.id);
    // PACE dashboard data is populated from IEP_Pace_Visits below. Never use
    // the legacy local operational cache as an analytics source.
    const filteredPaceLogs = [];
    const filteredChecks = user.isAdmin ? studentChecks : studentChecks.filter(c => {
      const ps = PILOT_STUDENTS.find(s => s.id === c.studentId);
      return ps?.teacherId === user.id;
    });

    // ── Walkthrough stats ──
    const followUp     = filteredRecords.filter(r => r.responses.supportNeeded !== "none");
    const helpSoon     = filteredRecords.filter(r => r.responses.supportNeeded === "help-soon");
    const urgent       = filteredRecords.filter(r => r.responses.supportNeeded === "urgent");
    const teacherIds   = new Set(filteredRecords.map(r => r.responses.teacherId).filter(Boolean));
    const classroomIds = new Set(filteredRecords.map(r => r.responses.classroomId).filter(Boolean));

    const statusCounts = { "on-track":0, "monitor":0, "needs-support":0, "immediate-follow-up":0 };
    filteredRecords.forEach(r => { if (statusCounts[r.responses.classroomStatus] !== undefined) statusCounts[r.responses.classroomStatus]++; });

    const supportCounts = {};
    filteredRecords.forEach(r => { (r.responses.supportsObserved||[]).forEach(s => { supportCounts[s]=(supportCounts[s]||0)+1; }); });
    const topSupports = Object.entries(supportCounts).sort((a,b)=>b[1]-a[1]).slice(0,6);
    const maxSupport  = topSupports[0]?.[1] || 1;
    const recent      = [...filteredRecords].sort((a,b)=>new Date(b.submittedAt)-new Date(a.submittedAt)).slice(0,5);

    // ── Daily Pulse stats ──
    const today = new Date().toISOString().slice(0,10);
    const flaggedToday = new Set(
      filteredPulses
        .filter(p => p.timestamp.slice(0,10) === today && (p.pulseStatus === "struggles" || p.pulseStatus === "concern"))
        .map(p => p.student)
    );
    const walkSupportReqs  = filteredRecords.filter(r => r.responses.supportNeeded && r.responses.supportNeeded !== "none").length;
    const pulseSupportReqs = filteredPulses.filter(p => p.supportLevel === "talk-weekly" || p.supportLevel === "need-help-now").length;

    const categoryCounts = {};
    filteredPulses.forEach(p => { (p.categories||[]).forEach(c => { categoryCounts[c]=(categoryCounts[c]||0)+1; }); });
    const topCategory = Object.entries(categoryCounts).sort((a,b)=>b[1]-a[1])[0];

    // PACE analytics render asynchronously from SharePoint after the local
    // dashboard shell; no PACE metric is derived from localStorage here.
    const paceBehaviorCounts = {};
    const topPaceIntervention = null;

    // ── Student Check-In stats ──
    const checkToday        = filteredChecks.filter(c => c.date === today);
    const checkinsToday     = checkToday.filter(c => c.type === "check-in").length;
    const checkoutsToday    = checkToday.filter(c => c.type === "check-out").length;
    const hardMorningsToday = checkToday.filter(c => c.status === "hard-morning").length;
    const roughEndingsToday = checkToday.filter(c => c.status === "rough-day").length;
    const recentChecks      = [...filteredChecks].sort((a,b)=>new Date(b.timestamp)-new Date(a.timestamp)).slice(0,10);
    const selfReportCounts  = {};
    filteredChecks.forEach(c => { if (c.statusLabel) selfReportCounts[c.statusLabel]=(selfReportCounts[c.statusLabel]||0)+1; });
    const topSelfReport     = Object.entries(selfReportCounts).sort((a,b)=>b[1]-a[1])[0];
    const hardMorningsByStudent = {};
    filteredChecks.filter(c => c.status === "hard-morning")
      .forEach(c => { hardMorningsByStudent[c.studentName]=(hardMorningsByStudent[c.studentName]||0)+1; });
    const repeatedHardMornings = Object.entries(hardMorningsByStudent)
      .filter(([,n])=>n>=2).sort((a,b)=>b[1]-a[1]).slice(0,3);

    // ── Student Pulse Trends (last 10 days, newest first) ──
    const dotColors = { great:"#16a34a", struggles:"#ca8a04", concern:"#dc2626" };
    const dotLabels = { great:"Great Day", struggles:"Some Struggles", concern:"Significant Concern" };
    const studentTrends = currentStudents.map(ps => {
      const entries = [...filteredPulses.filter(p => p.student === ps.name)]
        .sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp))
        .slice(0,10);
      return { student: ps.name, teacherName: ps.teacherName, entries };
    }).filter(s => s.entries.length > 0);

    // ── Convergence Engine (pulse + PACE + check-in signals per student) ──
    const convergenceData = currentStudents.map(ps => {
      const pulseEntries = [...filteredPulses.filter(p => p.student === ps.name)]
        .sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp))
        .slice(0,7);
      const paceEntries  = [...filteredPaceLogs.filter(p => p.studentId === ps.id)]
        .sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp))
        .slice(0,7);
      const checkEntries = [...filteredChecks.filter(c => c.studentId === ps.id)]
        .sort((a,b) => new Date(b.timestamp) - new Date(a.timestamp))
        .slice(0,7);
      if (pulseEntries.length === 0 && paceEntries.length === 0 && checkEntries.length === 0) return null;

      let green = 0, yellow = 0, red = 0;
      pulseEntries.forEach(e => {
        if      (e.pulseStatus === "great")    green++;
        else if (e.pulseStatus === "struggles") yellow++;
        else if (e.pulseStatus === "concern")   red++;
      });
      const hasElevated      = paceEntries.some(p =>
        (p.behaviors||[]).some(b => CONFIG.PACE_ELEVATED_BEHAVIORS.includes(b))
      );
      const paceCount        = paceEntries.length;
      const checkRedCount    = checkEntries.filter(c => ["hard-morning","rough-day"].includes(c.status)).length;
      const checkConcernCount = checkEntries.filter(c => ["hard-morning","rough-day","not-sure","same"].includes(c.status)).length;

      const adultSignals = (yellow + red) + (paceCount >= 1 ? 1 : 0);
      let label = "stable";
      if ((yellow + red) >= 3 || (paceCount >= 3 && (yellow + red) >= 1) || (hasElevated && red >= 1) ||
          (checkRedCount >= 3 && adultSignals >= 1)) {
        label = "escalating";
      } else if (yellow >= 2 || paceCount >= 1 || hasElevated || checkConcernCount >= 2) {
        label = "watch";
      }
      return { student: ps.name, teacherName: ps.teacherName, green, yellow, red, total: pulseEntries.length, paceCount, checkRedCount, checkConcernCount, label };
    }).filter(Boolean);

    // ── Emerging Patterns (category frequency per student) ──
    const patternData = currentStudents.map(ps => {
      const entries = filteredPulses.filter(p => p.student === ps.name);
      if (entries.length === 0) return null;
      const counts = {};
      entries.forEach(e => { (e.categories||[]).forEach(c => { counts[c]=(counts[c]||0)+1; }); });
      const sorted = Object.entries(counts).sort((a,b)=>b[1]-a[1]);
      if (sorted.length === 0) return null;
      return { student: ps.name, teacherName: ps.teacherName, topCategories: sorted.slice(0,3), topLabel: sorted[0][0] };
    }).filter(Boolean);

    // ── Merged Help Requests ──
    const helpRequests = [];
    filteredRecords.forEach(r => {
      const sn = r.responses.supportNeeded;
      if (!sn || sn === "none") return;
      const t = teachers.find(x => x.id === r.responses.teacherId);
      let urgency, label;
      if      (sn === "urgent")         { urgency = 3; label = "Need Help Now"; }
      else if (sn === "help-soon")      { urgency = 2; label = "Need Help Now"; }
      else if (sn === "talk-at-weekly") { urgency = 1; label = "Let's Talk"; }
      else return;
      helpRequests.push({ who: t ? t.name : "Unknown Teacher", source: "Walkthrough", urgency, label, ts: r.submittedAt });
    });
    filteredPulses.forEach(p => {
      if (!p.supportLevel || p.supportLevel === "got-it") return;
      let urgency, label;
      if      (p.supportLevel === "need-help-now") { urgency = 3; label = "Need Help Now"; }
      else if (p.supportLevel === "talk-weekly")   { urgency = 1; label = "Let's Talk"; }
      else return;
      helpRequests.push({ who: p.student, source: "Daily Pulse", urgency, label, ts: p.timestamp });
    });
    helpRequests.sort((a,b) => b.urgency - a.urgency || new Date(b.ts) - new Date(a.ts));
    const recentHelp = helpRequests.slice(0,10);

    el.innerHTML = `
      <div class="info-banner">
        <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18" style="flex-shrink:0"><path fill-rule="evenodd" d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" clip-rule="evenodd"/></svg>
        <div><strong>Visibility Snapshot</strong> — Support patterns and classroom trends. This is not a teacher evaluation tool. Use for coaching conversations and support planning.</div>
        <span id="d-live-badge" class="hidden" style="margin-left:auto;font-size:11px;font-weight:600;white-space:nowrap"></span>
      </div>

      <div class="stat-grid">
        <div class="stat-card accent-blue">
          <div class="stat-label">Total Walkthroughs</div>
          <div class="stat-value" id="d-walk-total">${filteredRecords.length}</div>
          <div class="stat-sub">All time</div>
        </div>
        <div class="stat-card accent-blue">
          <div class="stat-label">This Week</div>
          <div class="stat-value" id="d-walk-week">${weekRecords.length}</div>
          <div class="stat-sub">Since Sunday</div>
        </div>
        <div class="stat-card accent-green">
          <div class="stat-label">Teachers Visited</div>
          <div class="stat-value" id="d-walk-teachers">${teacherIds.size}</div>
          <div class="stat-sub">of ${teachers.length} in setup</div>
        </div>
        <div class="stat-card accent-green">
          <div class="stat-label">Classrooms Visited</div>
          <div class="stat-value" id="d-walk-classrooms">${classroomIds.size}</div>
          <div class="stat-sub">of ${classrooms.length} in setup</div>
        </div>
        <div class="stat-card accent-orange">
          <div class="stat-label">Follow-Up Opportunities</div>
          <div class="stat-value" id="d-walk-followups">${followUp.length}</div>
          <div class="stat-sub">Any support flag</div>
        </div>
        <div class="stat-card accent-orange">
          <div class="stat-label">Help Soon</div>
          <div class="stat-value" id="d-walk-helpsoon">${helpSoon.length}</div>
          <div class="stat-sub">Flagged this period</div>
        </div>
        <div class="stat-card accent-red">
          <div class="stat-label">Urgent</div>
          <div class="stat-value" id="d-walk-urgent">${urgent.length}</div>
          <div class="stat-sub">Needs immediate action</div>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat-card accent-blue">
          <div class="stat-label">Daily Pulse Entries</div>
          <div class="stat-value" id="d-pulse-total">${filteredPulses.length}</div>
          <div class="stat-sub">All time</div>
        </div>
        <div class="stat-card accent-orange">
          <div class="stat-label">Students Flagged Today</div>
          <div class="stat-value" id="d-pulse-flagged">${flaggedToday.size}</div>
          <div class="stat-sub">Struggles or concern</div>
        </div>
        <div class="stat-card accent-red">
          <div class="stat-label">Support Requests</div>
          <div class="stat-value" id="d-pulse-support">${walkSupportReqs + pulseSupportReqs}</div>
          <div class="stat-sub">Walkthroughs + pulses</div>
        </div>
        <div class="stat-card accent-green">
          <div class="stat-label">Most Common Concern</div>
          <div class="stat-value stat-value-text" id="d-pulse-concern">${topCategory ? escHtml(topCategory[0]) : "—"}</div>
          <div class="stat-sub" id="d-pulse-concern-sub">${topCategory ? topCategory[1] + " report" + (topCategory[1] !== 1 ? "s" : "") : "No categories yet"}</div>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat-card accent-orange">
          <div class="stat-label">PACE Visits Today</div>
          <div class="stat-value" id="d-pace-today">—</div>
          <div class="stat-sub">Completed visits</div>
        </div>
        <div class="stat-card accent-green">
          <div class="stat-label">Students Seen Today</div>
          <div class="stat-value" id="d-pace-students">—</div>
          <div class="stat-sub">Unique students</div>
        </div>
        <div class="stat-card accent-blue">
          <div class="stat-label">Average PACE Duration</div>
          <div class="stat-value" id="d-pace-duration">—</div>
          <div class="stat-sub">Completed visits · last 30 days</div>
        </div>
        <div class="stat-card accent-red">
          <div class="stat-label" id="d-pace-fourth-label">PACE Activity</div>
          <div class="stat-value stat-value-text" id="d-pace-fourth">—</div>
          <div class="stat-sub" id="d-pace-fourth-sub">Last 30 days</div>
        </div>
      </div>

      <div class="stat-grid">
        <div class="stat-card accent-green">
          <div class="stat-label">Student Check-Ins Today</div>
          <div class="stat-value">${checkinsToday}</div>
          <div class="stat-sub">Morning check-ins</div>
        </div>
        <div class="stat-card accent-blue">
          <div class="stat-label">Check-Outs Today</div>
          <div class="stat-value">${checkoutsToday}</div>
          <div class="stat-sub">End-of-day check-outs</div>
        </div>
        <div class="stat-card accent-orange">
          <div class="stat-label">Hard Mornings Today</div>
          <div class="stat-value">${hardMorningsToday}</div>
          <div class="stat-sub">Self-reported today</div>
        </div>
        <div class="stat-card accent-red">
          <div class="stat-label">Rough Endings Today</div>
          <div class="stat-value">${roughEndingsToday}</div>
          <div class="stat-sub">Self-reported today</div>
        </div>
      </div>

      <!-- Student Pulse Trends -->
      <div class="card" style="margin-bottom:20px">
        <div class="card-title">Student Pulse Trends</div>
        ${studentTrends.length === 0
          ? `<div class="empty-state"><div class="empty-icon">💚</div><p>No daily pulse entries yet. <a href="#pulse" style="color:var(--color-primary)">Record the first one.</a></p></div>`
          : studentTrends.map(s => `
              <div class="student-trend-card">
                <div class="student-trend-name">${escHtml(s.student)}${user.isAdmin ? `<span class="trend-teacher"> — ${escHtml(s.teacherName)}</span>` : ""}</div>
                <div style="font-size:11.5px;color:var(--text-muted);margin-bottom:6px">Last ${s.entries.length} Day${s.entries.length !== 1 ? "s" : ""}</div>
                <div class="pulse-strip">
                  ${s.entries.map(e => `<span class="pulse-dot" style="background:${dotColors[e.pulseStatus]||"#94a3b8"}" title="${escHtml(dotLabels[e.pulseStatus]||"")} — ${fmtDateShort(e.timestamp)}"></span>`).join("")}
                </div>
              </div>`).join("")}
      </div>

      <!-- Convergence Engine -->
      <div class="card" style="margin-bottom:20px">
        <div class="card-title">Convergence Engine</div>
        ${convergenceData.length === 0
          ? `<div class="empty-state"><div class="empty-icon">📊</div><p>No data yet. Record daily pulses to see convergence signals.</p></div>`
          : `<div class="convergence-grid">${convergenceData.map(s => `
              <div class="convergence-item">
                <div>
                  <div class="convergence-student">${escHtml(s.student)}${user.isAdmin ? `<span class="trend-teacher"> — ${escHtml(s.teacherName)}</span>` : ""}</div>
                  <div class="convergence-counts">🟢 ${s.green} &nbsp;🟡 ${s.yellow} &nbsp;🔴 ${s.red}${s.paceCount > 0 ? ` &nbsp;<span class="convergence-pace-signal">📋 ×${s.paceCount}</span>` : ""}${s.checkConcernCount > 0 ? ` &nbsp;<span class="convergence-check-signal">💬 ×${s.checkConcernCount}</span>` : ""} <span style="color:var(--text-muted);margin-left:4px">last ${s.total}</span></div>
                </div>
                <span class="convergence-badge ${s.label}">${s.label === "escalating" ? "Escalating" : s.label === "watch" ? "Watch" : "Stable"}</span>
              </div>`).join("")}</div>`}
      </div>

      <!-- Emerging Patterns -->
      <div class="card" style="margin-bottom:20px">
        <div class="card-title">Emerging Patterns</div>
        ${patternData.length === 0 && filteredPaceLogs.length === 0 && filteredChecks.length === 0
          ? `<div class="empty-state"><div class="empty-icon">🔍</div><p>No category patterns detected yet.</p></div>`
          : `${patternData.map(p => `
              <div class="pattern-item">
                <div class="pattern-student">${escHtml(p.student)}${user.isAdmin ? `<span class="trend-teacher"> — ${escHtml(p.teacherName)}</span>` : ""}</div>
                <div class="pattern-detail">
                  ${p.topCategories.map(([cat,cnt]) => `<div>${escHtml(cat)}: <strong>${cnt}</strong></div>`).join("")}
                  <div class="pattern-top">Most Frequent: <strong>${escHtml(p.topLabel)}</strong></div>
                </div>
              </div>`).join("")}
            ${filteredPaceLogs.length > 0 ? `
              <div class="pace-patterns">
                <div class="pace-patterns-label">PACE Behavior Patterns</div>
                <div class="pace-patterns-grid">
                  ${Object.entries(paceBehaviorCounts).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([b,n]) =>
                    `<div class="pace-pattern-row"><span class="pace-pattern-name">${escHtml(b)}</span><span class="pace-pattern-count">${n}</span></div>`
                  ).join("")}
                </div>
                ${topPaceIntervention ? `
                  <div class="pace-patterns-label" style="margin-top:12px">Most Used Intervention</div>
                  <div class="pace-pattern-row"><span class="pace-pattern-name">${escHtml(topPaceIntervention[0])}</span><span class="pace-pattern-count">${topPaceIntervention[1]}</span></div>
                ` : ""}
              </div>` : ""}
            ${filteredChecks.length > 0 ? `
              <div class="pace-patterns">
                <div class="pace-patterns-label">Student Self-Report Patterns</div>
                ${topSelfReport ? `
                  <div class="pace-patterns-label" style="margin-bottom:4px;font-size:10.5px">Most Frequent</div>
                  <div class="pace-pattern-row"><span class="pace-pattern-name">${escHtml(topSelfReport[0])}</span><span class="pace-pattern-count">${topSelfReport[1]}</span></div>
                ` : ""}
                ${repeatedHardMornings.length > 0 ? `
                  <div class="pace-patterns-label" style="margin-top:10px">Repeated Hard Mornings (2+)</div>
                  <div class="pace-patterns-grid">
                    ${repeatedHardMornings.map(([name,n]) =>
                      `<div class="pace-pattern-row"><span class="pace-pattern-name">${escHtml(name)}</span><span class="pace-pattern-count">${n}×</span></div>`
                    ).join("")}
                  </div>
                ` : ""}
              </div>` : ""}
          `}
      </div>

      <!-- PACE Room Activity -->
      <div class="card" style="margin-bottom:20px">
        <div class="card-title">PACE Activity <a href="#pace" class="pace-card-link">Explore visits →</a></div>
        <div id="d-pace-flags"></div>
        <div id="d-pace-activity"><div class="reports-loading"><span class="spinner"></span> Loading completed visits…</div></div>
      </div>

      <!-- Student Voice Check-Ins -->
      <div class="card" style="margin-bottom:20px">
        <div class="card-title">Student Voice Check-Ins</div>
        ${recentChecks.length === 0
          ? `<div class="empty-state"><div class="empty-icon">💬</div><p>No student check-ins yet. <a href="#checkin" style="color:var(--color-primary)">Record the first one.</a></p></div>`
          : `<ul class="recent-list">${recentChecks.map(c => {
              const isCheckin = c.type === "check-in";
              const colorMap  = { "ready":"#16a34a","better-good":"#16a34a","not-sure":"#ca8a04","same":"#ca8a04","hard-morning":"#dc2626","rough-day":"#dc2626" };
              const dotColor  = colorMap[c.status] || "#94a3b8";
              const typeLabel = isCheckin ? "Student Check-In" : "Student Check-Out";
              return `<li class="recent-item">
                <div class="recent-dot" style="background:${dotColor}"></div>
                <div class="recent-meta">
                  <div class="recent-who">${escHtml(c.studentName)}${user.isAdmin && c.teacherName ? `<span class="trend-teacher"> — ${escHtml(c.teacherName)}</span>` : ""}</div>
                  <div class="recent-when">${escHtml(typeLabel)} · ${escHtml(c.statusLabel)} · ${fmtDateTime(c.timestamp)}</div>
                  ${c.note ? `<div style="font-size:12px;color:var(--text-secondary);margin-top:2px;font-style:italic">"${escHtml(c.note)}"</div>` : ""}
                </div>
              </li>`;
            }).join("")}</ul>`}
      </div>

      <!-- Help Requests (merged) -->
      <div class="card" style="margin-bottom:20px">
        <div class="card-title">Help Requests</div>
        ${recentHelp.length === 0
          ? `<div class="empty-state"><div class="empty-icon">🙌</div><p>No active help requests.</p></div>`
          : `<ul class="recent-list">${recentHelp.map(r => {
              const urgencyColor  = r.urgency >= 3 ? "var(--red)" : "var(--blue)";
              const urgencyBg     = r.urgency >= 3 ? "var(--red-bg)" : "var(--blue-bg)";
              const urgencyBorder = r.urgency >= 3 ? "var(--red-border)" : "var(--blue-border)";
              return `<li class="recent-item">
                <div class="recent-dot" style="background:${urgencyColor}"></div>
                <div class="recent-meta">
                  <div class="recent-who">${escHtml(r.who)}</div>
                  <div class="recent-when">${fmtDateTime(r.ts)} · ${escHtml(r.source)}</div>
                </div>
                <span class="badge" style="background:${urgencyBg};color:${urgencyColor};border:1px solid ${urgencyBorder};font-size:11.5px;white-space:nowrap">${escHtml(r.label)}</span>
              </li>`;
            }).join("")}</ul>`}
      </div>

      <div class="dash-grid">
        <div class="card">
          <div class="card-title">Most Common Classroom Status</div>
          <ul class="status-list" id="d-status-list">
            <li class="status-row"><span class="status-label-text">🟢 On Track</span><span class="status-count green">${statusCounts["on-track"]}</span></li>
            <li class="status-row"><span class="status-label-text">🟡 Monitor</span><span class="status-count yellow">${statusCounts["monitor"]}</span></li>
            <li class="status-row"><span class="status-label-text">🟠 Needs Support</span><span class="status-count orange">${statusCounts["needs-support"]}</span></li>
            <li class="status-row"><span class="status-label-text">🔴 Immediate Follow-Up</span><span class="status-count red">${statusCounts["immediate-follow-up"]}</span></li>
          </ul>
        </div>
        <div class="card">
          <div class="card-title">Supports Observed</div>
          <div id="d-supports-body">
            ${topSupports.length === 0
              ? `<div class="empty-state"><div class="empty-icon">📋</div><p>No walkthroughs yet.</p></div>`
              : `<ul class="support-bar-list">${topSupports.map(([s,n])=>`
                <li class="support-bar-row">
                  <div class="support-bar-meta"><span>${escHtml(s)}</span><span>${n}</span></div>
                  <div class="support-bar-track"><div class="support-bar-fill" style="width:${Math.round(n/maxSupport*100)}%"></div></div>
                </li>`).join("")}</ul>`}
          </div>
        </div>
      </div>

      <div class="card">
        <div class="card-title">Recent Walkthroughs</div>
        <div id="d-recent-walks-body">
          ${recent.length === 0
            ? `<div class="empty-state"><div class="empty-icon">👣</div><p>No walkthroughs yet. <a href="#walkthrough" style="color:var(--color-primary)">Record the first one.</a></p></div>`
            : `<ul class="recent-list">${recent.map(r=>{
                const t = teachers.find(x=>x.id===r.responses.teacherId);
                const c = classrooms.find(x=>x.id===r.responses.classroomId);
                const st = r.responses.classroomStatus;
                const dotColor = {"on-track":"#16a34a","monitor":"#ca8a04","needs-support":"#ea580c","immediate-follow-up":"#dc2626"}[st]||"#94a3b8";
                return `<li class="recent-item">
                  <div class="recent-dot" style="background:${dotColor}"></div>
                  <div class="recent-meta">
                    <div class="recent-who">${t?escHtml(t.name):"Unknown"} · ${c?escHtml(c.name):"Unknown"}</div>
                    <div class="recent-when">${fmtDateTime(r.submittedAt)} · ${escHtml(r.responses.focus||"")}</div>
                  </div>
                  ${statusBadgeHtml(st)}
                </li>`;}).join("")}</ul>`}
        </div>
      </div>`;

    // Background: update stat cards with live SharePoint data
    if (AUTH.isAuthenticated) {
      try {
        const sp  = await DashboardSync.refresh(user);
        const upd = (id, v) => { const n = document.getElementById(id); if (n) n.textContent = v; };

        upd("d-walk-total",      sp.walk.total);
        upd("d-walk-week",       sp.walk.thisWeek);
        upd("d-walk-teachers",   sp.walk.teacherCount);
        upd("d-walk-classrooms", sp.walk.classroomCount);
        upd("d-walk-followups",  sp.walk.followUps);
        upd("d-walk-helpsoon",   sp.walk.helpSoon);
        upd("d-walk-urgent",     sp.walk.urgent);

        upd("d-pulse-total",     sp.pulse.total);
        upd("d-pulse-flagged",   sp.pulse.flaggedToday);
        upd("d-pulse-support",   sp.walk.followUps + sp.pulse.pulseSupportReqs);
        upd("d-pulse-concern",   sp.pulse.topCategory ? sp.pulse.topCategory[0] : "—");
        if (sp.pulse.topCategory) {
          const n    = sp.pulse.topCategory[1];
          const csub = document.getElementById("d-pulse-concern-sub");
          if (csub) csub.textContent = `${n} report${n !== 1 ? "s" : ""}`;
        }

        // PACE administrator metrics and activity are always read from the
        // shared IEP_Pace_Visits list. Missing optional columns change what
        // is displayed rather than producing fabricated zeroes.
        if (sp.pace) {
          upd("d-pace-today", sp.pace.visitsToday);
          upd("d-pace-students", sp.pace.studentsSeenToday);
          upd("d-pace-duration", sp.pace.averageDuration !== null ? `${sp.pace.averageDuration} min` : "—");
          const fourthLabel = document.getElementById("d-pace-fourth-label");
          const fourthSub   = document.getElementById("d-pace-fourth-sub");
          if (sp.pace.availability.scm && sp.pace.scmEvents !== null) {
            if (fourthLabel) fourthLabel.textContent = "SCM Events";
            upd("d-pace-fourth", sp.pace.scmEvents);
            if (fourthSub) fourthSub.textContent = "Last 30 days";
          } else {
            if (fourthLabel) fourthLabel.textContent = "Most Common Reason";
            upd("d-pace-fourth", sp.pace.mostCommonReason ? sp.pace.mostCommonReason[0] : "—");
            if (fourthSub) fourthSub.textContent = sp.pace.mostCommonReason
              ? `${sp.pace.mostCommonReason[1]} visit${sp.pace.mostCommonReason[1] !== 1 ? "s" : ""} · last 30 days`
              : "No completed visits in this period";
          }

          const flagsEl = document.getElementById("d-pace-flags");
          if (flagsEl) {
            const flags = [];
            if (sp.pace.flags.extendedVisits) flags.push(`${sp.pace.flags.extendedVisits} extended visit${sp.pace.flags.extendedVisits !== 1 ? "s" : ""} (45+ min)`);
            if (sp.pace.flags.repeatVisitDays) flags.push(`${sp.pace.flags.repeatVisitDays} repeat-visit day${sp.pace.flags.repeatVisitDays !== 1 ? "s" : ""} (3+)`);
            flagsEl.innerHTML = flags.length ? `<div class="pace-dashboard-flags">${flags.map(flag => `<span>${escHtml(flag)}</span>`).join("")}</div>` : "";
          }

          const activity = document.getElementById("d-pace-activity");
          if (activity) {
            activity.innerHTML = sp.pace.recentVisits.length === 0
              ? `<div class="empty-state"><div class="empty-icon">🚪</div><p>No completed PACE visits in the last 30 days.</p></div>`
              : `<ul class="recent-list">${sp.pace.recentVisits.map(visit => `<li class="recent-item pace-activity-item">
                  <div class="recent-dot" style="background:#2563eb;flex-shrink:0"></div>
                  <div class="recent-meta" style="flex:1;min-width:0">
                    <div class="recent-who">${escHtml(visit.student || "Student not recorded")}</div>
                    <div class="recent-when">${visit.date ? escHtml(fmtDate(visit.date + "T12:00:00")) : "—"} · ${escHtml(visit.timeIn || "—")} → ${escHtml(visit.timeOut || "—")}${visit.durationMinutes !== null ? ` · ${visit.durationMinutes} min` : ""}</div>
                    ${visit.reasons.length ? `<div class="pace-tag-row">${visit.reasons.slice(0, 3).map(reason => `<span class="pace-tag">${escHtml(reason)}</span>`).join("")}</div>` : ""}
                  </div>
                </li>`).join("")}</ul>`;
          }
        } else {
          const activity = document.getElementById("d-pace-activity");
          if (activity) activity.innerHTML = `<div class="reports-error">PACE activity is temporarily unavailable from SharePoint.</div>`;
        }

        // Most Common Classroom Status
        const statusList = document.getElementById("d-status-list");
        if (statusList) {
          statusList.innerHTML = `
            <li class="status-row"><span class="status-label-text">🟢 On Track</span><span class="status-count green">${sp.walk.statusCounts["on-track"]}</span></li>
            <li class="status-row"><span class="status-label-text">🟡 Monitor</span><span class="status-count yellow">${sp.walk.statusCounts["monitor"]}</span></li>
            <li class="status-row"><span class="status-label-text">🟠 Needs Support</span><span class="status-count orange">${sp.walk.statusCounts["needs-support"]}</span></li>
            <li class="status-row"><span class="status-label-text">🔴 Immediate Follow-Up</span><span class="status-count red">${sp.walk.statusCounts["immediate-follow-up"]}</span></li>`;
        }

        // Supports Observed bar chart
        const supBody = document.getElementById("d-supports-body");
        if (supBody) {
          supBody.innerHTML = sp.walk.topSupports.length === 0
            ? `<div class="empty-state"><div class="empty-icon">📋</div><p>No walkthroughs yet.</p></div>`
            : `<ul class="support-bar-list">${sp.walk.topSupports.map(([s, n]) => `
                <li class="support-bar-row">
                  <div class="support-bar-meta"><span>${escHtml(s)}</span><span>${n}</span></div>
                  <div class="support-bar-track"><div class="support-bar-fill" style="width:${Math.round(n / sp.walk.maxSupport * 100)}%"></div></div>
                </li>`).join("")}</ul>`;
        }

        // Recent Walkthroughs from SharePoint
        const recentBody = document.getElementById("d-recent-walks-body");
        if (recentBody) {
          const srLabel = { "none": "None", "talk-at-weekly": "Weekly Discussion", "help-soon": "Help Soon", "urgent": "Urgent" };
          const srDot   = { "none": "#16a34a", "talk-at-weekly": "#2563eb", "help-soon": "#ca8a04", "urgent": "#dc2626" };
          recentBody.innerHTML = sp.walk.recentItems.length === 0
            ? `<div class="empty-state"><div class="empty-icon">👣</div><p>No walkthroughs yet. <a href="#walkthrough" style="color:var(--color-primary)">Record the first one.</a></p></div>`
            : `<ul class="recent-list">${sp.walk.recentItems.map(item => {
                const sr      = item["SupportRequested"] || "none";
                const dotColor = srDot[sr] || "#94a3b8";
                const who     = escHtml(item["Classroom"] || item["Teacher"] || "—");
                const student = item["Student"] ? escHtml(item["Student"]) + " · " : "";
                const engage  = escHtml(item["Engagement"] || "");
                const win     = item["Observed Win"] || "";
                const tsRaw   = item["ObservationTimestamp"] || item["ObservationDate"] || "";
                const whenStr = (() => {
                  if (!tsRaw) return "—";
                  const d = new Date(tsRaw);
                  if (isNaN(d)) return escHtml(tsRaw);
                  const datePart = d.toLocaleDateString([], { month: "short", day: "numeric", year: "numeric" });
                  const timePart = item["ObservationTime"] || d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
                  return `${datePart} · ${timePart}`;
                })();
                return `<li class="recent-item">
                  <div class="recent-dot" style="background:${dotColor}"></div>
                  <div class="recent-meta">
                    <div class="recent-who">${who}</div>
                    <div class="recent-when">${whenStr}${engage ? " · " + engage : ""}</div>
                    ${student || sr !== "none" ? `<div style="font-size:11.5px;color:var(--text-secondary);margin-top:2px">${student}Support: ${escHtml(srLabel[sr] || sr)}</div>` : ""}
                    ${win ? `<div style="font-size:11.5px;color:var(--text-muted);margin-top:2px;font-style:italic">Win: ${escHtml(win.slice(0, 80))}${win.length > 80 ? "…" : ""}</div>` : ""}
                  </div>
                </li>`;
              }).join("")}</ul>`;
        }

        const badge = document.getElementById("d-live-badge");
        if (badge) {
          badge.textContent = "● Live from SharePoint";
          badge.style.color = "var(--color-primary)";
          badge.classList.remove("hidden");
        }
      } catch (e) {
        console.warn("Dashboard SharePoint refresh failed:", e.message);
        const badge = document.getElementById("d-live-badge");
        if (badge) {
          badge.textContent = "SharePoint unavailable";
          badge.style.color = "var(--text-muted)";
          badge.classList.remove("hidden");
        }
      }
    }
  },

  /* ── SETUP ──────────────────────────────────────────────────────────────────── */

  renderSetup(activeTab) {
    if (!AUTH.isAdmin) {
      const el = document.getElementById("page-setup");
      el.innerHTML = `<div class="page-header"><h2>Setup</h2></div>
        <div class="warning-banner">You do not have permission to access Setup.</div>`;
      return;
    }
    activeTab = activeTab || "users";
    const el  = document.getElementById("page-setup");
    const tabs = ["users","teachers","students","classrooms","data"];
    const tabLabels = { users:"Users", teachers:"Teachers", students:"Students", classrooms:"Classrooms", data:"Data Tools" };

    el.innerHTML = `
      <div class="page-header">
        <h2>Setup</h2>
        <p>Manage pilot users, teachers, students, and classrooms from SharePoint.</p>
        <button class="btn btn-secondary btn-sm setup-refresh-btn" id="setupRefreshBtn">↻ Refresh from SharePoint</button>
      </div>
      <div class="tab-bar">
        ${tabs.map(t => `<button class="tab-btn${activeTab===t?" active":""}" data-tab="${t}">${tabLabels[t]}</button>`).join("")}
      </div>
      <div id="tabContent" class="setup-loading">
        <div class="empty-state-text">Loading from SharePoint…</div>
      </div>`;

    el.querySelector(".tab-bar").addEventListener("click", e => {
      const btn = e.target.closest(".tab-btn");
      if (!btn) return;
      el.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      btn.classList.add("active");
      this._setupActiveTab = btn.dataset.tab;
      document.getElementById("tabContent").innerHTML = this._renderSetupTab(btn.dataset.tab);
      this._bindSetupTabEvents(btn.dataset.tab);
    });

    document.getElementById("setupRefreshBtn").addEventListener("click", () => {
      APP_CACHE.clear();
      document.getElementById("tabContent").innerHTML = `<div class="setup-loading"><div class="empty-state-text">Refreshing…</div></div>`;
      SETUP_DATA.refresh();
    });

    this._setupActiveTab = activeTab;

    if (SETUP_DATA.users.length || SETUP_DATA.teachers.length) {
      document.getElementById("tabContent").innerHTML = this._renderSetupTab(activeTab);
      this._bindSetupTabEvents(activeTab);
    } else {
      SETUP_DATA.refresh();
    }
  },

  _refreshSetupUI() {
    const tabContent = document.getElementById("tabContent");
    if (!tabContent) return;
    const activeTab = this._setupActiveTab || "users";
    tabContent.innerHTML = this._renderSetupTab(activeTab);
    this._bindSetupTabEvents(activeTab);
  },

  _renderSetupTab(tab) {
    if (tab === "users")       return this._renderUsersTab();
    if (tab === "teachers")    return this._renderTeachersTab();
    if (tab === "students")    return this._renderStudentsTab();
    if (tab === "classrooms")  return this._renderClassroomsTab();
    if (tab === "data")        return this._renderDataTab();
    return "";
  },

  _renderUsersTab() {
    const users       = SETUP_DATA.users;
    const activeUsers = users.filter(u => u.active);
    const myEmail     = (AUTH.account?.username || "").toLowerCase().trim();
    return `
      <div class="setup-source-bar">
        <span class="setup-source-label">Source: ${escHtml(SETUP_LISTS.users)}</span>
        <span class="text-muted" style="font-size:12px">${activeUsers.length} active user${activeUsers.length!==1?"s":""}</span>
      </div>
      <div class="form-card">
        <h3>Add User</h3>
        <div class="form-grid">
          <div class="form-field">
            <label class="form-label">Display Name <span class="req">*</span></label>
            <input class="form-input" id="newUserName" placeholder="Last, First" maxlength="80">
          </div>
          <div class="form-field">
            <label class="form-label">Email <span class="req">*</span></label>
            <input class="form-input" id="newUserEmail" type="email" placeholder="user@iu29.org" maxlength="120">
          </div>
          <div class="form-field">
            <label class="form-label">Role <span class="req">*</span></label>
            <select class="form-select" id="newUserRole">
              <option value="">— Select role —</option>
              <option value="Administrator">Administrator</option>
              <option value="Teacher">Teacher</option>
            </select>
          </div>
          <div class="form-field" style="display:flex;align-items:flex-end">
            <button class="btn btn-primary" id="addUserBtn" style="width:100%">Add User</button>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">All Users (${users.length})</div>
        ${users.length === 0
          ? `<div class="empty-state"><div class="empty-icon">👤</div><p>No users found in ${escHtml(SETUP_LISTS.users)}.</p></div>`
          : `<div class="table-wrap"><table><thead><tr><th>Name</th><th>Email</th><th>Role</th><th>Status</th><th></th></tr></thead><tbody>
              ${users.map(u => {
                const isSelf = (u.email || "").toLowerCase().trim() === myEmail;
                return `<tr class="${u.active ? "" : "row-inactive"}">
                  <td><strong>${escHtml(u.name)}</strong>${isSelf ? ' <span class="badge badge-blue">You</span>' : ""}</td>
                  <td class="text-muted">${escHtml(u.email)}</td>
                  <td>${escHtml(u.role)}</td>
                  <td>${u.active ? '<span class="badge badge-green">Active</span>' : '<span class="badge badge-slate">Inactive</span>'}</td>
                  <td>
                    ${isSelf ? "" : u.active
                      ? `<button class="btn btn-danger btn-sm deactivate-user-btn" data-sp-id="${escHtml(u.spId)}" data-name="${escHtml(u.name)}">Deactivate</button>`
                      : `<button class="btn btn-secondary btn-sm reactivate-user-btn" data-sp-id="${escHtml(u.spId)}" data-name="${escHtml(u.name)}">Reactivate</button>`}
                  </td>
                </tr>`;
              }).join("")}
            </tbody></table></div>`}
      </div>`;
  },

  _renderTeachersTab() {
    const teachers = SETUP_DATA.teachers;
    return `
      <div class="setup-source-bar">
        <span class="setup-source-label">Source: ${escHtml(SETUP_LISTS.teachers)}</span>
        <span class="text-muted" style="font-size:12px">${teachers.length} teacher${teachers.length!==1?"s":""}</span>
      </div>
      <div class="form-card">
        <h3>Add Teacher</h3>
        <div class="inline-add">
          <input class="form-input" id="newTeacherName" placeholder="Teacher last name" maxlength="80">
          <button class="btn btn-primary" id="addTeacherBtn">Add</button>
        </div>
        <p class="form-help" style="margin-top:8px">Adds teacher to the "Who Are You Visiting?" list used by walkthroughs.</p>
      </div>
      <div class="card">
        <div class="card-title">Teachers (${teachers.length})</div>
        ${teachers.length === 0
          ? `<div class="empty-state"><div class="empty-icon">👤</div><p>No teachers found.</p></div>`
          : `<div class="table-wrap"><table><thead><tr><th>Name</th><th>App Access</th><th></th></tr></thead><tbody>
              ${teachers.map(t => `<tr>
                <td><strong>${escHtml(t.name)}</strong></td>
                <td>${t.hasAccess
                  ? '<span class="badge badge-green">Has Access</span>'
                  : `<button class="btn btn-secondary btn-sm grant-access-btn" data-name="${escHtml(t.name)}">Grant Access</button>`}
                </td>
                <td><button class="btn btn-danger btn-sm del-teacher-btn" data-sp-id="${escHtml(t.spId)}" data-name="${escHtml(t.name)}">Remove</button></td>
              </tr>`).join("")}
            </tbody></table></div>`}
      </div>`;
  },

  _renderStudentsTab() {
    const students  = SETUP_DATA.students;
    const teachers  = SETUP_DATA.teachers;
    const nextId    = getNextPilotId(students);
    return `
      <div class="setup-source-bar">
        <span class="setup-source-label">Source: ${escHtml(SETUP_LISTS.students)}</span>
        <span class="text-muted" style="font-size:12px">${students.length} student${students.length!==1?"s":""}</span>
      </div>
      <div class="form-card">
        <h3>Add Student</h3>
        <div class="form-grid">
          <div class="form-field">
            <label class="form-label">Pilot ID <span class="req">*</span></label>
            <input class="form-input" id="newStudentPilotId" value="${escHtml(nextId)}" maxlength="20">
          </div>
          <div class="form-field">
            <label class="form-label">Student Name (First + Last Initial) <span class="req">*</span></label>
            <input class="form-input" id="newStudentName" placeholder="e.g. Antonio H." maxlength="60">
          </div>
          <div class="form-field">
            <label class="form-label">Teacher</label>
            <select class="form-select" id="newStudentTeacher">
              <option value="">— Select teacher —</option>
              ${teachers.map(t => `<option value="${escHtml(t.name)}">${escHtml(t.name)}</option>`).join("")}
            </select>
          </div>
          <div class="form-field">
            <label class="form-label">Classroom (optional)</label>
            <input class="form-input" id="newStudentClassroom" placeholder="e.g. Room 12" maxlength="80">
          </div>
          <div class="form-field" style="display:flex;align-items:flex-end">
            <button class="btn btn-primary" id="addStudentBtn" style="width:100%">Add Student</button>
          </div>
        </div>
      </div>
      <div class="card">
        <div class="card-title">Students (${students.length})</div>
        ${students.length === 0
          ? `<div class="empty-state"><div class="empty-icon">🪪</div><p>No students found in SharePoint.</p></div>`
          : `<div class="table-wrap"><table><thead><tr><th>Pilot ID</th><th>Name</th><th>Teacher</th><th>Classroom</th><th>Status</th><th></th></tr></thead><tbody>
              ${students.map(s => `<tr class="${s.active ? "" : "row-inactive"}">
                <td><strong>${escHtml(s.pilotId)}</strong></td>
                <td>${escHtml(s.name)}</td>
                <td class="text-muted">${escHtml(s.teacher) || "—"}</td>
                <td class="text-muted">${escHtml(s.classroom) || "—"}</td>
                <td>${s.active ? '<span class="badge badge-green">Active</span>' : '<span class="badge badge-slate">Inactive</span>'}</td>
                <td>${s.active
                  ? `<button class="btn btn-danger btn-sm deactivate-student-btn" data-sp-id="${escHtml(s.spId)}" data-name="${escHtml(s.name)}">Deactivate</button>`
                  : `<button class="btn btn-secondary btn-sm reactivate-student-btn" data-sp-id="${escHtml(s.spId)}" data-name="${escHtml(s.name)}">Reactivate</button>`}
                </td>
              </tr>`).join("")}
            </tbody></table></div>`}
      </div>`;
  },

  _renderClassroomsTab() {
    const classrooms = SETUP_DATA.classrooms;
    const hasSource  = classrooms.length > 0;
    return `
      <div class="setup-source-bar">
        <span class="setup-source-label">Source: ${escHtml(SETUP_LISTS.classrooms)}</span>
        <span class="text-muted" style="font-size:12px">${classrooms.length} record${classrooms.length!==1?"s":""}</span>
      </div>
      ${!hasSource ? `<div class="warning-banner">
        <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18" style="flex-shrink:0"><path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"/></svg>
        <div>The list <strong>${escHtml(SETUP_LISTS.classrooms)}</strong> returned no records. Classroom management may need a dedicated SharePoint list.</div>
      </div>` : ""}
      <div class="card">
        <div class="card-title">School Settings / Classrooms (${classrooms.length})</div>
        ${classrooms.length === 0
          ? `<div class="empty-state"><div class="empty-icon">🏫</div><p>No classroom records found.</p></div>`
          : `<div class="table-wrap"><table><thead><tr><th>Name</th><th>Teacher</th><th>School</th></tr></thead><tbody>
              ${classrooms.map(c => `<tr>
                <td><strong>${escHtml(c.name)}</strong></td>
                <td class="text-muted">${escHtml(c.teacher) || "—"}</td>
                <td class="text-muted">${escHtml(c.school) || "—"}</td>
              </tr>`).join("")}
            </tbody></table></div>`}
      </div>`;
  },

  _renderDataTab() {
    const records   = DB.getRecords();
    const pulses    = DB.getDailyPulses ? DB.getDailyPulses() : [];
    const paceItems = DB.getPaceLogs    ? DB.getPaceLogs()    : [];
    return `
      <div class="form-card">
        <h3>Export Data</h3>
        <p class="text-muted mb-12">Export all local walkthrough records as a JSON backup.</p>
        <button class="btn btn-primary" id="exportJsonBtn">Export JSON Backup</button>
      </div>
      <div class="form-card">
        <h3>Import Data</h3>
        <p class="text-muted mb-12">Restore from a previously exported JSON file. Replaces existing local data.</p>
        <button class="btn btn-secondary" id="importJsonBtn">Import JSON File</button>
        <input type="file" id="importFileInput" accept=".json" style="display:none">
      </div>
      <div class="form-card" style="border-color:var(--red-border)">
        <h3 style="color:var(--red)">⚠ Clear All Local Data</h3>
        <p class="text-muted mb-12">Permanently removes all local walkthrough records. <strong>Cannot be undone.</strong> Export a backup first. SharePoint data is not affected.</p>
        <p class="text-muted mb-12">Local records: <strong>${records.length} walkthrough${records.length!==1?"s":""}</strong>, <strong>${pulses.length} pulse${pulses.length!==1?"s":""}</strong>, <strong>${paceItems.length} PACE log${paceItems.length!==1?"s":""}</strong></p>
        <button class="btn btn-danger" id="clearDataBtn">Clear All Local Data</button>
      </div>
      <div class="card" style="margin-top:16px">
        <div class="card-title">Local Data Summary (read-only)</div>
        <div class="table-wrap"><table><thead><tr><th>Store</th><th>Count</th></tr></thead><tbody>
          <tr><td>Walkthrough Observations</td><td>${records.length}</td></tr>
          <tr><td>Daily Pulse Entries</td><td>${pulses.length}</td></tr>
          <tr><td>PACE Log Entries</td><td>${paceItems.length}</td></tr>
        </tbody></table></div>
      </div>`;
  },

  _bindSetupTabEvents(tab) {
    const tc = document.getElementById("tabContent");
    if (!tc) return;

    if (tab === "users") {
      document.getElementById("addUserBtn")?.addEventListener("click", async () => {
        const name  = document.getElementById("newUserName").value.trim();
        const email = document.getElementById("newUserEmail").value.trim();
        const role  = document.getElementById("newUserRole").value;
        if (!name || !email || !role) { showToast("Name, email, and role are required.", "error"); return; }
        try {
          requireSetupAdmin();
          const userId = createSetupUserId(role, name);
          await GRAPH.createMappedListItem(SETUP_LISTS.users, {
            "Title":  userId,
            "Name":   name,
            "Role":   role,
            "Email":  email,
            "Active": "Yes"
          });
          showToast(`User "${name}" added.`);
          APP_CACHE.clear("users");
          await SETUP_DATA.refresh();
        } catch (err) { showToast("Failed to add user: " + err.message, "error"); }
      });

      tc.addEventListener("click", async e => {
        const deactivateBtn = e.target.closest(".deactivate-user-btn");
        const reactivateBtn = e.target.closest(".reactivate-user-btn");
        if (deactivateBtn) {
          const { spId, name } = deactivateBtn.dataset;
          if (!confirm(`Deactivate "${name}"? They will lose access to IEP Skook.`)) return;
          try {
            await GRAPH.updateMappedListItem(SETUP_LISTS.users, spId, { "Active": "No" });
            showToast(`"${name}" deactivated.`);
            APP_CACHE.clear("users");
            await SETUP_DATA.refresh();
          } catch (err) { showToast("Failed to deactivate: " + err.message, "error"); }
        }
        if (reactivateBtn) {
          const { spId, name } = reactivateBtn.dataset;
          try {
            await GRAPH.updateMappedListItem(SETUP_LISTS.users, spId, { "Active": "Yes" });
            showToast(`"${name}" reactivated.`);
            APP_CACHE.clear("users");
            await SETUP_DATA.refresh();
          } catch (err) { showToast("Failed to reactivate: " + err.message, "error"); }
        }
      });
    }

    if (tab === "teachers") {
      document.getElementById("addTeacherBtn")?.addEventListener("click", async () => {
        const input = document.getElementById("newTeacherName");
        const name  = input.value.trim();
        if (!name) { showToast("Please enter a teacher name.", "error"); return; }
        try {
          requireSetupAdmin();
          await GRAPH.createListItem(SETUP_LISTS.teachers, { teacher: name });
          showToast(`Teacher "${name}" added.`);
          APP_CACHE.clear("teachers");
          await SETUP_DATA.refresh();
        } catch (err) { showToast("Failed to add teacher: " + err.message, "error"); }
      });

      document.getElementById("newTeacherName")?.addEventListener("keydown", e => {
        if (e.key === "Enter") document.getElementById("addTeacherBtn")?.click();
      });

      tc.addEventListener("click", async e => {
        const btn = e.target.closest(".del-teacher-btn");
        if (!btn) return;
        const { spId, name } = btn.dataset;
        if (!confirm(`Remove "${name}" from the visitor list?`)) return;
        try {
          const siteId = await GRAPH.getSiteId();
          const listId = await GRAPH.getListId(SETUP_LISTS.teachers);
          await GRAPH._patch(`sites/${siteId}/lists/${listId}/items/${spId}`, {});
          showToast(`"${name}" removed.`);
          APP_CACHE.clear("teachers");
          await SETUP_DATA.refresh();
        } catch (err) { showToast("Failed to remove teacher: " + err.message, "error"); }
      });

      tc.addEventListener("click", e => {
        const btn = e.target.closest(".grant-access-btn");
        if (!btn) return;
        showToast("To grant app access, add this teacher to IEP_Users2 on the Users tab with role 'Teacher'.", "error");
      });
    }

    if (tab === "students") {
      document.getElementById("addStudentBtn")?.addEventListener("click", async () => {
        const pilotId   = document.getElementById("newStudentPilotId").value.trim();
        const name      = document.getElementById("newStudentName").value.trim();
        const teacher   = document.getElementById("newStudentTeacher").value;
        const classroom = document.getElementById("newStudentClassroom").value.trim();
        if (!pilotId || !name) { showToast("Pilot ID and student name are required.", "error"); return; }
        try {
          requireSetupAdmin();
          await GRAPH.createMappedListItem(SETUP_LISTS.students, {
            "Title":        pilotId,
            "Student Name": name,
            "Teacher":      teacher,
            "Classroom":    classroom,
            "Active":       "Yes"
          });
          showToast(`Student "${name}" (${pilotId}) added.`);
          APP_CACHE.clear("students");
          await SETUP_DATA.refresh();
        } catch (err) { showToast("Failed to add student: " + err.message, "error"); }
      });

      tc.addEventListener("click", async e => {
        const deactivateBtn = e.target.closest(".deactivate-student-btn");
        const reactivateBtn = e.target.closest(".reactivate-student-btn");
        if (deactivateBtn) {
          const { spId, name } = deactivateBtn.dataset;
          if (!confirm(`Deactivate "${name}"? Their historical records are preserved.`)) return;
          try {
            await GRAPH.updateMappedListItem(SETUP_LISTS.students, spId, { "Active": "No" });
            showToast(`"${name}" deactivated.`);
            APP_CACHE.clear("students");
            await SETUP_DATA.refresh();
          } catch (err) { showToast("Failed to deactivate: " + err.message, "error"); }
        }
        if (reactivateBtn) {
          const { spId, name } = reactivateBtn.dataset;
          try {
            await GRAPH.updateMappedListItem(SETUP_LISTS.students, spId, { "Active": "Yes" });
            showToast(`"${name}" reactivated.`);
            APP_CACHE.clear("students");
            await SETUP_DATA.refresh();
          } catch (err) { showToast("Failed to reactivate: " + err.message, "error"); }
        }
      });
    }

    if (tab === "data") {
      document.getElementById("exportJsonBtn")?.addEventListener("click", () => {
        const ts = new Date().toISOString().slice(0, 10);
        downloadFile(`mac-walkthrough-backup-${ts}.json`, JSON.stringify(DB.exportAll(), null, 2), "application/json");
        showToast("Backup exported.");
      });
      document.getElementById("importJsonBtn")?.addEventListener("click", () => {
        document.getElementById("importFileInput").click();
      });
      document.getElementById("importFileInput")?.addEventListener("change", e => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = ev => {
          try {
            const data = JSON.parse(ev.target.result);
            if (!confirm("Import this backup? Existing local data will be replaced.")) return;
            DB.importAll(data);
            showToast("Data imported successfully.");
            this.renderSetup("data");
          } catch { showToast("Invalid JSON file.", "error"); }
        };
        reader.readAsText(file);
        e.target.value = "";
      });
      document.getElementById("clearDataBtn")?.addEventListener("click", () => {
        if (!confirm("This will permanently delete ALL local data in this browser. Are you absolutely sure?")) return;
        if (!confirm("Last chance — this cannot be undone. Delete everything?")) return;
        DB.clearAll();
        showToast("All local data cleared.");
        this.renderSetup("data");
      });
    }
  },

  /* ── REPORTS ────────────────────────────────────────────────────────────────── */

  renderReports() {
    const el = document.getElementById("page-reports");
    el.innerHTML = `
      <div class="page-header">
        <h2>Reports</h2>
        <p>Walkthrough records loaded from SharePoint.</p>
      </div>

      <div id="reports-status"></div>

      <div class="filter-bar">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
          <div class="card-title" style="margin:0">Filters</div>
          <button class="btn btn-secondary btn-sm" id="refreshReportsBtn">↻ Refresh from SharePoint</button>
        </div>
        <div class="filter-row">
          <div class="form-field">
            <label class="form-label">From</label>
            <input class="form-input" type="date" id="rf-from">
          </div>
          <div class="form-field">
            <label class="form-label">To</label>
            <input class="form-input" type="date" id="rf-to">
          </div>
          <div class="form-field">
            <label class="form-label">Teacher</label>
            <select class="form-select" id="rf-teacher"><option value="">All Teachers</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Classroom</label>
            <select class="form-select" id="rf-classroom"><option value="">All Classrooms</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Student</label>
            <select class="form-select" id="rf-student"><option value="">All Students</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Observer</label>
            <select class="form-select" id="rf-observer"><option value="">All Observers</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Engagement</label>
            <select class="form-select" id="rf-engagement"><option value="">All</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Support Requested</label>
            <select class="form-select" id="rf-support"><option value="">All</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Priority</label>
            <select class="form-select" id="rf-priority"><option value="">All</option></select>
          </div>
          <div class="form-field">
            <label class="form-label">Follow-Up Needed</label>
            <select class="form-select" id="rf-followup">
              <option value="">All</option>
              <option value="yes">Yes</option>
              <option value="no">No</option>
            </select>
          </div>
          <div class="form-field" style="grid-column:1/-1">
            <label class="form-label">Search</label>
            <input class="form-input" type="text" id="rf-search" placeholder="Search notes, wins, concerns, names…">
          </div>
          <div class="form-field" style="display:flex;align-items:flex-end;gap:8px">
            <button class="btn btn-primary" id="applyReportFiltersBtn" style="flex:1">Apply</button>
            <button class="btn btn-secondary" id="clearReportFiltersBtn">Clear</button>
          </div>
        </div>
      </div>

      <div id="reports-summary"></div>

      <div class="report-card student-check-status-card" style="margin-top:20px">
        <div class="report-card-header">
          <div>
            <div class="card-title" style="margin-bottom:4px">Student Check-In Status</div>
            <p style="font-size:13px;color:var(--text-secondary);margin:0">Schoolwide green/yellow/red trends from student check-ins.</p>
          </div>
        </div>
        <div id="student-check-status-report" class="check-status-grid"></div>
      </div>

      <div id="walkthrough-reports-container" style="margin-top:20px"></div>

      <div id="report-detail-modal" class="report-modal hidden" role="dialog" aria-modal="true">
        <div class="report-modal-backdrop" id="report-modal-backdrop"></div>
        <div class="report-modal-panel">
          <div class="report-modal-header">
            <h3 class="report-modal-title">Walkthrough Details</h3>
            <button class="report-modal-close" id="report-modal-close" aria-label="Close">✕</button>
          </div>
          <div class="report-modal-body" id="report-modal-body"></div>
        </div>
      </div>`;

    document.getElementById("applyReportFiltersBtn").addEventListener("click", applyReportsFilters);
    document.getElementById("clearReportFiltersBtn").addEventListener("click", () => {
      ["rf-from","rf-to","rf-teacher","rf-classroom","rf-student","rf-observer",
       "rf-engagement","rf-support","rf-priority","rf-followup","rf-search"].forEach(id => {
        const inp = document.getElementById(id);
        if (inp) inp.value = "";
      });
      applyReportsFilters();
    });
    document.getElementById("refreshReportsBtn").addEventListener("click", () => REPORTS.load());
    document.getElementById("report-modal-close").addEventListener("click", closeReportDetailModal);
    document.getElementById("report-modal-backdrop").addEventListener("click", closeReportDetailModal);

    this._renderCheckStatusPanels();

    if (REPORTS.walkthroughs.length > 0) {
      populateReportsFilterOptions(REPORTS.walkthroughs);
      applyReportsFilters();
    } else if (REPORTS.error) {
      renderReportsErrorState(REPORTS.error);
    } else if (!REPORTS.loading) {
      REPORTS.load();
    } else {
      renderReportsLoadingState();
    }
  },

  _renderCheckStatusPanels() {
    const container = document.getElementById("student-check-status-report");
    if (!container) return;

    const entries = DB.getStudentChecks();
    const today   = new Date();
    const todayStr = today.toISOString().slice(0, 10);

    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
    startOfWeek.setHours(0, 0, 0, 0);

    const normalizeStatus = s => {
      if (s === "ready"        || s === "better-good") return "green";
      if (s === "not-sure"     || s === "same")        return "yellow";
      if (s === "hard-morning" || s === "rough-day")   return "red";
      return null;
    };

    const summarize = filterFn => {
      let green = 0, yellow = 0, red = 0;
      entries.forEach(e => {
        const d = (e.timestamp || e.date || "").slice(0, 10);
        if (!d || !filterFn(d)) return;
        const bucket = normalizeStatus(e.status);
        if (bucket === "green")  green++;
        else if (bucket === "yellow") yellow++;
        else if (bucket === "red")    red++;
      });
      return { green, yellow, red, total: green + yellow + red };
    };

    const windows = [
      { label: "Today",      fn: d => d === todayStr },
      { label: "This Week",  fn: d => d >= startOfWeek.toISOString().slice(0, 10) && d <= todayStr },
      { label: "This Month", fn: d => d.slice(0, 7) === todayStr.slice(0, 7) }
    ];

    container.innerHTML = windows.map(({ label, fn }) => {
      const { green, yellow, red, total } = summarize(fn);
      if (total === 0) {
        return `<div class="check-status-panel">
          <div class="check-status-header">
            <h4>${label}</h4>
            <span class="status-badge status-badge-neutral">No data</span>
          </div>
          <p class="check-status-empty">No student check-in data for this period yet.</p>
        </div>`;
      }
      const gPct = Math.round(green  / total * 100);
      const yPct = Math.round(yellow / total * 100);
      const rPct = 100 - gPct - yPct;
      let dominantLabel = "Mostly Green";
      let dominantMod   = "status-badge-green";
      if (yellow > green && yellow >= red) { dominantLabel = "Mixed / Monitor";     dominantMod = "status-badge-yellow"; }
      if (red    > green && red    > yellow){ dominantLabel = "High Support Need";  dominantMod = "status-badge-red";    }
      return `<div class="check-status-panel">
        <div class="check-status-header">
          <h4>${label}</h4>
          <span class="status-badge ${dominantMod}">${dominantLabel}</span>
        </div>
        <div class="status-counts">
          <span><i class="color-dot dot-green"></i>🟢 ${green} Green</span>
          <span><i class="color-dot dot-yellow"></i>🟡 ${yellow} Yellow</span>
          <span><i class="color-dot dot-red"></i>🔴 ${red} Red</span>
        </div>
        <div class="stacked-status-bar" role="img" aria-label="${label}: ${green} green, ${yellow} yellow, ${red} red">
          <span class="bar-segment bar-green"  style="width:${gPct}%"></span>
          <span class="bar-segment bar-yellow" style="width:${yPct}%"></span>
          <span class="bar-segment bar-red"    style="width:${rPct}%"></span>
        </div>
        <p class="check-status-total">${total} total check-in/check-out entries</p>
      </div>`;
    }).join("");
  },

  _exportCsv(records, teachers, classrooms) {
    const headers = ["Date","WeekOf","Teacher","Classroom","Focus","StudentId","Engagement","Supports","ObservedWin","ConcernGap","SupportNeeded","ClassroomStatus","FollowUpNotes","SubmittedAt"];
    const rows = records.map(r => {
      const t = teachers.find(x=>x.id===r.responses.teacherId);
      const c = classrooms.find(x=>x.id===r.responses.classroomId);
      const cell = v => `"${String(v||"").replace(/"/g,'""')}"`;
      return [
        cell(r.responses.date), cell(r.responses.weekOf),
        cell(t?.name||r.responses.teacherId), cell(c?.name||r.responses.classroomId),
        cell(r.responses.focus), cell(r.responses.studentId),
        cell(r.responses.engagementObserved), cell((r.responses.supportsObserved||[]).join("; ")),
        cell(r.responses.observedWin), cell(r.responses.concernGap),
        cell(r.responses.supportNeeded), cell(r.responses.classroomStatus),
        cell(r.responses.followUpNotes), cell(r.submittedAt)
      ].join(",");
    });
    const ts = new Date().toISOString().slice(0,10);
    downloadFile(`mac-walkthrough-${ts}.csv`, [headers.join(","),...rows].join("\n"), "text/csv");
    showToast("CSV exported.");
  },

  /* ── STORAGE ────────────────────────────────────────────────────────────────── */

  renderStorage() {
    const el         = document.getElementById("page-storage");
    const user       = this.getCurrentUser();
    const records    = DB.getRecords();
    const teachers   = DB.getTeachers();
    const classrooms = DB.getClassrooms();
    const students   = DB.getStudents();

    el.innerHTML = `
      <div class="page-header">
        <h2>Storage Settings</h2>
        <p>How data is stored, and the path to Microsoft 365 integration.</p>
      </div>

      <div class="storage-card" style="border-left:4px solid var(--color-primary)">
        <h3 style="color:var(--color-primary)">
          <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18"><path d="M4 4a2 2 0 00-2 2v1h16V6a2 2 0 00-2-2H4z"/><path fill-rule="evenodd" d="M18 9H2v5a2 2 0 002 2h12a2 2 0 002-2V9zM4 13a1 1 0 011-1h1a1 1 0 110 2H5a1 1 0 01-1-1zm5-1a1 1 0 100 2h1a1 1 0 100-2H9z" clip-rule="evenodd"/></svg>
          Local Prototype Mode — Active
        </h3>
        <p>All data is stored only in <strong>this browser</strong> using <code>localStorage</code>. Not shared. Not backed up automatically. Export a JSON backup before clearing browser data.</p>
        <div class="stat-grid" style="margin:16px 0">
          <div class="stat-card"><div class="stat-label">Walkthroughs</div><div class="stat-value">${records.length}</div></div>
          <div class="stat-card"><div class="stat-label">Teachers</div><div class="stat-value">${teachers.length}</div></div>
          <div class="stat-card"><div class="stat-label">Classrooms</div><div class="stat-value">${classrooms.length}</div></div>
          <div class="stat-card"><div class="stat-label">Placeholder IDs</div><div class="stat-value">${students.length}</div></div>
        </div>
        <div class="btn-group">
          <button class="btn btn-primary" id="storageExportBtn">Export JSON Backup</button>
          <a href="#setup" class="btn btn-secondary">Manage Data in Setup</a>
        </div>
      </div>

      <div class="storage-card" style="border-left:4px solid #94a3b8">
        <h3>Microsoft 365 Connected Mode <span class="coming-soon-badge">Coming Soon</span></h3>
        <p>A future version will connect to Microsoft 365 for centralized, backed-up, multi-device storage — reviewed and managed by IU29 technology leadership.</p>
        <p><strong>Planned architecture:</strong></p>
        <div class="arch-flow">Browser (This App)<br>&nbsp;&nbsp;&nbsp;&nbsp;↓<br>Microsoft Entra ID (Login)<br>&nbsp;&nbsp;&nbsp;&nbsp;↓<br>Power Automate (Workflow)<br>&nbsp;&nbsp;&nbsp;&nbsp;↓<br>Microsoft Lists / SharePoint</div>
        <p>Future storage: <strong>SharePoint Lists</strong>, <strong>Microsoft Lists</strong>, <strong>Microsoft Entra ID</strong>, <strong>Power Automate</strong>.</p>
      </div>

      ${user.isAdmin ? `
      <div class="storage-card" style="border-left:4px solid #7c3aed">
        <h3 style="color:#7c3aed">SharePoint Sync Test — Admin Only</h3>
        <p>Write a test record to <strong>IEP_Daily_Pulse</strong> in SharePoint to confirm Graph write permissions are working. Use the column inspector to see exact internal field names.</p>
        <div class="btn-group" style="margin-top:14px">
          <button class="btn btn-primary" id="testDailyPulseSyncBtn">Test Daily Pulse Sync</button>
          <button class="btn btn-secondary" id="debugSpSchemasBtn">Debug SharePoint Schemas</button>
        </div>
      </div>` : ""}

      <div class="warning-banner">
        <svg viewBox="0 0 20 20" fill="currentColor" width="18" height="18" style="flex-shrink:0"><path fill-rule="evenodd" d="M8.257 3.099c.765-1.36 2.722-1.36 3.486 0l5.58 9.92c.75 1.334-.213 2.98-1.742 2.98H4.42c-1.53 0-2.493-1.646-1.743-2.98l5.58-9.92zM11 13a1 1 0 11-2 0 1 1 0 012 0zm-1-8a1 1 0 00-1 1v3a1 1 0 002 0V6a1 1 0 00-1-1z" clip-rule="evenodd"/></svg>
        <div><strong>Important:</strong> Microsoft 365 integration must be reviewed with IU29 technology leadership before any real student-related information is entered.<br><br>For this local prototype: <strong>use anonymous IDs, initials, or placeholder labels only.</strong></div>
      </div>`;

    document.getElementById("storageExportBtn").addEventListener("click", () => {
      const ts = new Date().toISOString().slice(0,10);
      downloadFile(`mac-walkthrough-backup-${ts}.json`, JSON.stringify(DB.exportAll(),null,2), "application/json");
      showToast("Backup exported.");
    });

    if (user.isAdmin) {
      document.getElementById("debugSpSchemasBtn").addEventListener("click", async () => {
        const btn = document.getElementById("debugSpSchemasBtn");
        btn.disabled = true;
        btn.textContent = "Fetching…";
        try {
          await GRAPH.getListSchema("IEP_Daily_Pulse");
          await GRAPH.getListSchema("IEP_Walkthrough_Observations");
          await GRAPH.getListSchema("IEP_Pace_Visits");
          await GRAPH.getListSchema("IEP_Student_CheckIn");
          await GRAPH.getListSchema("IEP_Weekly_Reflections");
          await GRAPH.getListSchema("IEP_Support_Requests");
          alert("SharePoint schemas logged to console.");
        } catch (err) {
          console.error("getListSchema failed:", err);
          showToast("Schema fetch failed. Check console.", "error");
        } finally {
          btn.disabled = false;
          btn.textContent = "Debug SharePoint Schemas";
        }
      });

      document.getElementById("testDailyPulseSyncBtn").addEventListener("click", async () => {
        const btn = document.getElementById("testDailyPulseSyncBtn");
        btn.disabled = true;
        btn.textContent = "Syncing…";
        try {
          await GRAPH.saveDailyPulse({
            PulseID:        "TEST-" + Date.now(),
            SubmissionDate: new Date().toISOString(),
            StudentName:    "Antonio H.",
            PilotID:        "P001",
            Teacher:        "Stock",
            PulseStatus:    "yellow",
            Categories:     "Transition Difficulty, Self Regulation",
            Note:           "Graph save test from IEP Skook.",
            SupportNeeded:  "Let's talk",
            SubmittedBy:    AUTH.pilotUser?.Name || AUTH.displayName
          });
          showToast("Daily Pulse synced to SharePoint.");
        } catch (err) {
          console.error("Daily Pulse sync failed:", err);
          showToast("Sync failed. Check console.", "error");
        } finally {
          btn.disabled = false;
          btn.textContent = "Test Daily Pulse Sync";
        }
      });
    }
  },

  /* ── FORM LAB ────────────────────────────────────────────────────────────────── */

  renderFormLab() {
    const el = document.getElementById("page-formlab");
    el.innerHTML = `
      <div class="page-header">
        <h2>Form Lab</h2>
        <p>Explore different ways to collect walkthrough data. Nothing here is saved. Tap, select, and evaluate what feels fastest on a phone.</p>
      </div>

      <!-- 1. Standard Dropdown -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">1</span>
          <div>
            <div class="lab-title">Standard Dropdown</div>
            <div class="lab-subtitle">Native &lt;select&gt; element</div>
          </div>
        </div>
        <div class="lab-demo">
          <label class="form-label">Classroom</label>
          <select class="form-input">
            <option value="">— Select —</option>
            <option>Room 101 — Smith</option>
            <option>Room 102 — Johnson</option>
            <option>Room 103 — Williams</option>
            <option>Room 104 — Brown</option>
            <option>Room 105 — Davis</option>
          </select>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Zero code, universally understood, keyboard-accessible, works offline.</div>
          <div class="lab-cons"><strong>Cons:</strong> Opens OS picker on mobile — hard to search large lists, no type-ahead, different look per device.</div>
        </div>
      </div>

      <!-- 2. Button Selection (Option Cards) -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">2</span>
          <div>
            <div class="lab-title">Button Selection</div>
            <div class="lab-subtitle">Large tap-target option cards</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="lab-btn-grid" id="lab-engagement">
            <button class="lab-btn" data-group="engagement" data-val="engaged-independently">Engaged Independently</button>
            <button class="lab-btn" data-group="engagement" data-val="engaged-with-support">Engaged With Support</button>
            <button class="lab-btn" data-group="engagement" data-val="disengaged">Disengaged</button>
            <button class="lab-btn" data-group="engagement" data-val="unclear">Unclear</button>
          </div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> One tap, no scroll, no typing. Very fast on mobile. Ideal for fixed short lists.</div>
          <div class="lab-cons"><strong>Cons:</strong> Doesn't scale — more than 5–6 options gets crowded. Takes vertical space.</div>
        </div>
      </div>

      <!-- 3. Searchable Type-Ahead -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">3</span>
          <div>
            <div class="lab-title">Searchable Type-Ahead</div>
            <div class="lab-subtitle">Auto-show on focus, filter by typing</div>
          </div>
        </div>
        <div class="lab-demo" id="lab-typeahead-wrap">
          ${this._buildSearchSelect(
            DB.getTeachers().map(t => ({ value: t.id, label: t.name })).length > 0
              ? DB.getTeachers().map(t => ({ value: t.id, label: t.name }))
              : CONFIG.PRELOADED_TEACHERS.map((n,i) => ({ value: "t"+i, label: n })),
            "Tap to see all teachers…",
            "lab-teacher",
            false
          )}
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Works for long lists. Shows all options immediately on tap (no typing needed when list is short). Keyboard-searchable.</div>
          <div class="lab-cons"><strong>Cons:</strong> More implementation complexity. Requires JS. Extra tap to close.</div>
        </div>
      </div>

      <!-- 4. Quick Chips (Single-Select) -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">4</span>
          <div>
            <div class="lab-title">Quick Chips — Single Select</div>
            <div class="lab-subtitle">Compact pill toggles, one choice</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="chip-grid" id="lab-focus-chips">
            <label class="chip"><input type="radio" name="lab-focus" value="whole-class"><span>Whole Class</span></label>
            <label class="chip"><input type="radio" name="lab-focus" value="small-group"><span>Small Group</span></label>
            <label class="chip"><input type="radio" name="lab-focus" value="individual-student"><span>Individual Student</span></label>
          </div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Compact, scannable, fast. Good for 3–6 options that fit on one or two rows.</div>
          <div class="lab-cons"><strong>Cons:</strong> Small touch targets if options are long. Can look cluttered beyond 6 items.</div>
        </div>
      </div>

      <!-- 5. Multi-Select Chips -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">5</span>
          <div>
            <div class="lab-title">Multi-Select Chips</div>
            <div class="lab-subtitle">Toggle multiple values independently</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="chip-grid" id="lab-supports-chips">
            ${CONFIG.SUPPORTS_OPTIONS.slice(0,8).map(s =>
              `<label class="chip"><input type="checkbox" name="lab-supports" value="${escHtml(s)}"><span>${escHtml(s)}</span></label>`
            ).join("")}
          </div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Best for "check all that apply" lists. Visual and immediate. Accessible via keyboard.</div>
          <div class="lab-cons"><strong>Cons:</strong> Hard to see full selection state at a glance. Wraps unpredictably on narrow screens if labels are long.</div>
        </div>
      </div>

      <!-- 6. Star Rating -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">6</span>
          <div>
            <div class="lab-title">Star Rating</div>
            <div class="lab-subtitle">1–5 tap-to-rate scale</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="lab-star-label">Overall Lesson Effectiveness</div>
          <div class="lab-stars" id="lab-stars" role="radiogroup" aria-label="Rating 1 to 5">
            <button class="lab-star" data-val="1" aria-label="1 star">★</button>
            <button class="lab-star" data-val="2" aria-label="2 stars">★</button>
            <button class="lab-star" data-val="3" aria-label="3 stars">★</button>
            <button class="lab-star" data-val="4" aria-label="4 stars">★</button>
            <button class="lab-star" data-val="5" aria-label="5 stars">★</button>
          </div>
          <div class="lab-star-val" id="lab-star-val"></div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Intuitive gesture, fast, familiar. Low cognitive load.</div>
          <div class="lab-cons"><strong>Cons:</strong> Subjective scale — raters interpret stars differently. Not ideal for clinical/IEP-aligned data. Provides little actionable information.</div>
        </div>
      </div>

      <!-- 7. Slider -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">7</span>
          <div>
            <div class="lab-title">Slider</div>
            <div class="lab-subtitle">Drag for a numeric value</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="lab-slider-label">Student Engagement Level: <strong id="lab-slider-val">50</strong>%</div>
          <input type="range" id="lab-slider" class="lab-slider" min="0" max="100" value="50" step="5">
          <div class="lab-slider-ticks"><span>0%</span><span>50%</span><span>100%</span></div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Captures nuance. Easy to drag on mobile. Good for continuous scales.</div>
          <div class="lab-cons"><strong>Cons:</strong> Imprecise by default. Hard to reproduce — same person gives different values day to day. Not suited to IEP-objective language.</div>
        </div>
      </div>

      <!-- 8. Voice Note Placeholder -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">8</span>
          <div>
            <div class="lab-title">Voice Note</div>
            <div class="lab-subtitle">Speak your observation (concept demo)</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="lab-voice-area" id="lab-voice">
            <button class="lab-voice-btn" id="lab-voice-btn">
              <svg viewBox="0 0 24 24" fill="currentColor" width="28" height="28"><path d="M12 1a4 4 0 014 4v6a4 4 0 01-8 0V5a4 4 0 014-4zm0 2a2 2 0 00-2 2v6a2 2 0 004 0V5a2 2 0 00-2-2zm-7 9h2a5 5 0 0010 0h2a7 7 0 01-6 6.92V21h2v2H9v-2h2v-2.08A7 7 0 015 12z"/></svg>
              Tap to Record
            </button>
            <div class="lab-voice-transcript" id="lab-voice-transcript">Transcript will appear here…</div>
          </div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Fastest possible input — no typing, no choosing. Excellent for narrative observations. Hands-free while moving through a classroom.</div>
          <div class="lab-cons"><strong>Cons:</strong> Requires microphone permission. Accuracy degrades in noisy classrooms. Transcript needs editing. Privacy questions in a school setting. Not structured data.</div>
        </div>
      </div>

      <!-- 9. Observation Tags -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">9</span>
          <div>
            <div class="lab-title">Observation Tags</div>
            <div class="lab-subtitle">Tap to add tags; tap again to remove</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="lab-tag-pool" id="lab-tag-pool">
            ${["On task","Off task","Redirected","Peer interaction","Adult proximity","Sensory break","Material access","Visual schedule in use","Transition difficulty","Engaged verbally","Working independently","Waiting"].map(t =>
              `<button class="lab-tag" data-tag="${escHtml(t)}">${escHtml(t)}</button>`
            ).join("")}
          </div>
          <div class="lab-tag-selected-label">Selected:</div>
          <div class="lab-tag-selected" id="lab-tag-selected"><em>None</em></div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Structured vocabulary, extremely fast, no typing. Tags are consistent across observers. Easy to filter later.</div>
          <div class="lab-cons"><strong>Cons:</strong> Limited to predefined language. Can feel constraining for nuanced observations. Tag list needs upfront design work.</div>
        </div>
      </div>

      <!-- 10. Quick Note Card -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">10</span>
          <div>
            <div class="lab-title">Quick Note</div>
            <div class="lab-subtitle">Free-text, expandable textarea</div>
          </div>
        </div>
        <div class="lab-demo">
          <textarea class="form-input lab-textarea" id="lab-textarea" rows="3"
            placeholder="Type an observation… (auto-expands)"></textarea>
          <div class="lab-charcount"><span id="lab-charcount">0</span> characters</div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Maximum flexibility. Can capture anything. Natural language is familiar. Works on any device.</div>
          <div class="lab-cons"><strong>Cons:</strong> Slowest input method on mobile. Unstructured — hard to aggregate or filter. Quality varies by observer. Typing while walking is difficult.</div>
        </div>
      </div>

      <!-- 11. Matrix Selection -->
      <div class="lab-card">
        <div class="lab-card-header">
          <span class="lab-num">11</span>
          <div>
            <div class="lab-title">Matrix Selection</div>
            <div class="lab-subtitle">Row × Column grid of cells</div>
          </div>
        </div>
        <div class="lab-demo">
          <div class="lab-matrix-wrap">
            <table class="lab-matrix" id="lab-matrix">
              <thead>
                <tr>
                  <th></th>
                  <th>Independent</th>
                  <th>With Support</th>
                  <th>Not Observed</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td class="lab-matrix-row-label">Reading</td>
                  <td><button class="lab-matrix-cell" data-row="reading" data-col="independent"></button></td>
                  <td><button class="lab-matrix-cell" data-row="reading" data-col="support"></button></td>
                  <td><button class="lab-matrix-cell" data-row="reading" data-col="not-observed"></button></td>
                </tr>
                <tr>
                  <td class="lab-matrix-row-label">Math</td>
                  <td><button class="lab-matrix-cell" data-row="math" data-col="independent"></button></td>
                  <td><button class="lab-matrix-cell" data-row="math" data-col="support"></button></td>
                  <td><button class="lab-matrix-cell" data-row="math" data-col="not-observed"></button></td>
                </tr>
                <tr>
                  <td class="lab-matrix-row-label">Writing</td>
                  <td><button class="lab-matrix-cell" data-row="writing" data-col="independent"></button></td>
                  <td><button class="lab-matrix-cell" data-row="writing" data-col="support"></button></td>
                  <td><button class="lab-matrix-cell" data-row="writing" data-col="not-observed"></button></td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Captures two dimensions at once. Dense information in a small space. Good for progress-monitoring grids.</div>
          <div class="lab-cons"><strong>Cons:</strong> Cognitively demanding. Requires horizontal scrolling on narrow phones. Cells are small touch targets. Complex to build and maintain.</div>
        </div>
      </div>

      <!-- 12. One-Tap Walkthrough -->
      <div class="lab-card lab-card-highlight">
        <div class="lab-card-header">
          <span class="lab-num lab-num-highlight">12</span>
          <div>
            <div class="lab-title">One-Tap Walkthrough</div>
            <div class="lab-subtitle">Minimal form — teacher + status only</div>
          </div>
        </div>
        <div class="lab-demo" id="lab-onetap-wrap">
          <div id="lab-onetap-step1">
            <p class="lab-onetap-hint">Step 1 — Who are you observing?</p>
            <div class="lab-btn-grid" id="lab-onetap-teachers">
              ${(DB.getTeachers().length > 0 ? DB.getTeachers().map(t => t.name) : CONFIG.PRELOADED_TEACHERS).slice(0,6).map(n =>
                `<button class="lab-btn lab-btn-sm" data-onetap-teacher="${escHtml(n)}">${escHtml(n)}</button>`
              ).join("")}
              <button class="lab-btn lab-btn-sm lab-btn-muted" id="lab-onetap-more">+ more…</button>
            </div>
          </div>
          <div id="lab-onetap-step2" class="hidden">
            <p class="lab-onetap-hint">Step 2 — Classroom status?</p>
            <div class="lab-btn-grid">
              ${CONFIG.CLASSROOM_STATUS_OPTIONS.map(o =>
                `<button class="lab-btn lab-onetap-status" data-onetap-status="${escHtml(o.value)}">${CONFIG.STATUS_EMOJI[o.value] || ""} ${escHtml(o.label)}</button>`
              ).join("")}
            </div>
          </div>
          <div id="lab-onetap-step3" class="hidden">
            <div class="lab-onetap-confirm" id="lab-onetap-confirm"></div>
            <button class="lab-onetap-reset" id="lab-onetap-reset">Try Again</button>
          </div>
        </div>
        <div class="lab-meta">
          <div class="lab-pros"><strong>Pros:</strong> Maximum speed — two taps captures teacher + status. Perfect for hallway walk-bys. No cognitive load. Under 5 seconds.</div>
          <div class="lab-cons"><strong>Cons:</strong> Very minimal data. No supports, no notes, no student detail. Trade-off: speed vs. richness. Best as a "quick ping" alongside a richer form.</div>
        </div>
      </div>

      <!-- Summary Card -->
      <div class="lab-summary-card">
        <h3>Which method fits your workflow?</h3>
        <table class="lab-summary-table">
          <thead>
            <tr><th>Method</th><th>Speed</th><th>Data Quality</th><th>Best For</th></tr>
          </thead>
          <tbody>
            <tr><td>Button Selection</td><td>⚡⚡⚡</td><td>Good</td><td>Fixed short lists (engagement, focus)</td></tr>
            <tr><td>One-Tap</td><td>⚡⚡⚡</td><td>Minimal</td><td>High-frequency quick scans</td></tr>
            <tr><td>Quick Chips</td><td>⚡⚡</td><td>Good</td><td>3–6 structured options</td></tr>
            <tr><td>Observation Tags</td><td>⚡⚡</td><td>Good</td><td>Consistent vocabulary across observers</td></tr>
            <tr><td>Searchable Type-Ahead</td><td>⚡⚡</td><td>Good</td><td>Long lists (classrooms, teachers)</td></tr>
            <tr><td>Multi-Select Chips</td><td>⚡⚡</td><td>Good</td><td>"Check all that apply" (supports observed)</td></tr>
            <tr><td>Voice Note</td><td>⚡</td><td>Rich (unstructured)</td><td>Narrative walk-throughs, quiet classrooms</td></tr>
            <tr><td>Quick Note</td><td>Slow</td><td>Rich (unstructured)</td><td>Detailed clinical notes</td></tr>
            <tr><td>Slider / Stars</td><td>⚡⚡</td><td>Subjective</td><td>Impressionistic snapshots only</td></tr>
            <tr><td>Matrix</td><td>Moderate</td><td>Dense</td><td>Progress monitoring grids</td></tr>
          </tbody>
        </table>
        <p class="lab-summary-note">The current MAC Walkthrough uses <strong>Button Selection</strong>, <strong>Quick Chips</strong>, <strong>Multi-Select Chips</strong>, and <strong>Searchable Type-Ahead</strong> — the four fastest structured methods.</p>
      </div>`;

    this._bindFormLab(el);
  },

  _bindFormLab(el) {
    // button-group single select
    el.querySelectorAll(".lab-btn[data-group]").forEach(btn => {
      btn.addEventListener("click", () => {
        const group = btn.dataset.group;
        el.querySelectorAll(`.lab-btn[data-group="${group}"]`).forEach(b => b.classList.remove("lab-btn-active"));
        btn.classList.add("lab-btn-active");
      });
    });

    // bind the search-select in demo 3
    this._bindSearchSelects(el.querySelector("#lab-typeahead-wrap"));

    // chip radios
    el.querySelectorAll(".chip input[type=radio]").forEach(inp => {
      inp.addEventListener("change", () => {
        const name = inp.name;
        el.querySelectorAll(`.chip input[name="${name}"]`).forEach(i => i.closest(".chip").classList.remove("selected"));
        inp.closest(".chip").classList.add("selected");
      });
    });

    // chip checkboxes
    el.querySelectorAll(".chip input[type=checkbox]").forEach(inp => {
      inp.addEventListener("change", () => {
        inp.closest(".chip").classList.toggle("selected", inp.checked);
      });
    });

    // star rating
    const stars   = el.querySelectorAll(".lab-star");
    const starVal = el.getElementById ? el.getElementById("lab-star-val") : el.querySelector("#lab-star-val");
    const starValEl = el.querySelector("#lab-star-val");
    const ratings  = ["","Not Effective","Somewhat Effective","Effective","Very Effective","Highly Effective"];
    stars.forEach(star => {
      star.addEventListener("click", () => {
        const v = Number(star.dataset.val);
        stars.forEach((s,i) => s.classList.toggle("lab-star-active", i < v));
        if (starValEl) starValEl.textContent = ratings[v] || "";
      });
    });

    // slider
    const slider    = el.querySelector("#lab-slider");
    const sliderVal = el.querySelector("#lab-slider-val");
    if (slider && sliderVal) {
      slider.addEventListener("input", () => { sliderVal.textContent = slider.value; });
    }

    // voice note (Web Speech API)
    const voiceBtn        = el.querySelector("#lab-voice-btn");
    const voiceTranscript = el.querySelector("#lab-voice-transcript");
    if (voiceBtn && voiceTranscript) {
      let recognition = null;
      const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
      if (SR) {
        recognition = new SR();
        recognition.continuous = false;
        recognition.interimResults = true;
        recognition.lang = "en-US";
        recognition.onresult = e => {
          const t = Array.from(e.results).map(r => r[0].transcript).join(" ");
          voiceTranscript.textContent = t;
        };
        recognition.onend = () => { voiceBtn.classList.remove("lab-voice-recording"); voiceBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="currentColor" width="28" height="28"><path d="M12 1a4 4 0 014 4v6a4 4 0 01-8 0V5a4 4 0 014-4zm0 2a2 2 0 00-2 2v6a2 2 0 004 0V5a2 2 0 00-2-2zm-7 9h2a5 5 0 0010 0h2a7 7 0 01-6 6.92V21h2v2H9v-2h2v-2.08A7 7 0 015 12z"/></svg> Tap to Record'; };
      }
      voiceBtn.addEventListener("click", () => {
        if (!SR) { voiceTranscript.textContent = "Speech recognition not available in this browser."; return; }
        if (voiceBtn.classList.contains("lab-voice-recording")) {
          recognition.stop();
        } else {
          voiceBtn.classList.add("lab-voice-recording");
          voiceBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="currentColor" width="28" height="28"><path d="M12 1a4 4 0 014 4v6a4 4 0 01-8 0V5a4 4 0 014-4zm0 2a2 2 0 00-2 2v6a2 2 0 004 0V5a2 2 0 00-2-2zm-7 9h2a5 5 0 0010 0h2a7 7 0 01-6 6.92V21h2v2H9v-2h2v-2.08A7 7 0 015 12z"/></svg> Recording…';
          voiceTranscript.textContent = "Listening…";
          recognition.start();
        }
      });
    }

    // observation tags
    const tagPool     = el.querySelector("#lab-tag-pool");
    const tagSelected = el.querySelector("#lab-tag-selected");
    if (tagPool && tagSelected) {
      const selected = new Set();
      tagPool.querySelectorAll(".lab-tag").forEach(tag => {
        tag.addEventListener("click", () => {
          const t = tag.dataset.tag;
          if (selected.has(t)) { selected.delete(t); tag.classList.remove("lab-tag-active"); }
          else                 { selected.add(t);    tag.classList.add("lab-tag-active");    }
          tagSelected.innerHTML = selected.size
            ? [...selected].map(s => `<span class="lab-tag lab-tag-active">${escHtml(s)}</span>`).join("")
            : "<em>None</em>";
        });
      });
    }

    // textarea char count
    const textarea  = el.querySelector("#lab-textarea");
    const charcount = el.querySelector("#lab-charcount");
    if (textarea && charcount) {
      textarea.addEventListener("input", () => {
        charcount.textContent = textarea.value.length;
        textarea.style.height = "auto";
        textarea.style.height = textarea.scrollHeight + "px";
      });
    }

    // matrix
    el.querySelectorAll(".lab-matrix-cell").forEach(cell => {
      cell.addEventListener("click", () => {
        const row = cell.dataset.row;
        el.querySelectorAll(`.lab-matrix-cell[data-row="${row}"]`).forEach(c => c.classList.remove("lab-matrix-cell-active"));
        cell.classList.add("lab-matrix-cell-active");
      });
    });

    // one-tap walkthrough
    const step1    = el.querySelector("#lab-onetap-step1");
    const step2    = el.querySelector("#lab-onetap-step2");
    const step3    = el.querySelector("#lab-onetap-step3");
    const confirm  = el.querySelector("#lab-onetap-confirm");
    const resetBtn = el.querySelector("#lab-onetap-reset");
    let onetapTeacher = "";

    el.querySelectorAll("[data-onetap-teacher]").forEach(btn => {
      btn.addEventListener("click", () => {
        onetapTeacher = btn.dataset.onetapTeacher;
        step1.classList.add("hidden");
        step2.classList.remove("hidden");
      });
    });

    el.querySelectorAll(".lab-onetap-status").forEach(btn => {
      btn.addEventListener("click", () => {
        const status = CONFIG.CLASSROOM_STATUS_OPTIONS.find(o => o.value === btn.dataset.onetapStatus);
        const emoji  = CONFIG.STATUS_EMOJI[btn.dataset.onetapStatus] || "";
        step2.classList.add("hidden");
        step3.classList.remove("hidden");
        if (confirm) confirm.innerHTML = `<div class="lab-onetap-done">${emoji} <strong>${escHtml(onetapTeacher)}</strong><br>${escHtml(status?.label || btn.dataset.onetapStatus)}</div><p class="lab-onetap-note">In a real walkthrough this would be saved. Here it disappears when you reset.</p>`;
      });
    });

    if (resetBtn) {
      resetBtn.addEventListener("click", () => {
        step1.classList.remove("hidden");
        step2.classList.add("hidden");
        step3.classList.add("hidden");
        onetapTeacher = "";
      });
    }
  },

  /* ── PILOT DATA SEEDING ─────────────────────────────────────────────────────── */

  _seedPilotData() {
    if (DB.getTeachers().length > 0) return;
    const teachers = PILOT_TEACHERS.map(pt => ({
      id: pt.id, name: pt.name, createdAt: new Date().toISOString()
    }));
    const classrooms = PILOT_TEACHERS.map(pt => ({
      id:         "classroom-" + pt.id,
      name:       pt.name + "'s Room",
      roomNumber: "",
      teacherId:  pt.id,
      createdAt:  new Date().toISOString()
    }));
    DB._set(CONFIG.STORAGE_KEYS.TEACHERS,   teachers);
    DB._set(CONFIG.STORAGE_KEYS.CLASSROOMS, classrooms);
  },

  // Additive migration — adds Mamrosh without wiping existing records
  _ensureMamrosh() {
    const teachers = DB.getTeachers();
    if (!teachers.find(t => t.id === "teacher-mamrosh")) {
      DB._set(CONFIG.STORAGE_KEYS.TEACHERS, [
        ...teachers,
        { id: "teacher-mamrosh", name: "Mamrosh", createdAt: new Date().toISOString() }
      ]);
    }
    const classrooms = DB.getClassrooms();
    if (!classrooms.find(c => c.id === "classroom-teacher-mamrosh")) {
      DB._set(CONFIG.STORAGE_KEYS.CLASSROOMS, [
        ...classrooms,
        { id: "classroom-teacher-mamrosh", name: "Mamrosh's Room", roomNumber: "", teacherId: "teacher-mamrosh", createdAt: new Date().toISOString() }
      ]);
    }
  }
};

/* ── Simulation Infrastructure ──────────────────────────────────────────────── */

function generateSimulatedPulse(student, status, categories, note) {
  return DB.addDailyPulse({
    student:      student || "",
    pulseStatus:  status || "great",
    categories:   Array.isArray(categories) ? categories : [],
    note:         note || "",
    supportLevel: ""
  });
}

/* ── Boot ────────────────────────────────────────────────────────────────────── */
window.addEventListener("DOMContentLoaded", () => APP.init());
