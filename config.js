const CONFIG = {
  VERSION: "0.1",
  STORAGE_MODE: "local",
  APP_NAME: "MAC Walkthrough",
  ORG_NAME: "IU29",

  STORAGE_KEYS: {
    TEACHERS:       "skookWalkthrough_teachers",
    CLASSROOMS:     "skookWalkthrough_classrooms",
    STUDENTS:       "skookWalkthrough_students",
    RECORDS:        "skookWalkthrough_records",
    SETTINGS:       "skookWalkthrough_settings",
    DAILY_PULSES:   "skookWalkthrough_dailyPulses",
    PACE_LOGS:      "paceLogEntries",
    STUDENT_CHECKS: "studentCheckEntries"
  },

  SUPPORTS_OPTIONS: [
    "Extended Time",
    "Check-In / Check-Out",
    "Preferential Seating",
    "Sensory Break",
    "Modified Assignment",
    "Visual Schedule",
    "Adult Proximity",
    "Peer Support",
    "Chunked Task",
    "Assistive Technology",
    "Small Group Support",
    "Visual Prompt",
    "Verbal Prompt",
    "Movement Break",
    "Reduced Choices",
    "Repeated Directions",
    "Modeling",
    "Positive Reinforcement",
    "First/Then Board",
    "Calming Strategy",
    "Other"
  ],

  ENGAGEMENT_OPTIONS: [
    { value: "engaged-independently", label: "Engaged Independently" },
    { value: "engaged-with-support",  label: "Engaged With Support"  },
    { value: "disengaged",            label: "Disengaged"            },
    { value: "unclear",               label: "Unclear"               }
  ],

  FOCUS_OPTIONS: [
    { value: "whole-class",        label: "Whole Class"        },
    { value: "small-group",        label: "Small Group"        },
    { value: "individual-student", label: "Individual Student" }
  ],

  SUPPORT_NEEDED_OPTIONS: [
    { value: "none",           label: "None"                       },
    { value: "talk-at-weekly", label: "Discuss at Weekly Meeting"  },
    { value: "help-soon",      label: "Help Soon"                  },
    { value: "urgent",         label: "Urgent"                     }
  ],

  CLASSROOM_STATUS_OPTIONS: [
    { value: "on-track",            label: "On Track"            },
    { value: "monitor",             label: "Monitor"             },
    { value: "needs-support",       label: "Needs Support"       },
    { value: "immediate-follow-up", label: "Immediate Follow-Up" }
  ],

  STATUS_EMOJI: {
    "on-track":            "🟢",
    "monitor":             "🟡",
    "needs-support":       "🟠",
    "immediate-follow-up": "🔴"
  },

  STATUS_COLOR_CLASS: {
    "on-track":            "status-green",
    "monitor":             "status-yellow",
    "needs-support":       "status-orange",
    "immediate-follow-up": "status-red"
  },

  SUPPORT_NEEDED_COLOR_CLASS: {
    "none":           "support-none",
    "talk-at-weekly": "support-talk",
    "help-soon":      "support-soon",
    "urgent":         "support-urgent"
  },

  // ── PACE Log Options ──────────────────────────────────────────────────────────
  PACE_BEHAVIOR_OPTIONS: [
    "Disruption", "Defiance / refusal", "Physical aggression", "Verbal aggression",
    "Elopement", "Unsafe behavior", "Peer conflict", "Property damage",
    "Transition difficulty", "Emotional dysregulation", "Other"
  ],

  PACE_INTERVENTION_OPTIONS: [
    "De-escalation conversation", "Calm space", "Sensory support",
    "Restorative conversation", "Problem-solving conference", "Break / reset",
    "Check-in / check-out", "Parent/guardian contact", "Counselor support",
    "Admin support", "Modified task", "Return-to-class plan", "Other"
  ],

  PACE_RETURN_OPTIONS: [
    { value: "returned-to-class",     label: "Returned to class"    },
    { value: "returned-with-support", label: "Returned with support" },
    { value: "stayed-in-pace",        label: "Stayed in PACE"        },
    { value: "sent-to-office",        label: "Sent to office/admin"  },
    { value: "sent-home",             label: "Sent home"             },
    { value: "other",                 label: "Other"                 }
  ],

  PACE_ELEVATED_BEHAVIORS: ["Physical aggression", "Unsafe behavior", "Elopement", "Property damage"],

  // ── Student Check-In Options ──────────────────────────────────────────────────
  CHECKIN_STATUS_OPTIONS: [
    { value: "ready",        label: "Ready",                 emoji: "😊" },
    { value: "not-sure",     label: "Not Sure",              emoji: "😐" },
    { value: "hard-morning", label: "Having a Hard Morning", emoji: "☹️" }
  ],

  CHECKOUT_STATUS_OPTIONS: [
    { value: "better-good", label: "Better / Good",  emoji: "😊" },
    { value: "same",        label: "About the Same", emoji: "😐" },
    { value: "rough-day",   label: "Rough Day",      emoji: "☹️" }
  ],

  // ── Preloaded Staff Directory ─────────────────────────────────────────────────
  // Pilot roster managed via PILOT_TEACHERS in app.js
  PRELOADED_TEACHERS: [],

  // ── 2026–27 Student Roster (Daily Pulse + PACE) ─────────────────────────────
  // Single configuration point for the production 2026–27 SharePoint student
  // list. Do not hard-code students anywhere in source. See STUDENT_ROSTER in
  // app.js, which reads this list by SharePoint DISPLAY name via the existing
  // schema mapper — internal field names are resolved automatically, never
  // guessed.
  //
  // Actual production display columns on IEP_Students_2026_27:
  //   "Student First Name"  — given name
  //   "Student Last Name"   — surname
  //   "Teacher"             — owning teacher's name (Daily Pulse eligibility)
  //   "Classroom"           — display/grouping only, no eligibility effect
  //   "Active"              — Yes/No
  //   "PACE Enabled"        — Yes/No
  //   "Daily Pulse Enabled" — Yes/No
  //
  // There is no "Student Name" or "Student ID" column on this list — the
  // display name is built at runtime from first + last name, and the
  // SharePoint list item's own stable item id is used as the roster id.
  //
  // STUDENT_ROSTER also has a defensive fallback to the old pilot list
  // ("IEP_Skook_Pilot_Students") for use only while this value is still the
  // placeholder ("REPLACE_WITH_2026_27_STUDENT_LIST"). Now that a real
  // production list name is configured below, that fallback no longer runs
  // under normal operation — IEP_Students_2026_27 is authoritative.
  STUDENT_ROSTER_LIST: "IEP_Students_2026_27",

  // PATCH A: new cross-app permission registry (Daily Pulse/PACE/
  // Walkthrough/Admin Panel flags + Teacher/Classroom), separate from
  // IEP_Users2. See iep-app-users.js.
  APP_USERS_LIST: "IEP_App_Users"
};
