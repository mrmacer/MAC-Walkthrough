const DB = {

  _get(key) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      console.error("localStorage read error:", key, e);
      return [];
    }
  },

  _set(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
      return true;
    } catch (e) {
      console.error("localStorage write error:", key, e);
      return false;
    }
  },

  _id(prefix) {
    return prefix + "_" + Date.now() + "_" + Math.random().toString(36).slice(2, 7);
  },

  // ── Teachers ────────────────────────────────────────────────────────────────

  getTeachers() { return this._get(CONFIG.STORAGE_KEYS.TEACHERS); },

  addTeacher(name) {
    const list = this.getTeachers();
    const teacher = { id: this._id("teacher"), name: name.trim(), createdAt: new Date().toISOString() };
    list.push(teacher);
    this._set(CONFIG.STORAGE_KEYS.TEACHERS, list);
    return teacher;
  },

  deleteTeacher(id) {
    this._set(CONFIG.STORAGE_KEYS.TEACHERS, this.getTeachers().filter(t => t.id !== id));
  },

  getTeacherById(id) {
    return this.getTeachers().find(t => t.id === id) || null;
  },

  // ── Classrooms ──────────────────────────────────────────────────────────────

  getClassrooms() { return this._get(CONFIG.STORAGE_KEYS.CLASSROOMS); },

  addClassroom(data) {
    const list = this.getClassrooms();
    const classroom = {
      id: this._id("classroom"),
      name: data.name.trim(),
      roomNumber: (data.roomNumber || "").trim(),
      teacherId: data.teacherId || "",
      createdAt: new Date().toISOString()
    };
    list.push(classroom);
    this._set(CONFIG.STORAGE_KEYS.CLASSROOMS, list);
    return classroom;
  },

  deleteClassroom(id) {
    this._set(CONFIG.STORAGE_KEYS.CLASSROOMS, this.getClassrooms().filter(c => c.id !== id));
  },

  getClassroomById(id) {
    return this.getClassrooms().find(c => c.id === id) || null;
  },

  // ── Students (Placeholder IDs) ───────────────────────────────────────────────

  getStudents() { return this._get(CONFIG.STORAGE_KEYS.STUDENTS); },

  addStudent(data) {
    const list = this.getStudents();
    const student = {
      id: this._id("student"),
      initials: (data.initials || "").trim(),
      anonymousId: (data.anonymousId || "").trim(),
      classroomId: data.classroomId || "",
      createdAt: new Date().toISOString()
    };
    list.push(student);
    this._set(CONFIG.STORAGE_KEYS.STUDENTS, list);
    return student;
  },

  deleteStudent(id) {
    this._set(CONFIG.STORAGE_KEYS.STUDENTS, this.getStudents().filter(s => s.id !== id));
  },

  getStudentById(id) {
    return this.getStudents().find(s => s.id === id) || null;
  },

  // ── Records ──────────────────────────────────────────────────────────────────

  getRecords() { return this._get(CONFIG.STORAGE_KEYS.RECORDS); },

  addRecord(responses) {
    const list = this.getRecords();
    const record = {
      id: this._id("record"),
      submittedAt: new Date().toISOString(),
      version: CONFIG.VERSION,
      storageMode: CONFIG.STORAGE_MODE,
      responses
    };
    list.push(record);
    this._set(CONFIG.STORAGE_KEYS.RECORDS, list);
    return record;
  },

  deleteRecord(id) {
    this._set(CONFIG.STORAGE_KEYS.RECORDS, this.getRecords().filter(r => r.id !== id));
  },

  getRecordsThisWeek() {
    const records = this.getRecords();
    const now = new Date();
    const startOfWeek = new Date(now);
    startOfWeek.setDate(now.getDate() - now.getDay());
    startOfWeek.setHours(0, 0, 0, 0);
    return records.filter(r => new Date(r.submittedAt) >= startOfWeek);
  },

  // ── Settings ─────────────────────────────────────────────────────────────────

  getSettings() {
    try {
      const raw = localStorage.getItem(CONFIG.STORAGE_KEYS.SETTINGS);
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  },

  saveSettings(settings) {
    this._set(CONFIG.STORAGE_KEYS.SETTINGS, settings);
  },

  // ── Daily Pulse ──────────────────────────────────────────────────────────────

  getDailyPulses() { return this._get(CONFIG.STORAGE_KEYS.DAILY_PULSES); },

  addDailyPulse(data) {
    const list = this.getDailyPulses();
    const entry = {
      // Reuse a caller-supplied stable id (e.g. a per-draft submission id) so
      // a retry after a failed SharePoint sync is recognized as the same
      // submission rather than generating a fresh, un-deduplicated id.
      id:                data.id || this._id("pulse"),
      timestamp:         new Date().toISOString(),
      date:              data.date || new Date().toISOString().slice(0, 10),
      student:           data.student || "",
      studentId:         data.studentId || "",
      attendanceStatus:  data.attendanceStatus || "present",
      pulseStatus:       data.pulseStatus || "",
      categories:        data.categories || [],
      categoryResponses: data.categoryResponses || {},
      note:              data.note || "",
      supportLevel:      data.supportLevel || "",
      teacherId:         data.teacherId || "",
      teacherName:       data.teacherName || "",
      submittedById:     data.submittedById || "",
      submittedByName:   data.submittedByName || "",
      source:            data.source || "daily-pulse"
    };
    list.push(entry);
    this._set(CONFIG.STORAGE_KEYS.DAILY_PULSES, list);
    return entry;
  },

  deleteDailyPulse(id) {
    this._set(CONFIG.STORAGE_KEYS.DAILY_PULSES, this.getDailyPulses().filter(e => e.id !== id));
  },

  // ── PACE Log ─────────────────────────────────────────────────────────────────

  getPaceLogs() { return this._get(CONFIG.STORAGE_KEYS.PACE_LOGS); },

  addPaceLog(data) {
    const list    = this.getPaceLogs();
    const timeIn  = data.timeIn  || "";
    const timeOut = data.timeOut || "";
    let durationMinutes = null;
    if (timeIn && timeOut) {
      const [ih, im] = timeIn.split(":").map(Number);
      const [oh, om] = timeOut.split(":").map(Number);
      let dur = (oh * 60 + om) - (ih * 60 + im);
      if (dur < 0) dur += 24 * 60;
      durationMinutes = dur;
    }
    const entry = {
      // Reuse a caller-supplied stable id (e.g. a per-form submission id) so a
      // retry after a failed SharePoint sync is recognized as the same
      // submission rather than generating a fresh, un-deduplicated id.
      id:                data.id || ((typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : `pace-${Date.now()}`),
      source:            "pace-log",
      timestamp:         new Date().toISOString(),
      date:              data.date        || new Date().toISOString().slice(0, 10),
      studentId:         data.studentId   || "",
      studentName:       data.studentName || "",
      teacherId:         data.teacherId   || "",
      teacherName:       data.teacherName || "",
      paceRoom:          data.paceRoom    || "",
      timeIn,
      timeOut,
      durationMinutes,
      isOpen:            !timeOut,
      behaviors:         data.behaviors        || [],
      behaviorOther:     data.behaviorOther    || "",
      interventions:     data.interventions    || [],
      interventionOther: data.interventionOther || "",
      returnStatus:      data.returnStatus     || "",
      returnOther:       data.returnOther      || "",
      scmUsed:           data.scmUsed !== undefined ? data.scmUsed : null,
      notes:             data.notes            || "",
      submittedById:     data.submittedById    || "",
      submittedByName:   data.submittedByName  || ""
    };
    list.push(entry);
    this._set(CONFIG.STORAGE_KEYS.PACE_LOGS, list);
    return entry;
  },

  updatePaceLog(id, updates) {
    const list = this.getPaceLogs();
    const idx  = list.findIndex(e => e.id === id);
    if (idx === -1) return null;
    const entry = { ...list[idx], ...updates };
    if (entry.timeIn && entry.timeOut) {
      const [ih, im] = entry.timeIn.split(":").map(Number);
      const [oh, om] = entry.timeOut.split(":").map(Number);
      let dur = (oh * 60 + om) - (ih * 60 + im);
      if (dur < 0) dur += 24 * 60;
      entry.durationMinutes = dur;
      entry.isOpen = false;
    }
    list[idx] = entry;
    this._set(CONFIG.STORAGE_KEYS.PACE_LOGS, list);
    return entry;
  },

  deletePaceLog(id) {
    this._set(CONFIG.STORAGE_KEYS.PACE_LOGS, this.getPaceLogs().filter(e => e.id !== id));
  },

  // ── Student Check-In / Check-Out ─────────────────────────────────────────────

  getStudentChecks() { return this._get(CONFIG.STORAGE_KEYS.STUDENT_CHECKS); },

  addStudentCheck(data) {
    const list  = this.getStudentChecks();
    const entry = {
      id:              (typeof crypto !== "undefined" && crypto.randomUUID) ? crypto.randomUUID() : `check-${Date.now()}`,
      source:          "student-check",
      type:            data.type        || "check-in",
      timestamp:       new Date().toISOString(),
      date:            data.date        || new Date().toISOString().slice(0, 10),
      studentId:       data.studentId   || "",
      studentName:     data.studentName || "",
      teacherId:       data.teacherId   || "",
      teacherName:     data.teacherName || "",
      status:          data.status      || "",
      statusLabel:     data.statusLabel || "",
      note:            data.note        || "",
      entryMode:       "student-facing",
      submittedById:   data.submittedById   || "",
      submittedByName: data.submittedByName || ""
    };
    list.push(entry);
    this._set(CONFIG.STORAGE_KEYS.STUDENT_CHECKS, list);
    return entry;
  },

  deleteStudentCheck(id) {
    this._set(CONFIG.STORAGE_KEYS.STUDENT_CHECKS, this.getStudentChecks().filter(e => e.id !== id));
  },

  // ── Bulk Import / Export / Clear ─────────────────────────────────────────────

  exportAll() {
    return {
      exportedAt:   new Date().toISOString(),
      version:      CONFIG.VERSION,
      teachers:     this.getTeachers(),
      classrooms:   this.getClassrooms(),
      students:     this.getStudents(),
      records:      this.getRecords(),
      dailyPulses:  this.getDailyPulses(),
      paceLogs:      this.getPaceLogs(),
      studentChecks: this.getStudentChecks(),
      settings:      this.getSettings()
    };
  },

  importAll(data) {
    if (Array.isArray(data.teachers))    this._set(CONFIG.STORAGE_KEYS.TEACHERS,     data.teachers);
    if (Array.isArray(data.classrooms))  this._set(CONFIG.STORAGE_KEYS.CLASSROOMS,   data.classrooms);
    if (Array.isArray(data.students))    this._set(CONFIG.STORAGE_KEYS.STUDENTS,     data.students);
    if (Array.isArray(data.records))     this._set(CONFIG.STORAGE_KEYS.RECORDS,      data.records);
    if (Array.isArray(data.dailyPulses)) this._set(CONFIG.STORAGE_KEYS.DAILY_PULSES, data.dailyPulses);
    if (Array.isArray(data.paceLogs))       this._set(CONFIG.STORAGE_KEYS.PACE_LOGS,      data.paceLogs);
    if (Array.isArray(data.studentChecks)) this._set(CONFIG.STORAGE_KEYS.STUDENT_CHECKS, data.studentChecks);
    if (data.settings && typeof data.settings === "object") this._set(CONFIG.STORAGE_KEYS.SETTINGS, data.settings);
  },

  clearAll() {
    Object.values(CONFIG.STORAGE_KEYS).forEach(key => localStorage.removeItem(key));
  }
};
