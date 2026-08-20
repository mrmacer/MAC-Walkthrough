// dashboard-sync.js — Live SharePoint stats for the dashboard
// Walkthroughs + Daily Pulse come from SharePoint (fully synced).
// PACE + Student Check-In come from localStorage (not yet synced to SP).

const DashboardSync = {
  CACHE_MS: 90_000,
  _cache:   {},
  _cacheAt: {},

  invalidate(listName) {
    if (listName) {
      delete this._cache[listName];
      delete this._cacheAt[listName];
    } else {
      this._cache   = {};
      this._cacheAt = {};
    }
  },

  // Fetch all items from a SP list and normalize keys to display names.
  async _getItems(listName) {
    const now = Date.now();
    if (this._cache[listName] && now - this._cacheAt[listName] < this.CACHE_MS) {
      return this._cache[listName];
    }
    const [raw, schema] = await Promise.all([
      GRAPH.getListItems(listName),
      GRAPH.getListSchema(listName)
    ]);
    // Reverse map: internalName → displayName
    const inv = {};
    Object.entries(schema).forEach(([display, internal]) => { inv[internal] = display; });
    const items = raw.map(item => {
      const out = {};
      Object.entries(item).forEach(([k, v]) => { out[inv[k] || k] = v; });
      return out;
    });
    this._cache[listName]   = items;
    this._cacheAt[listName] = now;
    return items;
  },

  async loadWalkthroughStats(user) {
    const all = await this._getItems("IEP_Walkthrough_Observations");

    const items = user.isAdmin
      ? all
      : all.filter(i => (i["Teacher"] || "") === user.name);

    const now = new Date();
    const weekStart = new Date(now);
    weekStart.setDate(now.getDate() - now.getDay());
    weekStart.setHours(0, 0, 0, 0);
    const thisWeek = items.filter(i => {
      const d = i["ObservationDate"];
      return d && new Date(d + "T12:00:00") >= weekStart;
    });

    const followUps    = items.filter(i => { const sr = i["SupportRequested"] || ""; return sr && sr !== "none"; });
    const helpSoon     = items.filter(i => (i["SupportRequested"] || "") === "help-soon");
    const urgent       = items.filter(i => (i["SupportRequested"] || "") === "urgent");
    const teacherSet   = new Set(items.map(i => i["Teacher"]).filter(Boolean));
    const classroomSet = new Set(items.map(i => i["Classroom"]).filter(Boolean));

    const supportCounts = {};
    items.forEach(i => {
      String(i["SupportObserved"] || "").split(",").map(s => s.trim()).filter(Boolean)
        .forEach(s => { supportCounts[s] = (supportCounts[s] || 0) + 1; });
    });
    const topSupports = Object.entries(supportCounts).sort((a, b) => b[1] - a[1]).slice(0, 6);
    const maxSupport  = topSupports[0]?.[1] || 1;

    const statusCounts = { "on-track": 0, "monitor": 0, "needs-support": 0, "immediate-follow-up": 0 };
    items.forEach(i => {
      const cs = i["Classroom Status"] || "";
      if (cs in statusCounts) statusCounts[cs]++;
    });

    const recentItems = [...items]
      .sort((a, b) => {
        const ts = i => String(i["ObservationTimestamp"] || i["ObservationDate"] || "");
        return ts(b).localeCompare(ts(a));
      })
      .slice(0, 5);

    return {
      total:          items.length,
      thisWeek:       thisWeek.length,
      teacherCount:   teacherSet.size,
      classroomCount: classroomSet.size,
      followUps:      followUps.length,
      helpSoon:       helpSoon.length,
      urgent:         urgent.length,
      topSupports,
      maxSupport,
      statusCounts,
      recentItems
    };
  },

  async loadDailyPulseStats(user) {
    const all = await this._getItems("IEP_Daily_Pulse");

    const myStudentNames = user.isAdmin
      ? null
      : new Set(PILOT_STUDENTS.filter(s => s.teacherId === user.id).map(s => s.name));
    const items = user.isAdmin
      ? all
      : all.filter(i => myStudentNames.has(i["Student"] || ""));

    const today = new Date().toISOString().slice(0, 10);
    const normalizePulseStatus = value => {
      const status = String(value || "").trim().toLowerCase();
      if (status === "some struggles") return "struggles";
      if (status === "significant concern") return "concern";
      if (status === "great day") return "great";
      return status;
    };
    const flaggedToday = new Set(
      items
        .filter(i => String(i["Date"] || "").slice(0, 10) === today &&
                     ["struggles", "concern"].includes(normalizePulseStatus(i["Status"])))
        .map(i => i["Student"])
        .filter(Boolean)
    );

    const pulseSupportReqs = items.filter(i => {
      const sl = i["Support Level"] || "";
      return sl === "talk-weekly" || sl === "need-help-now";
    }).length;

    const catCounts = {};
    items.forEach(i => {
      String(i["Category"] || "").split(",").map(c => c.trim()).filter(Boolean)
        .forEach(c => { catCounts[c] = (catCounts[c] || 0) + 1; });
    });
    const topCategory = Object.entries(catCounts).sort((a, b) => b[1] - a[1])[0] || null;

    return {
      total:            items.length,
      flaggedToday:     flaggedToday.size,
      pulseSupportReqs,
      topCategory
    };
  },

  // PACE not yet synced to SharePoint — reads localStorage
  loadPaceStats(user) {
    const all   = DB.getPaceLogs();
    const items = user.isAdmin ? all : all.filter(p => {
      const ps = PILOT_STUDENTS.find(s => s.id === p.studentId);
      return ps?.teacherId === user.id;
    });
    const today     = new Date().toISOString().slice(0, 10);
    const paceToday = items.filter(p => p.date === today).length;
    const paceOpen  = items.filter(p => p.isOpen).length;
    const withDur   = items.filter(p => p.durationMinutes !== null && p.durationMinutes >= 0);
    const avgDur    = withDur.length > 0
      ? Math.round(withDur.reduce((s, p) => s + p.durationMinutes, 0) / withDur.length)
      : null;
    const bCounts = {};
    items.forEach(p => { (p.behaviors || []).forEach(b => { bCounts[b] = (bCounts[b] || 0) + 1; }); });
    const topBehavior = Object.entries(bCounts).sort((a, b) => b[1] - a[1])[0] || null;
    return { paceToday, paceOpen, withDurCount: withDur.length, avgDur, topBehavior };
  },

  // Student Check-In not yet synced to SharePoint — reads localStorage
  loadCheckinStats(user) {
    const all   = DB.getStudentChecks();
    const items = user.isAdmin ? all : all.filter(c => {
      const ps = PILOT_STUDENTS.find(s => s.id === c.studentId);
      return ps?.teacherId === user.id;
    });
    const today  = new Date().toISOString().slice(0, 10);
    const today_ = items.filter(c => c.date === today);
    return {
      checkinsToday:  today_.filter(c => c.type === "check-in").length,
      checkoutsToday: today_.filter(c => c.type === "check-out").length,
      hardMornings:   today_.filter(c => c.status === "hard-morning").length,
      roughEndings:   today_.filter(c => c.status === "rough-day").length
    };
  },

  async refresh(user) {
    const [walk, pulse] = await Promise.all([
      this.loadWalkthroughStats(user),
      this.loadDailyPulseStats(user)
    ]);
    const pace     = this.loadPaceStats(user);
    const checkins = this.loadCheckinStats(user);
    return { walk, pulse, pace, checkins };
  }
};
