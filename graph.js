function normalizeFollowUp(value) {
  const raw = String(value || "").toLowerCase().trim();
  if (!raw || raw === "none" || raw === "no") return "No";
  if (raw.includes("weekly")) return "Discuss at Weekly Meeting";
  if (raw.includes("soon"))   return "Help Soon";
  if (raw.includes("urgent")) return "Urgent";
  return value || "No";
}

function isWritableSharePointField(internalName) {
  const blocked = new Set([
    "id",
    "ContentType",
    "Modified",
    "Created",
    "Author",
    "Editor",
    "AuthorLookupId",
    "EditorLookupId",
    "_UIVersionString",
    "Attachments",
    "Edit",
    "LinkTitle",
    "LinkTitleNoMenu",
    "ItemChildCount",
    "FolderChildCount",
    "_ComplianceFlags",
    "_ComplianceTag",
    "_ComplianceTagWrittenTime",
    "_ComplianceTagUserId"
  ]);
  if (!internalName) return false;
  if (blocked.has(internalName)) return false;
  if (internalName.startsWith("_")) return false;
  return true;
}

const GRAPH = {
  _BASE:        "https://graph.microsoft.com/v1.0",
  _SITE:        "siu29.sharepoint.com:/sites/IEP_Skook:",
  _siteId:      null,
  _listIdCache:  {},
  _schemaCache:  {},
  _lists:        { users: "IEP_Users2" },

  async _get(path) {
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      console.error("Graph GET failed", path, resp.status, body);
      throw new Error(`Graph ${resp.status}: ${body}`);
    }
    return resp.json();
  },

  async getSiteId() {
    if (this._siteId) return this._siteId;
    const data   = await this._get(`sites/${this._SITE}`);
    window.IEP_AUTH_DEBUG = window.IEP_AUTH_DEBUG || {};
    window.IEP_AUTH_DEBUG.site = data;
    this._siteId = data.id;
    return data.id;
  },

  async getListId(listName) {
    if (this._listIdCache[listName]) return this._listIdCache[listName];
    const siteId = await this.getSiteId();
    const data   = await this._get(`sites/${siteId}/lists?$select=id,name,displayName`);
    window.IEP_AUTH_DEBUG = window.IEP_AUTH_DEBUG || {};
    window.IEP_AUTH_DEBUG.lists = data.value;
    const match  = (data.value || []).find(list => {
      const names = [list.name, list.displayName]
        .map(v => String(v || "").toLowerCase().trim());
      return names.includes(String(listName).toLowerCase().trim());
    });
    if (!match) throw new Error(`SharePoint list not found: ${listName}`);
    this._listIdCache[listName] = match.id;
    return match.id;
  },

  async getListItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data   = await this._get(`sites/${siteId}/lists/${listId}/items?$expand=fields`);
    const items  = (data.value || []).map(item => ({ id: item.id, ...item.fields }));
    // NOTE: this is a generic list reader used for users, students, walkthroughs,
    // Daily Pulse, and PACE. Never dump `items` itself to the console here —
    // that would print full student/staff records for whichever list is being
    // read. window.IEP_AUTH_DEBUG below intentionally stores only field NAMES
    // (used by the sign-in troubleshooting panel), not record values, except
    // for the IEP_Users2 lookup path in findUserByEmail which needs the actual
    // rows to match against the signed-in email.
    if (listName === this._lists.users) {
      window.IEP_AUTH_DEBUG = window.IEP_AUTH_DEBUG || {};
      window.IEP_AUTH_DEBUG.userFieldNames = items[0] ? Object.keys(items[0]) : [];
    }
    return items;
  },

  // PACE visits accumulate beyond Microsoft Graph's default page size. Keep
  // the existing one-page reader stable for current callers, and use this
  // explicit all-pages reader for administrative history/analytics.
  async getAllListItems(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    let path = `sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
    let items = [];
    while (path) {
      const data = await this._get(path);
      items = items.concat(data.value || []);
      const nextLink = data["@odata.nextLink"];
      path = nextLink ? nextLink.replace(`${this._BASE}/`, "") : null;
    }
    return items.map(item => ({ id: item.id, ...item.fields }));
  },

  // Normalize internal SharePoint field keys back to display names. PACE's
  // read model uses only these stable display names and never guesses an
  // internal column name.
  async getListItemsByDisplayName(listName, allPages = false) {
    const [items, schema] = await Promise.all([
      allPages ? this.getAllListItems(listName) : this.getListItems(listName),
      this.getListSchema(listName)
    ]);
    const displayByInternal = {};
    Object.entries(schema).forEach(([displayName, internalName]) => {
      displayByInternal[internalName] = displayName;
    });
    return items.map(item => {
      const row = { id: item.id };
      Object.entries(item).forEach(([key, value]) => {
        if (key !== "id") row[displayByInternal[key] || key] = value;
      });
      return row;
    });
  },

  async findListItemByDisplayField(listName, displayFieldName, value) {
    const [items, schema] = await Promise.all([
      this.getListItems(listName),
      this.getListSchema(listName)
    ]);
    const internalName = schema[displayFieldName];
    if (!internalName) throw new Error(`Column not found in ${listName}: ${displayFieldName}`);
    return items.find(item =>
      String(item[internalName] || "").trim() === String(value || "").trim()
    ) || null;
  },

  async findUserByEmail(email) {
    const users  = await this.getListItems(this._lists.users);
    const target = String(email || "").toLowerCase().trim();
    window.IEP_AUTH_DEBUG = window.IEP_AUTH_DEBUG || {};
    window.IEP_AUTH_DEBUG.targetEmail = target;
    const match = users.find(user => {
      const possibleEmails = [
        user.Email,
        user.email,
        user.EMail,
        user.Email0,
        user.EmailAddress,
        user.Email_x0020_Address,
        user["Email Address"]
      ];
      return possibleEmails.some(value =>
        String(value || "").toLowerCase().trim() === target
      );
    }) || null;
    if (!match) {
      throw new Error(
        "No matching IEP_Users2 Email found for: " + target +
        ". Returned fields: " + JSON.stringify(window.IEP_AUTH_DEBUG.userFieldNames)
      );
    }
    return match;
  },

  async _post(path, body) {
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      method:  "POST",
      headers: {
        Authorization:  `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });
    if (!resp.ok) {
      const text = await resp.text().catch(() => "");
      console.error("Graph POST failed", path, resp.status, text);
      throw new Error(`Graph POST ${resp.status}: ${text}`);
    }
    return resp.json();
  },

  async _patch(path, body) {
    const token    = await AUTH.acquireGraphToken();
    const response = await fetch(`${this._BASE}/${path}`, {
      method:  "PATCH",
      headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json" },
      body:    JSON.stringify(body)
    });
    if (!response.ok) {
      const text = await response.text().catch(() => "");
      console.error("Graph PATCH failed", { path, status: response.status, response: text, body });
      throw new Error(`Graph PATCH ${response.status}: ${text}`);
    }
    if (response.status === 204) return null;
    return response.json();
  },

  async updateMappedListItem(listName, itemId, displayFields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const fields = await this.mapFields(listName, displayFields);
    return this._patch(`sites/${siteId}/lists/${listId}/items/${itemId}/fields`, fields);
  },

  // Administrative correction of an EXISTING walkthrough — always PATCHes
  // IEP_Walkthrough_Observations/{itemId}, never creates a new item. The
  // caller must never include Observation ID / Session ID / Submission ID /
  // Created / Modified / AI Summary / AI Suggestions in `displayFields` —
  // this wrapper does not filter them out, it trusts the caller the same
  // way updateMappedListItem's other callers (Setup) already do.
  async updateWalkthrough(itemId, displayFields) {
    if (!itemId) throw new Error("A walkthrough SharePoint item id is required to save an edit.");
    return this.updateMappedListItem("IEP_Walkthrough_Observations", itemId, displayFields);
  },

  // Administrative correction of an EXISTING PACE visit — always PATCHes
  // IEP_Pace_Visits/{itemId}, never creates a new item. `displayFields`
  // must use the CONFIRMED-live display names already established in
  // pace-admin.js's FIELD_ALIASES (e.g. "Room", "Reason",
  // "Intervention Used") — never the stale "Behavior"/"Interventions"
  // names the legacy savePaceVisit() below still sends. Omit "Time Out"
  // entirely unless the administrator is intentionally changing it — this
  // wrapper does not touch any field not present in `displayFields`.
  async updatePaceVisit(itemId, displayFields) {
    if (!itemId) throw new Error("A PACE visit SharePoint item id is required to save an edit.");
    return this.updateMappedListItem("IEP_Pace_Visits", itemId, displayFields);
  },
// ...existing code...

  async deletePaceVisit(itemId) {
    if (!itemId) {
      throw new Error("A PACE visit SharePoint item ID is required.");
    }
    if (!MAC_ADMIN_PANEL_ALLOWED) {
      throw new Error("Administrator access required.");
    }

    const siteId = await this.getSiteId();
    const listId = await this.getListId("IEP_Pace_Visits");
    const token = await AUTH.acquireGraphToken();

    const response = await fetch(
      `${this._BASE}/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}`,
      {
        method: "DELETE",
        headers: { Authorization: `Bearer ${token}` }
      }
    );

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      throw new Error(`Graph DELETE ${response.status}: ${text}`);
    }

    return true;
  },

// ...existing code...
  async createListItem(listName, fields) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    try {
      return await this._post(`sites/${siteId}/lists/${listId}/items`, { fields });
    } catch (err) {
      // Log which list/field NAMES failed, not the record values (student/staff data).
      console.error("createListItem failed:", { listName, listId, siteId, fieldNames: Object.keys(fields || {}), error: err.message });
      throw err;
    }
  },

  async saveWalkthrough(entry) {
    const submissionId = entry.SubmissionID || entry.submissionId;
    if (!submissionId) throw new Error("Walkthrough Submission ID is required.");

    const existing = await this.findListItemByDisplayField(
      "IEP_Walkthrough_Observations",
      "Submission ID",
      submissionId
    ).catch(() => null);

    if (existing) {
      console.warn("Walkthrough already exists in SharePoint:", submissionId);
      return { duplicatePrevented: true, existingItem: existing };
    }

    console.count("GRAPH.saveWalkthrough called");
    return this.createMappedListItem("IEP_Walkthrough_Observations", {
      "Observation ID":       entry.ObservationID       || `walkthrough-${Date.now()}`,
      "ObservationDate":      entry.ObservationDate      || new Date().toISOString().slice(0, 10),
      "ObservationTime":      entry.ObservationTime      || "",
      "ObservationTimestamp": entry.ObservationTimestamp || new Date().toISOString(),
      "Observer":             entry.Observer        || AUTH.pilotUser?.Name || AUTH.displayName || "",
      "Observer Email":       AUTH.account?.username || "",
      "Teacher":              entry.Teacher         || "",
      "Student":              entry.StudentName     || entry.Student || "",
      "Classroom":            entry.Classroom       || "",
      "Environment":          entry.Environment     || "",
      "Observation Length":   Number(entry.Duration || 0),
      "Engagement":           entry.Engagement      || "",
      "SupportObserved":      Array.isArray(entry.SupportObserved)      ? entry.SupportObserved.join(", ")      : entry.SupportObserved      || "",
      "DisengagementReasons": Array.isArray(entry.DisengagementReasons) ? entry.DisengagementReasons.join(", ") : entry.DisengagementReasons || "",
      "SupportRequested":     Array.isArray(entry.SupportRequested)     ? entry.SupportRequested.join(", ")     : entry.SupportRequested     || "",
      "Observed Win":         entry.ObservedWin     || "",
      "Concern / Gap":        entry.ConcernGap      || "",
      "Observation Notes":    entry.Notes           || "",
      "Follow-Up Notes":      entry.FollowUpNotes   || "",
      "Priority":             entry.Priority        || "",
      "Follow-Up Needed":     !!entry.FollowUpNeeded,
      "AI Summary":           entry.AISummary       || "",
      "AI Suggestions":       entry.AISuggestions   || "",
      "Submission ID":        submissionId,
      "Synced":               true
    });
  },

  async savePaceVisit(entry) {
    // entry.id is the stable ID already generated once by DB.addPaceLog for this
    // specific visit — reuse it as the duplicate-detection key (same submission,
    // not "same student/date", so a student can legitimately visit PACE more
    // than once per day without being blocked).
    const entryId = entry.id || `pace-${Date.now()}`;
    const existing = await this.findListItemByDisplayField(
      "IEP_Pace_Visits",
      "Entry ID",
      entryId
    ).catch(() => null);

    if (existing) {
      console.warn("PACE visit already exists in SharePoint:", entryId);
      return { duplicatePrevented: true, existingItem: existing };
    }

    return this.createMappedListItem("IEP_Pace_Visits", {
      "Entry ID":      entryId,
      "PACE Room":     entry.paceRoom                                                     || "",
      "Student":       entry.studentName                                                  || "",
      "Date":          entry.date                                                         || new Date().toISOString().slice(0, 10),
      "Time In":       entry.timeIn                                                       || "",
      "Time Out":      entry.timeOut                                                      || "",
      "Behavior":      Array.isArray(entry.behaviors) ? entry.behaviors.join(", ") : entry.behaviors || "",
      "Interventions": Array.isArray(entry.interventions) ? entry.interventions.join(", ") : entry.interventions || "",
      "Return Status": entry.returnStatus                                                 || "",
      "SCM Used":      entry.scmUsed === true,
      "Notes":         entry.notes                                                        || "",
      "Submitted By":  entry.submittedByName                                             || AUTH.pilotUser?.Name || "",
      "Submitted At":  entry.timestamp                                                   || new Date().toISOString()
    });
  },

  async saveDailyPulse(entry) {
    const responses = entry.CategoryResponses || {};
    const cats      = Array.isArray(entry.Categories)
      ? entry.Categories.join(", ")
      : (entry.Categories || entry.Category || "");
    // entry.RecordID is the stable ID already generated once by DB.addDailyPulse
    // for this specific submission — reuse it as the duplicate-detection key.
    const recordId = entry.RecordID || entry.PulseID || `pulse-${Date.now()}`;
    const existing = await this.findListItemByDisplayField(
      "IEP_Daily_Pulse",
      "Record ID",
      recordId
    ).catch(() => null);

    if (existing) {
      console.warn("Daily Pulse entry already exists in SharePoint:", recordId);
      return { duplicatePrevented: true, existingItem: existing };
    }

    return this.createMappedListItem("IEP_Daily_Pulse", {
      "Record ID":                recordId,
      "Date":                     entry.Date || entry.SubmissionDate?.slice?.(0, 10)          || new Date().toISOString().slice(0, 10),
      "Teacher":                  entry.Teacher                                               || "",
      "Student":                  entry.Student || entry.StudentName                          || "",
      "Attendance Status":        entry.AttendanceStatus === "absent" ? "Absent" : "Present",
      "Present":                  entry.Present === true,
      "Status":                   entry.Status || entry.PulseStatus                          || "",
      "Category":                 cats,
      "Category Responses":       JSON.stringify(responses),
      "Academic Progress Rating": responses.academicProgress?.rating                         || "",
      "Behavior Rating":          responses.behavior?.rating                                 || "",
      "Communication Rating":     responses.communication?.rating                            || "",
      "Social Skills Rating":     responses.socialSkills?.rating                             || "",
      "Self-Regulation Rating":   responses.selfRegulation?.rating                           || "",
      "Transition Rating":        responses.transitionDifficulty?.rating                     || "",
      "IEP Goal Rating":          responses.iepGoalProgress?.rating                          || "",
      "Quick Note":               entry.QuickNote || entry.Note                              || "",
      "Submitted By":             entry.SubmittedBy                                          || "",
      "Submitted At":             entry.SubmittedAt                                          || new Date().toISOString()
    });
  },

  async getListColumns(listName) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data   = await this._get(
      `sites/${siteId}/lists/${listId}/columns?$select=name,displayName,text,choice,dateTime,boolean,number`
    );
    console.table(data.value.map(col => ({
      displayName: col.displayName,
      name:        col.name,
      type: col.text ? "text" : col.choice ? "choice" : col.dateTime ? "dateTime"
          : col.boolean ? "boolean" : col.number ? "number" : "other"
    })));
    return data.value || [];
  },

  async getListSchema(listName) {
    if (this._schemaCache[listName]) return this._schemaCache[listName];
    const columns = await this.getListColumns(listName);
    const schema  = {};
    columns.forEach(col => {
      if (col.displayName && col.name) schema[col.displayName] = col.name;
    });
    this._schemaCache[listName] = schema;
    console.log(`Schema for ${listName}:`, schema);
    return schema;
  },

  async mapFields(listName, displayFields) {
    const schema = await this.getListSchema(listName);
    const mapped = {};
    Object.entries(displayFields).forEach(([displayName, value]) => {
      if (value === undefined || value === null) return;
      const internalName = schema[displayName];
      if (!internalName) {
        // Field NAME only — never log `value`, it may be student/staff data.
        console.warn(`UNMAPPED FIELD — display: "${displayName}", internalName: (none)`);
        return;
      }
      if (!isWritableSharePointField(internalName)) {
        console.warn(`SKIPPED (read-only) — display: "${displayName}", internalName: "${internalName}"`);
        return;
      }
      mapped[internalName] = value;
    });
    return mapped;
  },

  async createMappedListItem(listName, displayFields) {
    const fields = await this.mapFields(listName, displayFields);
    return this.createListItem(listName, fields);
  },

  async saveWalkthroughV2(entry) {
    return this.createMappedListItem("IEP_Walkthrough_Observations_V2", {
      "Observation ID":       entry.ObservationID  || `v2-${Date.now()}`,
      "ObservationDate":      entry.ObservationDate || new Date().toISOString().slice(0, 10),
      "Observer":             entry.Observer        || AUTH.pilotUser?.Name || AUTH.displayName || "",
      "Observer Email":       AUTH.account?.username || "",
      "Teacher":              entry.Teacher         || "",
      "Student":              entry.StudentName     || entry.Student || "",
      "Classroom":            entry.Classroom       || "",
      "Status":               entry.Status          || "",
      "Engagement":           entry.Engagement      || "",
      "SupportObserved":      Array.isArray(entry.SupportObserved)  ? entry.SupportObserved.join(", ")  : entry.SupportObserved  || "",
      "SupportRequested":     Array.isArray(entry.SupportRequested) ? entry.SupportRequested.join(", ") : entry.SupportRequested || "",
      "Observed Win":         entry.ObservedWin     || "",
      "Concern / Gap":        entry.ConcernGap      || "",
      "Observation Notes":    entry.Notes           || "",
      "Follow-Up Notes":      entry.FollowUpNotes   || "",
      "Follow-Up Needed":     !!entry.FollowUpNeeded,
      "Submission ID":        entry.SubmissionID    || crypto.randomUUID(),
      "Synced":               true
    });
  },

  async getWhoAreYouVisiting() {
    const data = await this._get(
      `sites/${this._SITE}/lists/macwalkthroughwhoareyouvisiting/items?$expand=fields($select=teacher)`
    );
    return data.value
      .map(item => ({
        spId: item.id,
        name: (item.fields?.teacher || "").trim()
      }))
      .filter(t => t.name)
      .sort((a, b) => a.name.localeCompare(b.name));
  }
};
