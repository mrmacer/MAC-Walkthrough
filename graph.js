function normalizeFollowUp(value) {
  const raw = String(value || "").toLowerCase().trim();
  if (!raw || raw === "none" || raw === "no") return "No";
  if (raw.includes("weekly")) return "Discuss at Weekly Meeting";
  if (raw.includes("soon"))   return "Help Soon";
  if (raw.includes("urgent")) return "Urgent";
  return value || "No";
}

// Computed, read-only columns SharePoint derives from the Title column. When
// Title is renamed, these carry the new display name (see resolveWriteTarget).
const COMPUTED_TITLE_COLUMNS = new Set(["LinkTitle", "LinkTitleNoMenu"]);

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
  _columnCache:  {},
  _lists:        { users: "IEP_Users2" },

  async _get(path) {
    const token = await AUTH.acquireGraphToken();
    const resp  = await fetch(`${this._BASE}/${path}`, {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!resp.ok) {
      const body = await resp.text().catch(() => "");
      console.error("Graph GET failed", path, resp.status, body);
      throw Object.assign(new Error(`Graph ${resp.status}: ${body}`), { graphStatus: resp.status });
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

  // One item by its exact SharePoint id — used to verify a just-saved record
  // independently of list paging or name matching.
  async getListItem(listName, itemId) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const data   = await this._get(`sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}?$expand=fields`);
    return { id: data.id, ...(data.fields || {}) };
  },

  // Writable columns SharePoint marks Required (and gives no default) that a
  // create payload does not supply. `mappedFields` is keyed by INTERNAL name,
  // i.e. the output of mapFields(). Returns display names only. Never invents
  // a value — the caller must refuse to create instead.
  async missingRequiredColumns(listName, mappedFields) {
    await this.getListSchema(listName);
    return (this._columnCache[listName] || [])
      .filter(c => c.required && !c.hidden && !c.readOnly && !c.calculated &&
                   isWritableSharePointField(c.name) && c.defaultValue == null && !(c.name in (mappedFields || {})))
      .map(c => c.displayName || c.name);
  },

  // PACE visits accumulate beyond Microsoft Graph's default page size. Keep
  // the existing one-page reader stable for current callers, and use this
  // explicit all-pages reader for administrative history/analytics.
  // Reads EVERY item of a list, following @odata.nextLink until there is no
  // further page. It is all-or-nothing: if any page fails, comes back
  // malformed, a continuation link is invalid or repeats, or the page limit
  // is hit, it THROWS (error.incompleteLoad = true) instead of returning the
  // pages gathered so far — a partial list must never be mistaken for the
  // whole list. Items repeated across overlapping pages are kept once (by
  // SharePoint item id). The roster, PACE history and other "needs all rows"
  // readers use this; the single-page getListItems() below does not.
  async getAllListItems(listName, options = {}) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const maxPages = options.maxPages || 500;
    const incomplete = why => Object.assign(
      new Error(`Incomplete load of ${listName}: ${why}. The partial result was discarded.`), { incompleteLoad: true });

    const first = `sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
    const requested = new Set([first]);
    const byId = new Map();
    let path = first;
    let pages = 0;
    while (path) {
      if (++pages > maxPages) throw incomplete(`stopped after ${maxPages} pages without reaching the end`);
      let data;
      try {
        data = await this._get(path);
      } catch (err) {
        throw Object.assign(new Error(`Incomplete load of ${listName}: page ${pages} failed (${err.message}). The partial result was discarded.`),
          { incompleteLoad: true, graphStatus: err.graphStatus });
      }
      if (!data || !Array.isArray(data.value)) throw incomplete(`page ${pages} was malformed`);
      for (const item of data.value) {
        if (!item || item.id === undefined || item.id === null || item.id === "") throw incomplete(`page ${pages} contained an item with no id`);
        if (!byId.has(String(item.id))) byId.set(String(item.id), item);
      }
      const next = data["@odata.nextLink"];
      if (next === undefined || next === null || next === "") { path = null; continue; }
      // The continuation must be a Graph URL for THIS list's items — anything else is not trusted.
      if (typeof next !== "string" || !next.startsWith(`${this._BASE}/`) || !next.includes("/items") || !next.includes(listId)) {
        throw incomplete(`page ${pages} returned an invalid continuation link`);
      }
      path = next.slice(this._BASE.length + 1);
      if (requested.has(path)) throw incomplete(`page ${pages} returned a continuation link that was already requested`);
      requested.add(path);
    }
    return [...byId.values()].map(item => ({ id: item.id, ...item.fields }));
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
      // graphStatus marks "SharePoint answered and refused" — as opposed to a
      // network failure (no graphStatus), where the write may or may not have
      // happened. Student saves use this to decide whether a retry is safe.
      throw Object.assign(new Error(`Graph POST ${resp.status}: ${text}`), { graphStatus: resp.status });
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
      throw Object.assign(new Error(`Graph PATCH ${response.status}: ${text}`), { graphStatus: response.status });
    }
    if (response.status === 204) return null;
    return response.json();
  },

  async updateMappedListItem(listName, itemId, displayFields, options) {
    const siteId = await this.getSiteId();
    const listId = await this.getListId(listName);
    const fields = await this.mapFields(listName, displayFields, options);
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
      `sites/${siteId}/lists/${listId}/columns?$select=name,displayName,hidden,readOnly,required,defaultValue,text,choice,dateTime,boolean,number,lookup,personOrGroup,calculated`
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
    const chosenIsUsable = {};
    columns.forEach(col => {
      if (!col.displayName || !col.name) return;
      // SharePoint reports several columns under ONE display name when the
      // built-in Title column has been renamed: the real, writable "Title"
      // plus the computed, read-only "LinkTitle" / "LinkTitleNoMenu". The old
      // plain last-wins assignment let one of those computed columns shadow
      // the real one, so writes to that field were silently skipped as
      // read-only (and reads were inverted through the wrong column). A
      // usable (writable, visible) column always wins; a non-usable one only
      // fills the slot while nothing usable has claimed it. Among usable
      // columns the previous last-wins behaviour is unchanged.
      const usable = isWritableSharePointField(col.name) && !col.hidden && !col.readOnly;
      if (schema[col.displayName] === undefined || usable || !chosenIsUsable[col.displayName]) {
        schema[col.displayName] = col.name;
        chosenIsUsable[col.displayName] = usable;
      }
    });
    this._schemaCache[listName] = schema;
    this._columnCache[listName] = columns; // authoritative metadata, for resolveWriteTarget()
    console.log(`Schema for ${listName}:`, schema);
    return schema;
  },

  // True only for the one list this app manages students in
  // (CONFIG.STUDENT_ROSTER_LIST, i.e. IEP_Students_2026_27). Read at call
  // time; if CONFIG isn't loaded, nothing qualifies.
  _isStudentRosterList(listName) {
    const roster = typeof CONFIG !== "undefined" ? String(CONFIG.STUDENT_ROSTER_LIST || "").trim() : "";
    return !!roster && String(listName || "").trim().toLowerCase() === roster.toLowerCase();
  },

  // What kind of value a column actually holds, from SharePoint's own column
  // metadata (never inferred from its name or from existing data).
  _columnKind(col) {
    if (!col) return "unknown";
    if (col.calculated) return "calculated";
    if (col.boolean)  return "boolean";
    if (col.choice)   return "choice";
    if (col.text)     return "text";
    if (col.number)   return "number";
    if (col.dateTime) return "dateTime";
    if (col.lookup)   return "lookup";
    if (col.personOrGroup) return "personOrGroup";
    return "other";
  },

  // The exact value to send for a Yes/No flag, chosen from the target
  // column's real type: a Boolean column takes true/false; a Choice column
  // takes whichever of its own choices reads Yes/No (or True/False); a Text
  // column takes "Yes"/"No" (the representation this app has always used).
  // Any other type — or a choice list with no clear yes/no — throws instead
  // of guessing, before anything is written.
  async booleanValueFor(listName, displayName, flag) {
    const target = await this.resolveWriteTarget(listName, displayName);
    if (!target.internalName) throw new Error(target.reason);
    const col = (this._columnCache[listName] || []).find(c => c.name === target.internalName);
    const kind = this._columnKind(col);
    if (kind === "boolean") return !!flag;
    if (kind === "text")    return flag ? "Yes" : "No";
    if (kind === "choice") {
      const choices = col.choice.choices || [];
      const want = flag ? ["yes", "true"] : ["no", "false"];
      const hit = choices.find(c => want.includes(String(c).trim().toLowerCase()));
      if (hit === undefined) throw new Error(`"${displayName}" is a Choice column (${choices.join(", ") || "no choices"}) with no clear ${flag ? "Yes" : "No"} option.`);
      return hit;
    }
    throw new Error(`"${displayName}" is a ${kind} column; this app cannot safely store a Yes/No value in it.`);
  },

  // Strict-mode check that a value is something the target column accepts.
  _valueProblem(col, displayName, value) {
    const kind = this._columnKind(col);
    const bad = why => `"${displayName}" ${why}`;
    if (kind === "text") {
      if (typeof value !== "string") return bad("is a text column but the value is not text");
      const max = col.text.maxLength;
      if (max && value.length > max) return bad(`is longer than the column allows (${max} characters)`);
      return null;
    }
    if (kind === "boolean") return typeof value === "boolean" ? null : bad("is a Yes/No (boolean) column and needs true/false, not text");
    if (kind === "choice") {
      const choices = col.choice.choices || [];
      if (typeof value === "string" && (col.choice.allowTextEntry || choices.includes(value))) return null;
      return bad(`is a Choice column and "${typeof value === "string" ? value : typeof value}" is not one of its choices (${choices.join(", ")})`);
    }
    return bad(`is a ${kind} column, which this app cannot safely write`);
  },

  // Which real, writable internal column a display name must be written to.
  //
  // Normally that is just the schema entry. The one exception is a renamed
  // Title column: SharePoint then exposes its value under the new display name
  // ONLY through the computed, read-only "LinkTitle"/"LinkTitleNoMenu"
  // columns, while the writable column underneath keeps the internal name
  // "Title". (Reads already work through the computed column; writes to it are
  // rejected.) That relationship is defined by SharePoint, not guessed — and
  // it is still verified against the list's own column metadata: the
  // translation is made only when no usable column carries the display name
  // AND the metadata contains a real, writable, visible column named exactly
  // "Title". Anything else is returned as-is so the caller rejects it.
  async resolveWriteTarget(listName, displayName) {
    const schema  = await this.getListSchema(listName);
    const columns = this._columnCache[listName] || [];
    const internal = schema[displayName];
    if (!internal) return { internalName: null, reason: `"${displayName}" has no matching column` };
    if (isWritableSharePointField(internal)) return { internalName: internal, translated: false };

    if (COMPUTED_TITLE_COLUMNS.has(internal)) {
      // Deliberately limited to the configured student roster list. Every
      // other list keeps the original behaviour (a computed column is
      // reported as read-only and the field is skipped/rejected).
      if (!this._isStudentRosterList(listName)) {
        return { internalName: null, reason: `"${displayName}" maps to the read-only column "${internal}"` };
      }
      const title = columns.find(c => c.name === "Title");
      if (title && !title.hidden && !title.readOnly) {
        return { internalName: "Title", translated: true, via: internal };
      }
      return { internalName: null, reason: `"${displayName}" maps to the read-only column "${internal}" and the list has no writable "Title" column to write it to` };
    }
    return { internalName: null, reason: `"${displayName}" maps to the read-only column "${internal}"` };
  },

  // By default a field with no writable column is skipped with a console
  // warning (long-standing behaviour other lists rely on for optional
  // columns). With { strict: true } any such field throws BEFORE anything is
  // written, so a caller that must not lose data (the Students page) can
  // never get a half-saved record reported as a success. Two fields that
  // would land in the same column (e.g. a renamed-Title field plus a field
  // that is literally "Title") are always rejected rather than overwritten.
  async mapFields(listName, displayFields, options = {}) {
    const mapped = {};
    const source = {};
    const dropped = [];
    for (const [displayName, value] of Object.entries(displayFields)) {
      if (value === undefined || value === null) continue;
      const target = await this.resolveWriteTarget(listName, displayName);
      if (!target.internalName) {
        // Field NAME only — never log `value`, it may be student/staff data.
        console.warn(`SKIPPED — display: "${displayName}": ${target.reason}`);
        dropped.push(target.reason);
        continue;
      }
      if (target.translated) {
        console.warn(`Column "${displayName}" is the computed "${target.via}"; writing to its underlying "Title" column.`);
      }
      if (source[target.internalName] !== undefined && source[target.internalName] !== displayName) {
        dropped.push(`"${displayName}" and "${source[target.internalName]}" would both write to the column "${target.internalName}"`);
        continue;
      }
      if (options.strict) {
        const col = (this._columnCache[listName] || []).find(c => c.name === target.internalName);
        const problem = this._valueProblem(col, displayName, value);
        if (problem) { dropped.push(problem); continue; }
      }
      mapped[target.internalName] = value;
      source[target.internalName] = displayName;
    }
    if (options.strict && dropped.length) {
      throw new Error(`Column mapping problem in ${listName} — nothing was saved. ${dropped.join("; ")}.`);
    }
    return mapped;
  },

  async createMappedListItem(listName, displayFields, options) {
    const fields = await this.mapFields(listName, displayFields, options);
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
  }
  // NOTE: getWhoAreYouVisiting() (macwalkthroughwhoareyouvisiting) was
  // removed here — the New Walkthrough teacher directory and Setup →
  // Teachers now both derive teachers from IEP_Users2 (active, Role
  // includes Teacher) instead. See TEACHER_DIRECTORY in app.js. That
  // SharePoint list itself has not been touched or deleted.
};
