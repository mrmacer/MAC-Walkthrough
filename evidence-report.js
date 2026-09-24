/* Teacher Observation Evidence Report.
 *
 * Builds a chronological, per-teacher evidence document from the walkthrough
 * records the Reports page has already loaded and normalized
 * (normalizeSharePointWalkthrough() in app.js) — no second SharePoint read,
 * no data invented. Every value shown here comes from a real stored field;
 * a field that was never collected (Focus/Classroom Status on most records,
 * Disengagement Reasons, Environment, etc.) is simply omitted, the same way
 * the existing Reports detail modal already omits blanks.
 *
 * This module never assigns an evaluation rating (Distinguished/Proficient/
 * Needs Improvement/Failing) and never labels evidence with a specific
 * Danielson domain — it only regroups the walkthrough tool's own "Supports
 * Observed" checklist under plain-language categories, via the explicit,
 * fully-visible SUPPORT_CATEGORY_MAP below. That grouping is a convenience
 * for pasting into PA-ETEP, not a judgment.
 *
 * Everything here runs client-side. Nothing is sent to a third-party
 * service; Word/CSV files are assembled in the browser exactly like
 * pace-export.js does for PACE.
 */

(function (root, factory) {
  const api = factory(root);
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.EVIDENCE_REPORT = api;
})(typeof globalThis !== "undefined" ? globalThis : window, function (root) {
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  // Must stay in sync with CONFIG's real option lists (config.js) — kept as
  // plain local constants so this module has no load-order dependency on
  // CONFIG and can be exercised standalone in tests.
  const ENGAGEMENT_LABELS = {
    "engaged-independently": "Engaged Independently",
    "engaged-with-support":  "Engaged With Support",
    "disengaged":            "Disengaged",
    "unclear":               "Unclear"
  };
  const SUPPORT_NEEDED_LABELS = {
    "none":           "None",
    "talk-at-weekly": "Discuss at Weekly Meeting",
    "help-soon":      "Help Soon",
    "urgent":         "Urgent"
  };
  const FOCUS_LABELS = {
    "whole-class":        "Whole Class",
    "small-group":        "Small Group",
    "individual-student": "Individual Student"
  };

  // Explicit, auditable regrouping of the walkthrough tool's own 21-item
  // "Supports Observed" checklist (CONFIG.SUPPORTS_OPTIONS) into plain-
  // language evidence categories a PA-ETEP entry can be filed under. This is
  // a relabeling of admin-checked boxes, never an inferred rating — every
  // category exists only because the specific box was checked on a real
  // visit. Any option not listed here falls back to "Other".
  const SUPPORT_CATEGORY_MAP = {
    "Extended Time":         "Accommodations & Modifications",
    "Preferential Seating":  "Accommodations & Modifications",
    "Modified Assignment":   "Accommodations & Modifications",
    "Visual Schedule":       "Accommodations & Modifications",
    "Chunked Task":          "Accommodations & Modifications",
    "Visual Prompt":         "Accommodations & Modifications",
    "Verbal Prompt":         "Accommodations & Modifications",
    "Reduced Choices":       "Accommodations & Modifications",
    "Repeated Directions":   "Accommodations & Modifications",
    "Modeling":              "Accommodations & Modifications",
    "Check-In / Check-Out":  "Behavior & Environment Supports",
    "Sensory Break":         "Behavior & Environment Supports",
    "Adult Proximity":       "Behavior & Environment Supports",
    "Movement Break":        "Behavior & Environment Supports",
    "Positive Reinforcement":"Behavior & Environment Supports",
    "First/Then Board":      "Behavior & Environment Supports",
    "Calming Strategy":      "Behavior & Environment Supports",
    "Peer Support":          "Peer & Small-Group Support",
    "Small Group Support":   "Peer & Small-Group Support",
    "Assistive Technology":  "Assistive Technology"
  };
  const CATEGORY_ORDER = [
    "Accommodations & Modifications", "Behavior & Environment Supports",
    "Peer & Small-Group Support", "Assistive Technology", "Other"
  ];
  function categoryFor(support) { return SUPPORT_CATEGORY_MAP[support] || "Other"; }

  const REDACTED_LABEL = "Student (name withheld)";

  /* ── selection ──────────────────────────────────────────────────────── */

  // `records` is REPORTS.walkthroughs (or any subset) — already normalized
  // by app.js's normalizeSharePointWalkthrough(). Returned chronologically
  // ascending (oldest first), the natural reading order for an evidence report.
  function selectRecords(records, filters) {
    const f = filters || {};
    const teacher = String(f.teacher || "").trim().toLowerCase();
    const focus   = String(f.focus   || "").trim().toLowerCase();
    const from    = f.from ? new Date(f.from + "T00:00:00") : null;
    const to      = f.to   ? new Date(f.to   + "T23:59:59") : null;

    return (records || [])
      .filter(r => {
        if (teacher && String(r.teacher || "").trim().toLowerCase() !== teacher) return false;
        if (focus && String(r.focus || "").trim().toLowerCase() !== focus) return false;
        const d = recordDate(r);
        if (from && (!d || d < from)) return false;
        if (to && (!d || d > to)) return false;
        return true;
      })
      .sort((a, b) => (recordDate(a) || 0) - (recordDate(b) || 0));
  }

  function recordDate(r) {
    const raw = r.observationTimestamp
      ? r.observationTimestamp
      : r.observationTime ? `${r.observationDate}T${r.observationTime}` : r.observationDate;
    const d = new Date(raw);
    return isNaN(d) ? null : d;
  }

  function formatDateLong(d) {
    return d ? `${MONTHS[d.getMonth()]} ${d.getDate()}, ${d.getFullYear()}` : "";
  }
  function formatTime(r) {
    return String(r.observationTime || "").trim();
  }

  /* ── entries ────────────────────────────────────────────────────────── */

  // One evidence entry per selected observation. `redactStudents` (default
  // true — see the Privacy note in app.js) replaces a real student name with
  // a generic label; it never alters the source record itself.
  function buildEntries(records, opts) {
    const redact = opts?.redactStudents !== false;
    return records.map(r => {
      const d = recordDate(r);
      const supports = (r.supportsObserved || []).filter(Boolean);
      const byCategory = {};
      supports.forEach(s => {
        const cat = categoryFor(s);
        (byCategory[cat] = byCategory[cat] || []).push(s);
      });
      const studentRaw = String(r.student || "").trim();
      return {
        sharePointId:     r.sharePointId,
        date:             d,
        dateLabel:        formatDateLong(d) || (r.observationDate || "Date not recorded"),
        time:             formatTime(r),
        observer:         r.observer || "",
        classroom:        r.classroom || "",
        studentRaw,
        student:          studentRaw ? (redact ? REDACTED_LABEL : studentRaw) : "",
        focusLabel:       r.focus ? (FOCUS_LABELS[r.focus] || r.focus) : "",
        engagementLabel:  r.engagement ? (ENGAGEMENT_LABELS[r.engagement] || r.engagement) : "",
        supports,
        byCategory,
        disengagementReasons: (r.disengagementReasons || []).filter(Boolean),
        supportRequestedLabel: r.supportRequested ? (SUPPORT_NEEDED_LABELS[r.supportRequested] || r.supportRequested) : "",
        observedWin:      r.observedWin || "",
        concernGap:       r.concernGap || "",
        notes:            r.observationNotes || "",
        followUpNotes:    r.followUpNotes || "",
        followUpNeeded:   !!r.followUpNeeded,
        priority:         r.priority || ""
      };
    });
  }

  function summarize(entries, meta) {
    const dates = entries.map(e => e.date).filter(Boolean);
    const observers = [...new Set(entries.map(e => e.observer).filter(Boolean))].sort((a, b) => a.localeCompare(b));
    const periodFrom = meta?.from ? new Date(meta.from + "T00:00:00") : (dates[0] || null);
    const periodTo   = meta?.to   ? new Date(meta.to   + "T00:00:00") : (dates[dates.length - 1] || null);
    return {
      teacherName: meta?.teacher || "",
      observers,
      periodLabel: periodFrom && periodTo
        ? `${formatDateLong(periodFrom)} – ${formatDateLong(periodTo)}`
        : periodFrom ? `From ${formatDateLong(periodFrom)}` : periodTo ? `Through ${formatDateLong(periodTo)}` : "All dates",
      count: entries.length
    };
  }

  // Aggregates every category mention across all entries, each item tagged
  // with the date it was observed — the PA-ETEP-friendly rollup. Never
  // fabricates a category that wasn't actually checked on a visit.
  function groupByCategory(entries) {
    const map = {};
    entries.forEach(e => {
      Object.entries(e.byCategory).forEach(([cat, supports]) => {
        (map[cat] = map[cat] || []).push({ dateLabel: e.dateLabel, supports });
      });
    });
    return CATEGORY_ORDER
      .filter(cat => map[cat])
      .map(cat => ({ category: cat, items: map[cat] }));
  }

  /* ── CSV ────────────────────────────────────────────────────────────── */

  function csvCell(value) {
    if (value === null || value === undefined) return "";
    let s = String(value);
    if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }

  const CSV_COLUMNS = [
    "Date", "Time", "Observer", "Classroom", "Student", "Focus", "Engagement",
    "Supports Observed", "Disengagement Reasons", "Support Requested",
    "Observed Win", "Concern / Gap", "Observation Notes", "Follow-Up Needed",
    "Follow-Up Notes", "Priority"
  ];

  function buildCsv(entries) {
    const rows = entries.map(e => [
      e.dateLabel, e.time, e.observer, e.classroom, e.student, e.focusLabel, e.engagementLabel,
      e.supports.join("; "), e.disengagementReasons.join("; "), e.supportRequestedLabel,
      e.observedWin, e.concernGap, e.notes, e.followUpNeeded ? "Yes" : "No",
      e.followUpNotes, e.priority
    ].map(csvCell).join(","));
    return "﻿" + [CSV_COLUMNS.join(","), ...rows].join("\r\n") + "\r\n";
  }

  /* ── PA-ETEP plain-text view ───────────────────────────────────────── */

  function line(label, value) { return value ? `${label}: ${value}` : ""; }

  function entrySectionText(e) {
    const lines = [
      `Date: ${e.dateLabel}${e.time ? " · " + e.time : ""}`,
      line("Observer", e.observer),
      line("Classroom", e.classroom),
      line("Focus", e.focusLabel),
      line("Student", e.student),
      "",
      "Evidence:",
      line("  Engagement", e.engagementLabel),
      e.supports.length ? `  Supports Observed: ${e.supports.join(", ")}` : "",
      e.disengagementReasons.length ? `  Disengagement Reasons: ${e.disengagementReasons.join(", ")}` : "",
      line("  Support Requested", e.supportRequestedLabel),
      "",
      "Observer Notes:",
      line("  Observed Win", e.observedWin),
      line("  Concern / Gap", e.concernGap),
      line("  Notes", e.notes),
      line("  Follow-Up Notes", e.followUpNotes),
      e.followUpNeeded ? "  Follow-Up Needed: Yes" : ""
    ].filter(Boolean);
    return lines.join("\n");
  }

  function categorySectionText(groups) {
    if (!groups.length) return "No Supports Observed entries recorded for this period.";
    return groups.map(g =>
      `${g.category}:\n` + g.items.map(i => `  ${i.dateLabel} — ${i.supports.join(", ")}`).join("\n")
    ).join("\n\n");
  }

  function headerSectionText(summary) {
    return [
      `Teacher: ${summary.teacherName}`,
      `Observer(s): ${summary.observers.join(", ") || "Not recorded"}`,
      `Report Period: ${summary.periodLabel}`,
      `Number of Observations: ${summary.count}`
    ].join("\n");
  }

  // Returns { sections: [{id, title, text}], allText } — one section per
  // observation (plus a header and a category rollup), and everything
  // concatenated in reading order for a single "Copy All".
  function buildPaEtepSections(entries, summary) {
    const sections = [{ id: "header", title: "Report Header", text: headerSectionText(summary) }];
    entries.forEach((e, i) => sections.push({
      id: `obs-${i}`,
      title: `Observation — ${e.dateLabel}${e.time ? " " + e.time : ""}`,
      text: entrySectionText(e)
    }));
    sections.push({
      id: "categories",
      title: "Evidence by Category (from Supports Observed — not a domain rating)",
      text: categorySectionText(groupByCategory(entries))
    });
    const allText = sections.map(s => `${s.title}\n${"-".repeat(s.title.length)}\n${s.text}`).join("\n\n");
    return { sections, allText };
  }

  /* ── ZIP (store-only) + DOCX ────────────────────────────────────────── */

  const CRC_TABLE = (() => {
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c >>> 0;
    }
    return table;
  })();
  function crc32(bytes) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }
  function utf8(text) { return new TextEncoder().encode(text); }

  function zip(files, now) {
    const when = now || new Date();
    const dosTime = (when.getHours() << 11) | (when.getMinutes() << 5) | (when.getSeconds() >> 1);
    const dosDate = ((when.getFullYear() - 1980) << 9) | ((when.getMonth() + 1) << 5) | when.getDate();
    const entries = files.map(f => {
      const name = utf8(f.name);
      const data = typeof f.data === "string" ? utf8(f.data) : f.data;
      return { name, data, crc: crc32(data) };
    });
    let size = 22;
    entries.forEach(e => { size += 30 + e.name.length + e.data.length + 46 + e.name.length; });
    const out = new Uint8Array(size);
    const view = new DataView(out.buffer);
    let pos = 0;
    const offsets = [];
    entries.forEach(e => {
      offsets.push(pos);
      view.setUint32(pos, 0x04034b50, true);
      view.setUint16(pos + 4, 20, true);
      view.setUint16(pos + 6, 0x0800, true);
      view.setUint16(pos + 8, 0, true);
      view.setUint16(pos + 10, dosTime, true);
      view.setUint16(pos + 12, dosDate, true);
      view.setUint32(pos + 14, e.crc, true);
      view.setUint32(pos + 18, e.data.length, true);
      view.setUint32(pos + 22, e.data.length, true);
      view.setUint16(pos + 26, e.name.length, true);
      view.setUint16(pos + 28, 0, true);
      out.set(e.name, pos + 30);
      out.set(e.data, pos + 30 + e.name.length);
      pos += 30 + e.name.length + e.data.length;
    });
    const centralStart = pos;
    entries.forEach((e, i) => {
      view.setUint32(pos, 0x02014b50, true);
      view.setUint16(pos + 4, 20, true);
      view.setUint16(pos + 6, 20, true);
      view.setUint16(pos + 8, 0x0800, true);
      view.setUint16(pos + 10, 0, true);
      view.setUint16(pos + 12, dosTime, true);
      view.setUint16(pos + 14, dosDate, true);
      view.setUint32(pos + 16, e.crc, true);
      view.setUint32(pos + 20, e.data.length, true);
      view.setUint32(pos + 24, e.data.length, true);
      view.setUint16(pos + 28, e.name.length, true);
      view.setUint16(pos + 30, 0, true);
      view.setUint16(pos + 32, 0, true);
      view.setUint16(pos + 34, 0, true);
      view.setUint16(pos + 36, 0, true);
      view.setUint32(pos + 38, 0, true);
      view.setUint32(pos + 42, offsets[i], true);
      out.set(e.name, pos + 46);
      pos += 46 + e.name.length;
    });
    view.setUint32(pos, 0x06054b50, true);
    view.setUint16(pos + 4, 0, true);
    view.setUint16(pos + 6, 0, true);
    view.setUint16(pos + 8, entries.length, true);
    view.setUint16(pos + 10, entries.length, true);
    view.setUint32(pos + 12, pos - centralStart, true);
    view.setUint32(pos + 16, centralStart, true);
    view.setUint16(pos + 20, 0, true);
    return out;
  }

  // Strip XML 1.0's forbidden control characters (everything below 0x20
  // except tab/LF/CR) — a stray one in a typed note would otherwise make
  // the whole .docx "corrupt" in Word. Uses explicit \x escapes only, never
  // a raw embedded byte, so the character class can't be silently mangled.
  function xml(text) {
    return String(text === null || text === undefined ? "" : text)
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  function docRun(text, fmt) {
    const f = fmt || {};
    const rPr = (f.bold || f.italic || f.color || f.size)
      ? `<w:rPr>${f.bold ? "<w:b/>" : ""}${f.italic ? "<w:i/>" : ""}${f.color ? `<w:color w:val="${f.color}"/>` : ""}${f.size ? `<w:sz w:val="${f.size}"/>` : ""}</w:rPr>`
      : "";
    const lines = String(text === null || text === undefined ? "" : text).split(/\r\n|\r|\n/);
    return `<w:r>${rPr}${lines.map((l, i) => (i ? "<w:br/>" : "") + `<w:t xml:space="preserve">${xml(l)}</w:t>`).join("")}</w:r>`;
  }
  function docPara(runs, style) {
    return `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ""}${runs.join("")}</w:p>`;
  }
  function labelled(label, value) {
    if (!value) return "";
    return docPara([docRun(label + ": ", { bold: true }), docRun(value)]);
  }

  const DOCX_STYLES = XML_DECL +
    `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:color w:val="64748B"/><w:sz w:val="22"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="1E293B"/></w:pBdr><w:spacing w:before="360" w:after="140"/><w:outlineLvl w:val="0"/></w:pPr>' +
    '<w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
    '</w:styles>';

  function buildDocx(entries, summary, opts) {
    const o = opts || {};
    const body = [
      docPara([docRun("IU29 — Teacher Observation Evidence Report")], "Title"),
      docPara([docRun("Compiled from IEP Skook walkthrough records — for administrator review and evaluation preparation.")], "Subtitle"),
      labelled("Teacher", summary.teacherName),
      labelled("Observer(s)", summary.observers.join(", ") || "Not recorded"),
      labelled("Report Period", summary.periodLabel),
      labelled("Number of Observations", String(summary.count))
    ];

    body.push(docPara([docRun("Chronological Observation Evidence")], "Heading1"));
    if (!entries.length) {
      body.push(docPara([docRun("No walkthrough observations were found for this teacher and date range.", { italic: true })]));
    }
    entries.forEach(e => {
      body.push(docPara([docRun(`${e.dateLabel}${e.time ? " · " + e.time : ""}`)], "Heading2"));
      body.push(labelled("Observer", e.observer));
      body.push(labelled("Classroom", e.classroom));
      body.push(labelled("Focus", e.focusLabel));
      body.push(labelled("Student", e.student));
      body.push(labelled("Engagement", e.engagementLabel));
      body.push(labelled("Supports Observed", e.supports.join(", ")));
      body.push(labelled("Disengagement Reasons", e.disengagementReasons.join(", ")));
      body.push(labelled("Support Requested", e.supportRequestedLabel));
      body.push(labelled("Observed Win", e.observedWin));
      body.push(labelled("Concern / Gap", e.concernGap));
      body.push(labelled("Observation Notes", e.notes));
      body.push(labelled("Follow-Up Notes", e.followUpNotes));
      if (e.followUpNeeded) body.push(labelled("Follow-Up Needed", "Yes"));
    });

    const groups = groupByCategory(entries);
    body.push(docPara([docRun("Evidence by Category")], "Heading1"));
    body.push(docPara([docRun(
      "Regrouped from this teacher's “Supports Observed” selections across the period above. " +
      "This is a literal relabeling of checked items, not a Danielson domain rating or an evaluation score.",
      { italic: true, color: "64748B" }
    )]));
    if (!groups.length) {
      body.push(docPara([docRun("No Supports Observed entries recorded for this period.", { italic: true })]));
    }
    groups.forEach(g => {
      body.push(docPara([docRun(g.category)], "Heading2"));
      g.items.forEach(i => body.push(docPara([docRun(`${i.dateLabel} — ${i.supports.join(", ")}`)])));
    });

    body.push(docPara([docRun("Administrator Synthesis")], "Heading1"));
    body.push(docPara([docRun(
      "The sections below are intentionally blank. Ratings, domain alignment, and narrative synthesis are the administrator's professional judgment and are not generated by this report.",
      { italic: true, color: "64748B" }
    )]));
    ["Strengths Observed", "Areas for Growth", "Recommended Next Steps", "Administrator Notes"].forEach(h => {
      body.push(docPara([docRun(h + ":", { bold: true })]));
      body.push(docPara([docRun("")]));
      body.push(docPara([docRun("")]));
    });

    body.push(docPara([docRun(
      "Confidential — contains protected teacher and student information. Handle according to district policy and evaluation procedures.",
      { italic: true, color: "64748B", size: 18 }
    )]));

    const document = XML_DECL +
      `<w:document xmlns:w="${W_NS}"><w:body>${body.join("")}` +
      '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>' +
      '</w:body></w:document>';

    return zip([
      { name: "[Content_Types].xml", data: XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
        '</Types>' },
      { name: "_rels/.rels", data: XML_DECL +
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DOC_REL}/officeDocument" Target="word/document.xml"/></Relationships>` },
      { name: "word/_rels/document.xml.rels", data: XML_DECL +
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DOC_REL}/styles" Target="styles.xml"/></Relationships>` },
      { name: "word/styles.xml", data: DOCX_STYLES },
      { name: "word/document.xml", data: document }
    ], o.now);
  }

  /* ── filenames ──────────────────────────────────────────────────────── */

  function safeName(text) {
    return String(text || "").normalize("NFC").replace(/[^\p{L}\p{N}]+/gu, "_").replace(/^_+|_+$/g, "");
  }
  function buildFilename(kind, summary) {
    const teacher = safeName(summary.teacherName) || "Teacher";
    const period = safeName(summary.periodLabel) || "All_Dates";
    const ext = kind === "docx" ? "docx" : kind === "csv" ? "csv" : "txt";
    return `Evidence_${teacher}_${period}.${ext}`;
  }

  return {
    ENGAGEMENT_LABELS, SUPPORT_NEEDED_LABELS, FOCUS_LABELS, SUPPORT_CATEGORY_MAP, CATEGORY_ORDER,
    REDACTED_LABEL,
    selectRecords, buildEntries, summarize, groupByCategory,
    buildCsv, buildPaEtepSections, buildDocx, buildFilename,
    zip, crc32
  };
});
