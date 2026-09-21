/* PACE administrator export.
 *
 * Builds CSV, Excel (.xlsx), Word (.docx) and JSON downloads from the
 * ALREADY-FILTERED in-memory PACE visits the #pace page is showing. It never
 * queries SharePoint, never writes anything, and never sends data anywhere:
 * every file is generated locally in the browser. .xlsx and .docx are plain
 * OOXML packages assembled here with a small store-only ZIP writer, so no
 * third-party library (and no CDN request carrying student data) is involved.
 *
 * SharePoint item ids are deliberately not exported — they are internal
 * plumbing, not report content.
 */

(function (root, factory) {
  const api = factory(root);
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.PACE_EXPORT = api;
})(typeof globalThis !== "undefined" ? globalThis : window, function (root) {
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  // Human-readable columns, in order. `optional` columns are dropped when no
  // exported visit has a value for them (e.g. Return Status, which the PACE
  // Room Tracker never fills in).
  const COLUMNS = [
    { key: "date",            header: "Date",                   type: "date" },
    { key: "startTime",       header: "Start Time" },
    { key: "endTime",         header: "End Time" },
    { key: "student",         header: "Student" },
    { key: "duration",        header: "Duration (min)",         type: "number" },
    { key: "reason",          header: "Reason" },
    { key: "support",         header: "Support / Intervention" },
    { key: "notes",           header: "Notes",                  wrap: true, width: 60 },
    { key: "specialist",      header: "Behavior Specialist" },
    { key: "teacherCameFrom", header: "Teacher Came From" },
    { key: "room",            header: "PACE Room" },
    { key: "scm",             header: "SCM" },
    { key: "status",          header: "Visit Status" },
    { key: "returnStatus",    header: "Return Status",          optional: true }
  ];

  /* ── formatting ─────────────────────────────────────────────────────── */

  function formatDateLong(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ""));
    return m ? `${MONTHS[Number(m[2]) - 1]} ${Number(m[3])}, ${m[1]}` : "";
  }

  function formatTime12(value) {
    const raw = String(value || "").trim();
    const m = /^(\d{1,2}):(\d{2})/.exec(raw);
    if (!m) return raw;
    const h = Number(m[1]);
    return `${h % 12 === 0 ? 12 : h % 12}:${m[2]} ${h >= 12 ? "PM" : "AM"}`;
  }

  function roomLabel(value) {
    const admin = root.PACE_ADMIN || (typeof require === "function" ? require("./pace-admin.js") : null);
    const info = admin ? admin.roomInfo(value) : null;
    if (!info) return String(value || "");
    return info.hallway ? `${info.label} · ${info.hallway}` : info.label;
  }

  function scmLabel(value) {
    return value === true ? "Yes" : value === false ? "No" : "";
  }

  /* ── rows ───────────────────────────────────────────────────────────── */

  function buildRows(visits) {
    return (visits || []).map(v => ({
      date:            v.date || "",
      startTime:       formatTime12(v.timeIn),
      endTime:         formatTime12(v.timeOut),
      student:         v.student || "",
      duration:        typeof v.durationMinutes === "number" ? v.durationMinutes : null,
      reason:          (v.reasons || []).join("; "),
      support:         (v.supports || []).join("; "),
      notes:           v.notes || "",
      specialist:      (v.specialists || []).join("; "),
      teacherCameFrom: v.teacherCameFrom || "",
      room:            roomLabel(v.paceRoom),
      scm:             scmLabel(v.scmUsed),
      status:          v.isCompleted ? "Completed" : "Open",
      returnStatus:    v.returnStatus || ""
    }));
  }

  function columnsFor(rows) {
    return COLUMNS.filter(col => !col.optional || rows.some(r => r[col.key] !== "" && r[col.key] !== null));
  }

  /* ── CSV ────────────────────────────────────────────────────────────── */

  // A cell starting with = + - @ (or a control char) is interpreted as a
  // formula by Excel/Sheets when a CSV is opened. Free-text fields (notes)
  // can legitimately start with "-", so neutralise with a leading apostrophe
  // rather than dropping data. Numbers are never text and are unaffected.
  function csvCell(value) {
    if (value === null || value === undefined) return "";
    let s = typeof value === "number" ? String(value) : String(value);
    if (typeof value === "string" && /^[=+\-@\t\r]/.test(s)) s = "'" + s;
    return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  }

  function toCsv(rows) {
    const cols = columnsFor(rows);
    const lines = [cols.map(c => csvCell(c.header)).join(",")];
    rows.forEach(r => lines.push(cols.map(c => csvCell(r[c.key])).join(",")));
    // UTF-8 BOM so Excel opens accented names correctly; CRLF per RFC 4180.
    return "﻿" + lines.join("\r\n") + "\r\n";
  }

  /* ── JSON ───────────────────────────────────────────────────────────── */

  function toJson(visits) {
    return JSON.stringify((visits || []).map(v => {
      const copy = { ...v };
      delete copy.id; // internal SharePoint item id — not report content
      return copy;
    }), null, 2);
  }

  /* ── ZIP (store-only) ───────────────────────────────────────────────── */

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
      view.setUint16(pos + 6, 0x0800, true);   // UTF-8 file names
      view.setUint16(pos + 8, 0, true);        // stored
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

  /* ── XML helpers ────────────────────────────────────────────────────── */

  // XML 1.0 forbids most control characters; a stray one in a note would make
  // the whole file "corrupt" in Excel/Word, so strip them.
  function xml(text) {
    return String(text === null || text === undefined ? "" : text)
      .replace(/[ --￾￿]/g, "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const REL_NS   = "http://schemas.openxmlformats.org/package/2006/relationships";
  const DOC_REL  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

  /* ── XLSX ───────────────────────────────────────────────────────────── */

  function colLetter(index) {
    let n = index + 1, s = "";
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // Excel date serial for a calendar date, computed from the ISO string
  // itself (no timezone involved), so the date can never shift by a day.
  function excelSerial(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ""));
    if (!m) return null;
    return (Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) - Date.UTC(1899, 11, 30)) / 86400000;
  }

  // Style ids (see STYLES_XML): 1 header, 2 date, 3 wrapped text, 4 text, 5 number
  function xlsxCell(col, value, ref) {
    if (col.type === "date") {
      const serial = excelSerial(value);
      return serial === null ? "" : `<c r="${ref}" s="2"><v>${serial}</v></c>`;
    }
    if (col.type === "number") {
      return typeof value === "number" ? `<c r="${ref}" s="5"><v>${value}</v></c>` : "";
    }
    if (value === "" || value === null || value === undefined) return "";
    return `<c r="${ref}" t="inlineStr" s="${col.wrap ? 3 : 4}"><is><t xml:space="preserve">${xml(value)}</t></is></c>`;
  }

  const STYLES_XML = XML_DECL +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/></numFmts>' +
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
    '<fill><patternFill patternType="solid"><fgColor rgb="FFE2E8F0"/><bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>' +
    '<border><left/><right/><top/><bottom style="thin"><color rgb="FF94A3B8"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right" vertical="top"/></xf>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function columnWidth(col, rows) {
    if (col.width) return col.width;
    if (col.type === "date") return 12;
    let longest = col.header.length;
    rows.forEach(r => {
      const v = r[col.key];
      const firstLine = v === null || v === undefined ? "" : String(v).split(/\r?\n/)[0];
      if (firstLine.length > longest) longest = firstLine.length;
    });
    return Math.min(Math.max(longest + 3, 10), 45);
  }

  function buildXlsx(rows, options) {
    const cols = columnsFor(rows);
    const lastCol = colLetter(cols.length - 1);
    const lastRow = rows.length + 1;

    const header = `<row r="1">${cols.map((c, i) =>
      `<c r="${colLetter(i)}1" t="inlineStr" s="1"><is><t>${xml(c.header)}</t></is></c>`).join("")}</row>`;
    const body = rows.map((r, ri) => {
      const n = ri + 2;
      return `<row r="${n}">${cols.map((c, ci) => xlsxCell(c, r[c.key], colLetter(ci) + n)).join("")}</row>`;
    }).join("");

    const sheet = XML_DECL +
      `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
      `<dimension ref="A1:${lastCol}${lastRow}"/>` +
      `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>` +
      `<selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>` +
      `<cols>${cols.map((c, i) => `<col min="${i + 1}" max="${i + 1}" width="${columnWidth(c, rows)}" customWidth="1"/>`).join("")}</cols>` +
      `<sheetData>${header}${body}</sheetData>` +
      `<autoFilter ref="A1:${lastCol}${lastRow}"/>` +
      `</worksheet>`;

    return zip([
      { name: "[Content_Types].xml", data: XML_DECL +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>' },
      { name: "_rels/.rels", data: XML_DECL +
        `<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${DOC_REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
      { name: "xl/workbook.xml", data: XML_DECL +
        `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${DOC_REL}">` +
        '<sheets><sheet name="PACE Log" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { name: "xl/_rels/workbook.xml.rels", data: XML_DECL +
        `<Relationships xmlns="${REL_NS}">` +
        `<Relationship Id="rId1" Type="${DOC_REL}/worksheet" Target="worksheets/sheet1.xml"/>` +
        `<Relationship Id="rId2" Type="${DOC_REL}/styles" Target="styles.xml"/></Relationships>` },
      { name: "xl/styles.xml", data: STYLES_XML },
      { name: "xl/worksheets/sheet1.xml", data: sheet }
    ], options && options.now);
  }

  /* ── DOCX (single-student summary) ──────────────────────────────────── */

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
    return docPara([docRun(label + ": ", { bold: true }), docRun(value)]);
  }

  const DOCX_STYLES = XML_DECL +
    `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>` +
    '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>' +
    '<w:pPrDefault><w:pPr><w:spacing w:after="80" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:spacing w:after="160"/></w:pPr><w:rPr><w:b/><w:sz w:val="40"/><w:szCs w:val="40"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:spacing w:before="320" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:style>' +
    '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="CBD5E1"/></w:pBdr><w:spacing w:before="280" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>' +
    '<w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:style>' +
    '</w:styles>';

  function visitSortKey(v) { return `${v.date || ""}T${v.timeIn || "00:00"}`; }

  function averageDuration(visits) {
    const durations = visits.map(v => v.durationMinutes).filter(n => typeof n === "number");
    return durations.length ? Math.round(durations.reduce((a, b) => a + b, 0) / durations.length) : null;
  }

  function buildStudentSummaryDocx(visits, options) {
    const opts = options || {};
    const student = String(opts.student || "").trim();
    if (!student) throw new Error("A student must be selected for a Student Summary Document.");
    const list = visits || [];
    // Hard guard: this document is single-student by design. Refuse rather
    // than ever place another student's record in it.
    if (list.some(v => v.student !== student)) {
      throw new Error("Student Summary Document may only contain the selected student's visits.");
    }

    const ordered = [...list].sort((a, b) => visitSortKey(a).localeCompare(visitSortKey(b)));
    const avg = averageDuration(ordered);
    const from = formatDateLong(opts.from), to = formatDateLong(opts.to);
    const range = from && to ? `${from} – ${to}` : from ? `From ${from}` : to ? `Through ${to}` : "All dates";

    const body = [
      docPara([docRun("PACE Student Summary")], "Title"),
      labelled("Student", student),
      labelled("Date Range", range),
      labelled("Total Visits", String(ordered.length)),
      labelled("Average Duration", avg === null ? "Not available" : `${avg} minutes`),
      docPara([docRun("Visits")], "Heading1")
    ];

    if (!ordered.length) body.push(docPara([docRun("No PACE visits in this date range.", { italic: true })]));

    ordered.forEach(v => {
      const start = formatTime12(v.timeIn), end = formatTime12(v.timeOut);
      const when = [formatDateLong(v.date) || "Date not recorded",
        start ? (end ? `${start} – ${end}` : `${start} – in progress`) : ""].filter(Boolean).join(" · ");
      const dur = typeof v.durationMinutes === "number" ? `${v.durationMinutes} min` : (v.isCompleted ? "duration not recorded" : "open visit");
      const room = roomLabel(v.paceRoom);
      body.push(docPara([docRun(`${when}  (${dur})`)], "Heading2"));
      body.push(labelled("Reason", (v.reasons || []).join("; ") || "Not recorded"));
      body.push(labelled("Support / Intervention", (v.supports || []).join("; ") || "Not recorded"));
      body.push(labelled("Behavior Specialist", (v.specialists || []).join("; ") || "Not recorded"));
      body.push(labelled("Teacher Came From", v.teacherCameFrom || "Not recorded"));
      body.push(labelled("PACE Room", room || "Not recorded"));
      body.push(labelled("SCM", v.scmUsed === true ? "Yes" : v.scmUsed === false ? "No" : "Not recorded"));
      body.push(docPara([docRun("Notes: ", { bold: true }), docRun(v.notes || "No notes recorded.")]));
    });

    body.push(docPara([docRun("Confidential — contains protected student information. Handle according to district policy.",
      { italic: true, color: "64748B", size: 18 })]));

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
    ], opts.now);
  }

  /* ── filenames ──────────────────────────────────────────────────────── */

  function safeName(text) {
    return String(text || "").normalize("NFC").replace(/[^\p{L}\p{N}]+/gu, "_").replace(/^_+|_+$/g, "");
  }

  // e.g. PACE_Analee_Cruz_2026-08-23_to_2026-09-21.xlsx. Falls back to the
  // dates actually present in the exported visits when a bound is not set.
  function buildFilename(options) {
    const o = options || {};
    const dates = (o.visits || []).map(v => v.date).filter(Boolean).sort();
    const from = o.from || dates[0] || "";
    const to   = o.to   || dates[dates.length - 1] || "";
    const scope = o.student ? (safeName(o.student) || "Student") : "All_Students";
    const range = from && to ? `${from}_to_${to}` : from ? `from_${from}` : to ? `through_${to}` : "All_Dates";
    return `PACE_${scope}_${range}.${o.ext}`;
  }

  const MIME = {
    csv:  "text/csv;charset=utf-8",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    json: "application/json"
  };

  // Single entry point used by the UI. `visits` MUST be the same filtered,
  // in-memory array the results table is rendering.
  function build(kind, visits, options) {
    const o = options || {};
    const fileOptions = { visits, student: o.student, from: o.from, to: o.to, ext: kind };
    let content;
    if (kind === "csv")       content = toCsv(buildRows(visits));
    else if (kind === "xlsx") content = buildXlsx(buildRows(visits), o);
    else if (kind === "docx") content = buildStudentSummaryDocx(visits, o);
    else if (kind === "json") content = toJson(visits);
    else throw new Error("Unknown export type: " + kind);
    return { filename: buildFilename(fileOptions), mime: MIME[kind], content };
  }

  return {
    COLUMNS, MIME,
    formatDateLong, formatTime12, buildRows, columnsFor,
    csvCell, toCsv, toJson, buildXlsx, buildStudentSummaryDocx,
    buildFilename, build, zip, crc32
  };
});
