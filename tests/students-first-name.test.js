/* ─────────────────────────────────────────────────────────────────────────
   Regression: a new student's FIRST NAME was not persisting

   Symptom (production): adding "Adelaide Testwell" from the Students page
   listed the student as just "Testwell"; Edit showed First Name blank, and
   typing it again changed nothing.

   Root cause: graph.js's getListSchema() let the computed, read-only
   LinkTitle / LinkTitleNoMenu columns — which SharePoint reports under the
   SAME display name as a renamed Title column — overwrite the real Title
   mapping. mapFields() then found that "Student First Name" mapped to a
   read-only column, skipped it with only a console warning, and the item was
   created (or PATCHed) without it. Both Setup → Students and the Students
   page reported success regardless.

   These tests run the REAL graph.js and app.js end to end against a fake
   SharePoint (separate stored vs computed columns, rejects unknown/read-only
   writes, assigns ids, returns computed columns like LinkTitle) across the
   realistic ways a roster list's columns can be set up. No production
   SharePoint is touched.

   To watch the original bug reproduce, run against the pre-fix files:
     git show 1873bf0:graph.js > /tmp/graph.old.js
     git show 1873bf0:app.js   > /tmp/app.old.js
     GRAPH_JS=/tmp/graph.old.js APP_JS=/tmp/app.old.js node tests/students-first-name.test.js

   Run with: node tests/students-first-name.test.js
   ───────────────────────────────────────────────────────────────────────── */

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const crypto = require("node:crypto");

const ROOT = path.resolve(__dirname, "..");
const readSrc = file => fs.readFileSync(path.join(ROOT, file), "utf8");
const GRAPH_SRC = process.env.GRAPH_JS ? fs.readFileSync(process.env.GRAPH_JS, "utf8") : readSrc("graph.js");
const APP_SRC   = process.env.APP_JS   ? fs.readFileSync(process.env.APP_JS, "utf8")   : readSrc("app.js");
const plain = obj => JSON.parse(JSON.stringify(obj));

/* ── fake SharePoint (Graph HTTP surface the app actually uses) ───────────── */

const COMPUTED_FROM_TITLE = ["LinkTitle", "LinkTitleNoMenu"];

function makeSharePoint(scn) {
  const state = {
    items: scn.items.map(i => {
      const f = { ...i.fields };
      if (scn.flagKind === "boolean") ["Active", "PACEEnabled", "DailyPulseEnabled"].forEach(k => { if (k in f) f[k] = f[k] === "Yes"; });
      return { id: i.id, fields: f };
    }),
    nextId: 500, posts: [], patches: []
  };
  const stored = new Set(scn.stored);
  const respond = (status, body) => ({
    ok: status >= 200 && status < 300, status,
    json: async () => body, text: async () => JSON.stringify(body)
  });
  const withComputed = fields => ({ ...fields, ...(fields.Title !== undefined && !scn.omitComputed ? { LinkTitle: fields.Title, LinkTitleNoMenu: fields.Title } : {}) });
  const unknown = keys => keys.find(k => !stored.has(k));
  // Mirror Graph: a write whose value doesn't fit the column's type is a 400.
  const misfit = fields => {
    for (const [k, v] of Object.entries(fields)) {
      const c = scn.columns.find(x => x.name === k);
      if (!c) continue;
      if (c.boolean && typeof v !== "boolean") return `${k}: expected boolean`;
      if (c.choice && !c.choice.allowTextEntry && !c.choice.choices.includes(v)) return `${k}: not a valid choice`;
      if (c.text && typeof v !== "string") return `${k}: expected text`;
    }
    return null;
  };
  const ignored = new Set(scn.silentlyIgnored || []);
  const keep = fields => Object.fromEntries(Object.entries(fields).filter(([k]) => !ignored.has(k)).map(([k, v]) => [k, scn.transform ? scn.transform(k, v) : v]));

  const USERS = [
    { id: "u1", fields: { field_1: "Amber Bossons", field_2: "Teacher", field_3: "Yes" } },
    { id: "u2", fields: { field_1: "Some Admin", field_2: "Administrator", field_3: "Yes" } }
  ];

  async function fetch(url, opts = {}) {
    const method = opts.method || "GET";
    const body = opts.body ? JSON.parse(opts.body) : null;
    if (/\/sites\/siu29\.sharepoint\.com:/.test(url)) return respond(200, { id: "site-1" });
    if (/\/sites\/site-1\/lists\?/.test(url)) {
      return respond(200, { value: [
        { id: "list-roster", name: "IEP_Students_2026_27", displayName: "IEP_Students_2026_27" },
        { id: "list-users",  name: "IEP_Users2",           displayName: "IEP_Users2" },
        { id: "list-other",  name: "IEP_Other_List",       displayName: "IEP_Other_List" }
      ] });
    }
    if (/\/lists\/list-(roster|other)\/columns/.test(url)) return respond(200, { value: scn.columns });
    if (/\/lists\/list-users\/items/.test(url)) return respond(200, { value: USERS });
    const one = url.match(/\/lists\/list-roster\/items\/([^/?]+)\?/);
    if (one && method === "GET") {
      state.itemReads = (state.itemReads || []).concat(one[1]);
      if (scn.getItemFails) return respond(500, { error: { message: "boom" } });
      const item = state.items.find(i => i.id === decodeURIComponent(one[1]));
      return item ? respond(200, { id: item.id, fields: withComputed(item.fields) }) : respond(404, { error: { message: "not found" } });
    }
    if (/\/lists\/list-roster\/items\?/.test(url) && method === "GET") {
      (state.rosterRequests ||= []).push(url);
      const size = scn.pageSize || Infinity;
      const tokenMatch = url.match(/skiptoken=(\d+)/);
      const start = tokenMatch ? Number(tokenMatch[1]) : 0;
      const pageNo = Number.isFinite(size) ? Math.floor(start / size) + 1 : 1;
      if (scn.pageFail === pageNo && !(scn.pageFailOnce && state.pageFailed)) { state.pageFailed = true; return respond(503, { error: { message: `page ${pageNo} unavailable` } }); }
      if (scn.malformedPage === pageNo) return respond(200, { value: "not-an-array" });
      const slice = state.items.slice(start, start + size);
      const body = { value: slice.map(i => ({ id: i.id, fields: withComputed(i.fields) })) };
      if (scn.noIdOnPage === pageNo && body.value.length) delete body.value[0].id;
      const nextStart = start + size - (scn.overlap ? 1 : 0);   // overlap: the next page repeats the last item
      const base = "https://graph.microsoft.com/v1.0/sites/site-1/lists/list-roster/items?$expand=fields&$top=" + (Number.isFinite(size) ? size : 200);
      const linkTo = n => base + "&$skiptoken=" + n;
      const hasMore = Number.isFinite(size) && start + size < state.items.length;
      if (scn.nextLinkMode === "endless") {
        body["@odata.nextLink"] = linkTo(start + 1);          // never ends, every link new
      } else if (hasMore) {
        let link = linkTo(nextStart);
        if (scn.nextLinkMode && pageNo >= (scn.badLinkFromPage || 1)) {
          link = {
            repeat:    url,   // points back at the page just served
            foreign:   "https://evil.example.com/v1.0/sites/site-1/lists/list-roster/items?$skiptoken=" + nextStart,
            otherList: "https://graph.microsoft.com/v1.0/sites/site-1/lists/list-other/items?$skiptoken=" + nextStart,
            relative:  "sites/site-1/lists/list-roster/items?$skiptoken=" + nextStart
          }[scn.nextLinkMode];
        }
        body["@odata.nextLink"] = link;
      }
      return respond(200, body);
    }
    if (/\/lists\/list-roster\/items$/.test(url) && method === "POST") {
      state.posts.push(body.fields);
      const bad = unknown(Object.keys(body.fields));
      if (bad) return respond(400, { error: { message: `Field '${bad}' is not recognized` } });
      const mis = misfit(body.fields);
      if (mis) return respond(400, { error: { message: `Invalid value — ${mis}` } });
      if (scn.postBehavior === "dropBeforeStore") throw new TypeError("Failed to fetch");
      const item = { id: String(state.nextId++), fields: keep(body.fields) };
      state.items.push(item);
      if (scn.postBehavior === "dropAfterStore") throw new TypeError("Failed to fetch");
      if (scn.postBehavior === "noId") return respond(201, { fields: withComputed(item.fields) });
      return respond(201, { id: item.id, fields: withComputed(item.fields) });
    }
    const patch = url.match(/\/lists\/list-roster\/items\/([^/]+)\/fields$/);
    if (patch && method === "PATCH") {
      state.patches.push({ id: patch[1], fields: body });
      const bad = unknown(Object.keys(body));
      if (bad) return respond(400, { error: { message: `Field '${bad}' is not recognized` } });
      const mis = misfit(body);
      if (mis) return respond(400, { error: { message: `Invalid value — ${mis}` } });
      const item = state.items.find(i => i.id === patch[1]);
      Object.assign(item.fields, keep(body));
      return respond(200, withComputed(item.fields));
    }
    throw new Error(`Unexpected Graph request in test: ${method} ${url}`);
  }
  return { state, fetch };
}

/* ── realistic roster column layouts ──────────────────────────────────────── */

// Every column carries a real Graph type facet (text unless stated), exactly
// as the live columns endpoint reports them.
const FACETS = ["text", "boolean", "choice", "number", "dateTime", "lookup", "personOrGroup", "calculated"];
const col = (name, displayName, extra = {}) =>
  ({ name, displayName, ...(FACETS.some(f => f in extra) ? {} : { text: { maxLength: 255 } }), ...extra });
const YN_CHOICE = { choice: { choices: ["Yes", "No"], allowTextEntry: false } };
const flagCols = kind => [
  col("Active", "Active", kind === "boolean" ? { boolean: {} } : kind === "choice" ? YN_CHOICE : {}),
  col("PACEEnabled", "PACE Enabled", kind === "boolean" ? { boolean: {} } : kind === "choice" ? YN_CHOICE : {}),
  col("DailyPulseEnabled", "Daily Pulse Enabled", kind === "boolean" ? { boolean: {} } : kind === "choice" ? YN_CHOICE : {})
];
const commonCols = (kind = "text") => [
  col("Teacher", "Teacher"), col("Classroom", "Classroom"), ...flagCols(kind),
  col("ContentType", "Content Type"), col("Modified", "Modified"), col("Created", "Created")
];
const COMMON = commonCols();
const COMMON_STORED = ["Teacher", "Classroom", "Active", "PACEEnabled", "DailyPulseEnabled"];
const rowFor = (id, keys, first, last, teacher, flags = {}) => ({
  id,
  fields: {
    ...(first === null ? {} : { [keys.first]: first }),
    [keys.last]: last, Teacher: teacher, Classroom: "", Active: "Yes",
    PACEEnabled: flags.pace ?? "No", DailyPulseEnabled: flags.pulse ?? "No"
  }
});

function scenario(label, keys, columns, stored) {
  return {
    label, keys, columns, stored,
    items: [
      rowFor("11", keys, "Jane",  "Doe", "Bossons, A", { pace: "Yes" }),
      rowFor("12", keys, "Maya",  "Ortiz", "Bossons, A", { pulse: "Yes" }),
      rowFor("13", keys, null,    "Quill", "Bossons, A")   // pre-existing record with NO first name
    ]
  };
}

const SCENARIOS = [
  // The Title column was renamed "Student First Name"; SharePoint also reports
  // the computed LinkTitle/LinkTitleNoMenu columns under that same display name.
  scenario("renamed Title = first name (+ computed LinkTitle columns)", { first: "Title", last: "StudentLastName" }, [
    col("Title", "Student First Name"), col("StudentLastName", "Student Last Name"), ...COMMON,
    col("LinkTitleNoMenu", "Student First Name"), col("LinkTitle", "Student First Name")
  ], ["Title", "StudentLastName", ...COMMON_STORED]),

  // Same thing the other way round: Title is the LAST name.
  // PRODUCTION SHAPE (the failure Melissa hit): the writable Title column
  // keeps the display name "Title"; ONLY the computed, read-only LinkTitle /
  // LinkTitleNoMenu columns carry "Student First Name".
  scenario("production: Title unrenamed, computed columns carry the name", { first: "Title", last: "StudentLastName" }, [
    col("Title", "Title"), col("StudentLastName", "Student Last Name"), ...COMMON,
    col("LinkTitleNoMenu", "Student First Name", { readOnly: true, hidden: true }), col("LinkTitle", "Student First Name", { readOnly: true })
  ], ["Title", "StudentLastName", ...COMMON_STORED]),

  // Same, but Graph does not include the computed fields on items at all.
  // Same production shape, but the three flags are SharePoint Yes/No (Boolean) columns.
  Object.assign(scenario("production + Boolean flag columns", { first: "Title", last: "StudentLastName" }, [
    col("Title", "Title"), col("StudentLastName", "Student Last Name"), ...commonCols("boolean"),
    col("LinkTitleNoMenu", "Student First Name", { readOnly: true, hidden: true }), col("LinkTitle", "Student First Name", { readOnly: true })
  ], ["Title", "StudentLastName", ...COMMON_STORED]), { flagKind: "boolean" }),

  // ...or Choice columns whose choices are Yes / No.
  Object.assign(scenario("production + Choice flag columns", { first: "Title", last: "StudentLastName" }, [
    col("Title", "Title"), col("StudentLastName", "Student Last Name"), ...commonCols("choice"),
    col("LinkTitleNoMenu", "Student First Name", { readOnly: true, hidden: true }), col("LinkTitle", "Student First Name", { readOnly: true })
  ], ["Title", "StudentLastName", ...COMMON_STORED]), { flagKind: "choice" }),

  Object.assign(scenario("production: computed fields omitted from item reads", { first: "Title", last: "StudentLastName" }, [
    col("Title", "Title"), col("StudentLastName", "Student Last Name"), ...COMMON,
    col("LinkTitleNoMenu", "Student First Name", { readOnly: true, hidden: true }), col("LinkTitle", "Student First Name", { readOnly: true })
  ], ["Title", "StudentLastName", ...COMMON_STORED]), { omitComputed: true }),

  scenario("renamed Title = last name (+ computed LinkTitle columns)", { first: "StudentFirstName", last: "Title" }, [
    col("Title", "Student Last Name"), col("StudentFirstName", "Student First Name"), ...COMMON,
    col("LinkTitleNoMenu", "Student Last Name"), col("LinkTitle", "Student Last Name")
  ], ["Title", "StudentFirstName", ...COMMON_STORED]),

  scenario("plain custom columns", { first: "StudentFirstName", last: "StudentLastName" }, [
    col("Title", "Title"), col("StudentFirstName", "Student First Name"), col("StudentLastName", "Student Last Name"), ...COMMON
  ], ["Title", "StudentFirstName", "StudentLastName", ...COMMON_STORED]),

  scenario("field_N internal names", { first: "field_1", last: "field_2" }, [
    col("Title", "Title"), col("field_1", "Student First Name"), col("field_2", "Student Last Name"), ...COMMON
  ], ["Title", "field_1", "field_2", ...COMMON_STORED]),

  scenario("shorter display names (First Name / Last Name)", { first: "FirstName", last: "LastName" }, [
    col("Title", "Title"), col("FirstName", "First Name"), col("LastName", "Last Name"), ...COMMON
  ], ["Title", "FirstName", "LastName", ...COMMON_STORED])
];

const PLAIN_SCENARIO = SCENARIOS.find(x => x.label === "plain custom columns");

/* ── app harness (REAL graph.js + app.js) ─────────────────────────────────── */

function makeDocument() {
  const registry = {};
  function elementFor(id) {
    if (!registry[id]) {
      const el = {
        id, value: "", checked: false, textContent: "", dataset: {},
        classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
        style: {}, _handlers: {},
        addEventListener(ev, fn) { (el._handlers[ev] ||= []).push(fn); },
        removeEventListener() {},
        querySelector: () => elementFor(Symbol()), querySelectorAll: () => [],
        closest() { return null; }, setAttribute() {}, getAttribute() { return null; }, removeAttribute() {},
        remove() {}, appendChild() {}, matches() { return false; }, focus() {}, blur() {},
        click() {}, scrollIntoView() {}, scrollTo() {}
      };
      Object.defineProperty(el, "innerHTML", { get: () => el._html || "", set: v => { el._html = v; } });
      registry[id] = el;
    }
    return registry[id];
  }
  return {
    registry, getElementById: id => elementFor(id),
    querySelector: () => elementFor(Symbol()), querySelectorAll: () => [],
    createElement: () => elementFor(Symbol()), addEventListener() {}, removeEventListener() {},
    body: elementFor("__body__")
  };
}

// A fresh "browser" talking to the given fake SharePoint. A second call with
// the same `sp` is the same data after a page refresh.
async function openBrowser(sp) {
  const document = makeDocument();
  const toasts = [];
  const context = {
    console: { ...console, log() {}, warn() {}, table() {}, error() {} }, document, navigator: {},
    location: { origin: "http://localhost", hash: "" }, crypto: crypto.webcrypto, TextEncoder, Uint8Array,
    localStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    sessionStorage: { getItem() { return null; }, setItem() {}, removeItem() {} },
    fetch: sp.fetch, msal: { PublicClientApplication: class {} },
    setTimeout, clearTimeout, structuredClone, addEventListener() {}, removeEventListener() {}, confirm: () => true,
    __toasts: toasts
  };
  context.window = context;
  vm.createContext(context);
  const run = (src, filename) => vm.runInContext(src, context, { filename });
  run(readSrc("config.js"), "config.js");
  run(readSrc("auth.js"), "auth.js");
  run(GRAPH_SRC, "graph.js");
  run(readSrc("data.js"), "data.js");
  run(readSrc("pace-admin.js"), "pace-admin.js");
  run(readSrc("pace-export.js"), "pace-export.js");
  run(readSrc("evidence-report.js"), "evidence-report.js");
  run(APP_SRC, "app.js");
  run('MAC_ADMIN_PANEL_ALLOWED = true; AUTH.role = "Administrator"; AUTH.acquireGraphToken = async () => "test-token"; showToast = (m, t) => __toasts.push({ m, t });', "setup");
  const b = {
    context, document, toasts,
    GRAPH: run("GRAPH", "x"), APP: run("APP", "x"), STUDENT_ROSTER: run("STUDENT_ROSTER", "x"),
    TEACHER_DIRECTORY: run("TEACHER_DIRECTORY", "x"), SETUP_DATA: run("SETUP_DATA", "x")
  };
  await b.STUDENT_ROSTER.refresh();
  await b.TEACHER_DIRECTORY.refresh();
  b.APP.currentPage = "students";
  b.APP.renderStudents();
  return b;
}

const fill = (b, v) => {
  const d = b.document;
  d.getElementById("stuFirst").value = v.first ?? "";
  d.getElementById("stuLast").value = v.last ?? "";
  d.getElementById("stuTeacher").value = v.teacher ?? "Bossons, A";
  d.getElementById("stuClassroom").value = v.classroom ?? "";
  d.getElementById("stuActive").checked = v.active ?? true;
  d.getElementById("stuPace").checked = v.pace ?? false;
  d.getElementById("stuPulse").checked = v.pulse ?? false;
};
const listHtml = b => b.document.registry["stuList"].innerHTML;
const formHtml = b => b.document.registry["stuFormHost"].innerHTML;
const successToasts = b => b.toasts.filter(t => t.t !== "error");

/* ── graph.js unit checks (the actual defect) ─────────────────────────────── */

async function schemaBuilderNeverLetsAComputedColumnShadowTheRealOne() {
  const sp = makeSharePoint(SCENARIOS[0]);
  const b = await openBrowser(sp);
  const schema = plain(await b.GRAPH.getListSchema("IEP_Students_2026_27"));
  assert.equal(schema["Student First Name"], "Title", "the real, writable Title column wins over LinkTitle / LinkTitleNoMenu");
  assert.equal(schema["Student Last Name"], "StudentLastName");

  // Behaviour that must NOT change: among usable columns, last still wins.
  const sp2 = makeSharePoint({ ...PLAIN_SCENARIO, columns: [col("A1", "Dup"), col("A2", "Dup"), col("LinkTitle", "Dup")], stored: ["A1", "A2"], items: [] });
  const b2 = await openBrowser(sp2);
  assert.equal((await b2.GRAPH.getListSchema("IEP_Students_2026_27"))["Dup"], "A2");
  // A read-only column alone is still reported (callers decide what to do with it).
  const sp3 = makeSharePoint({ ...PLAIN_SCENARIO, columns: [col("LinkTitle", "Only")], stored: [], items: [] });
  const b3 = await openBrowser(sp3);
  assert.equal((await b3.GRAPH.getListSchema("IEP_Students_2026_27"))["Only"], "LinkTitle");
}

async function strictMappingRefusesToDropFieldsButDefaultStaysLenient() {
  const sp = makeSharePoint({ ...PLAIN_SCENARIO, columns: [col("Title", "Title"), col("StudentLastName", "Student Last Name")], stored: ["Title", "StudentLastName"], items: [] });
  const b = await openBrowser(sp);
  const GRAPH = b.GRAPH;
  await assert.rejects(
    () => GRAPH.createMappedListItem("IEP_Students_2026_27", { "Student First Name": "Adelaide", "Student Last Name": "Testwell" }, { strict: true }),
    /Column mapping problem.*nothing was saved.*"Student First Name" has no matching column/
  );
  assert.equal(sp.state.posts.length, 0, "strict mode aborts BEFORE writing anything");
  // Lenient default is untouched — other lists rely on skipping optional columns.
  await GRAPH.createMappedListItem("IEP_Students_2026_27", { "Student First Name": "Adelaide", "Student Last Name": "Testwell" });
  assert.deepEqual(plain(sp.state.posts[0]), { StudentLastName: "Testwell" });
}

/* ── the full user journey, per schema layout ─────────────────────────────── */

async function addEditAndRefreshPreserveBothNamesInEveryLayout() {
  for (const scn of SCENARIOS) {
    const L = scn.label;
    const sp = makeSharePoint(scn);
    const before = plain(sp.state.items);
    let b = await openBrowser(sp);

    // Existing students read correctly (and the record with no first name stays that way — nothing is invented).
    const roster = plain(b.STUDENT_ROSTER.getAll().map(s => s.name));
    assert.deepEqual(roster, ["Jane Doe", "Maya Ortiz", "Quill"], `${L}: existing records read back unchanged`);

    // 1–4. Add Adelaide Testwell.
    b.APP._openStudentAdd();
    fill(b, { first: "Adelaide", last: "Testwell", classroom: "Room 4", pace: true });
    await b.APP._saveStudentForm(false);
    assert.equal(sp.state.posts.length, 1, `${L}: exactly one create`);
    assert.ok(sp.state.posts.every(f => !("LinkTitle" in f) && !("LinkTitleNoMenu" in f)), `${L}: a computed/read-only column is never written`);
    const created = sp.state.items.find(i => i.fields[scn.keys.last] === "Testwell");
    assert.equal(created.fields[scn.keys.first], "Adelaide", `${L}: FIRST name is stored in SharePoint (${scn.keys.first})`);
    assert.equal(created.fields[scn.keys.last], "Testwell", `${L}: LAST name is stored (${scn.keys.last})`);
    assert.equal(created.fields.Teacher, "Bossons, A");
    const flag = v => (scn.flagKind === "boolean" ? v === "Yes" : v);   // how this column type stores Yes/No
    assert.equal(created.fields.Active, flag("Yes"), `${L}: Active stored in the column's own type`);
    assert.equal(created.fields.PACEEnabled, flag("Yes"));
    assert.equal(created.fields.DailyPulseEnabled, flag("No"));
    assert.match(listHtml(b), /Adelaide Testwell/, `${L}: directory shows the full name`);
    assert.equal(b.toasts.filter(t => t.t === "error").length, 0, `${L}: no error`);
    const newId = created.id;

    // 8. "Refresh the browser": a brand-new session reading the same SharePoint.
    b = await openBrowser(sp);
    const rec = b.STUDENT_ROSTER.find(newId);
    assert.deepEqual([rec.firstName, rec.lastName, rec.name], ["Adelaide", "Testwell", "Adelaide Testwell"], `${L}: names survive normalization after a refresh`);
    assert.match(listHtml(b), /Adelaide Testwell/);

    // 5. Edit prepopulates both names.
    b.APP._openStudentEdit(newId);
    assert.match(formHtml(b), /id="stuFirst"[^>]*value="Adelaide"/, `${L}: Edit shows the first name`);
    assert.match(formHtml(b), /id="stuLast"[^>]*value="Testwell"/, `${L}: Edit shows the last name`);

    // 6. Edit the first name → only that column is sent; id/teacher/flags unchanged.
    fill(b, { first: "Adeline", last: "Testwell", classroom: "Room 4", pace: true });
    await b.APP._saveStudentForm(false);
    const patch1 = sp.state.patches.at(-1);
    assert.equal(patch1.id, newId, `${L}: same record id`);
    assert.deepEqual(plain(patch1.fields), { [scn.keys.first]: "Adeline" }, `${L}: PATCH carries just the first name`);
    assert.equal(sp.state.items.find(i => i.id === newId).fields[scn.keys.first], "Adeline");
    // 7. Edit the last name.
    b.APP._openStudentEdit(newId);
    fill(b, { first: "Adeline", last: "Testwelp", classroom: "Room 4", pace: true });
    await b.APP._saveStudentForm(false);
    assert.deepEqual(plain(sp.state.patches.at(-1).fields), { [scn.keys.last]: "Testwelp" });
    const after = sp.state.items.find(i => i.id === newId).fields;
    assert.equal(after[scn.keys.first], "Adeline");
    assert.equal(after[scn.keys.last], "Testwelp");
    // 9–11. id, teacher and flags untouched by name edits.
    assert.equal(after.Teacher, "Bossons, A");
    assert.deepEqual([after.Active, after.PACEEnabled, after.DailyPulseEnabled], [flag("Yes"), flag("Yes"), flag("No")]);
    assert.equal(sp.state.items.filter(i => i.fields[scn.keys.last] === "Testwelp").length, 1, `${L}: still one record, not a re-create`);

    // 12. A pre-existing record with NO first name can be repaired through Edit.
    b = await openBrowser(sp);
    b.APP._openStudentEdit("13");
    assert.match(formHtml(b), /id="stuFirst"[^>]*value=""/, `${L}: the missing first name shows blank, not invented`);
    fill(b, { first: "Alma", last: "Quill", classroom: "", active: true });
    await b.APP._saveStudentForm(false);
    assert.deepEqual(plain(sp.state.patches.at(-1)), { id: "13", fields: { [scn.keys.first]: "Alma" } });
    assert.equal(sp.state.items.find(i => i.id === "13").fields[scn.keys.first], "Alma");

    // Nothing else in the list was touched — no bulk change, no deletes, ids intact.
    const final = plain(sp.state.items);
    assert.deepEqual(final.filter(i => ["11", "12"].includes(i.id)), before.filter(i => ["11", "12"].includes(i.id)), `${L}: other students untouched`);
    assert.deepEqual(final.map(i => i.id), [...before.map(i => i.id), newId], `${L}: only the one new record exists; existing ids unchanged`);
  }
}

async function missingFirstNameColumnFailsLoudlyAndWritesNothing() {
  const scn = scenario("no first-name column", { first: "x", last: "StudentLastName" },
    [col("Title", "Title"), col("StudentLastName", "Student Last Name"), ...COMMON], ["Title", "StudentLastName", ...COMMON_STORED]);
  scn.items = [];
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 0, "nothing written when the first-name column cannot be found");
  assert.equal(successToasts(b).length, 0, "no false success");
  assert.match(b.toasts.at(-1).m, /no column for the student's first name\. Columns found:.*Student Last Name/);
  assert.equal(b.APP._studentsUi.mode, "add", "form stays open with the typed values");
}

async function aWriteSharePointAcceptsButDoesNotKeepIsNeverReportedAsSaved() {
  const scn = { ...PLAIN_SCENARIO, silentlyIgnored: ["StudentFirstName"] };
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell" });
  await b.APP._saveStudentForm(false);

  assert.equal(sp.state.posts.length, 1);
  assert.equal(successToasts(b).length, 0, "no success message for a save that did not stick");
  assert.match(b.toasts.at(-1).m, /SharePoint did not keep: first name/);
  assert.match(formHtml(b), /SharePoint did not keep: first name/, "the problem is shown on the form");
  assert.match(formHtml(b), /First name: &quot;Student First Name&quot; → column StudentFirstName/, "names the exact column it tried, so the cause can be read off the screen");
  assert.equal(b.APP._studentsUi.mode, "edit", "the form now targets the record that was created");
  assert.equal(b.APP._studentsUi.editId, sp.state.items.at(-1).id);

  // Retrying updates that record — it never creates a second one.
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 1, "no duplicate record from the retry");
  assert.equal(sp.state.patches.length, 1);
}


/* ── the production failure, pinned ───────────────────────────────────────── */

const STUDENT_FIELDS = {
  "Student First Name": "Adelaide", "Student Last Name": "Testwell", "Teacher": "Bossons, A",
  "Classroom": "Room 4", "Active": "Yes", "PACE Enabled": "No", "Daily Pulse Enabled": "No"
};

async function productionShapeOriginalMappingFailsAndCorrectedPayloadIsValid() {
  const scn = SCENARIOS.find(x => x.label.startsWith("production: Title unrenamed"));

  // (a) The error exactly as Melissa saw it: the display name resolves ONLY to the read-only LinkTitle.
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  const schema = plain(await b.GRAPH.getListSchema("IEP_Students_2026_27"));
  assert.equal(schema["Student First Name"], "LinkTitle", "authoritative metadata: only computed columns carry this display name");

  // (b) The previous mapping (strict, no Title translation) fails with that exact message.
  const prev = await openBrowser(makeSharePoint(scn));
  prev.GRAPH.resolveWriteTarget = async (list, display) => {
    const internal = (await prev.GRAPH.getListSchema(list))[display];
    return vm.runInContext("isWritableSharePointField", prev.context)(internal)
      ? { internalName: internal } : { internalName: null, reason: `"${display}" maps to the read-only column "${internal}"` };
  };
  await assert.rejects(
    () => prev.GRAPH.mapFields("IEP_Students_2026_27", STUDENT_FIELDS, { strict: true }),
    /Column mapping problem in IEP_Students_2026_27 — nothing was saved\. "Student First Name" maps to the read-only column "LinkTitle"\./
  );

  // (c) The corrected mapping produces a valid payload — inspected WITHOUT submitting it.
  const payload = plain(await b.GRAPH.mapFields("IEP_Students_2026_27", STUDENT_FIELDS, { strict: true }));
  assert.deepEqual(payload, {
    Title: "Adelaide", StudentLastName: "Testwell", Teacher: "Bossons, A", Classroom: "Room 4",
    Active: "Yes", PACEEnabled: "No", DailyPulseEnabled: "No"
  });
  assert.ok(!("LinkTitle" in payload) && !("LinkTitleNoMenu" in payload), "never a computed column");
  assert.equal(sp.state.posts.length + sp.state.patches.length, 0, "inspecting the payload wrote nothing");
  const target = await b.GRAPH.resolveWriteTarget("IEP_Students_2026_27", "Student First Name");
  assert.deepEqual(plain(target), { internalName: "Title", translated: true, via: "LinkTitle" });
  // every other field resolves to its own real, writable column, untranslated
  for (const d of ["Student Last Name", "Teacher", "Classroom", "Active", "PACE Enabled", "Daily Pulse Enabled"]) {
    const t = await b.GRAPH.resolveWriteTarget("IEP_Students_2026_27", d);
    assert.ok(t.internalName && !t.translated, `${d} → ${t.internalName}`);
  }
}

async function ambiguousOrUnverifiableMappingsAreRejectedNotGuessed() {
  const base = SCENARIOS.find(x => x.label.startsWith("production: Title unrenamed"));

  // No writable "Title" in the metadata → cannot translate → reject, write nothing.
  const noTitle = { ...base, columns: base.columns.filter(c => c.name !== "Title"), stored: base.stored.filter(n => n !== "Title") };
  const sp1 = makeSharePoint(noTitle);
  const b1 = await openBrowser(sp1);
  b1.APP._openStudentAdd();
  fill(b1, { first: "Adelaide", last: "Testwell" });
  await b1.APP._saveStudentForm(false);
  assert.equal(sp1.state.posts.length, 0);
  assert.match(b1.toasts.at(-1).m, /read-only column "LinkTitle" and the list has no writable "Title" column/);
  assert.equal(successToasts(b1).length, 0);

  // Title present but itself read-only/hidden → reject.
  for (const flag of [{ readOnly: true }, { hidden: true }]) {
    const sp = makeSharePoint({ ...base, columns: base.columns.map(c => c.name === "Title" ? { ...c, ...flag } : c) });
    const b = await openBrowser(sp);
    await assert.rejects(() => b.GRAPH.mapFields("IEP_Students_2026_27", STUDENT_FIELDS, { strict: true }), /no writable "Title" column/);
  }

  // A non-Title computed column (not LinkTitle*) is never translated.
  const odd = makeSharePoint({ ...base, columns: [...base.columns, col("Edit", "Odd Field", { readOnly: true })] });
  const bo = await openBrowser(odd);
  await assert.rejects(() => bo.GRAPH.mapFields("IEP_Students_2026_27", { "Odd Field": "x" }, { strict: true }), /"Odd Field" maps to the read-only column "Edit"/);

  // Two fields that would land in the same column are rejected.
  const b2 = await openBrowser(makeSharePoint(base));
  await assert.rejects(
    () => b2.GRAPH.mapFields("IEP_Students_2026_27", { "Student First Name": "A", "Title": "B" }, { strict: true }),
    /would both write to the column "Title"/
  );
}

async function sameDefectWasAlsoBreakingSetupStudentsAndEditAndNowReadsAreHardened() {
  const scn = SCENARIOS.find(x => x.label.startsWith("production: computed fields omitted"));
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  // Roster retrieval: with no computed fields on the items, first names are still read (from Title).
  assert.deepEqual(plain(b.STUDENT_ROSTER.getAll().map(s => s.name)), ["Jane Doe", "Maya Ortiz", "Quill"]);
  assert.equal(b.STUDENT_ROSTER.find("13").firstName, "", "a record with no first name stays blank — nothing invented");

  // Setup → Students (lenient path, unchanged UI code) used to drop the first name silently.
  const d = b.document;
  b.APP._bindSetupTabEvents("students");
  d.getElementById("newStudentFirstName").value = "Sofia";
  d.getElementById("newStudentLastName").value = "Reyes";
  d.getElementById("newStudentTeacher").value = "Bossons, A";
  d.getElementById("newStudentClassroom").value = "";
  d.getElementById("newStudentActive").checked = true;
  d.getElementById("newStudentPace").checked = false;
  d.getElementById("newStudentPulse").checked = false;
  await d.registry["addStudentBtn"]._handlers.click[0]();
  assert.equal(sp.state.items.at(-1).fields.Title, "Sofia");

  // Students page edit repairs the record with the missing first name via Title.
  b.APP._openStudentEdit("13");
  fill(b, { first: "Alma", last: "Quill", classroom: "", active: true });
  await b.APP._saveStudentForm(false);
  assert.deepEqual(plain(sp.state.patches.at(-1)), { id: "13", fields: { Title: "Alma" } });
  assert.equal(sp.state.items.find(i => i.id === "13").fields.Title, "Alma");
}

async function otherListsAreUnaffectedByTheTitleTranslation() {
  // A list whose display name resolves to a normal writable column behaves exactly as before.
  const scn = SCENARIOS.find(x => x.label === "plain custom columns");
  const b = await openBrowser(makeSharePoint(scn));
  const out = plain(await b.GRAPH.mapFields("IEP_Students_2026_27", { "Student First Name": "Z", "Teacher": "T" }));
  assert.deepEqual(out, { StudentFirstName: "Z", Teacher: "T" });
  // Lenient mode still skips (and does not throw for) a genuinely unmapped optional field.
  const lenient = plain(await b.GRAPH.mapFields("IEP_Students_2026_27", { "Nope": "x", "Teacher": "T" }));
  assert.deepEqual(lenient, { Teacher: "T" });
}


/* ── final safety review ─────────────────────────────────────────────────── */

async function titleTranslationIsRestrictedToTheStudentRosterList() {
  const scn = SCENARIOS.find(x => x.label.startsWith("production: Title unrenamed"));
  const b = await openBrowser(makeSharePoint(scn));
  const fields = { "Student First Name": "Adelaide" };

  // The roster list translates...
  assert.deepEqual(plain(await b.GRAPH.mapFields("IEP_Students_2026_27", fields, { strict: true })), { Title: "Adelaide" });
  // ...an identically-shaped OTHER list does not: strict rejects, lenient skips.
  await assert.rejects(() => b.GRAPH.mapFields("IEP_Other_List", fields, { strict: true }),
    /"Student First Name" maps to the read-only column "LinkTitle"\./);
  assert.deepEqual(plain(await b.GRAPH.mapFields("IEP_Other_List", fields)), {}, "lenient mode still silently skips, as before this work");
  const t = await b.GRAPH.resolveWriteTarget("IEP_Other_List", "Student First Name");
  assert.equal(t.internalName, null);

  // If the roster setting is not the configured list (e.g. placeholder fallback), nothing translates either.
  vm.runInContext("CONFIG", b.context).STUDENT_ROSTER_LIST = "REPLACE_WITH_2026_27_STUDENT_LIST";
  assert.equal((await b.GRAPH.resolveWriteTarget("IEP_Students_2026_27", "Student First Name")).internalName, null);
  // ...and the helper agrees: only the configured roster list ever qualifies.
  assert.equal(b.GRAPH._isStudentRosterList("IEP_Students_2026_27"), false);
}

async function addAndEditShareOneVerifiedMappingAndBothAreStrict() {
  const scn = SCENARIOS.find(x => x.label.startsWith("production: Title unrenamed"));
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  const seen = [];
  const real = b.GRAPH.mapFields.bind(b.GRAPH);
  b.GRAPH.mapFields = async (list, fields, options) => { const out = await real(list, fields, options); seen.push({ list, options, out: plain(out) }); return out; };

  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell", classroom: "" });
  await b.APP._saveStudentForm(false);
  const newId = sp.state.items.at(-1).id;
  b.APP._openStudentEdit(newId);
  fill(b, { first: "Adeline", last: "Testwell", classroom: "" });
  await b.APP._saveStudentForm(false);

  assert.equal(seen.length, 4, "each save = one strict validation pass + one strict write");
  seen.forEach(c => { assert.equal(c.list, "IEP_Students_2026_27"); assert.deepEqual(plain(c.options), { strict: true }); });
  assert.equal(seen[0].out.Title, "Adelaide", "Add → Title");
  assert.deepEqual(seen[0].out, seen[1].out, "the validated payload is exactly the payload written (Add)");
  assert.deepEqual(seen[2].out, { Title: "Adeline" }, "Edit → same Title mapping");
  assert.deepEqual(seen[2].out, seen[3].out, "the validated payload is exactly the payload written (Edit)");
}

async function readBackChecksTheActualSavedFirstNameNotJustGraphSuccess() {
  // SharePoint accepts the request (200/201) but stores a DIFFERENT value.
  const scn = { ...SCENARIOS.find(x => x.label.startsWith("production: Title unrenamed")), transform: (k, v) => (k === "Title" ? v.slice(0, 3) : v) };
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 1, "Graph itself reported success");
  assert.equal(sp.state.items.at(-1).fields.Title, "Ade", "but the stored value differs");
  assert.equal(successToasts(b).length, 0, "so the page does NOT report success");
  assert.match(b.toasts.at(-1).m, /SharePoint did not keep: first name/);
  // The comparison is against the value read back from SharePoint, character for character.
  assert.equal(b.STUDENT_ROSTER.find(sp.state.items.at(-1).id).firstName, "Ade");
}

async function flagPayloadsMatchTheRealColumnTypes() {
  const cases = [
    ["text",    "production: Title unrenamed, computed columns carry the name", "Yes", "No"],
    ["boolean", "production + Boolean flag columns",                            true,  false],
    ["choice",  "production + Choice flag columns",                             "Yes", "No"]
  ];
  for (const [kind, label, yes, no] of cases) {
    const sp = makeSharePoint(SCENARIOS.find(x => x.label === label));
    const b = await openBrowser(sp);
    b.APP._openStudentAdd();
    fill(b, { first: "Adelaide", last: "Testwell", active: true, pace: true, pulse: false });
    await b.APP._saveStudentForm(false);
    assert.equal(b.toasts.filter(t => t.t === "error").length, 0, `${kind}: saved without error`);
    const sent = plain(sp.state.posts[0]);
    assert.deepEqual([sent.Active, sent.PACEEnabled, sent.DailyPulseEnabled], [yes, yes, no], `${kind}: flags sent as ${JSON.stringify(yes)}/${JSON.stringify(no)}`);
    assert.equal(typeof sent.Active, kind === "boolean" ? "boolean" : "string");
    // the values read back mean the same thing to PACE / Daily Pulse
    const rec = b.STUDENT_ROSTER.find(sp.state.items.at(-1).id);
    assert.deepEqual([rec.active, rec.paceEnabled, rec.dailyPulseEnabled], [true, true, false]);
  }
}

async function booleanColumnsNeverReceiveYesNoStringsAndUnclearTypesAreRejected() {
  const bool = makeSharePoint(SCENARIOS.find(x => x.label === "production + Boolean flag columns"));
  const b = await openBrowser(bool);
  // Old-style string into a Boolean column is refused up front by strict validation...
  await assert.rejects(() => b.GRAPH.mapFields("IEP_Students_2026_27", { "Active": "Yes" }, { strict: true }),
    /"Active" is a Yes\/No \(boolean\) column and needs true\/false, not text/);
  // ...and the fake SharePoint proves such a write would have been a 400 anyway.
  const r = await bool.fetch("https://graph.microsoft.com/v1.0/sites/site-1/lists/list-roster/items", { method: "POST", body: JSON.stringify({ fields: { Active: "Yes" } }) });
  assert.equal(r.status, 400);

  // Choice column with no yes/no option → throws rather than guessing.
  const odd = SCENARIOS.find(x => x.label === "production + Choice flag columns");
  const oddSp = makeSharePoint({ ...odd, columns: odd.columns.map(c => c.name === "PACEEnabled" ? col("PACEEnabled", "PACE Enabled", { choice: { choices: ["Enrolled", "Exited"], allowTextEntry: false } }) : c) });
  const bo = await openBrowser(oddSp);
  await assert.rejects(() => bo.GRAPH.booleanValueFor("IEP_Students_2026_27", "PACE Enabled", true), /Choice column \(Enrolled, Exited\) with no clear Yes option/);
  // A choice column that happens to use True/False picks its own spelling.
  const tf = makeSharePoint({ ...odd, columns: odd.columns.map(c => c.name === "Active" ? col("Active", "Active", { choice: { choices: ["True", "False"], allowTextEntry: false } }) : c) });
  const bt = await openBrowser(tf);
  assert.equal(await bt.GRAPH.booleanValueFor("IEP_Students_2026_27", "Active", false), "False");
  // Number / unknown-type columns are not guessed at.
  const num = makeSharePoint({ ...odd, columns: odd.columns.map(c => c.name === "Active" ? col("Active", "Active", { number: {} }) : c) });
  const bn = await openBrowser(num);
  await assert.rejects(() => bn.GRAPH.booleanValueFor("IEP_Students_2026_27", "Active", true), /number column; this app cannot safely store a Yes\/No/);
  // Saving through the page with such a column writes nothing and says so.
  bn.APP._openStudentAdd();
  fill(bn, { first: "Adelaide", last: "Testwell" });
  await bn.APP._saveStudentForm(false);
  assert.equal(num.state.posts.length, 0);
  assert.match(bn.toasts.at(-1).m, /number column; this app cannot safely store a Yes\/No/);

  // Teacher as a lookup/person column can't take plain text → strict rejects.
  const look = makeSharePoint({ ...odd, columns: odd.columns.map(c => c.name === "Teacher" ? col("Teacher", "Teacher", { lookup: {} }) : c) });
  const bl = await openBrowser(look);
  await assert.rejects(() => bl.GRAPH.mapFields("IEP_Students_2026_27", { "Teacher": "Bossons, A" }, { strict: true }), /"Teacher" is a lookup column/);
  // Over-long text is caught before the write.
  await assert.rejects(() => b.GRAPH.mapFields("IEP_Students_2026_27", { "Classroom": "x".repeat(300) }, { strict: true }), /longer than the column allows \(255 characters\)/);
}

async function eligibilityRulesAndExistingRecordsAreUntouched() {
  const scn = SCENARIOS.find(x => x.label === "production + Boolean flag columns");
  const sp = makeSharePoint(scn);
  const before = plain(sp.state.items);
  const b = await openBrowser(sp);
  assert.deepEqual(plain(vm.runInContext("QUICK_ADD_STUDENT_DEFAULTS", b.context)), { active: true, paceEnabled: false, dailyPulseEnabled: false }, "defaults unchanged");
  assert.deepEqual(plain(b.STUDENT_ROSTER.getPaceEnabled().map(s => s.name)), ["Jane Doe"]);
  assert.deepEqual(plain(b.STUDENT_ROSTER.getDailyPulseEnabled().map(s => s.name)), ["Maya Ortiz"]);
  assert.deepEqual(plain(sp.state.items), before, "opening/reading the page changed no record");
  assert.equal(sp.state.posts.length + sp.state.patches.length, 0);
  // blank-flag fail-safe is still in force
  assert.equal(vm.runInContext('normalizeRosterBoolean("", false)', b.context), false);
}


/* ── release readiness: one verified path for both workflows ──────────────── */

const byLabel = label => SCENARIOS.find(x => x.label === label);
const PROD   = "production: Title unrenamed, computed columns carry the name";
const PROD_B = "production + Boolean flag columns";
const PROD_C = "production + Choice flag columns";

async function setupAdd(b, v) {
  const d = b.document;
  b.APP._bindSetupTabEvents("students");
  d.getElementById("newStudentFirstName").value = v.first;
  d.getElementById("newStudentLastName").value = v.last;
  d.getElementById("newStudentTeacher").value = v.teacher ?? "Bossons, A";
  d.getElementById("newStudentClassroom").value = v.classroom ?? "";
  d.getElementById("newStudentActive").checked = v.active ?? true;
  d.getElementById("newStudentPace").checked = v.pace ?? false;
  d.getElementById("newStudentPulse").checked = v.pulse ?? false;
  await d.registry["addStudentBtn"]._handlers.click[0]();
}
async function setupToggle(b, cls, id, name) {
  const btn = { dataset: { spId: id, name } };
  const evt = { target: { closest: sel => (sel === cls ? btn : null) } };
  b.APP._bindSetupTabEvents("students");
  for (const h of b.document.registry["tabContent"]._handlers.click) await h(evt);
}

async function setupStudentsWorksWithBooleanEligibilityColumns() {
  const sp = makeSharePoint(byLabel(PROD_B));
  const b = await openBrowser(sp);
  await setupAdd(b, { first: "Sofia", last: "Reyes", pace: true });
  assert.equal(b.toasts.filter(t => t.t === "error").length, 0, "no error");
  assert.match(b.toasts.find(t => /added/.test(t.m)).m, /Student "Sofia Reyes" added/);
  const sent = plain(sp.state.posts[0]);
  assert.deepEqual([sent.Active, sent.PACEEnabled, sent.DailyPulseEnabled], [true, true, false], "real booleans, never 'Yes'/'No' text");
  assert.equal(sent.Title, "Sofia", "first name → the writable Title column");
  assert.ok(!("LinkTitle" in sent) && !("LinkTitleNoMenu" in sent));

  const id = sp.state.items.at(-1).id;
  await setupToggle(b, ".deactivate-student-btn", id, "Sofia Reyes");
  assert.deepEqual(plain(sp.state.patches.at(-1)), { id, fields: { Active: false } }, "deactivate = boolean false, that record only");
  assert.equal(sp.state.items.find(i => i.id === id).fields.Active, false);
  await setupToggle(b, ".reactivate-student-btn", id, "Sofia Reyes");
  assert.deepEqual(plain(sp.state.patches.at(-1)), { id, fields: { Active: true } });
  assert.equal(b.toasts.filter(t => t.t === "error").length, 0);
  assert.equal(sp.state.posts.length, 1, "toggling never creates a record");
}

async function setupStudentsWorksWithChoiceEligibilityColumns() {
  const sp = makeSharePoint(byLabel(PROD_C));
  const b = await openBrowser(sp);
  await setupAdd(b, { first: "Sofia", last: "Reyes", pulse: true });
  assert.equal(b.toasts.filter(t => t.t === "error").length, 0);
  const sent = plain(sp.state.posts[0]);
  assert.deepEqual([sent.Active, sent.PACEEnabled, sent.DailyPulseEnabled], ["Yes", "No", "Yes"], "valid configured choices");
  const id = sp.state.items.at(-1).id;
  await setupToggle(b, ".deactivate-student-btn", id, "Sofia Reyes");
  assert.deepEqual(plain(sp.state.patches.at(-1).fields), { Active: "No" });
  assert.equal(b.toasts.filter(t => t.t === "error").length, 0);
}

async function setupAndStudentsPageProduceTheSamePayloadForTheSameInput() {
  for (const label of [PROD, PROD_B, PROD_C]) {
    const a = makeSharePoint(byLabel(label)), c = makeSharePoint(byLabel(label));
    const viaSetup = await openBrowser(a);
    await setupAdd(viaSetup, { first: "Adelaide", last: "Testwell", classroom: "Room 4", pace: true });
    const viaPage = await openBrowser(c);
    viaPage.APP._openStudentAdd();
    fill(viaPage, { first: "Adelaide", last: "Testwell", classroom: "Room 4", pace: true });
    await viaPage.APP._saveStudentForm(false);
    assert.deepEqual(plain(a.state.posts), plain(c.state.posts), `${label}: identical write payload from both screens`);
    assert.equal(a.state.posts.length, 1);
    // ...and a Setup deactivate matches the page's Active edit
    const id = a.state.items.at(-1).id;
    await setupToggle(viaSetup, ".deactivate-student-btn", id, "Adelaide Testwell");
    viaPage.APP._openStudentEdit(c.state.items.at(-1).id);
    fill(viaPage, { first: "Adelaide", last: "Testwell", classroom: "Room 4", active: false, pace: true });
    await viaPage.APP._saveStudentForm(false);
    assert.deepEqual(plain(a.state.patches.at(-1).fields), plain(c.state.patches.at(-1).fields), `${label}: identical Active patch`);
  }
}

async function successfulCreateIsVerifiedByExactItemId() {
  const sp = makeSharePoint(byLabel(PROD_B));
  const b = await openBrowser(sp);
  const res = await b.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
  const newId = sp.state.items.at(-1).id;
  assert.equal(res.status, "verified");
  assert.equal(res.id, newId, "the id comes from the create response");
  assert.deepEqual(sp.state.itemReads, [newId], "the saved record was fetched by that exact id");
  assert.deepEqual([res.record.firstName, res.record.lastName], ["Adelaide", "Testwell"]);
  assert.equal(sp.state.posts.length, 1);
}

async function identicalNamesAreVerifiedByIdNotByName() {
  const sp = makeSharePoint(byLabel(PROD_B));
  const b = await openBrowser(sp);
  // The roster already has a "Jane Doe". An admin confirms a genuinely separate second one.
  b.APP._openStudentAdd();
  fill(b, { first: "Jane", last: "Doe", teacher: "Bossons, A" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 0, "duplicate protection holds it first");
  await b.APP._saveStudentForm(false, true);
  const newId = sp.state.items.at(-1).id;
  assert.equal(sp.state.posts.length, 1);
  assert.deepEqual(sp.state.itemReads, [newId], "verification read the NEW record's id, not the existing Jane Doe");
  assert.equal(b.STUDENT_ROSTER.getAll().filter(s => s.name === "Jane Doe").length, 2);
  assert.equal(sp.state.items.find(i => i.id === "11").fields.Title, "Jane", "the original record is untouched");
}

async function missingCreateResponseIdIsAnUnverifiedSaveAndNeverRetried() {
  const scn = { ...byLabel(PROD_B), postBehavior: "noId" };
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell" });
  await b.APP._saveStudentForm(false);

  assert.equal(sp.state.posts.length, 1);
  assert.equal(sp.state.items.length, 4, "SharePoint did create it");
  // The roster now contains a "Adelaide Testwell" — that must NOT be accepted as verification.
  assert.ok(b.STUDENT_ROSTER.getAll().some(s => s.name === "Adelaide Testwell"));
  assert.equal(successToasts(b).length, 0, "no success message");
  assert.match(b.toasts.at(-1).m, /NOT VERIFIED.*returned no item ID.*check the roster before adding it again/);
  assert.match(formHtml(b), /Saved, but NOT verified/);
  assert.deepEqual(sp.state.itemReads || [], [], "nothing to read by id, and the name was not used instead");

  // Saving is paused; nothing retries it.
  assert.match(formHtml(b), /id="stuSave" disabled/);
  await b.APP._saveStudentForm(false);
  await new Promise(r => setTimeout(r, 20));
  assert.equal(sp.state.posts.length, 1, "no automatic or accidental second creation");
  // After the admin re-checks the roster, a re-save is still caught by duplicate protection.
  await b.APP._recheckStudentRoster();
  assert.doesNotMatch(formHtml(b), /id="stuSave" disabled/);
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 1);
  assert.match(formHtml(b), /may already be on the roster/);
}

async function uncertainCreateIsNeverRetriedAndNeverDuplicated() {
  for (const behavior of ["dropAfterStore", "dropBeforeStore"]) {
    const sp = makeSharePoint({ ...byLabel(PROD_B), postBehavior: behavior });
    const b = await openBrowser(sp);
    const before = sp.state.items.length;
    const res = await b.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
    await new Promise(r => setTimeout(r, 20));
    assert.equal(res.status, "uncertain", `${behavior}: reported as uncertain, not failed or saved`);
    assert.equal(sp.state.posts.length, 1, `${behavior}: exactly one attempt — no retry`);
    assert.equal(sp.state.items.length, behavior === "dropAfterStore" ? before + 1 : before, `${behavior}: roster state is whatever SharePoint really holds`);
    assert.match(b.STUDENT_ROSTER.describeResult(res, "Adelaide Testwell").message, /UNCERTAIN.*NOT retried/);
  }
  // Through the page: no success, paused, never auto-retried.
  const sp = makeSharePoint({ ...byLabel(PROD_B), postBehavior: "dropAfterStore" });
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell" });
  await b.APP._saveStudentForm(false);
  await new Promise(r => setTimeout(r, 20));
  assert.equal(sp.state.posts.length, 1);
  assert.equal(successToasts(b).length, 0);
  assert.match(formHtml(b), /The result of this save is UNKNOWN/);
  assert.match(formHtml(b), /id="stuSave" disabled/);
  assert.ok(b.STUDENT_ROSTER.getAll().some(s => s.name === "Adelaide Testwell"), "the refreshed roster shows it, so the admin can see it exists");
}

async function failedReadBackIsReportedAsUnverifiedAndRetryUpdatesTheSameRecord() {
  const sp = makeSharePoint({ ...byLabel(PROD_B), getItemFails: true });
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Adelaide", last: "Testwell" });
  await b.APP._saveStudentForm(false);
  const newId = sp.state.items.at(-1).id;
  assert.equal(successToasts(b).length, 0, "an unreadable record is never reported as saved");
  assert.match(b.toasts.at(-1).m, /NOT VERIFIED.*could not be read back/);
  assert.equal(b.APP._studentsUi.mode, "edit", "the form now targets the record that exists");
  assert.equal(b.APP._studentsUi.editId, newId);
  fill(b, { first: "Adelaide", last: "Testwell", classroom: "Room 9" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 1, "the retry is an UPDATE of that record, never a second create");
  assert.equal(sp.state.patches.at(-1).id, newId);
}

async function sharePointRefusalIsFailedAndSafeToRetry() {
  const sp = makeSharePoint(byLabel(PROD_B));
  const b = await openBrowser(sp);
  const orig = sp.fetch;
  let refuse = true;
  const wrapped = async (url, opts = {}) => (refuse && opts.method === "POST" ? { ok: false, status: 403, text: async () => "forbidden", json: async () => ({}) } : orig(url, opts));
  b.context.fetch = wrapped;
  const res = await b.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
  assert.equal(res.status, "failed");
  assert.equal(res.phase, "write");
  refuse = false;
  const again = await b.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
  assert.equal(again.status, "verified", "a definite refusal wrote nothing, so a retry creates exactly one");
  assert.equal(sp.state.items.filter(i => i.fields.Title === "Adelaide").length, 1);
}

async function requiredWritableColumnsThisFormDoesNotFillBlockTheCreate() {
  const base = byLabel(PROD_B);
  const withRequired = extra => ({ ...base, columns: [...base.columns, extra], stored: [...base.stored, extra.name] });

  // A required writable column the form doesn't supply → refused before any write; no value is invented.
  const sp = makeSharePoint(withRequired(col("Grade", "Grade", { required: true })));
  const b = await openBrowser(sp);
  const res = await b.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
  assert.equal(res.status, "failed");
  assert.equal(res.phase, "prepare");
  assert.match(res.error, /requires "Grade".*a value will not be made up/);
  assert.equal(sp.state.posts.length, 0, "nothing sent");
  assert.deepEqual(await b.GRAPH.missingRequiredColumns("IEP_Students_2026_27", { Title: "x" }), ["Grade"]);

  // Not a problem: has a default, is hidden/read-only/computed, or is the Title we supply (even though Title is required).
  for (const ok of [
    col("Grade", "Grade", { required: true, defaultValue: { value: "9" } }),
    col("Grade", "Grade", { required: true, hidden: true }),
    col("Grade", "Grade", { required: true, readOnly: true }),
    col("Grade", "Grade", { required: true, calculated: {} })
  ]) {
    const s2 = makeSharePoint(withRequired(ok));
    const b2 = await openBrowser(s2);
    const r2 = await b2.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
    assert.equal(r2.status, "verified", JSON.stringify(ok));
  }
  const titleReq = makeSharePoint({ ...base, columns: base.columns.map(c => c.name === "Title" ? { ...c, required: true } : c) });
  const bt = await openBrowser(titleReq);
  assert.equal((await bt.STUDENT_ROSTER.create({ first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false })).status, "verified");

  // Required columns only gate creating; editing an existing record is unaffected.
  const s3 = makeSharePoint(withRequired(col("Grade", "Grade", { required: true })));
  const b3 = await openBrowser(s3);
  const upd = await b3.STUDENT_ROSTER.update("11", { classroom: "Room 7" });
  assert.equal(upd.status, "verified");
}

async function editVerifiesTheExactExistingRecordAndSetupRespectsAuthorization() {
  const sp = makeSharePoint(byLabel(PROD_B));
  const b = await openBrowser(sp);
  const res = await b.STUDENT_ROSTER.update("12", { classroom: "Room 7", teacher: "Stock, N" });
  assert.equal(res.status, "verified");
  assert.deepEqual(sp.state.itemReads, ["12"], "read back by the exact existing id");
  assert.deepEqual(plain(sp.state.patches.at(-1)), { id: "12", fields: { Classroom: "Room 7", Teacher: "Stock, N" } });
  assert.equal(sp.state.items.find(i => i.id === "11").fields.Teacher, "Bossons, A", "other records untouched");
  // an update with no id is refused before sending
  assert.equal((await b.STUDENT_ROSTER.update("", { classroom: "x" })).status, "failed");
  // no Admin Panel permission → the shared path refuses, from either screen
  vm.runInContext("MAC_ADMIN_PANEL_ALLOWED = false", b.context);
  const before = sp.state.patches.length + sp.state.posts.length;
  const denied = await b.STUDENT_ROSTER.create({ first: "No", last: "Access", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false });
  assert.equal(denied.status, "failed");
  assert.match(denied.error, /Admin Panel access required/);
  assert.equal(sp.state.patches.length + sp.state.posts.length, before);
}


/* ── complete-roster pagination ───────────────────────────────────────────── */

// 530 students over three pages of 200 (200 / 200 / 130). The LAST row is the one that only a
// complete load can see.
function pagedScenario(extra = {}) {
  const base = byLabel(PROD_B);
  const items = [];
  for (let i = 1; i <= 529; i++) items.push({ id: String(1000 + i), fields: { Title: `First${i}`, StudentLastName: `Last${i}`, Teacher: "Bossons, A", Classroom: "", Active: true, PACEEnabled: i % 2 === 0, DailyPulseEnabled: false } });
  items.push({ id: "9999", fields: { Title: "Zed", StudentLastName: "Zimmer", Teacher: "Stock, N", Classroom: "", Active: true, PACEEnabled: true, DailyPulseEnabled: true } });
  return { ...base, flagKind: undefined, items, pageSize: 200, ...extra };
}
const NEW_STUDENT = { first: "Adelaide", last: "Testwell", teacher: "Bossons, A", active: true, paceEnabled: false, dailyPulseEnabled: false };
const stuHtml = b => b.document.registry["page-students"].innerHTML;

async function aRosterSpanningSeveralPagesIsReadInFull() {
  const sp = makeSharePoint(pagedScenario());
  const b = await openBrowser(sp);
  assert.equal(b.STUDENT_ROSTER.status(), "complete");
  assert.equal(b.STUDENT_ROSTER.getAll().length, 530, "every page was read");
  assert.equal(sp.state.rosterRequests.length, 3, "three requests: 200 + 200 + 130");
  assert.match(sp.state.rosterRequests[0], /\$top=200/);
  assert.match(sp.state.rosterRequests[1], /skiptoken=200/);
  assert.match(sp.state.rosterRequests[2], /skiptoken=400/);
  const last = b.STUDENT_ROSTER.findByName("Zed Zimmer");
  assert.ok(last, "the student on the FINAL page is present");
  assert.equal(last.id, "9999");
  assert.equal(last.paceEnabled, true, "normalization/eligibility unchanged for later-page records");
  assert.equal(b.STUDENT_ROSTER.getPaceEnabled().some(s => s.name === "Zed Zimmer"), true);
  assert.equal(b.STUDENT_ROSTER.getDailyPulseForTeacher("Stock, N").length, 1);
  assert.equal(b.document.registry["stuCount"].textContent, "530 students");
  assert.match(listHtml(b), /Zed Zimmer/);
}

async function overlappingPagesDoNotCreateDuplicateRecords() {
  const sp = makeSharePoint(pagedScenario({ overlap: true }));
  const b = await openBrowser(sp);
  assert.equal(b.STUDENT_ROSTER.status(), "complete");
  assert.equal(b.STUDENT_ROSTER.getAll().length, 530, "repeated rows across pages are kept once");
  assert.equal(new Set(b.STUDENT_ROSTER.getAll().map(s => s.id)).size, 530);
}

async function duplicateDetectionFindsAStudentOnALaterPage() {
  const sp = makeSharePoint(pagedScenario());
  const b = await openBrowser(sp);
  assert.equal(b.STUDENT_ROSTER.findDuplicates("zed", "ZIMMER").length, 1, "found on page 3");
  b.APP._openStudentAdd();
  fill(b, { first: "Zed", last: "Zimmer", teacher: "Stock, N" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 0, "held for confirmation, not created");
  assert.match(formHtml(b), /may already be on the roster/);
  assert.match(formHtml(b), /Zed Zimmer — Stock, N · Active/);
}

async function aFailedSecondOrThirdPageIsAnIncompleteRosterNeverAPartialOne() {
  for (const failPage of [2, 3]) {
    const sp = makeSharePoint(pagedScenario({ pageFail: failPage }));
    const b = await openBrowser(sp);
    assert.equal(b.STUDENT_ROSTER.status(), "incomplete", `page ${failPage} failed`);
    assert.match(b.STUDENT_ROSTER.error, new RegExp(`Incomplete load of IEP_Students_2026_27: page ${failPage} failed`));
    assert.match(b.STUDENT_ROSTER.error, /partial result was discarded/);
    assert.equal(b.STUDENT_ROSTER.getAll().length, 0, "none of the pages that DID load are kept");
    assert.equal(b.STUDENT_ROSTER.isComplete(), false);
    // Every screen refuses to show it.
    assert.match(stuHtml(b), /The student roster is incomplete/);
    assert.match(stuHtml(b), /id="stuRetry"/, "actionable: manual retry offered");
    assert.doesNotMatch(stuHtml(b), /First1|Last1/);
    assert.match(b.APP._rosterStatusHtml("Daily Pulse"), /roster is incomplete[\s\S]*Try Again/);
    assert.match(b.APP._renderStudentsTab(), /The student roster is incomplete/);
    assert.doesNotMatch(b.APP._renderStudentsTab(), /id="addStudentBtn"|Students \(/, "Setup shows neither a list nor an Add form");
  }
}

async function malformedOrRepeatedContinuationLinksAreRejected() {
  const cases = [
    ["repeat",    { nextLinkMode: "repeat", badLinkFromPage: 2 },                /continuation link that was already requested/],
    ["foreign",   { nextLinkMode: "foreign", badLinkFromPage: 1 },               /invalid continuation link/],
    ["otherList", { nextLinkMode: "otherList", badLinkFromPage: 2 },             /invalid continuation link/],
    ["relative",  { nextLinkMode: "relative", badLinkFromPage: 1 },              /invalid continuation link/],
    ["badBody",   { malformedPage: 2 },                                          /page 2 was malformed/],
    ["noId",      { noIdOnPage: 3 },                                             /page 3 contained an item with no id/]
  ];
  for (const [name, extra, re] of cases) {
    const sp = makeSharePoint(pagedScenario(extra));
    const b = await openBrowser(sp);
    assert.equal(b.STUDENT_ROSTER.status(), "incomplete", name);
    assert.match(b.STUDENT_ROSTER.error, re, name);
    assert.equal(b.STUDENT_ROSTER.getAll().length, 0, `${name}: nothing partial kept`);
    assert.ok(sp.state.rosterRequests.length <= 3, `${name}: stopped, did not loop`);
  }
  // A link chain that never ends is cut off by the page limit instead of running forever.
  const endless = makeSharePoint(pagedScenario({ nextLinkMode: "endless" }));
  const b = await openBrowser(endless);
  assert.equal(b.STUDENT_ROSTER.status(), "incomplete");
  await assert.rejects(() => b.GRAPH.getAllListItems("IEP_Students_2026_27", { maxPages: 4 }), error => error.incompleteLoad === true && /stopped after 4 pages/.test(error.message));
}

async function aRefreshFailureAfterASuccessfulLoadNeverReplacesTheCompleteRoster() {
  const sp = makeSharePoint(pagedScenario({ pageFail: 3, pageFailOnce: true }));
  sp.state.pageFailed = true;   // let the initial load succeed; arm the failure afterwards
  const b = await openBrowser(sp);
  assert.equal(b.STUDENT_ROSTER.status(), "complete");
  assert.equal(b.STUDENT_ROSTER.getAll().length, 530);

  sp.state.pageFailed = false;  // next refresh: page 3 fails once
  await b.STUDENT_ROSTER.refresh();
  assert.equal(b.STUDENT_ROSTER.status(), "incomplete");
  assert.equal(b.STUDENT_ROSTER.getAll().length, 530, "the earlier complete roster was not replaced by pages 1-2");
  assert.equal(b.STUDENT_ROSTER.loaded, true);

  // ...but it is not presented as current: list hidden, retry offered, saving disabled, creation refused.
  b.APP.renderStudents();
  assert.match(stuHtml(b), /The student roster is incomplete/);
  b.APP._openStudentAdd();   // (page is on the error screen; drive the form directly)
  b.document.registry["page-students"].innerHTML = "";
  b.APP.renderStudents();
  const res = await b.STUDENT_ROSTER.create(NEW_STUDENT);
  assert.equal(res.status, "failed");
  assert.match(res.error, /roster is not fully loaded, so a duplicate can't be ruled out/);
  assert.equal(sp.state.posts.length, 0);

  // Manual retry recovers (the fake fails only once).
  await b.APP._retryRosterLoad();
  assert.equal(b.STUDENT_ROSTER.status(), "complete");
  assert.equal(b.STUDENT_ROSTER.getAll().length, 530);
  assert.match(stuHtml(b), /Zed Zimmer|stuAddBtn/);
}

async function noCreateIsAttemptedWhileDuplicateCheckingIsIncomplete() {
  // Incomplete from the start.
  const sp = makeSharePoint(pagedScenario({ pageFail: 2 }));
  const b = await openBrowser(sp);
  assert.throws(() => b.STUDENT_ROSTER.findDuplicates("Zed", "Zimmer"), /not fully loaded, so duplicate names can't be checked/, "an incomplete roster cannot declare a name unique");
  const res = await b.STUDENT_ROSTER.create(NEW_STUDENT);
  assert.equal(res.status, "failed");
  assert.equal(sp.state.posts.length, 0);

  // Students page: the form can be opened, but Save is disabled and a forced save is refused before any write.
  b.STUDENT_ROSTER.loaded = true;        // (simulate "a roster exists" so the page renders its form)
  b.STUDENT_ROSTER.error = "Incomplete load of IEP_Students_2026_27: page 2 failed (x). The partial result was discarded.";
  b.APP._studentsUi.mode = "add";
  b.APP._studentsUi.draft = { first: "Adelaide", last: "Testwell", teacher: "Bossons, A", classroom: "", active: true, paceEnabled: false, dailyPulseEnabled: false };
  b.APP._renderStudentsForm();
  assert.match(formHtml(b), /Saving is disabled/);
  assert.match(formHtml(b), /id="stuSave" disabled/);
  assert.match(formHtml(b), /id="stuSaveAnother" disabled/);
  fill(b, { first: "Adelaide", last: "Testwell", teacher: "Bossons, A" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 0);
  assert.match(b.toasts.at(-1).m, /Nothing was saved/);
  assert.equal(successToasts(b).length, 0);

  // Setup → Students: refused as well (and no confirm() duplicate question is asked on made-up data).
  await setupAdd(b, { first: "Adelaide", last: "Testwell" });
  assert.equal(sp.state.posts.length, 0);
  assert.match(b.toasts.at(-1).m, /Nothing was saved: The student roster is not fully loaded/);
}

async function editingAndExactIdVerificationStillWorkOnAMultiPageRoster() {
  const sp = makeSharePoint(pagedScenario());
  const b = await openBrowser(sp);
  // Edit the student that only exists on page 3.
  b.APP._openStudentEdit("9999");
  assert.match(formHtml(b), /id="stuFirst"[^>]*value="Zed"/);
  fill(b, { first: "Zed", last: "Zimmer", teacher: "Stock, N", classroom: "Room 8", active: true, pace: true, pulse: true });
  await b.APP._saveStudentForm(false);
  assert.deepEqual(plain(sp.state.patches.at(-1)), { id: "9999", fields: { Classroom: "Room 8" } });
  assert.deepEqual(sp.state.itemReads, ["9999"], "verified by the exact id");
  assert.equal(b.toasts.filter(t => t.t === "error").length, 0);
  assert.equal(b.STUDENT_ROSTER.find("9999").classroom, "Room 8");

  // A new student is created, verified by its id, and found again on the re-read of all three pages.
  const before = sp.state.rosterRequests.length;
  const res = await b.STUDENT_ROSTER.create(NEW_STUDENT);
  assert.equal(res.status, "verified");
  assert.deepEqual(sp.state.itemReads.at(-1), res.id);
  assert.equal(b.STUDENT_ROSTER.getAll().length, 531);
  assert.equal(sp.state.rosterRequests.length - before, 3, "the post-save refresh also reads every page");
  assert.equal(b.STUDENT_ROSTER.status(), "complete");
}

async function allRosterReadersNowUseTheCompleteReader() {
  const code = APP_SRC.split("\n").filter(l => !l.trim().startsWith("//")).join("\n");
  assert.ok(!/getListItems\(\s*(listName|STUDENT_ROSTER\._listName\(\)|SETUP_LISTS\.students|CONFIG\.STUDENT_ROSTER_LIST)/.test(code), "no single-page read of the roster list remains");
  assert.match(APP_SRC, /GRAPH\.getAllListItems\(listName\)/);
  assert.equal((code.match(/GRAPH\.getListItems\(SETUP_LISTS\.classrooms\)/g) || []).length, 1, "other lists are untouched");
}

/* ── other workflows still behave ─────────────────────────────────────────── */

async function setupStudentsPaceAndDailyPulseStillWork() {
  const sp = makeSharePoint(SCENARIOS[0]);
  const b = await openBrowser(sp);

  // Setup → Students' own Add Student (unchanged code) now also keeps the first name.
  const d = b.document;
  b.APP._bindSetupTabEvents("students");
  d.getElementById("newStudentFirstName").value = "Sofia";
  d.getElementById("newStudentLastName").value = "Reyes";
  d.getElementById("newStudentTeacher").value = "Bossons, A";
  d.getElementById("newStudentClassroom").value = "";
  d.getElementById("newStudentActive").checked = true;
  d.getElementById("newStudentPace").checked = true;
  d.getElementById("newStudentPulse").checked = true;
  await d.registry["addStudentBtn"]._handlers.click[0]();
  const sofia = sp.state.items.find(i => i.fields.StudentLastName === "Reyes");
  assert.equal(sofia.fields.Title, "Sofia", "Setup → Students: first name stored");

  // PACE / Daily Pulse read the same roster with intact names and flags.
  const fresh = await openBrowser(sp);
  assert.deepEqual(plain(fresh.STUDENT_ROSTER.getPaceEnabled().map(s => s.name).sort()), ["Jane Doe", "Sofia Reyes"]);
  assert.deepEqual(plain(fresh.STUDENT_ROSTER.getDailyPulseEnabled().map(s => s.name).sort()), ["Maya Ortiz", "Sofia Reyes"]);
  assert.deepEqual(plain(fresh.STUDENT_ROSTER.getForTeacher("Bossons, A").map(s => s.name).sort()), ["Jane Doe", "Maya Ortiz", "Quill", "Sofia Reyes"]);
}

async function newWalkthroughIsUnchanged() {
  const b = await openBrowser(makeSharePoint(SCENARIOS[0]));
  const reads = { n: 0 };
  const before = b.STUDENT_ROSTER.getAll().length;
  b.TEACHER_DIRECTORY._teachers = [{ id: "u1", name: "Amber Bossons" }];
  b.TEACHER_DIRECTORY.loaded = true;
  b.APP.currentPage = "walkthrough";
  b.APP.renderWalkthrough();
  const html = b.document.registry["page-walkthrough"].innerHTML;
  assert.match(html, /<input type="text" class="form-input" name="studentName" id="walkthroughStudent"/, "Student is still optional free text");
  assert.ok(!/<select[^>]*studentId/.test(html));
  assert.equal(b.STUDENT_ROSTER.getAll().length, before);
  const src = APP_SRC.slice(APP_SRC.indexOf("renderWalkthrough() {"), APP_SRC.indexOf("renderWalkthroughV2() {"));
  const code = src.split("\n").filter(l => !l.trim().startsWith("//")).join("\n");
  assert.ok(!/STUDENT_ROSTER/.test(code), "no roster lookup in New Walkthrough");
  void reads;
}

const tests = {
  aRosterSpanningSeveralPagesIsReadInFull, overlappingPagesDoNotCreateDuplicateRecords,
  duplicateDetectionFindsAStudentOnALaterPage, aFailedSecondOrThirdPageIsAnIncompleteRosterNeverAPartialOne,
  malformedOrRepeatedContinuationLinksAreRejected, aRefreshFailureAfterASuccessfulLoadNeverReplacesTheCompleteRoster,
  noCreateIsAttemptedWhileDuplicateCheckingIsIncomplete, editingAndExactIdVerificationStillWorkOnAMultiPageRoster,
  allRosterReadersNowUseTheCompleteReader,
  setupStudentsWorksWithBooleanEligibilityColumns, setupStudentsWorksWithChoiceEligibilityColumns,
  setupAndStudentsPageProduceTheSamePayloadForTheSameInput, successfulCreateIsVerifiedByExactItemId,
  identicalNamesAreVerifiedByIdNotByName, missingCreateResponseIdIsAnUnverifiedSaveAndNeverRetried,
  uncertainCreateIsNeverRetriedAndNeverDuplicated, failedReadBackIsReportedAsUnverifiedAndRetryUpdatesTheSameRecord,
  sharePointRefusalIsFailedAndSafeToRetry, requiredWritableColumnsThisFormDoesNotFillBlockTheCreate,
  editVerifiesTheExactExistingRecordAndSetupRespectsAuthorization,
  titleTranslationIsRestrictedToTheStudentRosterList, addAndEditShareOneVerifiedMappingAndBothAreStrict,
  readBackChecksTheActualSavedFirstNameNotJustGraphSuccess, flagPayloadsMatchTheRealColumnTypes,
  booleanColumnsNeverReceiveYesNoStringsAndUnclearTypesAreRejected, eligibilityRulesAndExistingRecordsAreUntouched,
  productionShapeOriginalMappingFailsAndCorrectedPayloadIsValid, ambiguousOrUnverifiableMappingsAreRejectedNotGuessed,
  sameDefectWasAlsoBreakingSetupStudentsAndEditAndNowReadsAreHardened, otherListsAreUnaffectedByTheTitleTranslation,
  schemaBuilderNeverLetsAComputedColumnShadowTheRealOne, strictMappingRefusesToDropFieldsButDefaultStaysLenient,
  addEditAndRefreshPreserveBothNamesInEveryLayout, missingFirstNameColumnFailsLoudlyAndWritesNothing,
  aWriteSharePointAcceptsButDoesNotKeepIsNeverReportedAsSaved, setupStudentsPaceAndDailyPulseStillWork,
  newWalkthroughIsUnchanged
};

(async () => {
  let failed = 0;
  for (const [name, fn] of Object.entries(tests)) {
    try { await fn(); console.log("  ✓ " + name); }
    catch (err) { failed++; console.error("  ✗ " + name + "\n" + (err.stack || err)); }
  }
  if (failed) { console.error(`\n${failed} failing`); process.exit(1); }
  console.log(`\nAll ${Object.keys(tests).length} first-name regression checks passed.`);
})();
