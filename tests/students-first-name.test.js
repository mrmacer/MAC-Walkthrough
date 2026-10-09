/* ─────────────────────────────────────────────────────────────────────────
   Regression: a new student's FIRST NAME was not persisting

   Symptom (production): adding "Lorenzo Cassidy" from the Students page
   listed the student as just "Cassidy"; Edit showed First Name blank, and
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
    items: scn.items.map(i => ({ id: i.id, fields: { ...i.fields } })),
    nextId: 500, posts: [], patches: []
  };
  const stored = new Set(scn.stored);
  const respond = (status, body) => ({
    ok: status >= 200 && status < 300, status,
    json: async () => body, text: async () => JSON.stringify(body)
  });
  const withComputed = fields => ({ ...fields, ...(fields.Title !== undefined ? { LinkTitle: fields.Title, LinkTitleNoMenu: fields.Title } : {}) });
  const unknown = keys => keys.find(k => !stored.has(k));
  const ignored = new Set(scn.silentlyIgnored || []);
  const keep = fields => Object.fromEntries(Object.entries(fields).filter(([k]) => !ignored.has(k)));

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
        { id: "list-users",  name: "IEP_Users2",           displayName: "IEP_Users2" }
      ] });
    }
    if (/\/lists\/list-roster\/columns/.test(url)) return respond(200, { value: scn.columns });
    if (/\/lists\/list-users\/items/.test(url)) return respond(200, { value: USERS });
    if (/\/lists\/list-roster\/items\?/.test(url) && method === "GET") {
      return respond(200, { value: state.items.map(i => ({ id: i.id, fields: withComputed(i.fields) })) });
    }
    if (/\/lists\/list-roster\/items$/.test(url) && method === "POST") {
      state.posts.push(body.fields);
      const bad = unknown(Object.keys(body.fields));
      if (bad) return respond(400, { error: { message: `Field '${bad}' is not recognized` } });
      const item = { id: String(state.nextId++), fields: keep(body.fields) };
      state.items.push(item);
      return respond(201, { id: item.id, fields: withComputed(item.fields) });
    }
    const patch = url.match(/\/lists\/list-roster\/items\/([^/]+)\/fields$/);
    if (patch && method === "PATCH") {
      state.patches.push({ id: patch[1], fields: body });
      const bad = unknown(Object.keys(body));
      if (bad) return respond(400, { error: { message: `Field '${bad}' is not recognized` } });
      const item = state.items.find(i => i.id === patch[1]);
      Object.assign(item.fields, keep(body));
      return respond(200, withComputed(item.fields));
    }
    throw new Error(`Unexpected Graph request in test: ${method} ${url}`);
  }
  return { state, fetch };
}

/* ── realistic roster column layouts ──────────────────────────────────────── */

const col = (name, displayName, extra = {}) => ({ name, displayName, ...extra });
const COMMON = [
  col("Teacher", "Teacher"), col("Classroom", "Classroom"), col("Active", "Active"),
  col("PACEEnabled", "PACE Enabled"), col("DailyPulseEnabled", "Daily Pulse Enabled"),
  col("ContentType", "Content Type"), col("Modified", "Modified"), col("Created", "Created")
];
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
  const sp2 = makeSharePoint({ ...SCENARIOS[2], columns: [col("A1", "Dup"), col("A2", "Dup"), col("LinkTitle", "Dup")], stored: ["A1", "A2"], items: [] });
  const b2 = await openBrowser(sp2);
  assert.equal((await b2.GRAPH.getListSchema("IEP_Students_2026_27"))["Dup"], "A2");
  // A read-only column alone is still reported (callers decide what to do with it).
  const sp3 = makeSharePoint({ ...SCENARIOS[2], columns: [col("LinkTitle", "Only")], stored: [], items: [] });
  const b3 = await openBrowser(sp3);
  assert.equal((await b3.GRAPH.getListSchema("IEP_Students_2026_27"))["Only"], "LinkTitle");
}

async function strictMappingRefusesToDropFieldsButDefaultStaysLenient() {
  const sp = makeSharePoint({ ...SCENARIOS[2], columns: [col("Title", "Title"), col("StudentLastName", "Student Last Name")], stored: ["Title", "StudentLastName"], items: [] });
  const b = await openBrowser(sp);
  const GRAPH = b.GRAPH;
  await assert.rejects(
    () => GRAPH.createMappedListItem("IEP_Students_2026_27", { "Student First Name": "Lorenzo", "Student Last Name": "Cassidy" }, { strict: true }),
    /Column mapping problem.*nothing was saved.*"Student First Name" has no matching column/
  );
  assert.equal(sp.state.posts.length, 0, "strict mode aborts BEFORE writing anything");
  // Lenient default is untouched — other lists rely on skipping optional columns.
  await GRAPH.createMappedListItem("IEP_Students_2026_27", { "Student First Name": "Lorenzo", "Student Last Name": "Cassidy" });
  assert.deepEqual(plain(sp.state.posts[0]), { StudentLastName: "Cassidy" });
}

/* ── the full user journey, per schema layout ─────────────────────────────── */

async function addEditAndRefreshPreserveBothNamesInEveryLayout() {
  for (const scn of SCENARIOS) {
    const L = scn.label;
    const sp = makeSharePoint(scn);
    const before = plain(sp.state.items);
    let b = await openBrowser(sp);

    // Existing students read correctly (and the record with no first name stays that way — nothing is invented).
    const roster = b.STUDENT_ROSTER.getAll().map(s => s.name);
    assert.deepEqual(roster, ["Jane Doe", "Maya Ortiz", "Quill"], `${L}: existing records read back unchanged`);

    // 1–4. Add Lorenzo Cassidy.
    b.APP._openStudentAdd();
    fill(b, { first: "Lorenzo", last: "Cassidy", classroom: "Room 4", pace: true });
    await b.APP._saveStudentForm(false);
    assert.equal(sp.state.posts.length, 1, `${L}: exactly one create`);
    const created = sp.state.items.find(i => i.fields[scn.keys.last] === "Cassidy");
    assert.equal(created.fields[scn.keys.first], "Lorenzo", `${L}: FIRST name is stored in SharePoint (${scn.keys.first})`);
    assert.equal(created.fields[scn.keys.last], "Cassidy", `${L}: LAST name is stored (${scn.keys.last})`);
    assert.equal(created.fields.Teacher, "Bossons, A");
    assert.equal(created.fields.Active, "Yes");
    assert.equal(created.fields.PACEEnabled, "Yes");
    assert.equal(created.fields.DailyPulseEnabled, "No");
    assert.match(listHtml(b), /Lorenzo Cassidy/, `${L}: directory shows the full name`);
    assert.equal(b.toasts.filter(t => t.t === "error").length, 0, `${L}: no error`);
    const newId = created.id;

    // 8. "Refresh the browser": a brand-new session reading the same SharePoint.
    b = await openBrowser(sp);
    const rec = b.STUDENT_ROSTER.find(newId);
    assert.deepEqual([rec.firstName, rec.lastName, rec.name], ["Lorenzo", "Cassidy", "Lorenzo Cassidy"], `${L}: names survive normalization after a refresh`);
    assert.match(listHtml(b), /Lorenzo Cassidy/);

    // 5. Edit prepopulates both names.
    b.APP._openStudentEdit(newId);
    assert.match(formHtml(b), /id="stuFirst"[^>]*value="Lorenzo"/, `${L}: Edit shows the first name`);
    assert.match(formHtml(b), /id="stuLast"[^>]*value="Cassidy"/, `${L}: Edit shows the last name`);

    // 6. Edit the first name → only that column is sent; id/teacher/flags unchanged.
    fill(b, { first: "Lawrence", last: "Cassidy", classroom: "Room 4", pace: true });
    await b.APP._saveStudentForm(false);
    const patch1 = sp.state.patches.at(-1);
    assert.equal(patch1.id, newId, `${L}: same record id`);
    assert.deepEqual(plain(patch1.fields), { [scn.keys.first]: "Lawrence" }, `${L}: PATCH carries just the first name`);
    assert.equal(sp.state.items.find(i => i.id === newId).fields[scn.keys.first], "Lawrence");
    // 7. Edit the last name.
    b.APP._openStudentEdit(newId);
    fill(b, { first: "Lawrence", last: "Cassidi", classroom: "Room 4", pace: true });
    await b.APP._saveStudentForm(false);
    assert.deepEqual(plain(sp.state.patches.at(-1).fields), { [scn.keys.last]: "Cassidi" });
    const after = sp.state.items.find(i => i.id === newId).fields;
    assert.equal(after[scn.keys.first], "Lawrence");
    assert.equal(after[scn.keys.last], "Cassidi");
    // 9–11. id, teacher and flags untouched by name edits.
    assert.equal(after.Teacher, "Bossons, A");
    assert.deepEqual([after.Active, after.PACEEnabled, after.DailyPulseEnabled], ["Yes", "Yes", "No"]);
    assert.equal(sp.state.items.filter(i => i.fields[scn.keys.last] === "Cassidi").length, 1, `${L}: still one record, not a re-create`);

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
  fill(b, { first: "Lorenzo", last: "Cassidy" });
  await b.APP._saveStudentForm(false);
  assert.equal(sp.state.posts.length, 0, "nothing written when the first-name column cannot be found");
  assert.equal(successToasts(b).length, 0, "no false success");
  assert.match(b.toasts.at(-1).m, /no column for the student's first name\. Columns found:.*Student Last Name/);
  assert.equal(b.APP._studentsUi.mode, "add", "form stays open with the typed values");
}

async function aWriteSharePointAcceptsButDoesNotKeepIsNeverReportedAsSaved() {
  const scn = { ...SCENARIOS[2], silentlyIgnored: ["StudentFirstName"] };
  const sp = makeSharePoint(scn);
  const b = await openBrowser(sp);
  b.APP._openStudentAdd();
  fill(b, { first: "Lorenzo", last: "Cassidy" });
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
  assert.deepEqual(fresh.STUDENT_ROSTER.getPaceEnabled().map(s => s.name).sort(), ["Jane Doe", "Sofia Reyes"]);
  assert.deepEqual(fresh.STUDENT_ROSTER.getDailyPulseEnabled().map(s => s.name).sort(), ["Maya Ortiz", "Sofia Reyes"]);
  assert.deepEqual(fresh.STUDENT_ROSTER.getForTeacher("Bossons, A").map(s => s.name).sort(), ["Jane Doe", "Maya Ortiz", "Quill", "Sofia Reyes"]);
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
