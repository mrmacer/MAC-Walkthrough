# Local kit (not part of the app)

This file tells an AI agent how to rebuild the local test kit for MAC Walkthrough on a new PC.
The kit lives outside the repo in `~/.mac-walkthrough/` and is never committed. Recreate each
file below exactly as shown.

## `~/.mac-walkthrough/sandbox/` — run the app in a browser without Microsoft 365

The sandbox serves a copy of this repo at http://localhost:8329 with a fake Microsoft sign-in
(always signed in as an Administrator) and a fake SharePoint that keeps its data in the browser's
localStorage. The repo's own files are only read, never changed. The copy is refreshed on every
page load, so edits in the repo show up with a browser refresh.

Start it:

```
node ~/.mac-walkthrough/sandbox/server.js
```

(Optional argument: the repo folder; default `~/Projects/MAC-Walkthrough`.) Needs Node 18+.

The orange **SANDBOX** panel (bottom right) shows the student list exactly as stored, resets the
sandbox to its starting data, and signs out/in. Starting data (all made up): admin "Sandbox Admin", teachers Pat Sample
and Robin Example, students Test StudentOne (all on), Test StudentTwo (PACE/Pulse stored No —
an older record), Test StudentThree (Pulse stored No) and Test Inactive (inactive). Only the users and student
lists have real columns, so saving walkthroughs, PACE visits or Daily Pulse entries in the
sandbox is not representative.

`server.js` also creates `~/.mac-walkthrough/sandbox/app/` (the served copy); it is rebuilt
automatically and does not need to be recreated by hand.

### `~/.mac-walkthrough/sandbox/server.js`

```js
// MAC Walkthrough browser sandbox — local only, never part of the repo.
// Copies the project into ./app (a fresh copy on every page load, so edits in
// the repo show up with a browser refresh), swaps the Microsoft sign-in
// library for fake-msal.js and adds fake-sharepoint.js, then serves it at
// http://localhost:8329. The project's own files are only ever read.
//
//   node server.js [path-to-repo]      (default: ~/Projects/MAC-Walkthrough)

const http = require("node:http");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const REPO = path.resolve(process.argv[2] || path.join(os.homedir(), "Projects", "MAC-Walkthrough"));
const HERE = __dirname;
const APP = path.join(HERE, "app");
const PORT = Number(process.env.PORT) || 8329;
const SKIP = new Set([".git", ".vercel", ".claude", "tests", "docs", "node_modules"]);
const MSAL_TAG = /<script src="https:\/\/cdn\.jsdelivr\.net\/npm\/@azure\/msal-browser[^"]*"><\/script>/;
const AUTH_TAG = '<script src="auth.js"></script>';

if (!fs.existsSync(path.join(REPO, "index.html"))) {
  console.error(`No index.html in ${REPO}. Pass the repo folder: node server.js <path-to-MAC-Walkthrough>`);
  process.exit(1);
}

function copyProject() {
  fs.rmSync(APP, { recursive: true, force: true });
  fs.cpSync(REPO, APP, {
    recursive: true,
    filter: src => {
      const rel = path.relative(REPO, src);
      return !rel || (!SKIP.has(rel.split(path.sep)[0]) && !rel.toLowerCase().endsWith(".md"));
    }
  });
  let html = fs.readFileSync(path.join(APP, "index.html"), "utf8");
  if (!MSAL_TAG.test(html) || !html.includes(AUTH_TAG)) {
    throw new Error("index.html no longer has the expected MSAL / auth.js script tags; update server.js.");
  }
  html = html
    .replace(MSAL_TAG, '<script src="/sandbox/fake-msal.js"></script>')
    .replace(AUTH_TAG, `<script src="/sandbox/fake-sharepoint.js"></script>\n  ${AUTH_TAG}`);
  fs.writeFileSync(path.join(APP, "index.html"), html);
}

const TYPES = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8",
  ".png": "image/png", ".svg": "image/svg+xml", ".json": "application/json", ".ico": "image/x-icon"
};

copyProject();

http.createServer((req, res) => {
  try {
    let urlPath = decodeURIComponent(new URL(req.url, "http://x").pathname);
    if (urlPath === "/" || urlPath === "/index.html") { copyProject(); urlPath = "/index.html"; }
    const root = urlPath.startsWith("/sandbox/") ? HERE : APP;
    const rel = urlPath.startsWith("/sandbox/") ? urlPath.slice("/sandbox/".length) : urlPath.slice(1);
    const file = path.resolve(root, rel);
    const allowed = root === HERE ? ["fake-msal.js", "fake-sharepoint.js"].includes(rel) : file.startsWith(APP + path.sep);
    if (!allowed || !fs.existsSync(file) || !fs.statSync(file).isFile()) { res.writeHead(404); return res.end("Not found"); }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(file).toLowerCase()] || "application/octet-stream", "Cache-Control": "no-store" });
    fs.createReadStream(file).pipe(res);
  } catch (err) {
    res.writeHead(500); res.end(String(err.message || err));
  }
}).listen(PORT, "127.0.0.1", () => {
  console.log(`MAC Walkthrough sandbox: http://localhost:${PORT}`);
  console.log(`Serving a copy of ${REPO} (re-copied on every page load). Ctrl+C to stop.`);
});
```

### `~/.mac-walkthrough/sandbox/fake-msal.js`

```js
// Sandbox stand-in for the Microsoft sign-in library (msal-browser).
// Always "signed in" as the sandbox administrator unless Sign out was pressed.
window.msal = {
  PublicClientApplication: class {
    async handleRedirectPromise() { return null; }
    getAllAccounts() {
      return localStorage.getItem("sandbox.signedOut") === "1"
        ? []
        : [{ username: "admin@sandbox.local", name: "Sandbox Admin", homeAccountId: "sandbox-admin" }];
    }
    async acquireTokenSilent() { return { accessToken: "sandbox-token" }; }
    async acquireTokenRedirect() {}
    loginRedirect() { localStorage.removeItem("sandbox.signedOut"); location.reload(); }
    logoutRedirect() { localStorage.setItem("sandbox.signedOut", "1"); location.reload(); }
  }
};
```

### `~/.mac-walkthrough/sandbox/fake-sharepoint.js`

```js
// Sandbox stand-in for SharePoint: answers the Microsoft Graph requests graph.js
// sends (lists, columns, items, create, update, delete) from data kept in this
// browser's localStorage. Includes a corner panel showing the raw student list.
(function () {
  const KEY = "sandbox.sharepoint.v2";
  const BASE = "https://graph.microsoft.com/v1.0/";
  const SITE_ID = "sandbox-site";
  const col = (name, displayName, type) => ({ name, displayName: displayName || name, hidden: false, readOnly: false, required: false, [type || "text"]: {} });
  const yesNo = (name, displayName) => col(name, displayName, "boolean");
  const user = (id, title, name, role, email) => ({ id, fields: { Title: title, field_1: name, field_2: role, field_3: "Yes", Email: email } });
  const student = (id, first, last, teacher, room, active, pace, pulse) => ({
    id, fields: { StudentFirstName: first, StudentLastName: last, Teacher: teacher, Classroom: room, Active: active, PACEEnabled: pace, DailyPulseEnabled: pulse }
  });
  const OTHER_LISTS = ["IEP_Daily_Pulse", "IEP_Walkthrough_Observations", "IEP_Walkthrough_Observations_V2", "IEP_Pace_Visits",
    "IEP_Student_CheckIn", "IEP_Weekly_Reflections", "IEP_Support_Requests", "IEP_School_Settings", "IEP_Skook_Pilot_Students"];

  function seed() {
    const lists = {
      IEP_Users2: {
        columns: [col("Title"), col("field_1", "Name"), col("field_2", "Role"), col("field_3", "Active"), col("Email")],
        items: [
          user("1", "admin-sandbox", "Sandbox Admin", "Administrator", "admin@sandbox.local"),
          user("2", "teacher-sample", "Pat Sample", "Teacher", "pat.sample@sandbox.local"),
          user("3", "teacher-example", "Robin Example", "Teacher", "robin.example@sandbox.local")
        ]
      },
      IEP_App_Users: { columns: [col("Title"), col("Email")], items: [] },
      IEP_Students_2026_27: {
        columns: [col("Title"), col("StudentFirstName", "Student First Name"), col("StudentLastName", "Student Last Name"),
          col("Teacher"), col("Classroom"), yesNo("Active"), yesNo("PACEEnabled", "PACE Enabled"), yesNo("DailyPulseEnabled", "Daily Pulse Enabled")],
        items: [
          student("101", "Test", "StudentOne", "Sample, P", "Room 1", true, true, true),
          student("102", "Test", "StudentTwo", "Sample, P", "Room 1", true, false, false),   // older record saved under the old "off" default
          student("103", "Test", "StudentThree", "Example, R", "Room 4", true, true, false),
          student("104", "Test", "Inactive", "Example, R", "", false, false, false)           // inactive
        ]
      }
    };
    OTHER_LISTS.forEach(n => { lists[n] = { columns: [col("Title")], items: [] }; });
    return { nextId: 1000, lists };
  }

  let db;
  try { db = JSON.parse(localStorage.getItem(KEY)) || seed(); } catch { db = seed(); }
  const save = () => localStorage.setItem(KEY, JSON.stringify(db));
  save();

  const reply = (status, body) => new Response(status === 204 ? null : JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
  const err = (status, message) => reply(status, { error: { code: String(status), message } });
  const listById = id => Object.keys(db.lists).find(n => "list-" + n === id);

  // Like Graph: unknown column or a value that doesn't fit the column type → 400.
  function problem(list, fields) {
    for (const [k, v] of Object.entries(fields || {})) {
      const c = list.columns.find(x => x.name === k);
      if (!c) return `Field '${k}' is not recognized`;
      if (c.boolean && typeof v !== "boolean") return `Field '${k}' expects true/false`;
      if (c.text && typeof v !== "string") return `Field '${k}' expects text`;
    }
    return null;
  }

  async function fakeGraph(url, opts) {
    const method = (opts.method || "GET").toUpperCase();
    const [p] = url.slice(BASE.length).split("?");
    if (p.startsWith("sites/") && !p.startsWith(`sites/${SITE_ID}`)) return reply(200, { id: SITE_ID });
    if (p === `sites/${SITE_ID}/lists`) return reply(200, { value: Object.keys(db.lists).map(n => ({ id: "list-" + n, name: n, displayName: n })) });
    const m = p.match(new RegExp(`^sites/${SITE_ID}/lists/([^/]+)/(columns|items)(?:/([^/]+))?(/fields)?$`));
    const name = m && listById(m[1]);
    if (!name) return err(404, "Not found: " + p);
    const list = db.lists[name];
    const body = opts.body ? JSON.parse(opts.body) : {};
    if (m[2] === "columns") return reply(200, { value: list.columns });
    const itemId = m[3] && decodeURIComponent(m[3]);
    const item = itemId && list.items.find(i => i.id === itemId);
    if (!itemId && method === "GET") return reply(200, { value: list.items.map(i => ({ id: i.id, fields: { ...i.fields } })) });
    if (!itemId && method === "POST") {
      const bad = problem(list, body.fields); if (bad) return err(400, bad);
      const created = { id: String(db.nextId++), fields: { ...body.fields } };
      list.items.push(created); save();
      return reply(201, { id: created.id, fields: { ...created.fields } });
    }
    if (!item) return err(404, "Item not found");
    if (method === "GET") return reply(200, { id: item.id, fields: { ...item.fields } });
    if (method === "PATCH" && m[4]) {
      const bad = problem(list, body); if (bad) return err(400, bad);
      Object.assign(item.fields, body); save();
      return reply(200, { ...item.fields });
    }
    if (method === "DELETE") { list.items = list.items.filter(i => i !== item); save(); return reply(204); }
    return err(405, "Method not allowed");
  }

  const realFetch = window.fetch.bind(window);
  window.fetch = (input, opts = {}) => {
    const url = typeof input === "string" ? input : input.url;
    return url.startsWith(BASE) ? fakeGraph(url, opts) : realFetch(input, opts);
  };

  // ── Corner panel ──────────────────────────────────────────────────────────
  function panel() {
    const box = document.createElement("div");
    box.style.cssText = "position:fixed;right:12px;bottom:12px;z-index:99999;font:12px/1.4 system-ui,sans-serif;background:#fff7ed;color:#7c2d12;border:2px solid #f97316;border-radius:8px;padding:8px 10px;max-width:min(560px,calc(100vw - 24px));max-height:70vh;overflow:auto;box-shadow:0 4px 14px rgba(0,0,0,.15)";
    const btn = (label, fn) => { const b = document.createElement("button"); b.textContent = label; b.style.cssText = "margin-left:6px;font:inherit;cursor:pointer"; b.onclick = fn; return b; };
    const head = document.createElement("div");
    head.innerHTML = "<strong>SANDBOX</strong> · fake SharePoint ";
    const table = document.createElement("div");
    const show = () => {
      const rows = db.lists.IEP_Students_2026_27.items;
      const yn = v => v === true ? "Yes" : v === false ? "No" : "(blank)";
      table.innerHTML = `<table style="margin-top:6px;border-collapse:collapse;width:100%">
        <tr><th align="left">Student (as stored in IEP_Students_2026_27)</th><th>Active</th><th>PACE</th><th>Pulse</th></tr>
        ${rows.map(r => `<tr><td>${r.fields.StudentFirstName || ""} ${r.fields.StudentLastName || ""}</td><td align="center">${yn(r.fields.Active)}</td><td align="center">${yn(r.fields.PACEEnabled)}</td><td align="center">${yn(r.fields.DailyPulseEnabled)}</td></tr>`).join("")}
      </table>`;
    };
    let open = false;
    head.append(
      btn("Show stored roster", e => { open = !open; e.target.textContent = open ? "Hide stored roster" : "Show stored roster"; open ? show() : (table.innerHTML = ""); }),
      btn("Reset data", () => { if (confirm("Reset the sandbox to its starting students and clear local app data?")) { localStorage.clear(); location.reload(); } }),
      btn(localStorage.getItem("sandbox.signedOut") === "1" ? "Sign in" : "Sign out", () => {
        if (localStorage.getItem("sandbox.signedOut") === "1") localStorage.removeItem("sandbox.signedOut"); else localStorage.setItem("sandbox.signedOut", "1");
        location.reload();
      })
    );
    box.append(head, table);
    document.body.appendChild(box);
    setInterval(() => { if (open) show(); }, 1000);
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", panel); else panel();
})();
```

