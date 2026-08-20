# MAC Walkthrough — IU29 Pilot

A professional walkthrough and support-visibility tool for educational administrators. Built for the special education supervisor pilot at IU29.

---

## Purpose

This tool helps administrators improve visibility into:

- Classroom successes and instructional strengths
- Student support patterns across classrooms
- Instructional support opportunities
- Coaching conversation starters
- Follow-up and support planning needs

**This is NOT an evaluation system.** It is NOT an IEP tool, NOT a student advocacy platform, and NOT connected to any student information system.

---

## Features

- **Dashboard** — At-a-glance summary of walkthroughs, support signals, and follow-up counts
- **Setup** — Manage teachers, classrooms, and placeholder student IDs
- **Walkthrough Form** — Complete a structured classroom visit record in under 2 minutes
- **Reports** — Filter, review, and export walkthrough history
- **Storage Settings** — Understand data storage and the future Microsoft 365 roadmap

---

## How to Run

1. Download or clone this folder to your computer
2. Open `index.html` in any modern web browser (Chrome, Edge, Firefox, Safari)
3. No internet connection required
4. No installation, no server, no login

The application runs entirely in your browser. All data is stored locally using `localStorage`.

---

## How Data Is Stored

**Version 1 — Local Prototype**

All data is stored in your browser's `localStorage` under these keys:

| Key | Contents |
|-----|----------|
| `skookWalkthrough_teachers` | Teacher roster |
| `skookWalkthrough_classrooms` | Classroom list |
| `skookWalkthrough_students` | Placeholder student IDs |
| `skookWalkthrough_records` | Walkthrough records |
| `skookWalkthrough_settings` | Application settings |

Data persists across browser sessions but will be lost if browser storage is cleared. Always export a JSON backup before clearing browser data.

---

## Pilot Testing Guidance

### Before Your First Walkthrough

1. Open the app and go to **Setup**
2. Add at least 2–3 teachers
3. Add at least 2–3 classrooms and assign them to teachers
4. Optionally add 2–3 placeholder student IDs (use initials like "J.D." or codes like "Pilot-001")

### Testing the Walkthrough Flow

1. Click **Walkthrough** in the left nav
2. Complete the form — aim for under 2 minutes
3. Submit and verify the success confirmation
4. Record 3–5 practice walkthroughs

### Testing Reports

1. Go to **Reports**
2. Try filtering by teacher or date range
3. Review the Support Status summary and Follow-Up Opportunities
4. Export to CSV and open in Excel or Google Sheets

### Testing Data Export/Import

1. Go to **Setup → Data Tools**
2. Export a JSON backup
3. Clear data, then re-import the backup
4. Verify all records are restored

### What to Evaluate with the Supervisor

- Does the form capture what you observe during walkthroughs?
- Is the Support Needed language clear and actionable?
- Does the Classroom Status language feel appropriate (not punitive)?
- Are there missing support types in the Supports Observed list?
- Is the Report view useful for identifying patterns?
- Would weekly review of the dashboard be part of your workflow?

---

## Privacy Notice

This prototype is for testing purposes only. Do not enter:

- Real student names
- Student ID numbers
- IEP or service information
- Any protected student data (FERPA-covered information)

Use initials, anonymous codes, or placeholder labels only.

---

## Technical Notes

- Built with: HTML, CSS, vanilla JavaScript — no frameworks, no build tools
- Browser support: All modern browsers (Chrome 90+, Firefox 88+, Safari 14+, Edge 90+)
- Storage: Browser `localStorage` — local only, not synced
- File structure is self-contained — works offline

---

## Next Steps

See `MICROSOFT_365_PLAN.md` for the roadmap toward Microsoft 365 integration.
