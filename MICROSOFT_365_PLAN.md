# Microsoft 365 Integration Plan — MAC Walkthrough

Future phases of the MAC Walkthrough pilot. This document describes planned — not implemented — capabilities.

---

## Phase 1 — Local Prototype (Current)

**Status: Active**

- Single HTML file application
- Data stored in browser `localStorage`
- No authentication, no server, no cloud
- Suitable for: small-scale pilot testing with a single administrator and supervisor
- Data export via JSON and CSV
- No student information should be entered — use anonymous IDs only

**Exit criteria:** Supervisor confirms the walkthrough workflow is useful and the data captured is actionable. At least 20 pilot walkthroughs completed. Data export tested successfully.

---

## Phase 2 — Microsoft 365 Connected Pilot

**Status: Planning**

**Goal:** Move from a single-browser prototype to a multi-user, centrally stored system backed by existing IU29 Microsoft 365 infrastructure.

### Architecture

```
Browser (MAC Walkthrough App)
         ↓
Microsoft Entra ID (Authentication)
         ↓
Power Automate (Workflow / API Layer)
         ↓
Microsoft Lists / SharePoint (Storage)
```

### Components

**Authentication**
- Microsoft Entra ID (formerly Azure AD)
- Single Sign-On using existing IU29 accounts
- Role-based access: administrator view vs. read-only view

**Storage**
- Microsoft Lists for structured walkthrough records
- SharePoint document library for exports and attachments
- Lists schema mirrors the local data model (date, teacher, classroom, supports, status, etc.)

**Workflow Layer**
- Power Automate flows to handle form submissions
- Optional: automated email digest for Urgent or Immediate Follow-Up flags
- Optional: weekly summary report delivered to supervisor inbox

**Application Changes Required**
- Replace `localStorage` reads/writes with Microsoft Graph API calls
- Add authentication flow (MSAL.js)
- Handle offline/sync scenarios gracefully
- Update the storage namespace for production use

### Data Migration

Walkthrough records exported as JSON from Phase 1 can be imported into SharePoint Lists. A one-time migration script would parse the JSON and create list items.

### Prerequisites Before Phase 2

- [ ] Review with IU29 technology leadership
- [ ] Confirm Microsoft 365 licensing covers Lists and Power Automate
- [ ] Define data governance policy for walkthrough records
- [ ] Confirm FERPA compliance approach before any student-linked data is entered
- [ ] Assign SharePoint site and permissions structure

---

## Phase 3 — Expanded Support Ecosystem

**Status: Future Consideration**

This phase would only be considered after Phase 2 is stable, governance is in place, and the pilot has been validated across multiple users.

Potential additions (to be evaluated by IU29 leadership):

- **Multi-administrator support** — Multiple building-level admins recording walkthroughs, with aggregated views for supervisor
- **Teacher visibility** (read-only) — Optional sharing of support signals with the teacher, positioned as coaching notes not evaluation
- **Pattern analytics** — Support trends over time across classrooms, grade levels, or building
- **Integration with scheduling systems** — Pre-populate walkthrough dates from calendar
- **Mobile-optimized interface** — Tablet/phone friendly for on-the-go recording
- **Notification workflows** — Automated follow-up reminders for flagged walkthroughs

### What Phase 3 Explicitly Does NOT Include

- Formal teacher evaluation or performance scoring
- IEP management or SPED documentation
- Student referral workflows
- AI-generated recommendations or analysis
- Integration with state or federal reporting systems without explicit legal review

---

## Technology Review Checklist

Before any Microsoft 365 integration, work with IU29 technology leadership to confirm:

1. **Data classification** — Is walkthrough data considered directory information, confidential, or protected?
2. **FERPA compliance** — What student-linked fields (if any) are permissible? What anonymization is required?
3. **Access control** — Who should have read/write/admin access to walkthrough records?
4. **Data retention** — How long should records be kept? What is the deletion policy?
5. **Audit logging** — Is there a requirement to log who accessed or modified records?
6. **Vendor review** — Does IU29 IT need to approve any new third-party integrations?

---

*This document is for planning purposes only. No Microsoft 365 integration exists in the current prototype.*
