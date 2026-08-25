# PACE Administrator Integration

## Data flow

`PACE Room Tracker → IEP_Pace_Visits → MAC-Walkthrough administrator dashboard / PACE Log`

`IEP_Pace_Visits` is the sole source of truth for PACE analytics. The MAC
reader uses SharePoint display names resolved through the existing Graph schema
mapper, follows every Graph result page, and does not create, patch, delete, or
duplicate PACE visits.

## Live schema currently confirmed

- `Student` (text)
- `Date` (dateTime)
- `Time In` / `Time Out` (text)
- `Duration` (number)
- `Reason` (choice)
- `Intervention Used` (choice)
- `Notes` (text)
- `Staff Member` (personOrGroup; tracker cannot currently persist its selected specialist)
- `Return Status` (choice; standalone tracker does not currently populate it)
- SharePoint system `Created` timestamp

Fields not confirmed in the live list are `PACE Room`, `SCM Used`, `Entry ID`,
`Submitted By`, `Submitted At`, and `Teacher Came From`. The admin UI detects
their presence from the schema. Missing filters/cards are hidden; missing values
are not converted to zero or `No`.

## Alert architecture

MAC-Walkthrough currently detects two read-only dashboard flags from fields that
already exist:

- completed visit duration at least 45 minutes;
- the same student with at least three completed visits on one day.

These helpers only summarize data already loaded for the dashboard and send no
notifications.

Production notification delivery should be implemented in Microsoft Power
Automate with an **When an item is created** trigger on `IEP_Pace_Visits`:

- immediate SCM alert (after an `SCM Used` Yes/No column exists);
- extended visit alert from `Duration >= 45`;
- same-day and rolling five-school-day repeat patterns;
- repeated `Needs a Break` patterns.

Power Automate is the correct execution point because it runs server-side when a
record is created, is not dependent on an administrator leaving a browser open,
and can use Microsoft 365 identities and delivery controls. A future admin rules
screen should configure a separate rules/settings list; it should not send mail
from browser JavaScript.

## Manual schema decisions

If these analytics are required, add columns manually—never from app code—and
then confirm their display/internal mappings through the existing schema tools:

- `Teacher Came From`: Single line of text.
- `PACE Room`: Choice or Single line of text, with values aligned to the tracker.
- `SCM Used`: Yes/No.
- Behavior Specialist: either make `Staff Member` writable by implementing a
  tested SharePoint user lookup/Person-field payload, or add a governed text
  field after deciding which identity model is authoritative.

Do not configure SCM or room-dependent flows until the corresponding values are
actually persisted by the standalone tracker.
