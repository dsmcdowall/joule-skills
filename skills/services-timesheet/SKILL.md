---
name: services-timesheet
description: >-
  Tracks consulting activity hours against SAP CATS cost objects (Internal Orders and Sales Documents) and produces a weekly Excel timesheet based on Outlook Calendar and Email access via Joule Desktop. Staffing assignments are loaded from an ISP export file - no live API connection required. Use for: logging time, calendar/email scans, weekly timesheet generation, and importing staffing assignments. Trigger phrases: 'log hours', 'add time entry', 'scan my calendar', 'scan my emails', 'scan sent items', 'weekly scan', 'combined scan', 'timesheet', 'weekly report', 'services timesheet', 'log my time', 'generate report', 'what did I log', 'services hours', 'import assignments', 'import staffing', 'load assignments', 'import my staffing', 'update assignments'.
metadata:
  author: SAP Solution Architect
  version: 2.1.0
  tags: timesheet cats sap consulting time-tracking reporting
---

Log consulting activity hours against SAP CATS cost objects (Internal Orders and Sales Documents), auto-match calendar events and sent emails to staffing assignments, and produce a weekly Excel timesheet.

---

## Execution Model

All operations run entirely within the Joule sandbox - no external API calls required.

| Type | Examples | Who runs it |
|---|---|---|
| **File-only** (no network) | generate_reports.py, gen_assignments_yaml.py | Joule sandbox |

---

## Initial Setup

One-time setup. After this your assignments are populated and reports are ready to generate.

### Step 1 - Import your staffing assignments

Open the ISP staffing portal at the following URL:

**https://isp.hec.net.sap/sap/bc/webdynpro/sap/zui_staffinglist?&sap-language=EN#**

Navigate to the **Export Excel** icon on the far right of the screen and export the `.xlsx` file. This file contains your current active staffing assignments.

Once downloaded, tell Joule:

> *"Import my assignments from filename.xlsx"*

Joule runs **Workflow 7** to convert the export to `assets/assignments.yaml`. After import, open the file to add any extra keywords (abbreviations, short names, email domains) for better auto-matching.

**Note:** WBS element rows in the export are automatically skipped - only Internal Orders and Sales Documents are imported.

### Step 2 - Install report dependencies (once, before first timesheet)

Run the following in a terminal:

```
python.exe -m pip install openpyxl pyyaml
```

Only needed for Workflow 4 (Excel generation).

---

## Data Storage

**`services-log.json`** - confirmed time entries (saved to working directory):

```json
{
  "id": "20260101-acme-corp-001",
  "date": "2026-01-01",
  "day_of_week": "Thursday",
  "customer_name": "Acme Corp",
  "cost_object": "176027692/80",
  "cost_object_type": "SalesDoc",
  "hours": 2.0,
  "notes": "Project kick-off call",
  "source": "calendar",
  "calendar_event_id": null,
  "email_message_id": null,
  "created_at": "2026-01-01T09:00:00Z",
  "updated_at": "2026-01-01T09:00:00Z",
  "cats_posted_at": null
}
```

> **Email entries always have `"hours": null`** - duration cannot be reliably inferred from a sent item. The user fills in the hours in the Excel. Never store a non-null hours value for a `source: "email"` entry.

**`services-scan-excluded.json`** - items seen during scans but not logged (saved to working directory). Written by Workflows 6, 8, and 9. Used by Workflow 4 to populate the NOT MATCHED and CLASSIFIED PERSONAL audit sections of the Excel.

Each entry:

```json
{
  "date": "2026-01-02",
  "day_of_week": "Friday",
  "notes": "Event or email subject",
  "source": "calendar",
  "bucket": "classified_personal",
  "reason": "Personal"
}
```

`bucket` values: `"not_matched"` or `"classified_personal"`.

Common `reason` values:
- **Classified personal:** `"Personal"`, `"Cancelled event"`, `"All-day event"`, `"Out of hours"`, `"Personal block"`, `"Auto-reply (Accepted)"`, `"Auto-reply (Declined)"`, `"System notification"`
- **Not matched:** `"No assignment keyword match"`, `"Customer not in assignments"`

Create both files as `[]` if they do not exist before first write.

---

## Assignment List

Assignments are stored in `assets/assignments.yaml`. Each entry has:
- `cost_object`, `type` (Order, SalesDoc, WBS, or CostCentre), `customer`, `keywords`
- Optional: `description`, `project`, `start_date`, `end_date`, `responsible`

Load at the start of any workflow that needs to identify or display assignments.

The file is populated by **Workflow 7** (import from ISP staffing export Excel), or edit it manually.

**Reserved virtual assignment - SAP Internal:**
Add this entry manually if it is not already present:

```yaml
- cost_object: "INTERNAL"
  type: "CostCentre"
  customer: "SAP Internal"
  keywords: []
```

---

## Workflow 1: Create a New Entry (Manual)

**Triggered by:** "log hours", "add time entry", "log my time"

1. Date - YYYY-MM-DD, default today.
2. Hours - positive multiple of 0.5.
3. Assignment - Assignment Matching Logic; confirm or show menu.
4. Notes (optional).
5. Validate: warn if daily > 8h; check duplicates. Set `source: "manual"`. Append and write.

---

## Workflow 2: Update an Existing Entry

**Triggered by:** "update hours", "change my entry"

Find by fuzzy customer + date. Warn if outside current week. Apply changes, update `updated_at`, save.

---

## Workflow 3: Query Entries

**Triggered by:** "how many hours", "what did I log", "show entries"

Display as Markdown table with footer: **Total: X.0 hrs | [N] entries | [M] email stubs pending**

---

## Workflow 4: Generate Weekly Timesheet

**Triggered by:** "weekly report", "generate timesheet", "build report", "reproduce the excel", "regenerate the excel"

1. Determine week (default current Mon-Fri).
2. Filter `services-log.json` for the range.

### Step 2a - Re-match excluded items (always run before generating)

Before running the report script, re-run **Assignment Matching Logic** against all `not_matched` entries in `services-scan-excluded.json` for the target week:

1. Load `assets/assignments.yaml` (current keywords).
2. For each `not_matched` entry in the target date range, check its `notes` field against all assignment keywords.
3. If any entries now match (because `assignments.yaml` has been updated since the scan):
   - Present the newly-matched items to the user in a compact table: Date | Notes | Src | Customer | Cost Object
   - Ask: *"These [N] previously unmatched items now have an assignment. Log them? ('log all', 'log 1–N', or 'skip')"*
   - For confirmed items: append to `services-log.json` (hours from original duration if calendar, null if email), remove from `services-scan-excluded.json`.
4. If no new matches are found, proceed silently.

3. Run script:

```
python3 "<skill-disk-path>/scripts/generate_reports.py"
  --log      "services-log.json"
  --start    "YYYY-MM-DD"
  --end      "YYYY-MM-DD"
  --excel    "services-timesheet-YYYY-Wnn.xlsx"
  --excluded "services-scan-excluded.json"
```

**Always pass `--excluded`** - if `services-scan-excluded.json` does not exist yet, create it as `[]` first.

**Excel layout** - Customer/Cost Object column + one Notes+Hrs pair per weekday + Weekly Total:

- One slot row per activity per day; alternating pale/white fill
- Navy customer subtotal rows with SUM formulas per day and week
- Navy TOTAL HOURS row at the bottom
- Email stub rows show "fill hrs" prompt; totals recalculate automatically once filled

*Audit sections (below totals, display only - never affect logged hours):*
- **NOT MATCHED** (amber) - calendar events and emails where no assignment keyword matched. Shows Date, Activity/Notes, Source, Reason.
- **CLASSIFIED PERSONAL** (grey) - events and emails excluded by personal/automated filter rules. Shows Date, Activity/Notes, Source, Reason.

These sections are populated from `services-scan-excluded.json` filtered to the report week. They appear only if excluded data has been written for that week (Workflows 6, 8, or 9 must have run first).

4. Confirm output, then: *"Review the Excel and fill in any email stub hours."*

---

## Workflow 5: Email Reports

**Triggered by:** "email the report", "send me the report"

Find address from `list_emails(folder="sentitems", top=1)`. Send Excel to yourself.

---

## Workflow 6: Calendar Scan and Auto-Log

**Triggered by:** Explicit request only - "scan my calendar", "log from calendar"

1. Determine period.
2. `list_calendar_events` (limit 50). Route each event to a bucket using **Event Filter Rules**.
3. `get_calendar_event` for all non-personal events - single parallel batch.
4. Apply **Assignment Matching Logic** (title + attendees). Self-organised - title is primary signal.
5. Hours = duration divided by nearest 0.5h (min 0.5h). All events. Never 0.
6. Present using **Scan Results Format**.
7. After the user confirms which matched entries to log:
   - Write confirmed entries to `services-log.json` with `source: "calendar"`.
   - **Mandatory: write ALL items from Sections 2 (Not Matched) and 3 (Classified Personal) to `services-scan-excluded.json`** (append, dedup by date + notes + source). Use the reason values from the **Data Storage** section above. Do not skip this step even if the user moves straight to generating the Excel.

---

## Event Filter Rules

Route every calendar event or email to one of three buckets.

### Matched / Not matched (after assignment check)
- Events organised BY the user - confirmed by definition; use title as primary matching signal
- Recurring meetings (same title appears multiple times or title contains scheduling indicators)
- SAP-internal meetings - organizer is @sap.com AND all attendee emails are @sap.com AND no customer keyword match - route to INTERNAL / CostCentre

### Classified personal
- Title starts with Canceled:, Cancelled:, Declined:, Tentative:, Abgesagt:
- RSVP cross-reference: sent items contain "Declined: [event title]"
- All-day events
- No attendees other than organiser (personal reminders) - except self-organised work blocks
- Title keywords: "birthday", "anniversary", "holiday", "payroll", "townhall", "refresh", "labour code", "upside tracker"
- Email: subject starts with "Automatic reply", "Out of Office", "Delivery:", "Read:", "Accepted:", "Declined:"; or system notification patterns

### Limitation
Outlook's RSVP field is unavailable; title-prefix and sent-item cross-referencing are used as proxies.

---

## Scan Results Format

Every scan presents three sections. **Only Matched items are ever logged, stored in services-log.json, or included in the Excel totals.** Items from Sections 2 and 3 are written to `services-scan-excluded.json` and appear in the Excel audit sections.

### Section 1 - Matched (will be logged on confirmation)

| # | Src | Date | Event / Email Subject | Hrs | Customer | Cost Object | Type |
|---|-----|------|-----------------------|-----|----------|-------------|------|

Email rows always show dash for Hrs - duration is never derived from emails. The hours column is left blank (null) so the user can fill it in on the Excel.

*"Found [N] matched entries. Tell me which to log - 'log all', 'log 1-5', or 'skip 3'. Rescue items from sections below with 'include NM1 as [customer]' or 'include P3 as [customer]'."*

### Section 2 - Not matched (display only - not logged)

| # | Src | Date | Event / Email Subject | Why |
|---|-----|------|-----------------------|-----|

*These will not be logged. No customer assignment found. They will appear in the NOT MATCHED audit section of the Excel.*

### Section 3 - Classified personal (display only - not logged)

| # | Src | Date | Event / Email Subject | Reason |
|---|-----|------|-----------------------|--------|

*These will not be logged. They will appear in the CLASSIFIED PERSONAL audit section of the Excel.*

---

## Workflow 7: Import Assignments from File

**Triggered by:** "import assignments", "import my staffing", "load assignments", "update assignments", "import staffing file", "import my assignments"

Converts an ISP staffing export Excel file into `assets/assignments.yaml`. WBS element rows are automatically skipped - only Internal Orders and Sales Documents are imported.

### How to export from ISP

1. Open the ISP staffing portal: **https://isp.hec.net.sap/sap/bc/webdynpro/sap/zui_staffinglist?&sap-language=EN#**
2. Navigate to the **Export Excel** icon on the far right of the screen
3. Click it to export and save the `.xlsx` file
4. Tell Joule: *"Import my assignments from staffing-export.xlsx"*

### What Joule does

Run the import script:

```
python3 "<skill-disk-path>/scripts/gen_assignments_yaml.py"
  "<path-to-staffing-export>.xlsx"
  --output "<skill-disk-path>/assets/assignments.yaml"
```

On success, confirm using the output from the script, e.g.:
*"Imported N assignments (M WBS rows skipped) to assignments.yaml. Open the file to add extra keywords for better auto-matching."*

If the file cannot be parsed - report the error and ask the user to check the file format.

**After import:** The INTERNAL entry for SAP-internal time is not included in the ISP export. Add it manually if needed (see Assignment List section above).

---

## Workflow 8: Email Scan and Auto-Log

**Triggered by:** "scan my emails", "scan sent items", "log from emails"

Same structure as Workflow 6 but sources from `list_emails(folder="sentitems", top=50)`.
Filter and match using the same **Event Filter Rules**.

**Email hours are always null (stub)** - emails have no reliable duration. Never estimate or assign hours to email entries, regardless of thread length or content. They are logged as stubs with `hours: null` and appear as unfilled rows in the Excel for the user to complete manually.

After the user confirms which matched entries to log:
- Write confirmed entries to `services-log.json` with `source: "email"`.
- **Mandatory: write ALL items from Sections 2 (Not Matched) and 3 (Classified Personal) to `services-scan-excluded.json`** (append, dedup by date + notes + source, source: "email"). Do not skip this step.

---

## Workflow 9: Combined Weekly Scan

**Triggered by:** "weekly scan", "combined scan", "scan everything"

Run Workflow 6 and Workflow 8 sequentially. Present merged results in **Scan Results Format**.

After the user confirms matched entries:

1. **Log confirmed matches** to `services-log.json`.
2. **Write excluded items** - ALL items from Sections 2 and 3 of the scan results to `services-scan-excluded.json` (both calendar and email). This is a mandatory step - it must happen before Excel generation, even if the user did not explicitly ask for it.
3. **Offer to generate the weekly timesheet** (Workflow 4) - always pass `--excluded` so the NOT MATCHED and CLASSIFIED PERSONAL audit sections appear in the Excel.

---

## Assignment Matching Logic

When logging an event/email, find the best matching assignment:

1. Check keywords in `assignments.yaml` against event title + attendee domains.
2. If unique match - confirm and log.
3. If multiple matches - show a numbered menu, ask user to choose.
4. If no match - put in Not matched bucket.

**SAP Internal Detection:**
If organizer domain is @sap.com AND all attendees are @sap.com AND no customer keyword matches - assign to INTERNAL.

---

## Validation Rules

- Hours must be a positive multiple of 0.5.
- Daily total > 8h - warn (do not block).
- Duplicate check: same date + cost_object + source already in log - warn.
- Date must be in YYYY-MM-DD format.