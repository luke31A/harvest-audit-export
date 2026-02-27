# Project: Harvest API → Excel Export

## Goal
Python script (compiled to .exe) that queries all Harvest time entries for a user-prompted
date range and exports a fully formatted Excel report exposing detail not visible in the UI.

## Status
- [ ] In progress

## Target Users
- Commit financial operations team (non-technical)

## Key Requirements
- Prompt user for date range at runtime
- Pull all time entries via Harvest API (handle pagination)
- Export to formatted Excel (.xlsx)
- Compile to standalone .exe (no Python install required)

## Harvest API Notes
- Base URL: https://api.harvestapp.com/v2
- Auth: Personal Access Token + Account ID (in headers)
- Pagination: per_page max 100, use next_page / links.next
- Key endpoint: GET /time_entries

## Fields Exposed (all of the below)
- Employee ID/Name, Client, Project, Project Code, Task
- Work Date, Hours, Rounded Hours, Notes
- Billable, Billable Rate/Amount, Cost Rate/Amount
- Timer Started At, Started Time, Ended Time, Is Running
- Is Locked, Locked Reason, Is Closed, Is Billed, Budgeted
- Created At, Updated At (audit timestamps)
- Invoice ID/Number, External Ref
- Timesheet approval data if on Harvest Pro plan (Approval Status, Approver Name, timestamps, notes)

## Derived Audit Columns
- Submission Lag (Days): days from work date to Created At
- Edit Lag (Hours): hours from Created At to Updated At
- Was Edited: bool (edit lag > 3 min threshold)
- Late Submission: bool (submission lag > 7 days)
- Days to Approval: if approval data available

## Output Format
- Tab 1 "Audit Summary": banner, 8 KPI cards, by-employee table, by-project table, flagged entries table
- Tab 2 "Raw Data": full entry table, auto-filter, frozen header, flagged rows tinted orange
- Flagged = late submission OR edited after being locked

## Decisions Made
- Auth: prompt for PAT + Account ID at runtime (no config file)
- Timesheet approvals: attempted silently; skipped gracefully if not on Pro plan
- .exe output lands in dist/, Excel reports land in output/ next to .exe
- Build: PyInstaller --onefile via build.bat
