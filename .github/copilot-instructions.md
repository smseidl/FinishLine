# GitHub Copilot Instructions — FinishLine

## Project Overview

FinishLine is a **Google Apps Script (GAS)** project for middle school track & field meet management.
It lives entirely in a single file: `scripts/FinishLine.gs`, pasted into a Google Sheet via
**Extensions → Apps Script**.

The script generates two printable reports from spreadsheet data:
- **Lineup_View** — pre-meet athlete lineup with PRs
- **Event_Form_Printable** — post-meet event forms with results, PR/record highlights

---

## Language & Runtime

- **Google Apps Script** — JavaScript (ES2019) running server-side in Google's V8 runtime.
- No build step. No npm. No modules. Single `.gs` file pasted into the Apps Script editor.
- GAS has a **6-minute execution limit** per run. Batch API calls where possible.
- `SpreadsheetApp`, `Sheet`, `Range` are the primary APIs. No DOM, no fetch.

---

## File Structure

```
scripts/FinishLine.gs   ← the entire application; paste into Apps Script
instructions/           ← development notes and AI context
README.md               ← coach-facing setup guide
```

---

## Critical GAS Constraints — Never Violate These

### Triggers
- `onEdit(e)` (simple trigger) **cannot** call `getUi()`, `alert()`, or `toast()`.
- The installable trigger function is named **`onEditInstallable`** — must be registered manually
  in Extensions → Apps Script → Triggers. It does NOT copy when a spreadsheet is duplicated.
- Do not add a simple `onEdit()` — it would double-fire alongside the installable trigger.

### API Gaps
- `sheet.clearRowBreaks()` **does not exist** in GAS. Do not use it.
- `sheet.clear()` does **NOT** clear cell notes. Must call `range.clearNote()` explicitly.
- Page breaks cannot be set programmatically — GAS has no API for this.
- `SpreadsheetApp.flush()` is required to push cell writes to the browser mid-execution.

### Data Safety Contract — NEVER violate in any edit
- **Data tabs** (`Schedule`, `Data_Entry`, `Roster`, `School_Records`): only row 1 headers
  are written by code. Row 2+ user data is **never cleared**.
- **Generated tabs** (`Lineup_View`, `Event_Form_Printable`): `sheet.clear()` +
  `sheet.getRange(1,1,sheet.getMaxRows(),9).clearNote()` on every regeneration.
- **Home tab**: fully rebuilt by `fullInitialize()`. No user data stored here.

---

## Versioning

- Single source of truth: `const VERSION = "vX.Y"` at the top of `FinishLine.gs`.
- **Bump the version on every code change** — even minor ones.
- Tag stable releases in git: `git tag -a vX.Y -m "description"`.
- If a change requires the coach to run **"1. Build / Rebuild Entire System"**, say so explicitly.
- If a change breaks compatibility with the existing sheet structure, mark the commit message
  with `BREAKING: requires full rebuild`.

---

## Sheet / Tab Architecture

| Tab | Color | Purpose |
|-----|-------|---------|
| `Home` | Black | Control panel — dropdowns, buttons, status cell |
| `Schedule` | Cyan `#00bcd4` | Meet calendar (Meet #, Date, Name, Location, etc.) |
| `Data_Entry` | Cyan `#00bcd4` | Results — one row per athlete per event per meet |
| `Roster` | Cyan `#00bcd4` | Athlete list with Display Name and all PR columns |
| `School_Records` | Cyan `#00bcd4` | School records by Gender + Event |
| `Lineup_View` | Orange `#ff6d00` | Generated — pre-meet lineup report |
| `Event_Form_Printable` | Orange `#ff6d00` | Generated — post-meet event forms |

### Data_Entry Column Reference
| Col | Field |
|-----|-------|
| A | Meet # |
| B | Gender |
| C | Event |
| D | Athlete Name (must match Roster Display Name) |
| E | Relay Team ID |
| F | Result/Mark |
| G | Splits/Attempts (comma-separated) |
| H | Notes |
| I | Place |

---

## Key Functions

| Function | Description |
|----------|-------------|
| `onOpen()` | Builds the 🏁 FINISH LINE menu |
| `onEditInstallable(e)` | Handles Home tab checkbox buttons; writes status to A8:B8 |
| `fullInitialize()` | Builds/rebuilds all 6 tabs, headers, validations, tab colors |
| `generateLineupReport()` | Writes Lineup_View for a selected meet/gender |
| `generateEventFormReport()` | Writes Event_Form_Printable for a selected meet/gender |
| `renderStandardBlock()` | Renders a standard track event block |
| `renderRelayBlock()` | Renders a relay event block (grouped by team ID) |
| `renderSplitBlock()` | Renders 800M/1600M with lap sub-rows |
| `renderAttemptBlock()` | Renders field events with 3 attempt sub-rows |
| `applyReportLayout()` | Sets column widths and number formats on output sheets |
| `formatCellValue(val)` | Converts GAS cell values to display strings; handles Date objects |
| `isBetter(val, bench, event)` | Compares two performance values (field=bigger, track=smaller) |
| `isNoMark(val)` | Returns true for DNS/DNR/DQ/-/NH/NM/etc. — skip highlighting |
| `findPR(rData, name, event)` | Looks up PR from Roster; Display Name first, Athlete Name fallback |
| `checkPRSetup()` | Debug: finds unmatched names and missing Display Names |
| `checkMeetRoster()` | Debug: validates all Data_Entry names for a meet against Roster |

---

## Report Layout Logic

Reports use a **two-column layout** (left = cols A–B, right = cols D–E, col C = 20px gutter):

- Track events 1–5 → left column
- Track events 6–9 → right column
- At field event transition: sync both columns to `Math.max(curL, curR)` before continuing
- Field events alternate left/right

Minimum 6 data rows per event (padded with empty bordered rows).

---

## Highlight Color Key

| Color | Meaning |
|-------|---------|
| Yellow `#ffe599` | School record broken |
| Green `#b6d7a8` | PR or first-time result |
| Blue `#1155cc` text | Existing PR shown inline |
| Light blue `#cfe2f3` | No prior PR (Lineup only) |
| Amber `#bf9000` | School record header cell (Event Forms) |

No highlighting is applied when `isNoMark(res)` is true (DNS, DNR, DQ, -, NH, NM, etc.).

---

## PR Lookup Logic

`findPR()` matches athlete names case-insensitively:
1. **Display Name** (Roster col B) — primary key; the name used in Data_Entry
2. **Athlete Name** (Roster col A) — fallback if Display Name is blank

No fuzzy matching. If a PR lookup fails, the coach checks that Data_Entry col D exactly
matches the Roster Display Name.

---

## Event Lists

```javascript
SPLIT_EVTS   = ["800 M Run", "1600 M Run"]           // uses lap sub-rows
ATTEMPT_EVTS = ["Shot Put", "Discus", "Long Jump"]    // uses 3 attempt sub-rows
RELAY_EVTS   = ["400 M Relay", "800 M Relay", ...]    // grouped by Relay Team ID
```

---

## Home Tab Layout

| Row | Content |
|-----|---------|
| 1 | Header banner (black, "🏁 FINISH LINE vX.Y") |
| 3 | Meet # dropdown (B3, yellow background) |
| 4 | Gender dropdown (B4, yellow background) |
| 5 | Delay hint text (italic, small) |
| 6 | Lineup button (green, checkbox in A6) |
| 7 | Event Forms button (blue, checkbox in A7) |
| 8 | Status cell A8:B8 (merged; written by onEditInstallable) |
| 9 | Future Features stub |
| 11–22 | Setup Checklist |

---

## What "Rebuild Required" Means

Running **"1. Build / Rebuild Entire System"** (`fullInitialize`) is required when:
- Tab colors or column widths change
- Home tab layout changes
- New data validation rules are added
- New tabs are introduced

It is **not** required when only report generation logic changes (`generateLineupReport`,
`generateEventFormReport`, render helpers, or utility functions).
