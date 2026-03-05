# FINISH LINE v1.0 — Project Context for Claude Agent

## What This Is
A Google Apps Script (`.gs`) file that manages a **middle school track & field program**.
It builds and manages a Google Sheets workbook with roster, scheduling, data entry, and
printable meet reports. The coach pastes the script into Google Apps Script editor
(Extensions → Apps Script) and runs it from a custom menu inside the sheet.

---

## The Google Sheet Structure

| Tab Name | Purpose |
|---|---|
| `Schedule` | One row per meet. Columns: Meet #, Date, Time, Type, Location, Meet Name, Boys Standing, Girls Standing |

> **Note on Schedule "Type" column:** Type has a dropdown (Dual Meet / Invitational / Championship / Scrimmage) but is **reference-only** — no report logic reads `meetRow[3]`. It's there for the coach's planning view only.
| `Data_Entry` | One row per athlete per event per meet. See column reference below. |
| `Roster` | One row per athlete. Columns: Athlete Name, **Display Name**, Gender, Grade, School, Email, Parent Cell, then one column per event (PR storage) |

> **Display Name (Roster col B):** This is the PRIMARY key for PR lookups. Enter the same short name / nickname here that the coach uses in Data_Entry col D. `findPR()` matches on Display Name first, then falls back to Athlete Name. If a PR isn’t showing up, verify the Display Name exactly matches the Data_Entry name.
| `School_Records` | Gender, Event, Athlete, Record, Year, Notes |

> **School_Records columns:** Only `Record` (col D) is read by report logic for highlighting. `Year` and `Notes` are history/reference only — not used in any report.
| `Lineup_View` | Printable lineup sheet. Has dropdowns for Meet # and Gender in B2/B3. |
| `Event_Form_Printable` | Printable results/event form. Same dropdowns. |

### Data_Entry Column Reference
| Col | Field | Notes |
|---|---|---|
| A | Meet # | Validated against Schedule tab |
| B | Gender | Boys / Girls |
| C | Event | Validated against FULL_EVT list |
| D | Athlete Name | |
| E | Relay Team ID | Relays only — "1", "2", etc. |
| F | Result/Mark | Final time or distance |
| G | Splits/Attempts | Comma-separated. See rules below. |
| H | Notes | Free text |
| I | Place | Numeric place finish |

### Column G Rules by Event Type
- **Relay events** — one row per athlete, G = individual leg split
- **800M Run** — single row per athlete, G = `"lap1, lap2"`
- **1600M Run** — single row per athlete, G = `"lap1, lap2, lap3, lap4"`
- **Shot Put / Discus / Long Jump** — single row per athlete, G = `"att1, att2, att3"`
- **All other events** — G is unused or free notes

---

## Event Lists

### FULL_EVT (used for Data_Entry dropdown validation)
```
100M Hurdles, 100 M Dash, 1600 M Run, 800 M Relay,
400 M Dash, 400 M Relay, 800 M Run, 200 M Dash, 1600 M Relay,
High Jump, Discus, Shot Put, Long Jump,
65M Hurdles, Hurdle Shuttle, Sprint Relay, Special Relay, Distance Relay
```

### PRINT_EVT (order used in both printable reports)
```
Track (in meet order):
  100M Hurdles, 100 M Dash, 1600 M Run, 800 M Relay,
  400 M Dash, 400 M Relay, 800 M Run, 200 M Dash, 1600 M Relay

Field (at bottom):
  High Jump, Discus, Shot Put, Long Jump
```

---

## How the Printable Reports Work

Both `Lineup_View` and `Event_Form_Printable` use the same two-column layout:
- Events alternate left (col A-B) and right (col D-E) as they go down the page
- Col C is a narrow spacer (20px)
- Each event block starts with a dark header row, then athlete rows below
- Minimum 6 rows per event block (padded with empty bordered rows)
- Report starts at row 7; rows 1-4 are reserved for dropdowns/title

### Lineup_View
Shows athlete name + their PR for each event. Relay events show team groupings.

### Event_Form_Printable
Shows athlete name + PR, result, place, and highlights:
- 🟢 Green cell = Personal Record (or first time)
- 🟡 Yellow cell = School Record
- Relay events: grouped by team ID, shows splits per leg + team total
- 800M/1600M: athlete row + indented lap rows
- Shot Put/Discus/Long Jump: athlete row + 3 indented attempt rows

---

## Event Type Classification (internal logic)

```javascript
RELAY_EVTS   = events containing "Relay" or in explicit relay list
SPLIT_EVTS   = ["800 M Run", "1600 M Run"]
ATTEMPT_EVTS = ["Shot Put", "Discus", "Long Jump"]
// Everything else = standard track event
```

---

## Key Functions

| Function | What it does |
|---|---|
| `onOpen()` | Builds the 🏁 FINISH LINE menu |
| `fullInitialize()` | Builds/rebuilds all tabs, headers, validation, filters |
| `generateLineupReport()` | Writes printable lineup to Lineup_View |
| `generateEventFormReport()` | Writes printable event form to Event_Form_Printable |
| `renderStandardBlock()` | Renders standard track event rows |
| `renderRelayBlock()` | Renders relay team grouped rows |
| `renderSplitBlock()` | Renders 800M/1600M with lap sub-rows |
| `renderAttemptBlock()` | Renders field events with attempt sub-rows |
| `findPR(rData, name, event)` | Looks up athlete PR from Roster tab. Matches Display Name (col B) first, then Athlete Name (col A) exact match. Case-insensitive. No fuzzy matching — if PR missing, check Display Name = Data_Entry name. |
| `isBetter(val, bench, event)` | Compares two marks — field = bigger is better, track = smaller is better |
| `findPR(rData, name, event)` | Looks up athlete PR from Roster tab |
| `clearReportSpace(sheet)` | Clears rows 5+ including notes and formatting |
| `applyReportLayout(sheet)` | Sets column widths for printable tabs |

---

## History / What Was Fixed Getting Here

- v1.1: Added Type dropdown to Schedule (reference only); added instruction row 4 to printable tabs; data safety contract comment
- v1.5: Added Display Name column (Roster col B) as explicit PR lookup key; removed fuzzy name matching from `findPR()`; Roster now has 7 fixed columns before event PRs
- v1.4: `formatCellValue()` for Date object fix; full-row highlighting for PR/School Record; Lineup no-PR blue highlight; `-` treated as no PR

This script went through ~38 versions with another AI before being restarted as FINISH LINE v1.0.
The key problems that were resolved:

- `fullInitialize()` was incomplete in earlier versions — didn't build all tabs/headers
- Meet name was being pulled from wrong column (Type vs Meet Name)
- PR green highlight was accidentally dropped
- `clearNote()` was missing from report clearing (old notes would persist on reruns)
- Relay team grouping was missing from Lineup report
- Event order didn't match actual meet schedule
- 800M/1600M splits and field event attempts had no rendering logic
- `TypeError: Cannot read properties of undefined` crash when meet row not found — now validated

---

## Known Remaining Work / Ideas to Consider

- [ ] High Jump rendering — currently uses standard block. May need custom logic
      (height-by-height attempts with pass/fail, not a distance mark)
- [ ] Email/text notification to parents (roster has Email + Parent Cell columns ready)
- [ ] Seed time import for pre-meet planning
- [ ] Season PR auto-update — write best result back to Roster after each meet
- [ ] Team scoring summary tab
- [ ] 65M Hurdles and other non-standard events may need review
- [ ] Schedule "Type" column is display-only; could wire it into report title or filter logic later if desired

---

## How to Install / Test

1. Open the Google Sheet
2. Extensions → Apps Script
3. Paste entire `.gs` file contents, replacing everything
4. Save (Ctrl+S)
5. Reload the Google Sheet
6. Use the **🏁 FINISH LINE** menu → "Build / Rebuild Entire System" first
7. Then add data and test the generators

> ⚠️ Running `fullInitialize()` on an existing sheet will re-set headers and validation
> but will NOT delete existing data rows.

---

## Data Safety Contract (DO NOT BREAK)

- **Data tabs** (Schedule, Data_Entry, Roster, School_Records): `fullInitialize()` only writes row 1 headers. **Row 2+ is never cleared.** Do not add `sh.clear()` or `sh.deleteRows()` to these tabs.
- **Generated tabs** (Lineup_View, Event_Form_Printable): `sh.clear()` is called on full rebuild — these are fully regenerated, no user data lives here.
- **`clearReportSpace()`** only clears row 5+ on printable tabs. Rows 1–4 (version label, dropdowns, instruction text) are preserved between report regenerations.
- Row 4 of both printable tabs holds user-facing instructions written during `fullInitialize()` and intentionally left intact by `clearReportSpace()`.

---

## Coding Conventions to Maintain

- All sheet access goes through `getOrCreateSheet()` — never `ss.insertSheet()` directly
- Render functions return the updated `row` number
- Event type is always checked via `isRelayEvent()`, `isSplitEvent()`, `isAttemptEvent()` helpers
- Place value from col I should be checked as `(place !== "" && place !== null && place !== undefined)`
  before appending to result string — it can be 0 (legitimate value) so don't just do `if (place)`
- `isBetter()` handles feet/inches (`18'3"`), MM:SS time, and decimal formats
