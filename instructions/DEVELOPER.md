# FinishLine — Developer Reference

> This document is for developers maintaining or extending `scripts/FinishLine.gs`.
> For coach-facing setup and usage, see [README.md](../README.md).

---

## Overview

FinishLine is a single-file **Google Apps Script** application pasted into a Google Sheet.
There is no build step, no npm, and no modules. The entire application lives in
`scripts/FinishLine.gs`.

- **Runtime:** Google Apps Script (GAS), JavaScript ES2019, V8 engine
- **Execution limit:** 6 minutes per run — batch API calls where possible
- **Deployment:** Paste into Extensions → Apps Script in the target Google Sheet

---

## File Structure

```
FinishLine/
├── scripts/
│   └── FinishLine.gs       ← entire application; paste into Apps Script
├── instructions/
│   ├── DEVELOPER.md        ← this file
│   └── FINISHLINE_CONTEXT.md
└── README.md               ← coach-facing setup guide
```

---

## Versioning

- Single constant: `const VERSION = "vX.Y"` at the top of `FinishLine.gs`
- **Bump the version on every code change** — even minor ones
- Tag stable releases: `git tag -a vX.Y -m "description"`
- If a change requires the coach to run "1. Build / Rebuild Entire System", say so explicitly
- Mark breaking commits with `BREAKING: requires full rebuild`

---

## GAS Constraints — Never Violate

| Constraint | Detail |
|-----------|--------|
| `sheet.clearRowBreaks()` | **Does not exist** in GAS — do not call it |
| `sheet.clear()` | Does **NOT** clear cell notes — call `range.clearNote()` separately |
| Page breaks | Cannot be set programmatically — GAS has no API for this |
| `getUi()` in simple triggers | Cannot be used in `onEdit(e)` — use the installable trigger |
| `SpreadsheetApp.flush()` | Required to push mid-execution cell writes to the browser |
| 6-minute limit | Batch all API calls; avoid row-by-row `setValue()` in loops |

### Trigger Architecture

- `onEditInstallable` is an **installable trigger** — must be registered manually via
  Extensions → Apps Script → Triggers → `onEditInstallable` / On edit.
  It does **not** copy when a spreadsheet is duplicated.
- There is intentionally **no** simple `onEdit()` — having both would double-fire.

---

## Constants

| Constant | Value | Purpose |
|----------|-------|---------|
| `VERSION` | `"v2.114"` | Displayed in UI; bump on every change |
| `HOME_STATUS_CELL` | `"A11:B11"` | Merged status cell on the Home tab |
| `PDF_FOLDER_PROP_KEY` | `"FINISHLINE_PDF_FOLDER_ID"` | Script property key for PDF Drive folder (deprecated) |
| `CONF_SEED_MODE_PROP_KEY` | `"FINISHLINE_CONF_SEED_MODE"` | Script property key for conference seeding mode |
| `CONF_SEED_MODE_DEFAULT` | `"championship"` | Default seeding mode if not set |
| `DRIVE_WRITE_SCOPE` | Drive OAuth scope URL | Used by deprecated PDF helpers |
| `ENABLE_PDF_EXPORT` | `false` | Feature flag; when `false`, PDF menu items are hidden |
| `PAGE_HEIGHT` | `45` | Rows per printed page (Event Forms layout hint) |

### Event List Constants

| Constant | Contents |
|----------|---------|
| `FULL_EVT` | All 18 recognized events including specialty relays |
| `PRINT_EVT` | 13 standard events used for printed reports |
| `RELAY_MEET_EVT` | Events for relay-style meets |
| `PENTATHLON_EVT` | Events for 8th Grade Pentathlon meets |
| `SPLIT_EVTS` | `["800 M Run", "1600 M Run"]` — render with lap sub-rows |
| `ATTEMPT_EVTS` | `["Shot Put", "Discus", "Long Jump"]` — render with 3 attempt sub-rows |
| `RELAY_EVTS` | All relay variants — grouped by Relay Team ID in output |

---

## Sheet / Tab Architecture

| Tab | Color | Type | Notes |
|-----|-------|------|-------|
| `Home` | Black | Control | Dropdowns, buttons, status cell |
| `Schedule` | Cyan `#00bcd4` | Data | Meet calendar; col D = Type drives event lists |
| `Data_Entry` | Cyan `#00bcd4` | Data | Results; one row per athlete per event per meet |
| `Roster` | Cyan `#00bcd4` | Data | Athletes with Display Name and all PR columns |
| `School_Records` | Cyan `#00bcd4` | Data | Records by Gender + Event |
| `Historical_PRs` | Cyan `#00bcd4` | Data | Past-year PRs; same event columns as Roster |
| `Lineup_View` | Orange `#ff6d00` | Generated | Pre-meet lineup report |
| `Event_Form_Printable` | Orange `#ff6d00` | Generated | Post-meet event forms |
| `Top_Marks` | Purple `#674ea7` | Generated | Year-to-date top performers by event |
| `Athlete_Recaps` | Purple `#674ea7` | Generated | Year-end per-athlete recap with historical PRs |
| `Filtered_Results` | — | Generated | Ad-hoc grade/event filtered result export |
| `Placers_Export` | — | Generated | Placed results formatted for email/AI use |

### Data Safety Contract

**Never violate this in any edit:**

- **Data tabs** (`Schedule`, `Data_Entry`, `Roster`, `School_Records`, `Historical_PRs`):
  Only row 1 headers are written by code. Row 2+ user data is **never** cleared.
- **Generated tabs** (`Lineup_View`, `Event_Form_Printable`):
  `sheet.clear()` + `range.clearNote()` on every regeneration — fully rebuilt.
- **Home tab**: Fully rebuilt by `fullInitialize()`. No user data stored here.

### Data_Entry Column Reference

| Col | Index | Field |
|-----|-------|-------|
| A | 0 | Meet # |
| B | 1 | Gender |
| C | 2 | Event |
| D | 3 | Athlete Name (must match Roster Display Name) |
| E | 4 | Relay Team ID (relays only) |
| F | 5 | Result/Mark |
| G | 6 | Splits/Attempts (comma-separated) |
| H | 7 | Notes |
| I | 8 | Place |

### Home Tab Layout

| Row | Content |
|-----|---------|
| 1 | Black header banner "🏁 FINISH LINE vX.Y" |
| 3 | Meet # dropdown (B3, yellow) |
| 4 | Gender dropdown (B4, yellow) |
| 5 | Top N field (B5, yellow) — used by Top Marks report |
| 6 | Hint text: "After clicking, wait 5–15 sec…" |
| 7 | Lineup button (green, checkbox in A7) |
| 8 | Event Forms button (blue, checkbox in A8) |
| 9 | Update PRs button (amber, checkbox in A9) |
| 10 | Athlete Recaps button (purple, checkbox in A10) |
| 11 | Status cell `A11:B11` (`HOME_STATUS_CELL`) |
| 12 | Future Features stub |
| 14–25 | Setup Checklist |

---

## Function Reference

### Triggers & Menu

---

#### `onOpen()`
**Line:** 81

Builds the `🏁 FINISH LINE` custom menu in the Google Sheet UI.
Menu items call the major report and utility functions by string name.
PDF export items are gated behind `ENABLE_PDF_EXPORT` — when `false` (current), those items are hidden.

---

#### `onEditInstallable(e)`
**Line:** 119

Handles checkbox button presses on the Home tab. Must be registered as an
**installable trigger** (not a simple trigger) so it can call `getUi()` and `flush()`.

- Listens for `true` values set in column A of the Home tab
- Resets the checkbox to `false` immediately (button-like behavior)
- Rows 7–10 map to: Lineup, Event Forms, Update PRs, Athlete Recaps
- Writes yellow "generating" status, calls the report function, then writes green "done" status
- Requires Meet # (B3) and Gender (B4) for rows 7–9; row 10 runs regardless

---

### System Builder

---

#### `fullInitialize()`
**Line:** 177

Builds or rebuilds the entire spreadsheet system. Safe to re-run at any time
(data tabs are never cleared — only headers are written).

Creates/updates all tabs with correct headers, column widths, data validation rules,
tab colors, and frozen rows/columns. Generates an installable-trigger reminder alert on completion.

**Tabs created:** `Home`, `Schedule`, `Data_Entry`, `Roster`, `School_Records`,
`Historical_PRs`, `Lineup_View`, `Event_Form_Printable`, `Top_Marks`, `Athlete_Recaps`

**Required after:** tab color changes, Home layout changes, new data validation rules,
new tab additions.
**Not required for:** report generation logic changes.

---

### Report Generators

---

#### `generateLineupReport()`
**Line:** 443

Generates the `Lineup_View` tab. Uses Meet # (B3) and Gender (B4) from Home tab.

Writes three sections:

1. **By-Event** — two-column layout (cols A–B left, D–E right). First ½ track events on left,
   second ½ on right; field events sync both columns then alternate. Each athlete row shows
   name + PR. No-PR cells highlighted light blue `#cfe2f3`. Athletes with PRs shown in blue text.

2. **By-Athlete** — alphabetical list of each athlete and all their entered events for the meet.
   Includes relay leg numbers and High Jump/Long Jump starting distances where available.

3. **Conference Lineup** — event-by-event list formatted for conference submission.
   Uses short display names. If seeding is enabled (see `shouldIncludeConferenceSeedEstimates`),
   shows seed estimates in parentheses and seed basis notes in col B (italic grey).
   Individual track events are sorted fastest-first.

**Validates before generating:**
- All Data_Entry athlete names exist in Roster (aborts with unmatched list if not)
- Athletes are not over the 4-event limit (prompts to continue or cancel)

**Returns:** `true` on success, `false` if validation fails or user cancels.

---

#### `generateEventFormReport()`
**Line:** 894

Generates the `Event_Form_Printable` tab. Uses Meet # (B3) and Gender (B4) from Home.

Renders one event block per event. Block type is selected by event name:
- `renderSplitBlock()` — 800 M Run, 1600 M Run (lap sub-rows)
- `renderAttemptBlock()` — Shot Put, Discus, Long Jump (3 attempt sub-rows)
- `renderRelayBlock()` — standard relays (grouped by Team ID)
- `renderRelayFieldStandardBlock()` / `renderRelayFieldAttemptBlock()` — specialty relay formats
- `renderStandardBlock()` — all other events

Uses `PAGE_HEIGHT` to estimate page boundaries and leaves a blank row between pages
(manual page breaks must still be set in File → Print).

**Pentathlon-specific behavior:**
- Suppresses the standard `TEAM PLACE` / `TEAM POINTS` row
- Suppresses `At-Meet Additions` lines
- Renders Shot Put and Long Jump with 2 attempts instead of 3
- Adds a standings table at the bottom of the sheet starting in column A
- Reads Pentathlon `place/points` from `Data_Entry` col I (`Place`), not from the Home gender selection
- Uses each athlete's roster gender for labels like `10th Place (Girls)`

**Highlight colors written:**
- Yellow `#ffe599` — school record broken
- Green `#b6d7a8` — PR or first-time result
- Blue `#1155cc` text — existing PR shown inline
- Amber `#bf9000` — school record header cell

---

#### `generateTopMarks()`
**Line:** 3881

Generates the `Top_Marks` tab. No meet/gender selection needed.

Scans all Data_Entry results, groups by athlete, keeps each athlete's personal best,
sorts by performance (track: fastest first; field: farthest first), and writes the
top N per event (N from Home B5, default 7). Shows Boys and Girls in parallel columns.
Skips relay events. Skips `isNoMark()` values.

---

#### `generateAllAthleteRecaps()`
**Line:** 4025

Generates the `Athlete_Recaps` tab. No meet/gender selection needed.

Writes one section per athlete from the Roster with:
- Meet-by-meet result table for the current season
- Year-over-year PR comparison table using `Historical_PRs` data

Inserts extra rows if the sheet would overflow (estimates ~50 rows per athlete).
Aborts with a warning if approaching the 9,000-row soft limit.

> **Note:** The school name ("Our Lady of Lourdes Track") and current year ("2026")
> are currently hardcoded in this function. Update before each season.

---

#### `generateFilteredResults()`
**Line:** 2661

Generates the `Filtered_Results` tab via interactive prompts.

Prompts for: grade(s), optional event filter, sort mode.
Filters Data_Entry rows matching the grade(s) via Roster Grade column.
Supports flexible grade input (e.g., "5", "5,6", "6th") and event input
(event numbers, partial names, full names).

**Sort modes (1–5):**
1. Gender → Meet → Event → Athlete (default)
2. Gender → Event → Best Performance → Athlete
3. Gender → Athlete → Event → Meet
4. Gender → Best Performance → Athlete (all events mixed)
5. Gender → Athlete → Best Performance

Excludes no-mark results (DNS/DNR/DQ/etc.). Warns if any Data_Entry names
have no Roster match.

---

#### `generatePlacersExport()`
**Line:** 2955

Exports placed results for a meet into the `Placers_Export` tab, formatted
for easy copy/paste into emails or AI prompts.

Prompts for: Meet #, Gender filter (Girls/Boys/Combined), max place to include (default 6).

Also runs **sanity checks** — flags inconsistencies where a faster time has a worse
or missing place than a slower time. Warnings are written in col I of the output.

Includes school record status in output. Suppresses relay team times in favor of
individual leg splits where available.

---

#### `buildTimeTrialRelayList()`
**Line:** 3485

Appends 100 M Dash rows to `Data_Entry` for athletes who are **not** entered in
a specified source meet (the "main" limited-entry meet).

Prompts for: source meet # to exclude, target time-trial meet #, gender.
Confirms before writing. Generates stub Data_Entry rows (no result) with an
auto-generated note. Useful for seeding athletes into 100 M time trial heats
alongside a relay meet.

---

#### `generateYearEndSummary()`

Generates a `Year_End_Summary` tab for end-of-season reporting.

Output includes:
- Total athletes, meets, and events with results
- Season PR improvement counts from `Data_Entry`
- School record break count
- Detail table for school records: gender, event, athlete(s), mark/time, meet #, meet name, and prior record context

Notes:
- PR counts are based on progression of entered season results.
- School record detection compares season performances against current `School_Records` baseline.

---

#### `createNextSeasonFile()`

Creates a new spreadsheet copy for the next season and performs a lightweight reset.

Actions:
- Prompts for new file name
- Copies the current spreadsheet file via Drive
- Clears `Data_Entry` rows 2+
- Clears generated output tabs (`Lineup_View`, `Event_Form_Printable`, `Top_Marks`, `Athlete_Recaps`, `Filtered_Results`, `Placers_Export`, `Year_End_Summary`)
- Resets Home tab selections (`B3`, `B4`, `B5`) and status cell

Displays the new file URL and reminder to re-create the installable trigger (`onEditInstallable`) because triggers do not copy.

---

### Event Form Render Helpers

All render helpers write directly to a `Sheet` object and return the next available row.

---

#### `renderStandardBlock(sheet, aths, row, col, rosterData, schoolRec, ev, recordsData, selectedGender, meetType)`
**Line:** 1536

Renders a standard event block (one result row per athlete).
Writes athlete name in col `col`, result in `col+1`, PR note in `col+2`.
Applies PR/record highlight colors. Pads to minimum 6 data rows.

---

#### `renderRelayBlock(sheet, aths, row, col, rosterData, schoolRec, ev)`
**Line:** 1591

Renders a relay event block. Groups athletes by Relay Team ID.
Each team has a header row followed by individual leg rows showing
athlete name, leg split, and PR.

---

#### `renderRelayFieldStandardBlock(sheet, aths, row, col, rosterData, schoolRec, ev)`
**Line:** 1647

Variant of `renderRelayBlock` for specialty relay field events (e.g., relay Shot Put
or relay Long Jump) where each athlete has a single distance result.

---

#### `renderRelayFieldAttemptBlock(sheet, aths, row, col, rosterData, schoolRec, ev)`
**Line:** 1699

Variant for relay field events where each athlete has 3 attempts.
Renders 3 sub-rows per athlete within the relay team group.

---

#### `renderSplitBlock(sheet, aths, row, col, rosterData, schoolRec, ev, recordsData, selectedGender, meetType)`
**Line:** 1762

Renders 800 M Run or 1600 M Run blocks. Writes the athlete name and final result
on the main row, then splits (from col G, comma-separated) on sub-rows below.
Highlights PRs and school records.

---

#### `renderAttemptBlock(sheet, aths, row, col, rosterData, schoolRec, ev, recordsData, selectedGender, meetType)`
**Line:** 1819

Renders Shot Put, Discus, or Long Jump blocks. Writes athlete name and best mark
on the main row, then 3 individual attempts on sub-rows (from col G).
Highlights PRs and school records.

For `8th Grade Pentathlon`, this helper renders only 2 attempt rows.

---

#### `renderPentathlonPlacePointsBlock(sheet, entryData, rosterData, meetNum, selectedGender, startRow, startCol)`

Builds the Pentathlon standings table shown at the bottom of `Event_Form_Printable`.

- Starts in column A and spans 5 columns:
   - col A: `Meet Place`
   - cols B-D merged: athlete name
   - col E: `Total Points`
- Reads combined `place/points` values from `Data_Entry` col I (`Place`)
- Uses `findRosterAthleteRow()` to label each row with the athlete's roster gender
- Supports combined generation while still showing per-athlete `Girls` / `Boys`

---

#### `parsePentathlonPlacePoints(placeVal)`

Parses Pentathlon combined standings values from the `Place` column.

Expected/common formats include:
- `10/2643`
- `5th/3060`
- `10th Place / 2643`

Returns `{ placeNum, pointsText }`.

---

### Conference Lineup Helpers

---

#### `formatConferenceNamesFromParts(nameParts, sortOutput)`
**Line:** 1893

Given an array of `{displayName, lastName}` objects, formats names as
`"FirstName L."` (last initial) and joins them with `", "`.
If `sortOutput` is true, sorts alphabetically by last name before formatting.

> Legacy function — the newer `formatConferenceEntriesFromParts` should be
> preferred for individual events.

---

#### `formatConferenceEntriesFromParts(nameParts, sortOutput)`
**Line:** 1934

Like `formatConferenceNamesFromParts` but carries `seedEstimate` and `seedSource`
through the formatted output. Returns an array of `{formattedName, seedEstimate, seedSource}`
objects for individual use (writing name and seed to separate cells).

---

#### `isHeatSeedingEvent(eventName)`
**Line:** 1973

Returns `true` if the event is a track (not field) event — i.e., seeding order by
time makes sense. Delegates to `!isFieldEvent(eventName)`.

---

#### `getConferenceSeedEstimate(rosterData, entryRow, eventName)`
**Line:** 1977

Returns `{value, source}` for an individual athlete's seed estimate.

Priority:
1. PR from Roster — `source: "PR: X"`
2. Best result from Data_Entry — `source: "Result: X (no PR on record)"`
3. Split from Data_Entry col G — `source: "Split: X"`
4. Nothing found — `source: "No PR or result on record — add to Roster or Data_Entry"`

---

#### `getRelayTeamSeedEstimate(teamEntries, rosterData, entryData, meetNum, selectedGender, eventName)`
**Line:** 2000

Returns `{value, source}` for a relay team's seed estimate.

Priority:
1. **Same-group past team mark** — identifies the relay team by sorted athlete keys
   (see `getRelayTeamGroupKey`) and finds their best past result.
   Source: `"Past team time: X"` or `"Past times: X, Y (using best)"`
2. **Sum of individual leg estimates** — sums each athlete's best individual PR/result.
   Source: `"Legs: Name1 X + Name2 Y + ..."`
3. **Partial estimate** — if some athletes have no individual mark.
   Source: `"Name1 X, Name2 Y; Missing: Name3, Name4"`

---

#### `getRelayTeamGroupKey(teamEntries, rosterData)`
**Line:** 2077

Builds a canonical group key for a relay team by sorting the normalized match keys
of all athletes. Used to identify "the same team" across different meets/years.

---

#### `getRelayAthleteMatchKey(rosterData, athleteName)`
**Line:** 2091

Normalizes an athlete name for relay group matching (lowercased, whitespace-trimmed,
falls back from Display Name to Athlete Name).

---

#### `getRelayTeamMarkFromEntries(teamEntries, eventName)`
**Line:** 2106

Finds the best (fastest) team result from a set of relay entry rows for a given event.
Returns the best mark as a string, or `null` if no valid marks exist.

---

#### `compareEventSeedValues(a, b, eventName)`
**Line:** 2121

Sort comparator for seed values. TBD/empty values sort to the end.
Field events sort descending (bigger first); track events sort ascending (faster first).

---

#### `formatConferenceSeedText(seedEstimate, eventName, includeConferenceSeeding)`
**Line:** 2133

Returns the formatted seed suffix for a conference lineup entry:
- `" (X)"` if there is a valid estimate
- `" (TBD)"` if no estimate and seeding is enabled
- `""` if seeding is disabled

---

#### `formatConferenceNames(athleteNames)`
**Line:** 2145

Legacy helper — formats a flat array of athlete display name strings into
`"FirstName L., FirstName L."` conference format.
Not used in current report logic; prefer `formatConferenceEntriesFromParts`.

---

### Conference Seeding Mode

---

#### `configureConferenceSeedingMode()`
**Line:** 1237

Prompts the user to select conference seeding mode via dialog:
- `1` = Championship meets only (default)
- `2` = All meets
- `3` = Off

Saves to Script Properties under `CONF_SEED_MODE_PROP_KEY`.

---

#### `getConferenceSeedingMode()`
**Line:** 1267

Reads the stored seeding mode from Script Properties.
Returns `CONF_SEED_MODE_DEFAULT` ("championship") if not set.

---

#### `shouldIncludeConferenceSeedEstimates(meetType)`
**Line:** 1274

Returns `true` if seed estimates should be included in the Conference Lineup section.
- Mode `"off"` → always false
- Mode `"all"` → always true
- Mode `"championship"` (default) → true only if `meetType === "Championship"`

---

### PR Update

---

#### `findAndUpdatePRs()`
**Line:** 3633

Scans all individual results for a meet/gender, compares to current Roster PRs
using `isBetter()`, shows a preview dialog (up to 20 entries, with a count for the
rest), then writes approved updates directly to the Roster tab.

- Excludes relay events and no-mark results
- Aborts with an error if any athlete name is not found in Roster
- Aborts if any event column is not found in Roster headers
- Shows an info message (no changes) if no PRs would be updated

**Preview format:** `Name | Event: old PR → new PR`

---

### Utility — Event Classification

---

#### `isFieldEvent(ev)`
**Line:** 2218

Returns `true` for High Jump, Shot Put, Discus, Long Jump.
Used throughout to determine sort direction (field = bigger is better).

---

#### `isRelayEvent(ev)`
**Line:** 2227

Returns `true` if `ev` is in `RELAY_EVTS`.

---

#### `isSplitEvent(ev)`
**Line:** 2231

Returns `true` if `ev` is in `SPLIT_EVTS` (800 M Run, 1600 M Run).

---

#### `isAttemptEvent(ev)`
**Line:** 2235

Returns `true` if `ev` is in `ATTEMPT_EVTS` (Shot Put, Discus, Long Jump).

---

#### `isSpecialRelayEvent(eventName)`
**Line:** 2262

Returns `true` for non-standard relay events (Hurdle Shuttle, Sprint Relay, etc.)
that have variable leg distances.

---

#### `getMeetType(meetRow)`
**Line:** 2208

Returns the `Type` value from a Schedule row (col D, index 3).
Falls back to `"Regular"` if blank.

---

#### `getMeetEventList(meetType)`
**Line:** 2212

Returns the appropriate event list constant for a given meet type:
- `"Relays"` → `RELAY_MEET_EVT`
- `"8th Grade Pentathlon"` → `PENTATHLON_EVT`
- Everything else → `PRINT_EVT`

---

#### `getTrackEventCount(eventList)`
**Line:** 2222

Returns the count of non-field events in an event list.
Used by the two-column layout logic to determine the split point.

---

### Utility — Name & PR Lookup

---

#### `findPR(rData, athleteName, eventName)`
**Line:** 3852

Looks up an athlete's PR from Roster data.
Match order: Display Name (col B) first, then Athlete Name (col A) as fallback.
Case-insensitive. Returns the value as a string, or `null` if not found.

No fuzzy matching — the name in Data_Entry must exactly match the Roster Display Name.

---

#### `findRosterAthleteRow(rosterData, athleteName)`
**Line:** 2239

Returns the Roster row array for a given athlete name (case-insensitive match on
Display Name first, then Athlete Name). Returns `null` if not found.

---

#### `getStartDistance(rosterData, athleteName, eventName)`
**Line:** 2249

Returns the athlete's starting distance for High Jump or Long Jump
from the Roster `High Jump Start Dist` / `Long Jump Start Dist` columns.
Returns `null` for all other events or if not set.

---

#### `validateRosterNames(entryData, rosterData, meetNum, gender)`
**Line:** 2508

Validates that all athlete names in Data_Entry for a given meet/gender have a
matching entry in the Roster. Returns a sorted array of unmatched name strings.

---

#### `checkAthleteEventCount(entryData, meetNum, gender)`
**Line:** 2544

Returns a list of athletes entered in more than 4 events for a meet.
Each entry is `{name, count, events}`. Relay entries are counted separately
per team slot.

---

### Utility — Performance Comparison

---

#### `isBetter(val, bench, event)`
**Line:** 3816

Returns `true` if `val` is a better performance than `bench` for the given event.
- Field events: bigger is better (distance/height)
- Track events: smaller is better (time)

Handles time strings (`"1:02.45"`), feet-inches strings (`"18'3"`), and plain numbers.
Returns `false` for no-mark values.

---

#### `isNoMark(val)`
**Line:** 2190

Returns `true` for DNS, DNR, DQ, NH, NM, `-`, empty, etc.
These values are excluded from highlights and comparisons throughout the system.

---

#### `parseComparablePerformance(val)`
**Line:** 2317

Converts a performance value to a number for comparison:
- Time string `"1:02.45"` → total seconds as float
- Feet-inches `"18'3\""` → total inches as float
- Plain number string → float
- Returns `null` for no-marks or unparseable values

---

#### `parseDistanceToInches(val)`
**Line:** 2337

Parses a distance value (feet-inches string or plain inches) to total inches.
Used in relay field event distance summation.

---

#### `formatFeetInchesFromInches(totalInches)`
**Line:** 2355

Converts a total-inches number back to a `feet'inches"` display string.

---

#### `parseDelimitedCellValues(val)`
**Line:** 2288

Splits a comma-separated string (from Data_Entry col G) into trimmed tokens.
Returns an empty array for blank values.

---

#### `getRelayComparableResult(splitVal)`
**Line:** 2294

Finds the best (smallest) time value from a comma-separated split string.
Used to determine a relay athlete's best leg split for PR comparison.

---

### Utility — Relay Helpers

---

#### `getRelayLegDistanceLabel(entryRow, eventName)`
**Line:** 2266

Returns a short distance label for a relay leg (e.g., `"100M"`, `"200M"`)
based on the athlete's entry row and the event. Returns `null` for standard relays.

---

#### `formatRelayPrDisplay(pr, eventName, entryRow)`
**Line:** 2283

Formats a relay athlete's PR for display in the Lineup_View — uses leg split
where available, falling back to team PR.

---

#### `getRelayLegSplitLabels(eventName, legIndex, teamSize)`
**Line:** 2428

Returns the display label(s) for a relay leg's split sub-row(s)
(e.g., `"Leg 1 Split"`, `"400M Split"`).

---

#### `sumRelayFieldTeamDistanceInches(members)`
**Line:** 2366

Sums the field event distances of all relay team members (in inches).
Used by relay field event blocks to calculate team totals.

---

#### `getRelayFieldTeamSchoolRec(recordsData, gender, eventName)`
**Line:** 2379

Looks up the school record for a relay field event by gender and event name.

---

#### `applyRelayPrHighlight(rowRange, resultCell, res, pr, ev)`
**Line:** 2415

Applies green (PR) or yellow (school record) highlight to a relay result cell
based on comparison with the athlete's current PR.

---

### Utility — School Records

---

#### `getEventFormSchoolRecDisplay(recordsData, selectedGender, eventName, meetType, defaultSchoolRec)`
**Line:** 2396

Returns the school record display value for an event in the Event Forms report.
Handles gender selection (Girls/Boys/Combined) and relay-vs-individual record lookup.

---

#### `getAthleteSchoolRecForEvent(recordsData, selectedGender, athleteGender, eventName, meetType, defaultSchoolRec)`
**Line:** 2408

Returns the appropriate school record for an individual athlete given their gender
and the event.

---

### Utility — Formatting

---

#### `formatCellValue(val)`
**Line:** 3605

Converts a raw GAS cell value to a display string.
Handles the critical case where time-formatted cells (e.g., `"1:02.45"`) are
returned by `getValues()` as JavaScript `Date` objects anchored at 1899-12-30.
Without this, `setValue()` would render them as `"12/30/1899"`.

Also ensures numeric values (e.g., `18.0`) are rendered with 2 decimal places.

---

#### `formatSecondsAsTimeString(seconds)`
**Line:** 2308

Converts a total-seconds float to a `M:SS.ss` display string.
Used when generating seed estimates from summed relay legs.

---

#### `formatMeetDateForTitle(meetDateVal)`
**Line:** 1525

Formats a meet date (Date object or string from Schedule col B) as `"Mon DD"` for
use in report title rows. Returns empty string if the value is blank.

---

#### `formatMeetDateForFileName(meetDateVal)`
**Line:** 1516 (**DEPRECATED** — used only by deprecated PDF helpers)

Formats a meet date as `"MMDD"` for use in PDF file names.

---

### Utility — Gender Helpers

---

#### `isCombinedGenderSelection(gender)`
**Line:** 2195

Returns `true` if the gender dropdown value is `"Combined"`.

---

#### `matchesSelectedGender(rowGender, selectedGender)`
**Line:** 2199

Returns `true` if a Data_Entry row's gender matches the selected gender.
For `"Combined"`, always returns `true`.

---

#### `getOutputGenderLabel(gender)`
**Line:** 2204

Returns a display label: `"Girls"`, `"Boys"`, or `"All"` for combined.

---

### Utility — Sheet Management

---

#### `getOrCreateSheet(ss, name)`
**Line:** 2438

Returns an existing sheet by name, or inserts a new one at the end of the spreadsheet.
Used throughout `fullInitialize()` to safely create tabs.

---

#### `applyReportLayout(sh)`
**Line:** 2444

Sets the standard column widths and number formats for output report sheets
(`Lineup_View`, `Event_Form_Printable`):
- Cols A, D: 190px
- Cols B, E: 100px
- Col C: 20px (gutter)
- Col F: number format `@` (force-text to preserve time strings)

---

### Debug Tools

---

#### `checkPRSetup()`
**Line:** 2462

Shows a dialog listing:
- Athletes in Data_Entry with no matching Roster entry (unmatched names)
- Roster rows where Display Name is blank

Useful when setting up a new season or after renaming athletes.

---

#### `checkMeetRoster()`
**Line:** 2583

For a selected meet/gender (prompted), validates all Data_Entry athlete names
against the Roster. Shows a list of any names that don't match.

---

#### `validateRosterNames(entryData, rosterData, meetNum, gender)`
**Line:** 2508

Internal helper called by `generateLineupReport()` and `generateEventFormReport()`
before generating a report. Returns an array of unmatched athlete names.

---

#### `checkAthleteEventCount(entryData, meetNum, gender)`
**Line:** 2544

Internal helper — checks if any athletes are entered in more than 4 events.
Returns `[{name, count, events}]` for over-limit athletes.

---

### Deprecated Functions (PDF Export)

> All functions below are **deprecated** as of v2.98.
> PDF export was disabled due to GAS limitations with page breaks and margins.
> The feature flag `ENABLE_PDF_EXPORT = false` hides these from the menu.
> Functions are retained for reference but should not be used or extended.

---

#### `saveLineupPdf()`
**Line:** 1086 — DEPRECATED

Entry point for saving Lineup_View as PDF to Drive.
Hard-stops if `ENABLE_PDF_EXPORT` is false.

#### `saveEventFormsPdf()`
**Line:** 1098 — DEPRECATED

Entry point for saving Event_Form_Printable as PDF.

#### `saveReportPdf(sheetName, reportLabel)`
**Line:** 1110 — DEPRECATED

Core PDF save routine — resolved Drive folder, prompted for suffix, exported via
GAS `getAs(MimeType.PDF)`. Unreliable due to GAS page-break/margin limitations.

#### `promptPdfExportMode(ui, reportLabel)`
**Line:** 1222 — DEPRECATED

Prompted user to select PDF export mode (Drive vs. browser download).

#### `ensureDriveAuthorization(ui)`
**Line:** 1285 — DEPRECATED

Checked and initiated Drive OAuth flow if needed.

#### `showDriveAuthorizationDialog(authUrl)`
**Line:** 1306 — DEPRECATED

Showed an HTML dialog with Drive authorization URL.

#### `authorizeDriveAccess()`
**Line:** 1328 — DEPRECATED

Triggered Drive OAuth authorization.

#### `resolvePdfFolder(ui)`
**Line:** 1335 — DEPRECATED

Resolved or prompted for a Drive folder to save PDFs into.

#### `checkDriveAccess()`
**Line:** 1400 — DEPRECATED

Debug tool — tested Drive access and reported status.

#### `getOrCreatePdfFallbackFolder()`
**Line:** 1463 — DEPRECATED

Created a fallback Drive folder ("FinishLine PDFs") if no folder was configured.

#### `promptPdfSuffix(ui, reportLabel)`
**Line:** 1470 — DEPRECATED

Prompted user for optional filename suffix.

#### `extractDriveFolderId(input)`
**Line:** 1500 — DEPRECATED

Extracted a folder ID from a Drive URL or raw ID string.

#### `sanitizeFilePart(val)`
**Line:** 1508 — DEPRECATED

Stripped unsafe characters from filename components.

---

## Report Layout Logic

### Two-Column Layout (Lineup_View & Event_Form_Printable)

Reports use a two-column layout across a 5-column (A–E) grid:
- **Left column:** cols A–B (athlete name + PR/result)
- **Gutter:** col C (20px, blank)
- **Right column:** cols D–E (athlete name + PR/result)

**Event placement algorithm:**
1. Compute `TRACK_COUNT` = number of non-field events in event list
2. `HALF_TRACK` = `Math.ceil(TRACK_COUNT / 2)` — left column events
3. Events 0 to `HALF_TRACK-1` → left column
4. Events `HALF_TRACK` to `TRACK_COUNT-1` → right column
5. At transition to field events: sync both columns to `Math.max(curL, curR) + 2`
6. Field events alternate left/right

Minimum 6 data rows per event (padded with empty bordered rows).

### Highlight Color Key

| Color | Hex | Meaning |
|-------|-----|---------|
| Yellow | `#ffe599` | School record broken |
| Green | `#b6d7a8` | PR or first-time result |
| Blue text | `#1155cc` | Existing PR shown inline |
| Light blue | `#cfe2f3` | No prior PR (Lineup only) |
| Amber header | `#bf9000` | School record header cell (Event Forms) |

No highlighting for `isNoMark()` values (DNS, DNR, DQ, -, NH, NM, etc.).

---

## Conference Seeding Logic

The Conference Lineup section of `generateLineupReport()` can optionally display
seed time estimates next to each athlete's name. Controlled by `getConferenceSeedingMode()`:

- `"championship"` (default) — estimates only shown for Championship meets
- `"all"` — estimates shown for all meets
- `"off"` — estimates never shown

### Individual Athlete Seed Priority (`getConferenceSeedEstimate`)

1. PR from Roster
2. Best result from Data_Entry (any meet)
3. Best split from Data_Entry col G
4. TBD

### Relay Team Seed Priority (`getRelayTeamSeedEstimate`)

1. Best past team mark for this same athlete group (cross-meet lookup)
2. Sum of each athlete's individual best (PR or Data_Entry result)
3. Partial sum with "Missing: Name" annotation

### Output Format

- Seed appended to name: `"Emma S. (1:12.34)"` or `"Emma S. (TBD)"`
- Seed basis written to col B in italic grey (8pt)
- Individual track events sorted fastest-first within each event

---

## PR Lookup Logic

`findPR()` matches names case-insensitively:
1. **Display Name** (Roster col B) — primary key
2. **Athlete Name** (Roster col A) — fallback if Display Name is blank

No fuzzy matching. If a PR lookup fails silently, the coach should verify that
the Data_Entry col D name exactly matches the Roster Display Name.

---

## Common Patterns

### Reading Sheet Data

```javascript
const entryData = ss.getSheetByName('Data_Entry').getDataRange().getValues();
// entryData[0] = headers, entryData.slice(1) = data rows
```

### Writing Without Breaking User Data

```javascript
// Safe: only write to row 1
sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
// Never: sheet.clear() on a data tab
```

### Flushing Mid-Execution

```javascript
statusCell.setValue("⏳ Generating...");
SpreadsheetApp.flush(); // pushes to browser immediately
// ... long-running work ...
statusCell.setValue("✅ Done!");
```

### Handling Date-Formatted Time Cells

```javascript
// BAD: returns "12/30/1899" for time cells
const val = row[5].toString();

// GOOD: use formatCellValue()
const val = formatCellValue(row[5]); // returns "1:02.45"
```

---

## Adding a New Feature

1. **New menu item:** Add `addItem(label, 'functionName')` in `onOpen()`. If it's
   an experimental feature, consider gating it behind a constant flag like `ENABLE_PDF_EXPORT`.

2. **New Home button:** Add a `btnStyle()` call in `fullInitialize()` and a new
   `if (row === N)` branch in `onEditInstallable()`. Update `HOME_STATUS_CELL` area if needed.
   **Requires fullInitialize rebuild.**

3. **New output tab:** Add to the `['Lineup_View', ...]` forEach in `fullInitialize()` and
   to the tab colors block. **Requires fullInitialize rebuild.**

4. **New event:** Add to `FULL_EVT` and the appropriate sub-list (`PRINT_EVT`, `RELAY_EVTS`, etc.).
   Add an entry to the Roster headers in `fullInitialize()`.
   **Requires fullInitialize rebuild** (Roster headers change).

5. **Bump `VERSION`** on every change.
