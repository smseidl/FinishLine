# 🏁 FinishLine — Track & Field Meet Management

Google Apps Script for middle school track & field meet management.
Generates printable Lineup and Event Form reports from Google Sheets data.

---

## New Season Setup (do this once per season)

### Step 1 — Create the Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) and create a **blank** spreadsheet.
2. Name it something like `Track 2026`.

### Step 2 — Add the Script

1. In the spreadsheet, go to **Extensions → Apps Script**.
2. Delete all existing code in `Code.gs` (select all, delete).
3. Open `scripts/FinishLine.gs` from this repo and copy the entire contents.
4. Paste it into the Apps Script editor.
5. Click **Save** (💾 or Ctrl+S).
6. Close the Apps Script tab and **reload** the Google Sheet.
   - The `🏁 FINISH LINE` menu should appear in the menu bar.

### Step 3 — Build the System

1. Click **🏁 FINISH LINE → 1. Build / Rebuild Entire System**.
2. Accept the authorization prompt if it appears (first run only).
3. Click through any "This app isn't verified" warnings — this is your own script.
4. A success alert will appear when done. These tabs will be created:
   - `Home` — control panel
   - `Schedule` — meet calendar
   - `Data_Entry` — results entry
   - `Roster` — athlete list with PRs
   - `School_Records` — record book
   - `Historical_PRs` — past-year PRs for year-over-year comparisons
   - `Lineup_View` and `Event_Form_Printable` — generated output
   - `Top_Marks` and `Athlete_Recaps` — additional generated reports

### Step 4 — Set Up the Installable Trigger ⚠️ (required)

This enables the Home tab buttons to show live status. **Must be done once per spreadsheet** — it does not copy when you duplicate a sheet.

1. Go to **Extensions → Apps Script**.
2. Click the **Triggers** icon (clock ⏰) in the left sidebar.
3. Click **+ Add Trigger** (bottom right).
4. Configure:
   - **Function:** `onEditInstallable`
   - **Event source:** `From spreadsheet`
   - **Event type:** `On edit`
5. Click **Save**.

### Step 5 — Enter Your Data

Fill in the data tabs (row 2 onward — **never edit row 1 headers**):

| Tab | What to enter |
|-----|---------------|
| `Schedule` | One row per meet: Meet #, Date, Meet Name, Location, Type, etc. |
| `Roster` | One row per athlete: Athlete Name, **Display Name**, Gender, Grade, PRs |
| `School_Records` | One row per record: Gender, Event, Athlete, Record |
| `Historical_PRs` | Past-year PRs (same format as Roster) — copy from Roster at end of each season |
| `Data_Entry` | Results after each meet: Meet #, Gender, Event, Athlete Name, Result |

> **Display Name** is the name the coach uses in `Data_Entry`. It's the primary key for PR lookups.
> If an athlete goes by "Emma" in results, put `Emma` as the Display Name even if her full name is `Emma Johnson`.

---

## Each Meet Workflow

1. Enter results in `Data_Entry` (filter by meet # for convenience).
2. On the `Home` tab, select the **Meet #** and **Gender** from the dropdowns.
3. Click **▶ Generate Printable Lineup** or **▶ Generate Printable Event Forms**.
4. Wait 5–15 seconds for the status message to appear (normal GAS startup delay).
5. Switch to `Lineup_View` or `Event_Form_Printable` to print.
   - For printing: **File → Print**, set to landscape, adjust page breaks manually as needed.
6. After the meet, click **▶ Update PRs from This Meet** to automatically update athlete PRs
   in the Roster. You'll see a preview of all changes before anything is saved.

### Pentathlon Notes

- Pentathlon Event/Result forms use an individual-results layout, not team scoring.
- No `TEAM PLACE` or `TEAM POINTS` row is shown for Pentathlon meets.
- No `At-Meet Additions` lines are shown on Pentathlon forms.
- Shot Put and Long Jump render with **2 attempts** instead of 3.
- A standings table is added at the bottom showing `Meet Place`, athlete, and `Total Points`.
- Pentathlon standings read the combined `place/points` value from the `Place` column in `Data_Entry`.
- In combined-generation mode, the standings table still labels each athlete with their actual roster gender.

---

## Copying to a New Season

1. **File → Make a copy** of the spreadsheet.
2. Clear `Data_Entry` rows 2+ (keep the header row).
3. Update `Schedule` with new meets.
4. Re-run **1. Build / Rebuild** to reset the output tabs.
5. **Re-create the installable trigger** (Step 4 above) — triggers do not copy with the sheet.

---

## Additional Tools (🏁 FINISH LINE menu)

### Reports

| Menu Item | Purpose |
|-----------|--------|
| **3. Generate Printable Lineup** | Pre-meet lineup with PRs; also writes a By-Athlete view and Conference Lineup section |
| **4. Generate Printable Event/Result Forms** | Post-meet event forms with results, PR/record highlights |
| **7. Generate Top Marks (YTD)** | Year-to-date top N performers per event (set N in the Home tab Top N field) |
| **8. Generate All Athlete Recaps** | Year-end recap per athlete: meet-by-meet results + year-over-year PR table |
| **9. Generate Filtered Results** | Filter results by grade and/or event with multiple sort options |
| **10. Export Event Placers (Email/AI)** | Export placed results formatted for email or AI use; includes sanity checks |
| **11. Build Time-Trial List (100M)** | Builds a 100 M Dash entry list for athletes not in a main meet (time-trial seeding) |

### Conference Seeding Mode

| Menu Item | Purpose |
|-----------|--------|
| **2. Set Conference Seeding Mode** | Configure when estimated seed times appear in the Conference Lineup section of the lineup report. Options: Championship only (default), All meets, or Off |

### Debug Tools

| Tool | Purpose |
|------|--------|
| **Check PR Setup (debug)** | Finds athletes in `Data_Entry` with no matching Roster entry, and Roster rows missing a Display Name |
| **Check Meet Roster (debug)** | For a specific meet/gender, lists any `Data_Entry` names that don't match the Roster |

---

## Repo Layout

```
FinishLine/
├── scripts/
│   └── FinishLine.gs    ← paste this into Apps Script
├── instructions/
│   ├── DEVELOPER.md     ← developer reference: all functions documented
│   └── FINISHLINE_CONTEXT.md
└── README.md
```

---

## Highlight Color Key (Event Forms)

| Color | Meaning |
|-------|---------|
| 🟡 Yellow row | School record broken |
| 🟢 Green row | Personal record or first-time result |
| 🔵 Blue PR text | Existing PR shown for reference |
| 🔵 Light blue cell | No prior PR on record (Lineup only) |

---

## Conference Lineup

The bottom section of the Lineup report (`CONFERENCE LINEUP`) is formatted for
digital conference submission. It lists each athlete by first name and last initial.

For Championship meets (default), estimated seed times appear next to each name
in parentheses (e.g., `Emma S. (1:12.34)`). A brief note explaining the estimate
(PR, past result, or relay leg sum) is shown alongside each entry.

Athletes with no estimate show `(TBD)` — add their result to `Data_Entry` or PR
to the `Roster` to generate an estimate.

To change when seeding estimates appear, use **🏁 FINISH LINE → 2. Set Conference Seeding Mode**.
