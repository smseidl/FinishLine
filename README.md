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
4. A success alert will appear when done. Six tabs will be created:
   - `Home` — control panel
   - `Schedule` — meet calendar
   - `Data_Entry` — results entry
   - `Roster` — athlete list with PRs
   - `School_Records` — record book
   - `Lineup_View` and `Event_Form_Printable` — generated output

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
| `Schedule` | One row per meet: Meet #, Date, Meet Name, Location, etc. |
| `Roster` | One row per athlete: Athlete Name, **Display Name**, Gender, Grade, PRs |
| `School_Records` | One row per record: Gender, Event, Athlete, Record |
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

---

## Copying to a New Season

1. **File → Make a copy** of the spreadsheet.
2. Clear `Data_Entry` rows 2+ (keep the header row).
3. Update `Schedule` with new meets.
4. Re-run **1. Build / Rebuild** to reset the output tabs.
5. **Re-create the installable trigger** (Step 4 above) — triggers do not copy with the sheet.

---

## Debug Tools (🏁 FINISH LINE menu)

| Tool | Purpose |
|------|---------|
| **Check PR Setup** | Finds athletes in `Data_Entry` with no matching Roster entry, and Roster rows missing a Display Name |
| **Check Meet Roster** | For a specific meet/gender, lists any `Data_Entry` names that don't match the Roster |

---

## Repo Layout

```
FinishLine/
├── scripts/
│   └── FinishLine.gs    ← paste this into Apps Script
├── instructions/        ← development notes and AI context
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
