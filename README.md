# India IPO Tracker

A self-updating IPO tracker for Indian Mainboard + SME IPOs, built with **Google Sheets + Google Apps Script**. No server, no API keys, no cost — the script scrapes public IPO data and keeps the spreadsheet current every day on its own.

## What it does

- Fetches current **Mainboard + SME IPO** data from Chittorgarh (free source) with NSE India as backup.
- Sorts every IPO into the right tab automatically: **Upcoming → Open → Closed → Listed**.
- When an IPO's subscription closes, it moves to the *Closed* tab; once listed, it archives to *Listed*.
- For Open IPOs it pulls price band, lot size, issue size, subscription (RII / QIB / NII) and GMP from each IPO's detail page.
- Appends a daily GMP snapshot to the *GMP Log* tab.
- A trigger runs `updateAll` every morning (~8–9 AM IST), so the sheet updates itself.

## Repository structure

```
apps-script/
  IPO_Tracker_AppsScript.gs   # The full Apps Script (paste into the sheet)
data/                          # CSV snapshot of every tab (as of 17/09/2026)
  open.csv                     # Currently open IPOs
  upcoming.csv                 # Announced, not yet open
  closed.csv                   # Subscription closed, awaiting listing
  listed.csv                   # All-time archive of listed IPOs
  gmp_log.csv                  # Daily grey-market premium snapshots
  config.csv                   # Source & automation settings
```

The Google Sheet stays the live system; this repo versions the script and keeps point-in-time CSV backups of the data.

## Setup (one time, ~5 minutes)

1. Open the **IPO Tracker - India** spreadsheet (or a new blank sheet).
2. Click **Extensions > Apps Script**, delete any code shown.
3. Paste the entire contents of [`apps-script/IPO_Tracker_AppsScript.gs`](apps-script/IPO_Tracker_AppsScript.gs).
4. Click the **Save** icon.
5. In the toolbar dropdown select the function **`updateAll`** and click **Run**.
6. Authorize when Google asks: *Review permissions > your account > Advanced > Go to project > Allow*.
7. For daily automation, run **`installTrigger`** once (or add a trigger manually: function `updateAll`, event source *Time-driven*, *Day timer*, 8am–9am).
8. Optional: in Apps Script **Project Settings**, set the time zone to *(GMT+05:30) Kolkata* so the trigger fires in the morning.

## How the tabs work

| Tab | Meaning | Maintained by |
|---|---|---|
| Dashboard | Quick stats + diagnostics from the last run | `updateAll` |
| Upcoming | IPOs announced but not yet open for subscription | `updateAll` |
| Open | IPOs whose subscription window is currently open | `updateAll` |
| Closed | Subscription over, awaiting listing | auto-move |
| Listed | Listed IPOs with issue price, listing price & listing-day gain | auto-archive |
| GMP Log | One row per open IPO per day: date, name, GMP | `updateAll` |
| Config | Sources, trigger schedule, disclaimer | manual |

## Data sources & disclaimer

- Primary source: Chittorgarh (IPO lists, dates, subscription data); secondary: NSE India.
- **GMP (grey market premium) is unofficial** grey-market data and can be manipulated. Always verify with the company's RHP before making any investment decision.
- If a source site changes its layout, the script logs errors instead of crashing — update the parsing functions and re-paste the script.

> This tracker is an informational tool, not investment advice.
