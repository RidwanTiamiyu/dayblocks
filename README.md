# DayBlocks

Personal time log. Your day as 96 quarter-hour slots, shown as 15, 30 or 60 minute blocks.

- **Log:** tap a start block, then an end block, then pick an activity. Or long-press and drag. The **+** button logs a time range.
- **Summary:** day, week and month totals, averages per day, change vs the previous period, and a day-by-day timeline.
- **Export:** blocks CSV, daily-totals CSV (opens in Excel/Sheets), a ready-made AI prompt, and a JSON backup/restore.
- **Optional Google Sheet sync + Monday email summary:** see `apps-script/Code.gs`.

## Install on Android
Open the GitHub Pages URL in Chrome, then tap ⋮ > **Add to Home screen** (or **Install app**).

## Google Sheet sync (optional, about 5 min)
1. New Google Sheet > Extensions > Apps Script > paste `apps-script/Code.gs` > Save.
2. Select `setup` > Run > approve. Copy the token from the execution log.
3. Deploy > New deployment > type **Web app** > Execute as **Me**, access **Anyone** > Deploy > copy the `/exec` URL.
4. In the app: Settings > paste the URL and token > **Sync everything now**.

The token lives in Script Properties, not in code. Data stays in your own Google account.
