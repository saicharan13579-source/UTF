# Data and daily record

Editing happens on the site itself (Data entry tab, editor name + PIN). Edits are stored in the UTF Google Sheet
through the Apps Script in `tools/google-sheet/Code.gs`; the sheet URL is set in `data/config.json` (`sheetUrl`).

Sheet tabs: `plots` (edited plots), `edits` (change log), `daily` (counts per day), `current` (all plots as last shown),
and `Snapshot YYYY-MM-DD` (copied every evening at 9 pm IST by the script's daily trigger).

To rebuild the site after changing `src/` or `data/`: `python3 tools/build.py`, commit and push.
