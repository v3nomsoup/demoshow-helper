# Demoshow helper

Scripts to prepare a BDSM (Berlin DemoScene Meetup) / demoparty demoshow sheet
inside `Demoshows.xlsx`.

## How the sheet is used

Each demoshow tab has these columns (row 1 is the header):

| col | name | notes |
|-----|------|-------|
| A | title | filled by `fill_demoshow.py` from the demozoo page |
| B | group(s) | filled by `fill_demoshow.py` (joined with ` + ` if multiple); may be blank if demozoo has no author |
| C | link | input: demozoo URL. Cell text and hyperlink are both set to the URL |
| D | platform | set manually (e.g. `pc 4k`, `Amiga Demo`, `wild`, ...) |
| E | Party Placement | set manually (1 = compo winner) |
| F | YouTube | filled by `fill_demoshow.py` — canonical `https://www.youtube.com/watch?v=...` URL, also as hyperlink |
| G | runtime | filled by `fill_demoshow.py` — stored as an Excel duration with number format `hh:mm:ss` |
| H+ | OK / endorsed by / vetoed by / comment | manual |

The target runtime for a BDSM show is ~2h (see the hint in column L of the header row).

## Preparing a new demoshow tab

1. Create a new tab named e.g. `BDSM - May 2026`.
2. Add the header row. Easiest is to copy it from the previous month's tab.
3. For each production you want to include, create a row with **columns C, D, E only**:
   - **C**: paste the demozoo production URL (Excel will usually auto-hyperlink;
     if not, Insert → Link). The cell text can be anything — the script uses the
     hyperlink target.
   - **D**: platform category (free text; used for grouping when sorting).
   - **E**: party placement as an integer.
4. Run the fill script:
   ```
   python3 fill_demoshow.py --sheet 'BDSM - May 2026'
   ```
   It fills A (title), B (group(s)), F (YouTube), G (runtime). Rows that already
   have all four filled are skipped — you can re-run safely.
5. Sort the show order:
   ```
   python3 sort_demoshow.py --sheet 'BDSM - May 2026'
   ```
   Keeps the platform groups in their current order; within each group, sorts by
   placement **descending** (4th, 3rd, 2nd, 1st) which is the usual demoshow
   build-up. Pass `--asc` for winner-first.

The scripts always write to `Demoshows.xlsx` in the current directory by default;
use `--file` to point elsewhere.

## Scripts

### `fill_demoshow.py`

- Input: rows with a demozoo hyperlink in column C.
- Fetches each demozoo page (parallel, 6 workers by default, backs off on HTTP 429).
- Extracts `<h2>` title, `<h3> by ...</h3>` group(s) / scener, and the YouTube
  URL from the carousel JSON (falls back to any youtube.com / youtu.be link on the page).
- **Pouët fallback**: if demozoo has no YouTube link but links to a Pouët
  production, the script fetches the Pouët page and pulls the YouTube URL from
  there instead.
- Fetches each YouTube watch page and reads `lengthSeconds` for the runtime.
- Writes A, B, F, G. Canonicalizes the demozoo link in C to the URL itself (text
  and hyperlink). **Does not touch D, E, or any H+ columns.**
- Idempotent: only rows missing title or YouTube (and thus runtime) are re-fetched.

Expected gaps the script can't fix:
- **No group(s) on demozoo** — some productions have no author listed. Leave blank.
- **No YouTube link on demozoo or Pouët yet** — common for freshly released
  compos. The sheet will show blank; add the URL manually in column F when the
  capture appears, then re-run the script to get the runtime.
- **Runtime is way longer than expected** — the demozoo "video" link sometimes
  points at the whole compo capture rather than a single entry. Fix the YouTube
  URL manually and re-run.

### `sort_demoshow.py`

- Sorts rows in place within the given sheet.
- Platform groups retain their first-appearance order (so you decide the category
  order by the order you paste rows in).
- Within a platform, rows are sorted by column E. Default `desc` (worst placement
  first); `--asc` flips to winner-first.
- Preserves cell values, number formats, and hyperlinks.

## Runtime display convention

`fill_demoshow.py` stores the actual video duration as an Excel duration
(timedelta) and applies the number format `hh:mm:ss`, so a 3-minute-45-second
demo shows as `00:03:45`. Both Excel and Google Sheets render this consistently.

(The older sheets in this workbook use `[h]":"mm` and store e.g. a 3:45 video
as `timedelta(hours=3, minutes=45)` — a workaround for the way Excel parses
typed `H:MM` input. New tabs filled by the script use the cleaner format above.)

## Requirements

- Python 3.9+
- `openpyxl` — `pip3 install openpyxl`

## Example file

`example.xlsx` contains the `BDSM - April 2026` sheet already filled, so you can
inspect the expected output. Running `fill_demoshow.py` again is a safe no-op on
already-filled rows (only rows missing title or YouTube get re-fetched):

```
python3 fill_demoshow.py --file example.xlsx
python3 sort_demoshow.py --file example.xlsx
```

## Backups

Before a destructive run, `cp Demoshows.xlsx Demoshows.xlsx.bak`. The scripts do
not create backups themselves.
