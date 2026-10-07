# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Django site for a delivery company's driver activity log. The data store is the Excel workbook
`Driver_Activity_Log.xlsx` (no database models, no migrations). Staff type trips into it by hand
or import SinoTrack GPS "Travel Report" CSV exports; the site lists drivers and shows each
driver's travel record.

## Commands

Use the project venv (Python 3.11, Django 5.2 LTS, openpyxl, pywin32): `.venv/Scripts/python` on Windows.

```bash
.venv/Scripts/python -m pip install -r requirements.txt
.venv/Scripts/python manage.py runserver                       # http://127.0.0.1:8000/
.venv/Scripts/python manage.py test activity                   # all tests
.venv/Scripts/python manage.py test activity.tests.ImportTests.test_reimport_skips_trips_already_in_the_log
.venv/Scripts/python manage.py import_travel_report "Travel Report_T05-GONZAGA.csv" [--driver X] [--plate Y] [--workbook PATH]
.venv/Scripts/python manage.py import_travel_report "Park report_T05-GONZAGA.csv"   # same command; report type is auto-detected
```

`ACTIVITY_LOG_PATH` (env var or setting) points the site and the import command at another
workbook; use a copy for manual testing so the real file is not modified.

## Architecture (`activity` app)

- `excel_log.py` reads the workbook (`load()` caches by file mtime/size, reads bytes so Excel
  can keep the file open). It **recomputes** travel time, stay, long stay and daily totals
  instead of reading the workbook's formula columns, because openpyxl-written files have no
  cached formula results. These rules must stay identical to the sheet formulas in Activity
  Log columns P–T:
  - travel = `MOD(arrive - depart, 1)`
  - stay = the *next sheet row's* depart minus this row's arrive, only if that row has the same
    date, driver and plate; if it departs earlier → "Check order"
  - long stay = stay ≥ `Lists!B3`
  - daily last-arrive uses depart + travel so legs past midnight sort last (column T).
- `sinotrack.py` parses SinoTrack CSV exports. `detect_report()` tells them apart by header:
  the Travel Report has one row per drive, the Park Report one row per stop with address and
  longitude/latitude. From the Travel Report the user wants only Device name, Start/End time,
  Start/End address and Start/End longitude/latitude. Drive mileage, speeds and drive time are
  deliberately not imported. Export quirks: UTF-8 BOM, every cell `"\t<value>"`, addresses ending in
  `" ."`, hundreds of trailing blank rows. Device names look like `T05_RJP 162 - Gonzaga`
  → unit, plate, driver. Coordinates are stored "lat, lon", the order Google Maps uses.
- `log_writer.py` writes both reports. Every change saves a copy in `backups/` first.
  - `import_trips()` appends trips to the next empty Activity Log rows and adds new
    drivers/plates to Lists. A trip already present (date + plate + depart + arrive) is not added
    again; only its **blank** cells are filled in. That is how new columns get completed for old
    rows, and it never overwrites typed values. From/To Coordinates cells get a Google Maps
    hyperlink.
  - `merge_parks()` fills Park Count / Park Time / Park Address / Park Coordinates of existing
    trip rows. Each stop goes to the row with the latest departure at or before the stop's
    start, so stops at the destination, stops mid-drive and overnight stops all land on that
    trip. Stops on days without trip rows are reported, not attached.
  - Several stops become multi-line cells, one line per stop. The coordinates cell gets an Excel
    hyperlink to the longest stop. Re-merging rewrites the same values.

  A planner (`_plan_trips` / `_plan_parks`) works on a `Snapshot` of the sheet, and two appliers
  write the resulting `Plan`:
  - **Workbook closed:** openpyxl, saved via a temp file + `os.replace`. Rows past the prepared
    formula rows get translated formulas, and the validation / conditional-formatting / autofilter
    ranges are stretched to match.
  - **Workbook open in Excel** (Windows denies other writers): `excel_com.py` finds the open
    workbook through the Running Object Table, so it never opens files or starts Excel. The
    applier writes cells with `Value2` and calls `Workbook.Save()`. The COM work runs in its own
    thread with its own CoInitialize/CoUninitialize; only plain values and messages cross back.
    Needs pywin32. Test this path in a hidden `DispatchEx` Excel instance on a copy, never in the
    user's open Excel.
- `views.py`: `drivers_activity` → `templates/activity/drivers_Activity.html` (all drivers
  plus the import form), `driver_travel_record` → `templates/activity/Driver_travel_record.html`,
  `import_travel_report` (POST), `export_travel_record` (`export/drivers/<name>/?start&end`).
  The export is built by `export.py`: an .xlsx with the record's legs and a Daily Summary, real
  date/time/duration values, and SUM/COUNTIF totals. Template filenames and their casing were chosen by the user;
  keep them. Formatting filters are in `templatetags/activity_extras.py`: `duration` (h:mm, for
  travel), `hms` (hh:mm:ss, which the user wants for stays, the long-stay rule and park time),
  `clock` and `number`. From/To and park coordinates link to
  `https://www.google.com/maps/search/?api=1&query=lat,lon` (`excel_log.maps_url`).
- Messages use `CookieStorage`, so there are no sessions/auth/admin apps.

## Workbook contract

Sheets `Activity Log` (headers on row 4, data from row 5, formula rows prepared to 2004),
`Lists` (long-stay threshold in B3; headers on row 6; drivers A, contacts B, plates D, vehicle
description E, SinoTrack device ID F, rows 7–56 feed the `DriverList`/`PlateList` named ranges),
`Daily Summary`, `Instructions`. Columns are located by **header text** (`LOG_HEADERS` /
`LISTS_HEADERS` in `excel_log.py`), not by letter.

The real workbook's Activity Log has these columns:
- Inputs A–O: Date, Driver Name, Plate No., From (Origin), From Coordinates, Depart Time,
  To (Destination), To Coordinates, Arrive Time, Park Count (whole number), Park Time
  (`[h]:mm:ss`), Park Address, Park Coordinates, Distance (km), Remarks / DR No.
- Formulas P–T: Travel, Stay, Long Stay, then hidden Key and Arrive After Depart.

The coordinate and park columns were inserted on 2026-10-07 through Excel, so formulas and Daily
Summary references shifted. To change the layout, insert or delete columns through Excel (COM)
in a hidden `DispatchEx` instance, with a backup first. openpyxl does not rewrite formula
references. The original template had "Purpose" where Park Count is, which is still accepted.

Distance (km) is no longer filled by imports. Rows 5–27 still hold the implausible SinoTrack
mileage copied earlier.
