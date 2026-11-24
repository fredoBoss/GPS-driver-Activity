import calendar
import datetime as dt
from collections import Counter
from urllib.parse import urlencode

from django.conf import settings
from django.contrib import messages
from django.http import Http404, HttpResponse
from django.shortcuts import redirect, render
from django.urls import reverse
from django.utils.http import content_disposition_header
from django.views.decorators.http import require_POST

from . import excel_log, export, log_writer, sinotrack

MAX_UPLOAD_BYTES = 10 * 1024 * 1024
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
SORTS = {'name': 'Name (A–Z)', 'added': 'Date added (newest first)'}
DEFAULT_SORT = 'name'
SORT_COOKIE = 'driver_sort'  # remembers the choice, e.g. after opening a driver and coming back
CALENDAR_MONTHS = 24  # the travel record's "did not travel" calendar shows at most this many months


def _date_range(request):
    """Read ?start=YYYY-MM-DD&end=YYYY-MM-DD; invalid values are ignored and reported."""
    dates, notices = {}, []
    for key in ('start', 'end'):
        raw = request.GET.get(key, '').strip()
        dates[key] = None
        if raw:
            try:
                dates[key] = dt.date.fromisoformat(raw)
            except ValueError:
                notices.append(f'Ignored the {key} date "{raw}" because it is not a valid date.')
    start, end = dates['start'], dates['end']
    if start and end and start > end:
        start, end = end, start
    return start, end, notices


def _preset(label, p_start, p_end, start, end, extra=None):
    """A quick-filter link for the dates `p_start`–`p_end` (None = all dates), keeping `extra` params."""
    params = dict(extra or {})
    if p_start:
        params.update(start=p_start.isoformat(), end=p_end.isoformat())
    return {'label': label, 'query': urlencode(params), 'active': (p_start, p_end) == (start, end)}


def _presets(start, end, extra=None):
    today = dt.date.today()
    ranges = [
        ('All dates', None, None),
        ('Today', today, today),
        ('Last 7 days', today - dt.timedelta(days=6), today),
        ('This month', today.replace(day=1), today),
    ]
    return [_preset(label, p_start, p_end, start, end, extra) for label, p_start, p_end in ranges]


def _month_filters(trip_dates, start, end, picked='', extra=None, keep=None):
    """The year dropdown and the month quick filters (the 1st to the last day) of one year.

    The dropdown offers the first year with trips to the last, always including the current year,
    newest first. The year shown is the one picked in the dropdown (?year=), else the year of the
    selected dates when they fall in one year (so a chosen month stays in view), else the current
    year. Its months run to December, or to the current month for the current year (to the last
    month with trips, if that is later). Each month carries its number of trip legs so the page
    can dim the months without trips. Picking a year keeps the current filters (`keep`).
    """
    today = dt.date.today()
    legs = Counter((d.year, d.month) for d in trip_dates)
    last_year, last_month = max([(today.year, today.month), *legs])
    years = list(range(last_year, min([today.year, *(y for y, _ in legs)]) - 1, -1))
    try:
        year = int(picked)
    except ValueError:  # blank or not a number
        year = None
    if year not in years:
        same_year = start and end and start.year == end.year and start.year in years
        year = start.year if same_year else today.year

    months = []
    for month in range(1, (last_month if year == last_year else 12) + 1):
        first_day = dt.date(year, month, 1)
        last_day = first_day.replace(day=calendar.monthrange(year, month)[1])
        preset = _preset(f'{first_day:%b}', first_day, last_day, start, end, extra)
        preset.update(name=f'{first_day:%B %Y}', legs=legs[year, month])
        months.append(preset)
    whole_year = _preset('Whole year', dt.date(year, 1, 1), dt.date(year, 12, 31), start, end, extra)
    whole_year.update(name=f'January 1 – December 31, {year}', legs=sum(m['legs'] for m in months))
    return {
        'year': year,
        'years': [{'year': y, 'selected': y == year, 'query': urlencode({**(keep or {}), 'year': y})}
                  for y in years],
        'whole_year': whole_year,
        'months': months,
    }


def _filter_params(start, end, q='', sort=DEFAULT_SORT):
    params = {}
    if start:
        params['start'] = start.isoformat()
    if end:
        params['end'] = end.isoformat()
    if q:
        params['q'] = q
    if sort != DEFAULT_SORT:
        params['sort'] = sort
    return params


def _filter_query(start, end, q='', sort=DEFAULT_SORT):
    return urlencode(_filter_params(start, end, q, sort))


def _sort(request):
    """?sort=..., else the last choice (cookie), else name."""
    sort = request.GET.get('sort') or request.COOKIES.get(SORT_COOKIE, '')
    return sort if sort in SORTS else DEFAULT_SORT


def _matching(log, q, sort=DEFAULT_SORT):
    """Drivers whose name contains `q`, in the chosen order (the list page and the "Export all" sheets)."""
    drivers = [d for d in log.drivers if q.casefold() in d.name.casefold()]
    if sort == 'added':
        # The workbook keeps no creation date. log.drivers is in the order drivers were added: Lists
        # sheet rows top to bottom (imports add new drivers to the next free row), then names that
        # are only in the Activity Log, by first row. Newest first is that order reversed.
        return drivers[::-1]
    return sorted(drivers, key=lambda d: d.name.casefold())


def _period(log, start, end):
    """The days that "days traveled / not traveled" count: the dates chosen, else the log's first
    to last day (any driver, so drivers are compared over the same days), never past today."""
    first, last = start or log.first_date, end or log.last_date
    if first is None or last is None:
        return None, None
    last = min(last, dt.date.today())
    return (first, last) if first <= last else (None, None)


def _days_not_traveled(driver, first, last):
    """Dates from `first` to `last` on which the driver made no trip (parking-only days included)."""
    if first is None:
        return []
    every_day = (first + dt.timedelta(days=i) for i in range((last - first).days + 1))
    return [day for day in every_day if day not in driver.travel_dates]


def _idle_calendar(first, last, not_traveled, card_dates=()):
    """The months from `first` to `last`, each with all its days (1 to 28–31): state 'idle' for a
    day not traveled, 'traveled' for the other days of the period, 'out' for the days outside it.
    `card` marks the days that have a day card on the page to jump to.

    Only the last CALENDAR_MONTHS months are built, so a far-off ?start= (year 1, say) cannot
    make the page list every day since then; the page says when months were left out.
    """
    if first is None:
        return []
    idle, cards, months = set(not_traveled), set(card_dates), []
    earliest = max(last.year * 12 + last.month - CALENDAR_MONTHS, 12)  # months since year 0; 12 = Jan of year 1
    month = max(first.replace(day=1), dt.date(earliest // 12, earliest % 12 + 1, 1))
    while month <= last:
        days = []
        for n in range(1, calendar.monthrange(month.year, month.month)[1] + 1):
            date = month.replace(day=n)
            state = 'idle' if date in idle else 'traveled' if first <= date <= last else 'out'
            days.append({'date': date, 'state': state, 'card': date in cards})
        months.append({
            'month': month,
            'days': days,
            'idle': sum(d['state'] == 'idle' for d in days),
            'period_days': sum(d['state'] != 'out' for d in days),
        })
        month = (month + dt.timedelta(days=31)).replace(day=1)
    return months


def _timeline(record, not_traveled):
    """The record's days oldest first (day 1 at the top, the 31st at the bottom), with a "no travel"
    item for each run of dates in between that has no day of its own (a parking-only day keeps its
    day card)."""
    if not record.days:
        return []
    by_date, idle = {day.date: day for day in record.days}, set(not_traveled)
    newest = max([record.days[-1].date, *idle])
    oldest = min([record.days[0].date, *idle])
    items, gap = [], []

    def close_gap():
        if gap:
            items.append({'gap': True, 'first': gap[0], 'last': gap[-1], 'days': len(gap)})
            gap.clear()

    for offset in range((newest - oldest).days + 1):
        date = oldest + dt.timedelta(days=offset)
        if date in by_date:
            close_gap()
            items.append({'day': by_date[date]})
        elif date in idle:
            gap.append(date)
        else:  # e.g. the arrival day of a trip past midnight
            close_gap()
    close_gap()
    return items


def drivers_activity(request):
    start, end, notices = _date_range(request)
    q = request.GET.get('q', '').strip()
    sort = _sort(request)
    context = {
        'start': start,
        'end': end,
        'q': q,
        'sort': sort,
        'sorts': SORTS.items(),
        'notices': notices,
        'presets': _presets(start, end, {'q': q} if q else None),
        'filter_query': _filter_query(start, end),
        # The export follows the search and the sort too: it holds the drivers shown on the page.
        'export_query': _filter_query(start, end, q, sort),
    }
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        context['error'] = str(exc)
        return render(request, 'activity/drivers_Activity.html', context, status=503)

    first, last = _period(log, start, end)
    period_days = (last - first).days + 1 if first else None
    # (driver limited to the selected dates, trips on any date, days traveled, days not traveled)
    rows = []
    for d in _matching(log, q, sort):
        not_traveled = len(_days_not_traveled(d, first, last)) if first else None
        traveled = period_days - not_traveled if first else None
        rows.append((d.between(start, end), len(d.trips), traveled, not_traveled))
    drivers = [row[0] for row in rows]
    context.update(
        log=log,
        months=_month_filters([leg.date for d in log.drivers for leg in d.trips], start, end,
                              request.GET.get('year', ''), {'q': q} if q else None,
                              _filter_params(start, end, q, sort)),
        rows=rows,
        period_first=first,
        period_last=last,
        drivers=drivers,
        total_drivers=len(log.drivers),
        active_drivers=sum(1 for d in drivers if d.trips),
        can_export=any(d.legs for d in drivers),  # parking-only rows count: they are exported too
        total_legs=sum(len(d.trips) for d in drivers),
        total_long_stays=sum(d.long_stays for d in drivers),
    )
    response = render(request, 'activity/drivers_Activity.html', context)
    if request.GET.get('sort') in SORTS:
        response.set_cookie(SORT_COOKIE, sort, max_age=365 * 24 * 3600, samesite='Lax')
    return response


def driver_travel_record(request, name):
    start, end, notices = _date_range(request)
    context = {
        'start': start,
        'end': end,
        'notices': notices,
        'presets': _presets(start, end),
        'filter_query': _filter_query(start, end),
        'name': name,
    }
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        context['error'] = str(exc)
        return render(request, 'activity/Driver_travel_record.html', context, status=503)

    driver = log.find_driver(name)
    if driver is None:
        raise Http404(f'No driver named "{name}" in {log.path.name}.')
    record = driver.between(start, end)
    first, last = _period(log, start, end)
    not_traveled = _days_not_traveled(driver, first, last)
    context.update(
        log=log,
        months=_month_filters([leg.date for leg in driver.trips], start, end,
                              request.GET.get('year', ''), keep=_filter_params(start, end)),
        driver=record,
        days=record.days,  # oldest first, like the timeline
        timeline=_timeline(record, not_traveled),
        period_first=first,
        period_last=last,
        period_days=(last - first).days + 1 if first else 0,
        days_traveled=(last - first).days + 1 - len(not_traveled) if first else 0,
        not_traveled=excel_log.date_ranges(not_traveled),
        idle_calendar=_idle_calendar(first, last, not_traveled, [day.date for day in record.days]),
        days_not_traveled=len(not_traveled),
        all_legs=len(driver.trips),
        vehicles=[log.vehicle(plate) for plate in record.plates],
    )
    return render(request, 'activity/Driver_travel_record.html', context)



def export_travel_record(request, name):
    """The driver's travel record for the selected dates, downloaded as an Excel workbook."""
    start, end, _ = _date_range(request)
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        messages.error(request, f'Cannot export: {exc}')
        return redirect('activity:drivers_activity')
    driver = log.find_driver(name)
    if driver is None:
        raise Http404(f'No driver named "{name}" in {log.path.name}.')
    record = driver.between(start, end)
    response = HttpResponse(export.travel_record_workbook(log, record, start, end), content_type=XLSX)
    response['Content-Disposition'] = content_disposition_header(
        as_attachment=True, filename=export.export_filename(record.name, start, end, record.days))
    return response


def export_all_travel_records(request):
    """Every driver on the list page with trips in the selected dates: one sheet per driver."""
    start, end, _ = _date_range(request)
    q = request.GET.get('q', '').strip()
    sort = _sort(request)
    query = _filter_query(start, end, q, sort)
    back = reverse('activity:drivers_activity') + (f'?{query}' if query else '')
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        messages.error(request, f'Cannot export: {exc}')
        return redirect(back)
    # Drivers without trips in these dates are left out rather than given empty sheets.
    records = [r for r in (d.between(start, end) for d in _matching(log, q, sort)) if r.legs]
    if not records:
        messages.info(request, 'Nothing to export: no driver has trips on these dates.')
        return redirect(back)
    response = HttpResponse(export.all_drivers_workbook(log, records, start, end), content_type=XLSX)
    response['Content-Disposition'] = content_disposition_header(
        as_attachment=True, filename=export.all_drivers_filename(records, start, end, q))
    return response

@require_POST
def import_travel_report(request):
    """Copy uploaded SinoTrack Travel Report and Park Report CSVs into the Activity Log sheet.

    Several files can be uploaded at once, in any order and mixing both kinds. The trips of all
    Travel Reports are imported together first, then the stops of all Park Reports are merged
    together, so stops land on trips from the same upload and a stop that appears in two
    overlapping exports is counted once.
    """
    back = redirect('activity:drivers_activity')
    uploads = request.FILES.getlist('report')
    if not uploads:
        messages.error(request, 'Choose one or more SinoTrack Travel Report or Park Report CSV files to import.')
        return back
    driver = request.POST.get('driver', '').strip() or None
    plate = request.POST.get('plate', '').strip() or None

    found = {'travel': ([], {}), 'park': ([], {})}  # kind -> (file names, {(device, start, end): item})
    for upload in uploads:
        if upload.size > MAX_UPLOAD_BYTES:
            messages.error(request, f'{upload.name} is too large to be a SinoTrack report (over 10 MB); '
                                    'it was not imported.')
            continue
        data = upload.read()
        try:
            kind = sinotrack.detect_report(data)
            if kind == 'park':
                items, invalid_rows = sinotrack.parse_park_report(data)
            elif kind == 'travel':
                items, invalid_rows = sinotrack.parse_travel_report(data)
            else:
                raise sinotrack.ReportError('This is not a SinoTrack Travel Report or Park Report CSV.')
        except sinotrack.ReportError as exc:
            messages.error(request, f'{upload.name}: {exc}' + (' It was not imported.' if len(uploads) > 1 else ''))
            continue
        if invalid_rows:
            messages.warning(request, f'{len(invalid_rows)} row(s) in {upload.name} could not be read '
                                      f'and were skipped (CSV rows {", ".join(map(str, invalid_rows[:20]))}).')
        names, items_by_key = found[kind]
        names.append(upload.name)
        for item in items:
            items_by_key.setdefault((item.device, item.start, item.end), item)

    changed = []
    for kind, write, report in (
            ('travel', lambda items: log_writer.import_trips(
                settings.ACTIVITY_LOG_PATH, items, driver=driver, plate=plate), _report_trip_import),
            ('park', lambda items: log_writer.merge_parks(
                settings.ACTIVITY_LOG_PATH, items, plate=plate, driver=driver), _report_park_merge)):
        names, items_by_key = found[kind]
        if not items_by_key:
            continue
        files = _files_text(names)
        try:
            result = write(sorted(items_by_key.values(), key=lambda item: (item.device, item.start)))
        except excel_log.ActivityLogError as exc:
            messages.error(request, f'{files}: {exc}')
            break  # e.g. the workbook is locked: the Park Reports would fail the same way
        if report(request, files, result):
            changed.append(result)
        for warning in result.warnings:
            messages.warning(request, warning)

    # No date filter here: limiting the pages to the uploaded files' dates hid every other trip.
    drivers = {name for result in changed for name in result.drivers}
    if len(drivers) == 1:
        return redirect('activity:driver_travel_record', name=drivers.pop())
    return back


def _files_text(names):
    if len(names) == 1:
        return names[0]
    shown = ', '.join(names[:3]) + (f' and {len(names) - 3} more' if len(names) > 3 else '')
    return f'{len(names)} files ({shown})'


def _dates_text(start, end):
    fmt = lambda d: f'{d:%b} {d.day}, {d:%Y}'
    return fmt(start) if start == end else f'{fmt(start)} – {fmt(end)}'


def _rows_text(rows):
    first, last = min(rows), max(rows)
    return f'row {first}' if first == last else f'rows {first}–{last}'


def _saved_text(result):
    text = ''
    if result.via_excel:
        text += ' The workbook was open, so the cells were typed into it in Excel and saved there.'
    return text + f' Backup of the previous workbook: {result.backup.name}.'


def _report_trip_import(request, name, result):
    if not (result.added or result.completed_rows):
        messages.info(request, f'Nothing new in {name}: all {result.duplicates} trips are '
                               'already in the Activity Log.')
        return False
    parts = []
    if result.added:
        parts.append(f'Copied {result.added} trip{"s" * (result.added != 1)} from {name} into the '
                     f'{excel_log.LOG_SHEET} sheet ({_rows_text([result.first_row, result.last_row])}) for '
                     f'{", ".join(result.drivers)} ({", ".join(result.plates)}), '
                     f'{_dates_text(result.start, result.end)}.')
    if result.completed_rows:
        count = len(result.completed_rows)
        parts.append(f'Filled in {", ".join(result.filled_columns)} for {count} trip{"s" * (count != 1)} '
                     f'already in the log ({_rows_text(result.completed_rows)}).')
    unchanged = result.duplicates - len(result.completed_rows)
    if unchanged:
        parts.append(f'{unchanged} trip{"s" * (unchanged != 1)} already in the log needed nothing.')
    if result.added_to_lists:
        parts.append(f'Added to the Lists sheet: {", ".join(result.added_to_lists)}.')
    messages.success(request, ' '.join(parts) + _saved_text(result))
    return True


def _plural(count, word):
    return f'{count} {word}{"s" * (count != 1)}'


def _report_park_merge(request, name, result):
    parts = []
    if result.matched:
        parts.append(f'Merged {result.matched} of {result.parks} stops from {name} into '
                     f'{_plural(len(result.rows), "trip row")} of the {excel_log.LOG_SHEET} sheet '
                     f'({_rows_text(result.rows)}): Park Count, Park Time, Park Address and Park Coordinates.')
    if result.parking_only:
        days = ', '.join(f'{d:%b} {d.day}' for d in result.parking_only_dates)
        rows = result.parking_only_rows
        on_days = 'a day' if len(result.parking_only_dates) == 1 else 'days'
        parts.append(f'{_plural(result.parking_only, "stop")} on {on_days} with no trip in the log ({days}) '
                     f'{"was" if result.parking_only == 1 else "were"} copied into '
                     f'{_plural(len(rows), "parking-only row")} ({_rows_text(rows)}), with no depart or arrive '
                     'time. If you import the Travel Report for those days later, import this Park Report '
                     'again to move the stops onto the trips.')
    if result.cleared_rows:
        parts.append(f'Cleared {_plural(len(result.cleared_rows), "parking-only row")} '
                     f'({_rows_text(result.cleared_rows)}) whose stops now belong to trips.')
    messages.success(request, ' '.join(parts) + _saved_text(result))
    return True
