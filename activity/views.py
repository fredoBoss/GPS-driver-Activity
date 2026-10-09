import datetime as dt
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


def _presets(start, end, extra=None):
    today = dt.date.today()
    ranges = [
        ('All dates', None, None),
        ('Today', today, today),
        ('Last 7 days', today - dt.timedelta(days=6), today),
        ('This month', today.replace(day=1), today),
    ]
    presets = []
    for label, p_start, p_end in ranges:
        params = dict(extra or {})
        if p_start:
            params.update(start=p_start.isoformat(), end=p_end.isoformat())
        presets.append({
            'label': label,
            'query': urlencode(params),
            'active': (p_start, p_end) == (start, end),
        })
    return presets


def _filter_query(start, end, q=''):
    params = {}
    if start:
        params['start'] = start.isoformat()
    if end:
        params['end'] = end.isoformat()
    if q:
        params['q'] = q
    return urlencode(params)


def _matching(log, q):
    return [d for d in log.drivers if q.casefold() in d.name.casefold()]


def drivers_activity(request):
    start, end, notices = _date_range(request)
    q = request.GET.get('q', '').strip()
    context = {
        'start': start,
        'end': end,
        'q': q,
        'notices': notices,
        'presets': _presets(start, end, {'q': q} if q else None),
        'filter_query': _filter_query(start, end),
        # The export follows the search too: it holds the drivers shown on the page.
        'export_query': _filter_query(start, end, q),
    }
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        context['error'] = str(exc)
        return render(request, 'activity/drivers_Activity.html', context, status=503)

    # (driver limited to the selected dates, number of trips on any date)
    rows = [(d.between(start, end), len(d.trips)) for d in _matching(log, q)]
    drivers = [driver for driver, _ in rows]
    context.update(
        log=log,
        rows=rows,
        drivers=drivers,
        total_drivers=len(log.drivers),
        active_drivers=sum(1 for d in drivers if d.trips),
        can_export=any(d.legs for d in drivers),  # parking-only rows count: they are exported too
        total_legs=sum(len(d.trips) for d in drivers),
        total_long_stays=sum(d.long_stays for d in drivers),
    )
    return render(request, 'activity/drivers_Activity.html', context)


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
    context.update(
        log=log,
        driver=record,
        days=list(reversed(record.days)),
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
    query = _filter_query(start, end, q)
    back = reverse('activity:drivers_activity') + (f'?{query}' if query else '')
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        messages.error(request, f'Cannot export: {exc}')
        return redirect(back)
    # Drivers without trips in these dates are left out rather than given empty sheets.
    records = [r for r in (d.between(start, end) for d in _matching(log, q)) if r.legs]
    if not records:
        messages.info(request, 'Nothing to export: no driver has trips on these dates.')
        return redirect(back)
    response = HttpResponse(export.all_drivers_workbook(log, records, start, end), content_type=XLSX)
    response['Content-Disposition'] = content_disposition_header(
        as_attachment=True, filename=export.all_drivers_filename(records, start, end, q))
    return response

@require_POST
def import_travel_report(request):
    """Copy an uploaded SinoTrack Travel Report or Park Report CSV into the Activity Log sheet."""
    back = redirect('activity:drivers_activity')
    upload = request.FILES.get('report')
    if upload is None:
        messages.error(request, 'Choose a SinoTrack Travel Report or Park Report CSV file to import.')
        return back
    if upload.size > MAX_UPLOAD_BYTES:
        messages.error(request, f'{upload.name} is too large to be a SinoTrack report (over 10 MB).')
        return back

    data = upload.read()
    driver = request.POST.get('driver', '').strip() or None
    plate = request.POST.get('plate', '').strip() or None
    try:
        kind = sinotrack.detect_report(data)
        if kind == 'park':
            parks, invalid_rows = sinotrack.parse_park_report(data)
            result = log_writer.merge_parks(settings.ACTIVITY_LOG_PATH, parks, plate=plate, driver=driver)
            changed = _report_park_merge(request, upload.name, result)
        elif kind == 'travel':
            trips, invalid_rows = sinotrack.parse_travel_report(data)
            result = log_writer.import_trips(settings.ACTIVITY_LOG_PATH, trips, driver=driver, plate=plate)
            changed = _report_trip_import(request, upload.name, result)
        else:
            raise sinotrack.ReportError('This is not a SinoTrack Travel Report or Park Report CSV.')
    except (sinotrack.ReportError, excel_log.ActivityLogError) as exc:
        messages.error(request, f'{upload.name}: {exc}')
        return back

    if invalid_rows:
        messages.warning(request, f'{len(invalid_rows)} row(s) in {upload.name} could not be read '
                                  f'and were skipped (CSV rows {", ".join(map(str, invalid_rows[:20]))}).')
    for warning in result.warnings:
        messages.warning(request, warning)

    if not changed:
        return back
    # No date filter here: limiting the pages to the uploaded file's dates hid every other trip.
    if len(result.drivers) == 1:
        return redirect('activity:driver_travel_record', name=result.drivers[0])
    return back


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
