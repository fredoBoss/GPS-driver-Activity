import datetime as dt
from urllib.parse import urlencode

from django.conf import settings
from django.http import Http404
from django.shortcuts import render

from . import excel_log


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


def _filter_query(start, end):
    params = {}
    if start:
        params['start'] = start.isoformat()
    if end:
        params['end'] = end.isoformat()
    return urlencode(params)


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
    }
    try:
        log = excel_log.load(settings.ACTIVITY_LOG_PATH)
    except excel_log.ActivityLogError as exc:
        context['error'] = str(exc)
        return render(request, 'activity/drivers_Activity.html', context, status=503)

    drivers = [d.between(start, end) for d in log.drivers if q.casefold() in d.name.casefold()]
    distances = [d.distance for d in drivers if d.distance is not None]
    context.update(
        log=log,
        drivers=drivers,
        total_drivers=len(log.drivers),
        active_drivers=sum(1 for d in drivers if d.legs),
        total_legs=sum(len(d.legs) for d in drivers),
        total_long_stays=sum(d.long_stays for d in drivers),
        total_distance=sum(distances) if distances else None,
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
        vehicles=[log.vehicle(plate) for plate in record.plates],
    )
    return render(request, 'activity/Driver_travel_record.html', context)
