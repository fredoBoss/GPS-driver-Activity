from django import template

register = template.Library()

EMPTY = '—'


@register.filter
def dhms(value):
    """timedelta -> "DD-HH-MM-SS" (days-hours-minutes-seconds), used for every duration on the pages."""
    if value is None or value == '':
        return EMPTY
    minutes, seconds = divmod(round(value.total_seconds()), 60)
    hours, minutes = divmod(minutes, 60)
    days, hours = divmod(hours, 24)
    return f'{days:02d}-{hours:02d}-{minutes:02d}-{seconds:02d}'


@register.filter
def clock(value):
    """Offset from midnight (timedelta) -> "8:05 AM", marking times that fall on a later day."""
    if value is None or value == '':
        return EMPTY
    days, minutes = divmod(round(value.total_seconds() / 60), 24 * 60)
    hour, minute = divmod(minutes, 60)
    text = f'{hour % 12 or 12}:{minute:02d} {"AM" if hour < 12 else "PM"}'
    return f'{text} (+{days}d)' if days else text


@register.filter
def number(value, decimals=1):
    """Numbers with thousands separators; whole numbers drop the decimals."""
    if value is None or value == '':
        return EMPTY
    if isinstance(value, str):
        return value
    if float(value).is_integer():
        return f'{int(value):,}'
    return f'{value:,.{int(decimals)}f}'
