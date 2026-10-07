from django import template

register = template.Library()

EMPTY = '—'


@register.filter
def duration(value):
    """timedelta -> "h:mm" (hours may pass 24)."""
    if value is None or value == '':
        return EMPTY
    minutes = round(value.total_seconds() / 60)
    return f'{minutes // 60}:{minutes % 60:02d}'


@register.filter
def hms(value):
    """timedelta -> "hh:mm:ss" (hours may pass 24), used for stays and the long-stay rule."""
    if value is None or value == '':
        return EMPTY
    minutes, seconds = divmod(round(value.total_seconds()), 60)
    hours, minutes = divmod(minutes, 60)
    return f'{hours:02d}:{minutes:02d}:{seconds:02d}'


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
