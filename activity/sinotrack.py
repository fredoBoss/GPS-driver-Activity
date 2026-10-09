"""Parse SinoTrack report exports (CSV): the Travel Report and the Park Report.

* Travel Report: one row per drive between two stops (start/end time, addresses, coordinates).
* Park Report: one row per stop (start/end time, park time, address, longitude, latitude).

SinoTrack wraps every cell as "\\t<value>" (so Excel does not reformat it), pads rows with
empty columns, ends addresses with " ." and appends hundreds of blank rows. All of that
is cleaned here; the values themselves are kept as exported.
"""
import csv
import datetime as dt
import io
import re
from dataclasses import dataclass

TRAVEL_HEADERS = {
    'device': ('device name',),
    'start': ('start time',),
    'end': ('end time',),
    'start_address': ('start address',),
    'end_address': ('end address',),
    'start_longitude': ('start longitude',),
    'start_latitude': ('start latitude',),
    'end_longitude': ('end longitude',),
    'end_latitude': ('end latitude',),
}
PARK_HEADERS = {
    'device': ('device name',),
    'start': ('start time',),
    'end': ('end time',),
    'park_time': ('park time',),
    'address': ('address',),
    'longitude': ('longitude',),
    'latitude': ('latitude',),
}
REQUIRED = ('device', 'start', 'end')
DATETIME_FORMATS = ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y/%m/%d %H:%M:%S', '%Y/%m/%d %H:%M')

# "T05_RJP 162 - Gonzaga" -> unit "T05", plate "RJP 162", driver "Gonzaga"
DEVICE_NAME = re.compile(r'^(?P<unit>[^_\s]+)_(?P<plate>.+?)\s+-\s+(?P<driver>.+)$')
# Some devices are named without the " - ": "T06_RLR 795 Capoy" (plate = letters, then digits).
DEVICE_NAME_NO_DASH = re.compile(r'^(?P<unit>[^_\s]+)_(?P<plate>[A-Za-z]{1,4}\s?\d{2,5})\s+(?P<driver>\S.*)$')


class ReportError(Exception):
    """The file is not a SinoTrack report this importer understands."""


class _Device:
    @property
    def unit(self):
        return parse_device_name(self.device)[0]

    @property
    def plate(self):
        return parse_device_name(self.device)[1]

    @property
    def driver(self):
        return parse_device_name(self.device)[2]


@dataclass
class Trip(_Device):
    """The Travel Report columns the Activity Log uses (mileage and speeds are left out)."""
    row: int
    device: str
    start: dt.datetime
    end: dt.datetime
    start_address: str
    end_address: str
    start_latitude: float | None = None
    start_longitude: float | None = None
    end_latitude: float | None = None
    end_longitude: float | None = None

    @property
    def start_coordinates(self):
        return lat_lon(self.start_latitude, self.start_longitude)

    @property
    def end_coordinates(self):
        return lat_lon(self.end_latitude, self.end_longitude)


@dataclass
class Park(_Device):
    row: int
    device: str
    start: dt.datetime
    end: dt.datetime
    address: str
    latitude: float | None
    longitude: float | None

    @property
    def duration(self):
        return self.end - self.start

    @property
    def coordinates(self):
        return lat_lon(self.latitude, self.longitude)


def lat_lon(latitude, longitude):
    """"lat, lon" as exported (Google Maps order), or '' when the export has none."""
    if latitude is None or longitude is None:
        return ''
    return f'{_coord(latitude)}, {_coord(longitude)}'


def parse_device_name(name):
    """Split a SinoTrack device name into (unit, plate, driver); unparsed names become the plate."""
    match = DEVICE_NAME.match(name) or DEVICE_NAME_NO_DASH.match(name)
    if not match:
        return '', name, ''
    return match['unit'], match['plate'].strip(), match['driver'].strip()


def detect_report(data):
    """'travel', 'park', or None, judged from the CSV header row."""
    header = _header(_text(data))
    if _columns(header, PARK_HEADERS).keys() >= {'start', 'end', 'park_time'}:
        return 'park'
    if _columns(header, TRAVEL_HEADERS).keys() >= {'start', 'end', 'start_address'}:
        return 'travel'
    return None


def parse_travel_report(data):
    """Return (trips sorted by device and start time, CSV row numbers that were not valid trips)."""
    trips, invalid = [], []
    for row_no, cells, start, end in _rows(data, TRAVEL_HEADERS, 'Travel Report', invalid):
        trips.append(Trip(
            row=row_no,
            device=cells['device'],
            start=start,
            end=end,
            start_address=_address(cells.get('start_address', '')),
            end_address=_address(cells.get('end_address', '')),
            start_latitude=_to_float(cells.get('start_latitude')),
            start_longitude=_to_float(cells.get('start_longitude')),
            end_latitude=_to_float(cells.get('end_latitude')),
            end_longitude=_to_float(cells.get('end_longitude')),
        ))
    if not trips:
        raise ReportError('No trips found in this file.')
    trips.sort(key=lambda t: (t.device, t.start))
    return trips, invalid


def parse_park_report(data):
    """Return (parks sorted by device and start time, CSV row numbers that were not valid parks)."""
    parks, invalid = [], []
    for row_no, cells, start, end in _rows(data, PARK_HEADERS, 'Park Report', invalid):
        parks.append(Park(
            row=row_no,
            device=cells['device'],
            start=start,
            end=end,
            address=_address(cells.get('address', '')),
            latitude=_to_float(cells.get('latitude')),
            longitude=_to_float(cells.get('longitude')),
        ))
    if not parks:
        raise ReportError('No parks found in this file.')
    parks.sort(key=lambda p: (p.device, p.start))
    return parks, invalid


def _rows(data, wanted, report_name, invalid):
    """Yield (CSV row number, cleaned cells, start, end) for every non-blank data row."""
    rows = csv.reader(io.StringIO(_text(data)))
    cols = _columns([_clean(h).casefold() for h in next(rows, [])], wanted)
    missing = [wanted[name][0] for name in REQUIRED if name not in cols]
    if missing:
        raise ReportError(
            f'This does not look like a SinoTrack {report_name} CSV: no '
            + ', '.join(f'"{m.title()}"' for m in missing) + ' column.')
    for row_no, values in enumerate(rows, start=2):
        cells = {name: _clean(values[i]) if i < len(values) else '' for name, i in cols.items()}
        if not any(cells.values()):
            continue
        start, end = _to_datetime(cells['start']), _to_datetime(cells['end'])
        if not cells['device'] or start is None or end is None:
            invalid.append(row_no)
            continue
        yield row_no, cells, start, end


def _text(data):
    if isinstance(data, bytes):
        try:
            return data.decode('utf-8-sig')
        except UnicodeDecodeError:
            return data.decode('cp1252', errors='replace')
    return data


def _header(text):
    return [_clean(h).casefold() for h in next(csv.reader(io.StringIO(text)), [])]


def _columns(header, wanted):
    """Map field -> column index. "Longitude(°)" matches "longitude"; units in brackets are ignored."""
    cols = {}
    for name, aliases in wanted.items():
        for i, h in enumerate(header):
            if any(h == a or h.startswith(a + '(') or h.startswith(a + ' (') for a in aliases):
                cols[name] = i
                break
    return cols


def _clean(value):
    return ' '.join(str(value).split())


def _address(value):
    return re.sub(r'\s+\.$', '', value)


def _coord(value):
    return f'{value:.7f}'.rstrip('0').rstrip('.')


def _to_float(value):
    try:
        return float(value.replace(',', '')) if value else None
    except ValueError:
        return None


def _to_datetime(value):
    for fmt in DATETIME_FORMATS:
        try:
            return dt.datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None
