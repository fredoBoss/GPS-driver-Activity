import datetime as dt
import shutil
import tempfile
from pathlib import Path
from unittest import skipUnless

from django.conf import settings
from django.test import SimpleTestCase, override_settings
from django.urls import reverse
from openpyxl import Workbook

from . import excel_log

D1, D2 = dt.date(2026, 10, 1), dt.date(2026, 10, 2)
LOG_HEADER = ['Date', 'Driver Name', 'Plate No.', 'From (Origin)', 'Depart Time',
              'To (Destination)', 'Arrive Time', 'Park Count', 'Distance (km)', 'Remarks / DR No.']


def T(h, m):
    return dt.time(h, m)


def minutes(td):
    return None if td is None else round(td.total_seconds() / 60)


def make_workbook(path, rows, drivers=(), vehicles=(), threshold=dt.time(0, 30), header=LOG_HEADER):
    """Write a workbook laid out like Driver_Activity_Log.xlsx (headers on row 4 / row 6)."""
    wb = Workbook()
    log = wb.active
    log.title = 'Activity Log'
    log['A1'] = 'Delivery Driver Activity Log'
    for c, title in enumerate(header, start=1):
        log.cell(row=4, column=c, value=title)
    for r, values in enumerate(rows, start=5):
        for c, value in enumerate(values, start=1):
            log.cell(row=r, column=c, value=value)
    lists = wb.create_sheet('Lists')
    lists['A3'] = 'Long-stay threshold (h:mm)'
    lists['B3'] = threshold
    for col, title in zip('ABDEF', ['Driver Name', 'Contact No.', 'Plate No.',
                                    'Vehicle Description', 'SinoTrack Device ID']):
        lists[f'{col}6'] = title
    for i, (name, contact) in enumerate(drivers, start=7):
        lists[f'A{i}'], lists[f'B{i}'] = name, contact
    for i, (plate, desc, device) in enumerate(vehicles, start=7):
        lists[f'D{i}'], lists[f'E{i}'], lists[f'F{i}'] = plate, desc, device
    wb.save(path)


MARK_DAY = [
    (D1, 'Mark Reyes', 'ABC 1234', 'Warehouse', T(8, 5), 'Customer A', T(8, 50), 2, 18.4, 'DR 1021'),
    (D1, 'Mark Reyes', 'ABC 1234', 'Customer A', T(9, 10), 'Customer B', T(9, 45), 1, 12.1, 'DR 1022'),
    (D1, 'Mark Reyes', 'ABC 1234', 'Customer B', T(11, 0), 'Warehouse', T(11, 50), 0, 20.6, None),
]


class WorkbookTestCase(SimpleTestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.tmp, ignore_errors=True)

    def build(self, rows, **kwargs):
        path = self.tmp / f'log{len(list(self.tmp.iterdir()))}.xlsx'
        make_workbook(path, rows, **kwargs)
        return path

    def load(self, rows, **kwargs):
        return excel_log.load(self.build(rows, **kwargs))


class ExcelLogTests(WorkbookTestCase):
    def test_travel_stay_and_long_stay_match_the_workbook_example(self):
        log = self.load(MARK_DAY, drivers=[('Mark Reyes', '0917 000 0000')])
        legs = log.find_driver('mark reyes').legs
        self.assertEqual([minutes(leg.travel) for leg in legs], [45, 35, 50])
        self.assertEqual([minutes(leg.stay) for leg in legs], [20, 75, None])
        self.assertEqual([leg.long_stay for leg in legs], [False, True, False])

    def test_daily_summary(self):
        day = self.load(MARK_DAY).find_driver('Mark Reyes').days[0]
        self.assertEqual(minutes(day.first_depart), 8 * 60 + 5)
        self.assertEqual(minutes(day.last_arrive), 11 * 60 + 50)
        self.assertEqual(minutes(day.time_on_road), 3 * 60 + 45)
        self.assertEqual(minutes(day.travel), 2 * 60 + 10)
        self.assertEqual(minutes(day.stay), 60 + 35)
        self.assertEqual(day.long_stays, 1)
        self.assertAlmostEqual(day.distance, 51.1)
        self.assertEqual(day.park_count, 3)

    def test_check_order_and_leg_past_midnight(self):
        log = self.load([
            (D1, 'Ana Cruz', 'XYZ 5678', 'Warehouse', T(13, 0), 'Customer C', T(13, 40)),
            (D1, 'Ana Cruz', 'XYZ 5678', 'Customer C', T(13, 20), 'Warehouse', T(14, 0)),
            (D1, 'Ana Cruz', 'XYZ 5678', 'Warehouse', T(23, 30), 'Customer D', T(0, 20)),
        ])
        legs = log.find_driver('Ana Cruz').legs
        self.assertTrue(legs[0].check_order)
        self.assertIsNone(legs[0].stay)
        self.assertEqual(minutes(legs[2].travel), 50)
        day = log.find_driver('Ana Cruz').days[0]
        self.assertEqual(minutes(day.last_arrive), 24 * 60 + 20)
        self.assertEqual(minutes(day.time_on_road), 11 * 60 + 20)

    def test_stay_needs_next_row_with_same_date_driver_and_plate(self):
        log = self.load([
            (D1, 'Mark Reyes', 'ABC 1234', 'Warehouse', T(8, 0), 'Customer A', T(8, 30)),
            (D1, 'Ana Cruz', 'XYZ 5678', 'Warehouse', T(9, 0), 'Customer C', T(9, 30)),
            (D1, 'Ana Cruz', 'DEF 9012', 'Customer C', T(10, 0), 'Warehouse', T(10, 30)),
            (D2, 'Ana Cruz', 'DEF 9012', 'Warehouse', T(11, 0), 'Customer E', T(11, 30)),
        ])
        self.assertEqual([leg.stay for d in log.drivers for leg in d.legs], [None] * 4)

    def test_drivers_from_lists_and_unlisted_names_from_the_log(self):
        log = self.load(MARK_DAY, drivers=[('Leo Tan', ''), ('Mark Reyes', '')])
        self.assertEqual([(d.name, d.in_lists, len(d.legs)) for d in log.drivers],
                         [('Leo Tan', True, 0), ('Mark Reyes', True, 3)])
        log = self.load(MARK_DAY, drivers=[('Leo Tan', '')])
        self.assertEqual([(d.name, d.in_lists) for d in log.drivers],
                         [('Leo Tan', True), ('Mark Reyes', False)])

    def test_rows_without_date_or_driver_are_reported(self):
        log = self.load([MARK_DAY[0], (None, 'Mark Reyes', 'ABC 1234', 'X', T(9, 0)), (), MARK_DAY[1]])
        self.assertEqual(log.skipped_rows, [6])
        self.assertEqual(len(log.find_driver('Mark Reyes').legs), 2)

    def test_columns_are_found_by_header(self):
        header = LOG_HEADER[:7] + ['Purpose'] + LOG_HEADER[8:]
        rows = [r[:7] + ('Delivery',) + r[8:] for r in MARK_DAY]
        log = self.load(rows, header=header)
        self.assertTrue(log.has_purpose)
        self.assertFalse(log.has_park_count)
        self.assertEqual(log.find_driver('Mark Reyes').legs[0].purpose, 'Delivery')

    def test_threshold_comes_from_lists_b3(self):
        log = self.load(MARK_DAY, threshold=dt.time(1, 30))
        self.assertEqual(minutes(log.threshold), 90)
        self.assertFalse(any(leg.long_stay for leg in log.find_driver('Mark Reyes').legs))

    def test_missing_required_column(self):
        with self.assertRaisesMessage(excel_log.ActivityLogError, '"Depart Time"'):
            self.load(MARK_DAY, header=[h if h != 'Depart Time' else 'Left' for h in LOG_HEADER])

    def test_missing_file(self):
        with self.assertRaisesMessage(excel_log.ActivityLogError, 'Cannot find'):
            excel_log.load(self.tmp / 'nope.xlsx')

    @skipUnless((settings.BASE_DIR / 'Driver_Activity_Log.xlsx').exists(), 'project workbook not present')
    def test_project_workbook_loads(self):
        excel_log.load(settings.BASE_DIR / 'Driver_Activity_Log.xlsx')


class ViewTests(WorkbookTestCase):
    def setUp(self):
        super().setUp()
        rows = MARK_DAY + [
            (D2, 'Mark Reyes', 'ABC 1234', 'Warehouse', T(7, 0), 'Customer E', T(7, 30), 1, 5.0),
            (D1, 'Ana Cruz', 'XYZ 5678', 'Warehouse', T(13, 0), 'Customer C', T(13, 40), 1, 9.0),
        ]
        path = self.build(rows, drivers=[('Mark Reyes', '0917 000 0000'), ('Ana Cruz', ''), ('Leo Tan', '')],
                          vehicles=[('ABC 1234', 'Isuzu Elf', '9171234567')])
        self.enterContext(override_settings(ACTIVITY_LOG_PATH=path))

    def test_drivers_activity_lists_every_driver(self):
        response = self.client.get(reverse('activity:drivers_activity'))
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'activity/drivers_Activity.html')
        for name in ('Mark Reyes', 'Ana Cruz', 'Leo Tan', '0917 000 0000', 'ABC 1234'):
            self.assertContains(response, name)
        self.assertContains(response, reverse('activity:driver_travel_record', kwargs={'name': 'Mark Reyes'}))

    def test_drivers_activity_filters(self):
        response = self.client.get(reverse('activity:drivers_activity'), {'start': '2026-10-02', 'end': '2026-10-02'})
        mark = next(d for d in response.context['drivers'] if d.name == 'Mark Reyes')
        self.assertEqual(len(mark.legs), 1)
        self.assertContains(response, 'start=2026-10-02')  # filter carried into driver links
        response = self.client.get(reverse('activity:drivers_activity'), {'q': 'ana'})
        self.assertEqual([d.name for d in response.context['drivers']], ['Ana Cruz'])
        response = self.client.get(reverse('activity:drivers_activity'), {'start': 'soon'})
        self.assertContains(response, 'not a valid date')

    def test_driver_travel_record(self):
        url = reverse('activity:driver_travel_record', kwargs={'name': 'Mark Reyes'})
        response = self.client.get(url)
        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, 'activity/Driver_travel_record.html')
        self.assertEqual([d.date for d in response.context['days']], [D2, D1])  # newest first
        for text in ('Customer B', '8:05 AM', '11:50 AM', '1:15', 'LONG STAY', 'Isuzu Elf', '9171234567'):
            self.assertContains(response, text)
        self.assertNotContains(response, 'Customer C')  # Ana's leg

    def test_driver_travel_record_date_filter(self):
        url = reverse('activity:driver_travel_record', kwargs={'name': 'Mark Reyes'})
        response = self.client.get(url, {'start': '2026-10-02'})
        self.assertEqual([d.date for d in response.context['days']], [D2])

    def test_unknown_driver_is_404(self):
        response = self.client.get(reverse('activity:driver_travel_record', kwargs={'name': 'Nobody'}))
        self.assertEqual(response.status_code, 404)

    def test_missing_workbook_shows_message(self):
        with override_settings(ACTIVITY_LOG_PATH=self.tmp / 'missing.xlsx'):
            response = self.client.get(reverse('activity:drivers_activity'))
        self.assertEqual(response.status_code, 503)
        self.assertContains(response, 'Cannot find the activity log', status_code=503)
