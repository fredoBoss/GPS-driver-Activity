from pathlib import Path

from django.conf import settings
from django.core.management.base import BaseCommand, CommandError

from activity import excel_log, log_writer, sinotrack


class Command(BaseCommand):
    help = ('Copy SinoTrack Travel Report CSVs (new trip rows) or Park Report CSVs (stops merged into '
            'the trip rows; stops on days without trips get parking-only rows) into the Activity Log sheet '
            'of the Excel workbook. Import the Travel Report first.')

    def add_arguments(self, parser):
        parser.add_argument('csv', nargs='+', type=Path, help='Travel Report or Park Report CSV file(s) from SinoTrack.')
        parser.add_argument('--driver', help='Driver name to use instead of the one in the device name.')
        parser.add_argument('--plate', help='Plate No. to use instead of the one in the device name.')
        parser.add_argument('--workbook', type=Path, default=settings.ACTIVITY_LOG_PATH,
                            help='Workbook to update (default: ACTIVITY_LOG_PATH).')

    def handle(self, *args, **options):
        for csv_path in options['csv']:
            try:
                data = csv_path.read_bytes()
                kind = sinotrack.detect_report(data)
                if kind == 'park':
                    parks, invalid_rows = sinotrack.parse_park_report(data)
                    result = log_writer.merge_parks(options['workbook'], parks, plate=options['plate'],
                                                    driver=options['driver'])
                    self._park_summary(csv_path, result)
                elif kind == 'travel':
                    trips, invalid_rows = sinotrack.parse_travel_report(data)
                    result = log_writer.import_trips(options['workbook'], trips,
                                                     driver=options['driver'], plate=options['plate'])
                    self._trip_summary(csv_path, result)
                else:
                    raise sinotrack.ReportError('not a SinoTrack Travel Report or Park Report CSV.')
            except FileNotFoundError:
                raise CommandError(f'{csv_path}: file not found.')
            except (sinotrack.ReportError, excel_log.ActivityLogError) as exc:
                raise CommandError(f'{csv_path.name}: {exc}')

            if result.backup:
                self.stdout.write(f'  Backup: {result.backup}')
            if invalid_rows:
                self.stdout.write(self.style.WARNING(f'  Skipped unreadable CSV rows: {invalid_rows[:20]}'))
            for warning in result.warnings:
                self.stdout.write(self.style.WARNING(f'  {warning}'))

    def _trip_summary(self, csv_path, result):
        if result.added:
            self.stdout.write(self.style.SUCCESS(
                f'{csv_path.name}: copied {result.added} trips into rows {result.first_row}-{result.last_row} '
                f'for {", ".join(result.drivers)} ({", ".join(result.plates)}), '
                f'{result.start:%b %d} to {result.end:%b %d, %Y}.'))
        if result.completed_rows:
            self.stdout.write(self.style.SUCCESS(
                f'{csv_path.name}: filled in {", ".join(result.filled_columns)} for {len(result.completed_rows)} '
                f'trips already in the log (rows {min(result.completed_rows)}-{max(result.completed_rows)}).'))
        unchanged = result.duplicates - len(result.completed_rows)
        if unchanged:
            self.stdout.write(f'  {unchanged} trips already in the log needed nothing.')
        if result.added_to_lists:
            self.stdout.write(f'  Added to the Lists sheet: {", ".join(result.added_to_lists)}')

    def _park_summary(self, csv_path, result):
        if result.matched:
            self.stdout.write(self.style.SUCCESS(
                f'{csv_path.name}: merged {result.matched} of {result.parks} stops into {len(result.rows)} '
                f'trip rows ({min(result.rows)}-{max(result.rows)}).'))
        if result.parking_only:
            days = ', '.join(f'{d:%b %d}' for d in result.parking_only_dates)
            rows = result.parking_only_rows
            self.stdout.write(self.style.SUCCESS(
                f'{csv_path.name}: copied {result.parking_only} stops on days with no trip ({days}) into '
                f'{len(rows)} parking-only rows ({min(rows)}-{max(rows)}).'))
        if result.cleared_rows:
            self.stdout.write(f'  Cleared parking-only rows whose stops now belong to trips: {result.cleared_rows}')
