"""Orchestrate the end-to-end schedule sync.

token -> download xlsx -> parse -> wipe this script's previous events on the
target calendar (within the schedule's date range) -> recreate from the new
file. Wipe & recreate keeps the calendar an exact mirror of the newest xlsx.
"""
from datetime import timedelta

import config
import schedule_downloader
import token_provider
from event_creator import EventCreator
from google_api import GoogleCalendarAPI
from schedule_downloader import TokenExpiredError
from xlsx_parser import ScheduleParser

DEFAULT_XLSX = 'assets/schedule.xlsx'


def _resolve_and_download(token):
    """Resolve the current file id by name (unless pinned) and download it."""
    if config.SCHEDULE_FILE_ID:
        file_id = config.SCHEDULE_FILE_ID
        print(f"  Using pinned SCHEDULE_FILE_ID -> {file_id}")
    else:
        file_id = schedule_downloader.find_file_id(token)
        print(f"  Resolved {config.SCHEDULE_FILE_NAME!r} -> {file_id}")
    return schedule_downloader.download_schedule(token, DEFAULT_XLSX, file_id=file_id)


def _download_with_retry():
    """Get a token, resolve the file, and download — refreshing once on a 401."""
    token = token_provider.get_token()
    try:
        return _resolve_and_download(token)
    except TokenExpiredError:
        print("Token rejected (401); refreshing session and retrying once…")
        token_provider.invalidate_session()
        token = token_provider.get_token()
        return _resolve_and_download(token)


def _range_for(parsed_events):
    """RFC3339 [time_min, time_max] window covering the parsed events, padded
    by a day on each side so timezone offsets never clip the edges."""
    starts = [e['start'] for e in parsed_events]
    ends = [e['end'] for e in parsed_events]
    lo = (min(starts).date() - timedelta(days=1))
    hi = (max(ends).date() + timedelta(days=2))
    return lo.strftime('%Y-%m-%dT00:00:00Z'), hi.strftime('%Y-%m-%dT00:00:00Z')


def sync(dry_run=False, download=True):
    calendar_id = config.CALENDAR_ID

    print("=" * 60)
    print("SCHEDULE SYNC" + (" (DRY RUN)" if dry_run else ""))
    print("=" * 60)

    # 1 + 2. Token + download.
    if download:
        print("\n[1/4] Acquiring token and downloading schedule…")
        path = _download_with_retry()
        print(f"✓ Downloaded {path} ({path.stat().st_size} bytes)")
    else:
        path = DEFAULT_XLSX
        print(f"\n[1/4] Using existing file {path} (download skipped)")

    # 3. Parse + build calendar events.
    print(f"\n[2/4] Parsing schedule…")
    parsed = ScheduleParser(str(path)).parse()
    if not parsed:
        print("No events found in schedule — aborting (calendar left untouched).")
        return
    calendar_events = EventCreator.create_batch(parsed)
    time_min, time_max = _range_for(parsed)
    print(f"✓ Parsed {len(parsed)} events spanning {time_min[:10]} … {time_max[:10]}")

    # 4. Reconcile with the calendar.
    print(f"\n[3/4] Checking existing events on calendar '{calendar_id}'…")
    api = GoogleCalendarAPI()
    api.authenticate()
    existing = api.list_managed_events(calendar_id, time_min, time_max)
    print(f"  Existing script events in range: {len(existing)}")
    print(f"  Will delete: {len(existing)}   |   Will create: {len(calendar_events)}")

    if dry_run:
        print("\nDRY RUN — no changes made. Events that would be created:")
        for event in calendar_events:
            print(f"  + {event['summary']}")
        return

    print(f"\n[4/4] Applying changes…")
    if existing:
        del_result = api.delete_events([e['id'] for e in existing], calendar_id)
        print(f"  Deleted {del_result['deleted']}/{del_result['total']} old events")

    result = api.create_events_batch(calendar_events, calendar_id)

    print("\n" + "=" * 60)
    print("RESULTS")
    print("=" * 60)
    print(f"Total events: {result['total']}")
    print(f"✓ Created: {result['success_count']}")
    print(f"✗ Failed:  {result['fail_count']}")
    if result['failed']:
        print("\nFailed events:")
        for event in result['failed']:
            print(f"  - {event['summary']}")


if __name__ == "__main__":
    sync()
