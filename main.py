import argparse
import sys
from xlsx_parser import ScheduleParser
from event_creator import EventCreator
from google_api import GoogleCalendarAPI

def test_parser(xlsx_file):

    print(f"Parsing {xlsx_file}...")
    parser = ScheduleParser(xlsx_file)
    events = parser.parse()

    print(f"\nFound {len(events)} events:\n")
    for event in events:
        mode = "ONLINE" if event['is_online'] else "OFFLINE"
        print(f"{mode} | {event['start']} - {event['end']}")
        print(f"  {event['subject']} ({event['type']})")
        print(f"  {event['professor']} | {event['classroom']}\n")

    return events

def test_event_creator(xlsx_file):
    """Test event creator"""
    print(f"Parsing {xlsx_file}...")
    parser = ScheduleParser(xlsx_file)
    parsed_events = parser.parse()

    print(f"\nCreating {len(parsed_events)} Google Calendar events...\n")
    events = EventCreator.create_batch(parsed_events)

    for event in events:
        print(f"Title: {event['summary']}")
        print(f"Start: {event['start']['dateTime']}")
        print(f"End: {event['end']['dateTime']}")
        print(f"Location: {event.get('location', 'N/A')}")
        print()

    return events

def test_google_api():
    # Test Google API authentication and calendar listing
    print("Testing Google Calendar API...")
    api = GoogleCalendarAPI()
    api.authenticate()

    print("\nAvailable calendars:")
    calendars = api.list_calendars()
    for cal in calendars:
        print(f"- {cal['summary']} (ID: {cal['id']})")

    print("\nUpcoming events:")
    events = api.get_events(max_results=5)
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(f"- {start}: {event['summary']}")

    return api

def full_pipeline(xlsx_file, calendar_id='primary', dry_run=False):
    # Full pipeline: parse XLSX -> create events -> upload to Google Calendar
    print("=" * 60)
    print("SCHEDULE TO CALENDAR - FULL PIPELINE")
    print("=" * 60)

    # 1. Parse XLSX
    print(f"\n[1/3] Parsing {xlsx_file}...")
    parser = ScheduleParser(xlsx_file)
    parsed_events = parser.parse()
    print(f"✓ Found {len(parsed_events)} events")

    # 2. Create Google Calendar events
    print(f"\n[2/3] Converting to Google Calendar format...")
    calendar_events = EventCreator.create_batch(parsed_events)
    print(f"✓ Converted {len(calendar_events)} events")

    # 3. Upload to Google Calendar
    if dry_run:
        print(f"\n[3/3] DRY RUN - Would create {len(calendar_events)} events:")
        for event in calendar_events:
            print(f"  - {event['summary']}")
        return

    print(f"\n[3/3] Uploading to Google Calendar...")
    api = GoogleCalendarAPI()
    api.authenticate()

    result = api.create_events_batch(calendar_events, calendar_id)

    print("\n" + "=" * 60)
    print("RESULTS")
    print("=" * 60)
    print(f"Total events: {result['total']}")
    print(f"✓ Successfully created: {result['success_count']}")
    print(f"✗ Failed: {result['fail_count']}")

    if result['failed']:
        print("\nFailed events:")
        for event in result['failed']:
            print(f"  - {event['summary']}")

def main():
    parser = argparse.ArgumentParser(
        description='Parse XLSX schedule and create Google Calendar events'
    )

    subparsers = parser.add_subparsers(dest='command', help='Commands')

    # Test parser
    parser_test = subparsers.add_parser('test-parser', help='Test XLSX parser')
    parser_test.add_argument('xlsx_file', help='Path to XLSX file')

    # Test event creator
    creator_test = subparsers.add_parser('test-creator', help='Test event creator')
    creator_test.add_argument('xlsx_file', help='Path to XLSX file')

    # Test Google API
    api_test = subparsers.add_parser('test-api', help='Test Google Calendar API')

    # Full pipeline
    pipeline = subparsers.add_parser('run', help='Run full pipeline')
    pipeline.add_argument('xlsx_file', help='Path to XLSX file')
    pipeline.add_argument('--calendar', default='primary', help='Calendar ID (default: primary)')
    pipeline.add_argument('--dry-run', action='store_true', help='Show what would be created without actually creating')

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return

    try:
        if args.command == 'test-parser':
            test_parser(args.xlsx_file)
        elif args.command == 'test-creator':
            test_event_creator(args.xlsx_file)
        elif args.command == 'test-api':
            test_google_api()
        elif args.command == 'run':
            full_pipeline(args.xlsx_file, args.calendar, args.dry_run)
    except Exception as e:
        print(f"\nError: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()