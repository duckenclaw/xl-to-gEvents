from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle

SCOPES = ['https://www.googleapis.com/auth/calendar']

class GoogleCalendarAPI:
    def __init__(self, credentials_file='credentials.json'):
        self.credentials_file = credentials_file
        self.creds = None
        self.service = None

    def authenticate(self):
        # Token file stores user's access and refresh tokens
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

        # If no valid credentials, let user log in
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, SCOPES)
                self.creds = flow.run_local_server(port=0)

            # Save credentials for next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('calendar', 'v3', credentials=self.creds)
        return self.service

    def create_event(self, event, calendar_id='primary'):
        # Create a single event
        try:
            created_event = self.service.events().insert(
                calendarId=calendar_id,
                body=event,
                conferenceDataVersion=1 if 'conferenceData' in event else 0
            ).execute()
            return created_event
        except Exception as e:
            print(f"Error creating event: {e}")
            return None

    def create_events_batch(self, events, calendar_id='primary'):
        # Create multiple events in batch
        created = []
        failed = []

        for idx, event in enumerate(events, 1):
            print(f"Creating event {idx}/{len(events)}: {event['summary']}")
            result = self.create_event(event, calendar_id)

            if result:
                created.append(result)
                print(f"✓ Created: {event['summary']}")
            else:
                failed.append(event)
                print(f"✗ Failed: {event['summary']}")

        return {
            'created': created,
            'failed': failed,
            'total': len(events),
            'success_count': len(created),
            'fail_count': len(failed)
        }

    def list_calendars(self):
        # List all available calendars
        calendar_list = self.service.calendarList().list().execute()
        return calendar_list.get('items', [])

    def delete_event(self, event_id, calendar_id='primary'):
        # Delete an event
        try:
            self.service.events().delete(
                calendarId=calendar_id,
                eventId=event_id
            ).execute()
            return True
        except Exception as e:
            print(f"Error deleting event: {e}")
            return False

    def list_managed_events(self, calendar_id='primary', time_min=None, time_max=None):
        # List events this script created (tagged appId=xl-to-gEvents), within
        # an optional [time_min, time_max] window. Handles pagination.
        events = []
        page_token = None

        while True:
            params = {
                'calendarId': calendar_id,
                'privateExtendedProperty': 'appId=xl-to-gEvents',
                'singleEvents': True,
                'maxResults': 2500,
                'pageToken': page_token,
            }
            if time_min:
                params['timeMin'] = time_min
            if time_max:
                params['timeMax'] = time_max

            result = self.service.events().list(**params).execute()
            events.extend(result.get('items', []))

            page_token = result.get('nextPageToken')
            if not page_token:
                break

        return events

    def delete_all_events(self, calendar_id='primary'):
        # Delete EVERY event on the calendar. Re-lists after each batch so it
        # fully drains regardless of count; stops if a pass makes no progress
        # (e.g. all deletes failing) to avoid looping forever.
        deleted = 0
        failed = 0
        while True:
            result = self.service.events().list(
                calendarId=calendar_id,
                maxResults=2500,
                showDeleted=False,
            ).execute()
            items = result.get('items', [])
            if not items:
                break

            progressed = False
            for ev in items:
                print(f"Deleting: {ev.get('summary', '(no title)')}")
                if self.delete_event(ev['id'], calendar_id):
                    deleted += 1
                    progressed = True
                else:
                    failed += 1

            if not progressed:
                break

        return {'deleted': deleted, 'failed': failed}

    def count_events(self, calendar_id='primary'):
        # Total number of events on the calendar (for dry-run reporting).
        total = 0
        page_token = None
        while True:
            result = self.service.events().list(
                calendarId=calendar_id,
                maxResults=2500,
                showDeleted=False,
                pageToken=page_token,
            ).execute()
            total += len(result.get('items', []))
            page_token = result.get('nextPageToken')
            if not page_token:
                break
        return total

    def delete_events(self, event_ids, calendar_id='primary'):
        # Delete many events by id, tracking successes/failures.
        deleted = 0
        failed = 0
        for idx, event_id in enumerate(event_ids, 1):
            print(f"Deleting event {idx}/{len(event_ids)}...")
            if self.delete_event(event_id, calendar_id):
                deleted += 1
            else:
                failed += 1
        return {'deleted': deleted, 'failed': failed, 'total': len(event_ids)}

    def get_events(self, calendar_id='primary', max_results=10):
        # Get upcoming events
        from datetime import datetime
        now = datetime.utcnow().isoformat() + 'Z'

        events_result = self.service.events().list(
            calendarId=calendar_id,
            timeMin=now,
            maxResults=max_results,
            singleEvents=True,
            orderBy='startTime'
        ).execute()

        return events_result.get('items', [])


if __name__ == "__main__":
    # Test authentication
    api = GoogleCalendarAPI()
    api.authenticate()

    print("Available calendars:")
    calendars = api.list_calendars()
    for cal in calendars:
        print(f"- {cal['summary']} (ID: {cal['id']})")