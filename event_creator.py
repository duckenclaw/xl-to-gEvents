from datetime import datetime

class EventCreator:
    @staticmethod
    def create_event(parsed_event):
        """
        Convert parsed event into Google Calendar event format

        Format: OFFLINE-Wykład | B401 | Subject | Professor
        """
        mode = "ONLINE" if parsed_event['is_online'] else "OFFLINE"
        event_type = parsed_event['type'].capitalize()

        # Build title
        parts = [f"{mode}-{event_type}"]

        if parsed_event['classroom']:
            parts.append(parsed_event['classroom'])

        parts.append(parsed_event['subject'])

        if parsed_event['professor']:
            parts.append(parsed_event['professor'])

        title = " | ".join(parts)

        # Create Google Calendar event structure
        event = {
            'summary': title,
            'start': {
                'dateTime': parsed_event['start'].isoformat(),
                'timeZone': 'Europe/Warsaw',
            },
            'end': {
                'dateTime': parsed_event['end'].isoformat(),
                'timeZone': 'Europe/Warsaw',
            },
            'description': EventCreator._build_description(parsed_event),
            'location': parsed_event['classroom'] if not parsed_event['is_online'] else 'Online',
        }

        if parsed_event['is_online']:
            event['conferenceData'] = {
                'createRequest': {
                    'requestId': f"online-{parsed_event['start'].timestamp()}",
                    'conferenceSolutionKey': {'type': 'hangoutsMeet'}
                }
            }

        return event

    @staticmethod
    def _build_description(parsed_event):
        # Build event description with details
        lines = []
        lines.append(f"Przedmiot: {parsed_event['subject']}")
        lines.append(f"Typ: {parsed_event['type']}")

        if parsed_event['professor']:
            lines.append(f"Prowadzący: {parsed_event['professor']}")

        if parsed_event['is_online']:
            lines.append("Forma: Online")
        else:
            lines.append(f"Forma: Stacjonarne")
            if parsed_event['classroom']:
                lines.append(f"Sala: {parsed_event['classroom']}")

        return "\n".join(lines)

    @staticmethod
    def create_batch(parsed_events):
        # Create multiple events from parsed events
        return [EventCreator.create_event(event) for event in parsed_events]


if __name__ == "__main__":
    # Test event creator
    test_event = {
        'subject': 'Wstęp do analizy matematycznej z elementami algebry liniowej',
        'type': 'wykład',
        'professor': 'dr inż. Katarzyna Łuczak',
        'classroom': 'B401',
        'start': datetime(2025, 10, 4, 10, 30),
        'end': datetime(2025, 10, 4, 11, 15),
        'is_online': False
    }

    event = EventCreator.create_event(test_event)
    print("Created event:")
    print(f"Title: {event['summary']}")
    print(f"Start: {event['start']['dateTime']}")
    print(f"End: {event['end']['dateTime']}")
    print(f"Description:\n{event['description']}")