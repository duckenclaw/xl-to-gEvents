# xl-to-gEvents

python xlsx parser for AŚ schedule tables to automatically create google calendar events

## Setup

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

```bash
deactivate # to exit venv virtual environment
```

## Google Calendar API Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable Google Calendar API
4. Create OAuth 2.0 credentials (Desktop app)
5. Download credentials as `credentials.json` and place in project root

## Usage

### Test individual modules

```bash
# Test XLSX parser. Outputs parsed schedule to console
python main.py test-parser schedule.xlsx

# Test event creator (parser + formatting)
python main.py test-creator schedule.xlsx

# Test Google API authentication. Creates `token.pickle` used for automatic authentication
python main.py test-api
```

### Run full pipeline

```bash
# Dry run (shows what would be created)
python main.py run schedule.xlsx --dry-run

# Create events in default calendar
python main.py run schedule.xlsx

# Create in specific calendar
python main.py run schedule.xlsx --calendar "your-calendar-id"
```

## Event Format

Events are created with the following format:
- **Title**: `OFFLINE-Wykład | B401 | Subject | Professor`
- **Online events**: All events where the header row ends with `ZAJĘCIA ONLINE` are marked as `ONLINE-` and include Google Meet link
- **Description**: Contains subject, type, professor, and location details
- **Time zone**: Europe/Warsaw