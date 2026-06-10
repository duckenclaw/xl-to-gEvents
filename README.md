# xl-to-gEvents

python xlsx parser for AŚ schedule tables to automatically create google calendar events

## Setup

```bash
# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Install the browser used for automated login (one-time)
python -m playwright install chromium
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

## Automated sync (download + update calendar)

`sync` logs in, downloads the newest schedule `.xlsx` from the university file
repository, and mirrors it onto your calendar — deleting the previous events it
created and recreating them from the new file, so the calendar always matches
the latest schedule.

### `.env` configuration

Create a `.env` file in the project root:

```
# Google calendar to mirror the schedule onto
CALENDAR_ID=your_calendar_id@group.calendar.google.com

# University Microsoft / Azure AD login
AZURE_LOGIN=you@uni.example
AZURE_PASSWORD=your_password

# The student-platform page URL (the "File repository" address in your browser)
PLATFORM_URL=https://...

# Optional: schedule file name to look up (defaults to the value below).
# The file's GUID rotates on every re-upload, so it is resolved by name each run.
# SCHEDULE_FILE_NAME=Informatyka - niestacjonarne - semestr 02

# Optional: override the repository listing endpoint if the default is wrong
# (paste the exact XHR URL from the File-repository page in DevTools).
# REPO_LIST_URL=https://api-dziekanat.wst.com.pl/api/repozytorium-plikow/platforma-studenta/pliki-do-pobrania

# Optional: pin a specific file GUID instead of name lookup
# SCHEDULE_FILE_ID=5c520ab7-e6e0-4346-9a60-39e7d1b2112b

# Optional: paste a bearer token manually to skip the browser login entirely
# SCHEDULE_TOKEN=eyJ...
```

The token is obtained automatically by driving the login in a headless browser
and capturing the session's bearer token. The session is cached in `.auth/`
(gitignored); the credential login only re-runs when that session expires. If
Microsoft shows an MFA prompt, a visible browser opens so you can complete it.

### Commands

```bash
# Preview: download + show what would change, without touching the calendar
python main.py sync --dry-run

# Download the newest schedule and update the calendar
python main.py sync

# Re-run against the already-downloaded file (skip the browser/download)
python main.py sync --no-download
```

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