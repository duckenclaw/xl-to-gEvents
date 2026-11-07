"""Central configuration loaded from the .env file.

Reads the environment once via python-dotenv and exposes the values the
download/sync flow needs. Required values raise a clear error when missing so
failures point at the .env file instead of a cryptic KeyError deep in the code.
"""
import os
from dotenv import load_dotenv

load_dotenv()

# --- Google Calendar -------------------------------------------------------
# Calendar that the schedule events live in (wiped & recreated on each sync).
CALENDAR_ID = os.getenv('CALENDAR_ID', 'primary')

# --- University login (Microsoft / Azure AD) -------------------------------
AZURE_LOGIN = os.getenv('AZURE_LOGIN')
AZURE_PASSWORD = os.getenv('AZURE_PASSWORD')

# URL of the student-platform SPA (the "File repository" page). Navigating here
# in the browser forces an authenticated API call, which is how the bearer token
# is captured — so this MUST be the SPA page (dziekanat.wst.com.pl), NOT the auth
# server. Defaults to the known repository page; overridable in .env.
PLATFORM_URL = os.getenv(
    'PLATFORM_URL',
    'https://dziekanat.wst.com.pl/pl/repozytorium-plikow/student/pliki-do-pobrania',
)

# --- Schedule file repository ----------------------------------------------
# Base host for the repository API.
API_BASE_URL = os.getenv('API_BASE_URL', 'https://api-dziekanat.wst.com.pl')

# The file is looked up by name, because its GUID rotates on every re-upload.
SCHEDULE_FILE_NAME = os.getenv(
    'SCHEDULE_FILE_NAME', 'Informatyka - niestacjonarne - semestr 03'
)

# Listing endpoint that returns {"data": [{"id", "nazwa", ...}]}. typKataloguId=1
# is the downloadable-plans catalog; pageSize is large so the file is never on a
# later page. Overridable in .env if the API changes.
REPO_LIST_URL = os.getenv(
    'REPO_LIST_URL',
    f'{API_BASE_URL}/api/repozytorium-plikow/platforma-studenta/pliki'
    '?orderBy=nazwa:asc&page=0&pageSize=200&typKataloguId=1',
)

# Optional manual override: pin a specific file GUID instead of name lookup.
SCHEDULE_FILE_ID = os.getenv('SCHEDULE_FILE_ID')

# Optional manual bearer-token override / fast-path. When set and still valid,
# the browser login is skipped entirely.
SCHEDULE_TOKEN = os.getenv('SCHEDULE_TOKEN')


def require(name):
    """Return the named config value, raising if it is unset/empty."""
    value = globals().get(name)
    if not value:
        raise RuntimeError(
            f"Missing required config '{name}'. Add it to your .env file."
        )
    return value
