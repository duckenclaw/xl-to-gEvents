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

# URL of the student-platform SPA (the "File repository" page address). The
# bearer token is minted when this app loads, so the browser flow starts here.
PLATFORM_URL = os.getenv('PLATFORM_URL')

# --- Schedule file repository ----------------------------------------------
# GUID of the schedule file in the repository API. Defaults to the known file
# ("Informatyka - niestacjonarne - semestr 02") but can be overridden in .env.
SCHEDULE_FILE_ID = os.getenv(
    'SCHEDULE_FILE_ID', '5c1d20da-dd43-4966-971d-5e76d59dc4d8'
)

# Optional manual bearer-token override / fast-path. When set and still valid,
# the browser login is skipped entirely.
SCHEDULE_TOKEN = os.getenv('SCHEDULE_TOKEN')

# Base host for the repository API.
API_BASE_URL = os.getenv('API_BASE_URL', 'https://api-dziekanat.wst.com.pl')


def require(name):
    """Return the named config value, raising if it is unset/empty."""
    value = globals().get(name)
    if not value:
        raise RuntimeError(
            f"Missing required config '{name}'. Add it to your .env file."
        )
    return value
