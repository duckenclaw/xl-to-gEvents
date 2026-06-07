"""Download the schedule .xlsx from the university repository API."""
import shutil
from pathlib import Path

import requests

import config


class TokenExpiredError(Exception):
    """Raised when the API rejects the bearer token (HTTP 401)."""


def download_schedule(token, dest_path='assets/schedule.xlsx', file_id=None):
    """Download the schedule file to dest_path and return its Path.

    Backs up any existing file to <name>.prev.xlsx first, so the previous
    version is still available for manual diffing. Raises TokenExpiredError on
    a 401 so the caller can refresh the session and retry.
    """
    file_id = file_id or config.SCHEDULE_FILE_ID
    url = (
        f"{config.API_BASE_URL}/api/repozytorium-plikow/"
        f"platforma-studenta/pliki/{file_id}"
    )
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'pl-PL',
    }

    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code == 401:
        raise TokenExpiredError("Repository API returned 401 (token expired).")
    resp.raise_for_status()

    dest = Path(dest_path)
    dest.parent.mkdir(parents=True, exist_ok=True)

    # Preserve the previous version for diffing before overwriting.
    if dest.exists():
        prev = dest.with_suffix('.prev.xlsx')
        shutil.copyfile(dest, prev)

    dest.write_bytes(resp.content)
    return dest


if __name__ == "__main__":
    import token_provider

    path = download_schedule(token_provider.get_token())
    print(f"Downloaded schedule to {path} ({path.stat().st_size} bytes)")
