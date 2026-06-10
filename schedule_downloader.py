"""Download the schedule .xlsx from the university repository API."""
import shutil
from pathlib import Path

import requests

import config


class TokenExpiredError(Exception):
    """Raised when the API rejects the bearer token (HTTP 401)."""


def _auth_headers(token):
    return {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'pl-PL',
    }


def find_file_id(token, name=None, list_url=None):
    """Resolve the current GUID of the schedule file by its name.

    The repository assigns a new GUID on every re-upload, so we look the file
    up by `nazwa` each time. Raises TokenExpiredError on 401 so the caller can
    refresh the session and retry.
    """
    name = (name or config.SCHEDULE_FILE_NAME).strip()
    list_url = list_url or config.REPO_LIST_URL

    resp = requests.get(list_url, headers=_auth_headers(token), timeout=60)
    if resp.status_code == 401:
        raise TokenExpiredError("Repository listing API returned 401 (token expired).")
    resp.raise_for_status()

    try:
        body = resp.json()
    except ValueError:
        raise RuntimeError(
            f"Listing endpoint {list_url} did not return JSON. Set REPO_LIST_URL "
            f"in .env to the exact XHR URL from the File-repository page (DevTools)."
        )

    items = body.get('data') if isinstance(body, dict) else body
    if not isinstance(items, list):
        raise RuntimeError(
            f"Listing endpoint {list_url} returned no 'data' array. Set REPO_LIST_URL "
            f"in .env to the exact XHR URL from the File-repository page (DevTools)."
        )

    target = name.casefold()
    for item in items:
        nazwa = str(item.get('nazwa', '')).strip()
        stem = str(item.get('nazwaPliku', '')).strip()
        if stem.lower().endswith('.xlsx'):
            stem = stem[:-5]
        if nazwa.casefold() == target or stem.strip().casefold() == target:
            return item['id']

    available = ', '.join(repr(str(i.get('nazwa', '?'))) for i in items)
    raise RuntimeError(
        f"No repository file named {name!r}. Available: {available}"
    )


def download_schedule(token, dest_path='assets/schedule.xlsx', file_id=None):
    """Download the schedule file to dest_path and return its Path.

    Backs up any existing file to <name>.prev.xlsx first, so the previous
    version is still available for manual diffing. Raises TokenExpiredError on
    a 401 so the caller can refresh the session and retry.
    """
    file_id = file_id or config.SCHEDULE_FILE_ID
    if not file_id:
        raise RuntimeError(
            "No file_id to download. Resolve one via find_file_id() or set "
            "SCHEDULE_FILE_ID in .env."
        )
    url = (
        f"{config.API_BASE_URL}/api/repozytorium-plikow/"
        f"platforma-studenta/pliki/{file_id}"
    )

    resp = requests.get(url, headers=_auth_headers(token), timeout=60)
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

    tok = token_provider.get_token()
    fid = find_file_id(tok)
    print(f"Resolved {config.SCHEDULE_FILE_NAME!r} -> {fid}")
    path = download_schedule(tok, file_id=fid)
    print(f"Downloaded schedule to {path} ({path.stat().st_size} bytes)")
