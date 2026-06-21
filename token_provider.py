"""Obtain a fresh API bearer token for the file-repository API.

Strategy, in priority order:
  1. Manual override: a still-valid SCHEDULE_TOKEN in .env (no browser at all).
  2. Saved session: reuse .auth/state.json headlessly and let the SPA mint a
     fresh token (no credentials typed).
  3. Full login: drive the Microsoft/Azure AD login with AZURE_LOGIN /
     AZURE_PASSWORD, headless first, falling back to a visible browser if an
     interactive challenge (MFA) appears.

The token is captured by intercepting the `Authorization: Bearer ...` header on
the SPA's calls to the repository API, with an OIDC-storage read as fallback.
"""
import base64
import json
import time
from pathlib import Path
from urllib.parse import urlparse

import config

STATE_PATH = Path('.auth/state.json')
# Refresh slightly early so we never hand out a token about to expire.
EXP_MARGIN_SECONDS = 120


class NeedsInteraction(Exception):
    """Raised when headless login hits a challenge needing a human."""


# --- JWT helpers -----------------------------------------------------------

def _token_exp(token):
    """Return the JWT `exp` (epoch seconds), or None if undecodable."""
    try:
        payload_b64 = token.split('.')[1]
        payload_b64 += '=' * (-len(payload_b64) % 4)  # restore padding
        payload = json.loads(base64.urlsafe_b64decode(payload_b64))
        return payload.get('exp')
    except Exception:
        return None


def _token_is_valid(token):
    exp = _token_exp(token)
    if exp is None:
        return False
    return exp > time.time() + EXP_MARGIN_SECONDS


# --- Playwright lazy import ------------------------------------------------

def _playwright():
    try:
        from playwright.sync_api import sync_playwright
        return sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run:\n"
            "  pip install -r requirements.txt\n"
            "  python -m playwright install chromium"
        ) from e


# --- Token capture ---------------------------------------------------------

def _attach_capture(context):
    """Capture the bearer token off any request to the repository API."""
    holder = {}

    def on_request(request):
        if 'token' in holder:
            return
        if request.url.startswith(config.API_BASE_URL):
            auth = request.headers.get('authorization', '')
            if auth.lower().startswith('bearer '):
                holder['token'] = auth[7:]

    context.on('request', on_request)
    return holder


def _read_token_from_storage(page):
    """Fallback: pull access_token out of the SPA's OIDC storage."""
    try:
        return page.evaluate(
            """() => {
                for (const store of [window.localStorage, window.sessionStorage]) {
                    for (let i = 0; i < store.length; i++) {
                        const key = store.key(i);
                        if (key && key.indexOf('oidc.user') === 0) {
                            try {
                                const u = JSON.parse(store.getItem(key));
                                if (u && u.access_token) return u.access_token;
                            } catch (e) {}
                        }
                    }
                }
                return null;
            }"""
        )
    except Exception:
        return None


def _settle(page):
    """Give the SPA a moment to bootstrap and fire its first authenticated API
    call. networkidle can legitimately time out on apps that long-poll, so it's
    best-effort."""
    try:
        page.wait_for_load_state('networkidle', timeout=10000)
    except Exception:
        pass


def _wait_for_token(page, holder, timeout):
    deadline = time.time() + timeout
    while time.time() < deadline:
        if 'token' in holder:
            return holder['token']
        tok = _read_token_from_storage(page)
        if tok:
            return tok
        page.wait_for_timeout(500)
    return holder.get('token')


# --- Microsoft login flow --------------------------------------------------

def _perform_login(page, headless):
    """Walk the IdentityServer -> Microsoft Azure AD login pages."""
    from playwright.sync_api import TimeoutError as PWTimeout

    ms_button = "a[href*='scheme=Azure']"
    try:
        page.wait_for_selector(ms_button, timeout=15000)
        page.click(ms_button)
    except PWTimeout:
        # Not on the IdentityServer login page (already authenticated, or the
        # SPA went straight through). Let the caller wait for a token.
        return

    # Microsoft: email -> Next
    page.wait_for_selector("input[name='loginfmt']", timeout=30000)
    page.fill("input[name='loginfmt']", config.require('AZURE_LOGIN'))
    page.click("#idSIButton9")

    # Microsoft: password -> Sign in
    page.wait_for_selector("input[name='passwd']", state='visible', timeout=30000)
    page.fill("input[name='passwd']", config.require('AZURE_PASSWORD'))
    page.click("#idSIButton9")

    _wait_through_post_password(page, headless)


def _wait_through_post_password(page, headless):
    """After the password, click through 'Stay signed in?' and wait for the
    SPA. Raise NeedsInteraction (headless) if a challenge stalls the flow."""
    platform_host = urlparse(config.PLATFORM_URL).netloc
    # Give a human plenty of time to finish MFA in a visible browser.
    timeout = 25 if headless else 240
    deadline = time.time() + timeout

    while time.time() < deadline:
        url = page.url
        if platform_host in urlparse(url).netloc:
            return  # back at the SPA -> success

        if 'login.microsoftonline.com' in url:
            # "Stay signed in?" (KMSI) keeps the session long-lived: click Yes.
            if page.query_selector("#KmsiCheckboxField") or _has_text(page, "Stay signed in"):
                _safe_click(page, "#idSIButton9")
                page.wait_for_timeout(1000)
                continue
            if headless and _looks_like_mfa(page):
                raise NeedsInteraction("MFA challenge detected during headless login")

        page.wait_for_timeout(1000)

    if headless:
        raise NeedsInteraction("Login did not complete headlessly (likely MFA)")
    # Headed but timed out — let the caller's token wait decide.


def _looks_like_mfa(page):
    selectors = [
        "#idDiv_SAOTCAS_Title",     # "Approve sign in request"
        "#idTxtBx_SAOTCC_OTC",       # enter verification code
        "#idDiv_SAOTCS_Title",       # choose verification method
    ]
    return any(page.query_selector(s) for s in selectors)


def _has_text(page, text):
    try:
        return page.get_by_text(text, exact=False).count() > 0
    except Exception:
        return False


def _safe_click(page, selector):
    try:
        page.click(selector, timeout=3000)
    except Exception:
        pass


# --- Browser orchestration -------------------------------------------------

def _try_saved_session():
    sync_playwright = _playwright()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(storage_state=str(STATE_PATH))
        try:
            page = context.new_page()
            holder = _attach_capture(context)
            page.goto(config.require('PLATFORM_URL'), wait_until='domcontentloaded', timeout=60000)
            _settle(page)
            token = _wait_for_token(page, holder, timeout=25)
            if token:
                _save_state(context)
            return token
        finally:
            context.close()
            browser.close()


def _full_login(headless):
    sync_playwright = _playwright()
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context()
        try:
            page = context.new_page()
            holder = _attach_capture(context)
            page.goto(config.require('PLATFORM_URL'), wait_until='domcontentloaded', timeout=60000)
            _perform_login(page, headless)
            _settle(page)
            token = _wait_for_token(page, holder, timeout=40)
            if not token:
                raise NeedsInteraction("No token captured after login")
            _save_state(context)
            return token
        finally:
            context.close()
            browser.close()


def _save_state(context):
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(STATE_PATH))


# --- Public API ------------------------------------------------------------

def invalidate_session():
    """Drop the saved session so the next call does a fresh login."""
    if STATE_PATH.exists():
        STATE_PATH.unlink()


def get_token():
    """Return a valid bearer token (without the 'Bearer ' prefix)."""
    # 1. Manual override.
    if config.SCHEDULE_TOKEN and _token_is_valid(config.SCHEDULE_TOKEN):
        print("Using SCHEDULE_TOKEN from .env (still valid).")
        return config.SCHEDULE_TOKEN

    # 2. Reuse a saved browser session.
    if STATE_PATH.exists():
        print("Reusing saved session to refresh token…")
        try:
            token = _try_saved_session()
            if token:
                return token
        except Exception as e:
            print(f"Saved session unusable ({e}); logging in fresh.")
        print("Saved session expired; logging in with credentials…")

    # 3. Full credential login: headless first, headed fallback for MFA.
    try:
        return _full_login(headless=True)
    except NeedsInteraction as e:
        print(f"{e}. Opening a visible browser — complete any prompts/MFA…")
        return _full_login(headless=False)


if __name__ == "__main__":
    tok = get_token()
    print(f"\nToken acquired ({len(tok)} chars). exp={_token_exp(tok)}")
