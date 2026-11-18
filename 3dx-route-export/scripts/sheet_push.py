"""Push the Route Tasks rows to the shared Google Sheet through the Apps Script web app
(docs/apps-script/plmRouteTasks.gs).

The web app is deployed with access "Anyone within rideriver.com" (the Workspace admin does not
allow "Anyone"), so Google only lets in a signed-in rideriver.com account. push_via_browser()
sends the rows through the project's Chrome profile (.auth/profile), where the Google sign-in is
kept: the first time (or after Google signs you out) a browser window opens on the Google sign-in
page (rideriver.com goes on to the JumpCloud login); sign in once and the push carries on. Later runs push in the background without a window.

.env:
  PLM_SHEET_WEB_APP_URL  the web app /exec URL
  PLM_SHEET_TOKEN        same value as TOKEN in the Apps Script
  PLM_SHEET_GID          tab id (the number after gid= in the sheet link)
"""
import csv
import json
import os
import time
import urllib.request

from datetime import datetime

CHUNK = 2000  # rows per request (Apps Script request and run-time limits)
SIGN_IN_WAIT = 600  # seconds to wait for the Google + JumpCloud sign-in in the browser window


class SheetPushError(RuntimeError):
    pass


class SignInNeeded(SheetPushError):
    """Google answered with its sign-in page / 401: the browser profile is not signed in."""


def sheet_settings():
    """(url, token, gid) from the environment, or None when the sheet is not set up."""
    url = os.getenv("PLM_SHEET_WEB_APP_URL", "").strip()
    token = os.getenv("PLM_SHEET_TOKEN", "").strip()
    gid = os.getenv("PLM_SHEET_GID", "").strip()
    return (url, token, gid) if url and token and gid else None


def cell(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d-%b-%Y %H:%M")
    return value


def payloads(token, gid, headers, records, hidden, chunk=CHUNK, updated=None):
    """The request bodies: one per `chunk` rows, the first one clears the tab, the last one finishes it."""
    rows = [[cell(r.get(h)) for h in headers] for r in records]
    updated = updated or datetime.now().strftime("%d-%b-%Y %H:%M")
    starts = list(range(0, len(rows), chunk)) or [0]
    return [{"token": token, "gid": gid, "headers": headers, "rows": rows[s:s + chunk], "start": s,
             "last": s == starts[-1], "hidden": hidden, "updated": updated} for s in starts]


def _post(url, body, attempts=3):
    data = json.dumps(body).encode("utf-8")
    for n in range(1, attempts + 1):
        try:
            req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
            # Apps Script answers with a redirect to the result; urllib follows it with a GET.
            with urllib.request.urlopen(req, timeout=300) as resp:
                answer = json.loads(resp.read().decode("utf-8"))
            if answer.get("ok"):
                return answer
            raise SheetPushError(answer.get("error") or "web app did not answer ok")
        except (OSError, ValueError, SheetPushError) as e:
            if n == attempts or (isinstance(e, SheetPushError) and "token" in str(e)):
                raise SheetPushError(f"sheet push failed: {e}") from e
            time.sleep(5 * n)


def push(headers, records, hidden=(), settings=None, post=_post):
    """Replace the tab's contents with `records`; returns the number of rows written."""
    url, token, gid = settings or sheet_settings() or (None, None, None)
    if not url:
        raise SheetPushError("PLM_SHEET_WEB_APP_URL / PLM_SHEET_TOKEN / PLM_SHEET_GID not set in .env")
    written = 0
    for body in payloads(token, gid, list(headers), records, list(hidden)):
        post(url, body)
        written += len(body["rows"])
    return written


# ---------------------------------------------------------------- hand-over through Google Drive

DRIVE_CSV = "Route_Tasks_latest.csv"


def drive_dir():
    """The Google Drive for desktop folder the CSV is saved in (PLM_SHEET_DRIVE_DIR), or None."""
    return os.getenv("PLM_SHEET_DRIVE_DIR", "").strip() or None


def write_drive_csv(headers, records, folder):
    """Save the rows as <folder>/Route_Tasks_latest.csv for the sheet's Apps Script
    (docs/apps-script/plmRouteTasksImport.gs) to import. Google Drive for desktop uploads it under
    the user's own account, so no web app access or sign-in is involved. Returns the row count."""
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, DRIVE_CSV)
    tmp = path + ".part"
    with open(tmp, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(headers)
        w.writerows([cell(r.get(h)) for h in headers] for r in records)
    os.replace(tmp, path)  # Drive never uploads a half-written file
    return len(records)


# ---------------------------------------------------------------- push through the signed-in browser

def answer_of(status, url, text):
    """The web app's JSON answer, or SignInNeeded / SheetPushError when Google answered something else."""
    if status in (401, 403) or "accounts.google.com" in url:
        raise SignInNeeded(f"HTTP {status}: not signed in to Google with a rideriver.com account")
    try:
        answer = json.loads(text)
    except ValueError:
        raise SheetPushError(f"HTTP {status}: web app did not answer JSON: {text[:200]!r}") from None
    if not answer.get("ok"):
        raise SheetPushError(answer.get("error") or "web app did not answer ok")
    return answer


def browser_post(request, attempts=3):
    """A post(url, body) for push() that sends through a Playwright APIRequestContext (browser cookies)."""
    from playwright.sync_api import Error as PlaywrightError

    def post(url, body):
        data = json.dumps(body)
        for n in range(1, attempts + 1):
            try:
                # Apps Script answers with a redirect to the result; it is followed with a GET.
                resp = request.post(url, data=data, headers={"Content-Type": "application/json"}, timeout=300_000)
                return answer_of(resp.status, resp.url, resp.text())
            except SignInNeeded:
                raise
            except (PlaywrightError, SheetPushError) as e:
                if n == attempts or "token" in str(e):
                    raise SheetPushError(f"sheet push failed: {e}") from e
                time.sleep(5 * n)
    return post


def _launch(p, profile_dir, headless):
    return p.chromium.launch_persistent_context(
        str(profile_dir), headless=headless, no_viewport=not headless,
        # Google refuses its sign-in page in a browser flagged as automated.
        ignore_default_args=["--enable-automation"], args=["--disable-blink-features=AutomationControlled"],
    )


def signed_in(request, url):
    """True when the web app lets this browser in: a post with an empty token is answered by the
    script itself ("token does not match") instead of Google's 401 / sign-in page. Nothing is written.
    The page address is no proof: rideriver.com sign-in passes through script.google.com on its way
    to the JumpCloud single sign-on page."""
    from playwright.sync_api import Error as PlaywrightError
    try:
        resp = request.post(url, data=json.dumps({"token": ""}), headers={"Content-Type": "application/json"},
                            timeout=60_000)
        answer_of(resp.status, resp.url, resp.text())
        return True
    except SignInNeeded:
        return False
    except SheetPushError:
        return True  # the script answered (token check), so Google let the request in
    except PlaywrightError:
        return False


def _sign_in(context, url):
    if signed_in(context.request, url):
        return
    page = context.new_page()
    page.goto(url)  # the web app address itself starts the rideriver.com (JumpCloud) sign-in
    print(f"Google Sheet: sign in with your rideriver.com Google account in the browser window, "
          f"including the JumpCloud login (waiting up to {SIGN_IN_WAIT // 60} minutes) ...")
    deadline = time.time() + SIGN_IN_WAIT
    while not signed_in(context.request, url):
        if time.time() > deadline or page.is_closed():
            raise SheetPushError("sheet push failed: Google sign-in not completed "
                                 + ("(browser window closed)" if page.is_closed() else "in time"))
        page.wait_for_timeout(3000)
    print("Google Sheet: signed in.")
    page.close()


def push_via_browser(headers, records, hidden, profile_dir, settings=None):
    """push() through the browser profile; opens a window for the Google sign-in only when needed."""
    from playwright.sync_api import sync_playwright

    settings = settings or sheet_settings()
    if not settings:
        raise SheetPushError("PLM_SHEET_WEB_APP_URL / PLM_SHEET_TOKEN / PLM_SHEET_GID not set in .env")
    with sync_playwright() as p:
        context = _launch(p, profile_dir, headless=True)
        try:
            return push(headers, records, hidden, settings, post=browser_post(context.request))
        except SignInNeeded:
            context.close()
            context = _launch(p, profile_dir, headless=False)
            _sign_in(context, settings[0])
            try:
                return push(headers, records, hidden, settings, post=browser_post(context.request))
            except SignInNeeded as e:
                raise SheetPushError(f"sheet push failed: {e} (sign in with a rideriver.com account; "
                                     "web app access must be 'Anyone within rideriver.com')") from e
        finally:
            context.close()
