"""3DEXPERIENCE helpers: self-healing locators, read-only guard, login and app search.

Every locator used for an action or a wait has backups. While the primary matches it is
used unchanged; when it finds nothing, the first backup with a visible match is used,
logged as [HEAL] and written to output/healed-locators.json so it can be fixed for good.
"""
import json
import os
import re
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

from playwright.sync_api import Error as PlaywrightError

PROJECT_DIR = Path(__file__).resolve().parent.parent
HEAL_FILE = PROJECT_DIR / "output" / "healed-locators.json"
I = re.IGNORECASE


@dataclass
class Config:
    base_url: str
    username: str
    password: str
    search_text: str
    app_name: str
    headless: bool

    @classmethod
    def from_env(cls):
        from dotenv import load_dotenv
        load_dotenv(PROJECT_DIR / ".env")
        base = os.getenv("THREEDX_BASE_URL", "").strip()
        if not base:
            raise SystemExit("THREEDX_BASE_URL is missing in .env (copy .env.example to .env).")
        return cls(
            base_url=base,
            username=os.getenv("THREEDX_USERNAME", "").strip(),
            password=os.getenv("THREEDX_PASSWORD", ""),
            search_text=os.getenv("THREEDX_SEARCH_TEXT", "Route").strip(),
            app_name=os.getenv("THREEDX_APP_NAME", "Route Management").strip(),
            headless=os.getenv("HEADLESS", "false").lower() == "true",
        )

    @property
    def platform_host(self):
        """Dashboard host. The start URL may be the 3DEXPERIENCE ID login page; then it is in ?service=."""
        url = urlparse(self.base_url)
        service = parse_qs(url.query).get("service")
        return urlparse(unquote(service[0])).netloc if service and ".iam." in url.netloc else url.netloc


class LocatorNotFound(Exception):
    pass


class LoginError(Exception):
    pass


# ---------------------------------------------------------------- self-healing locators

HEALED = []


def _page_of(scope):
    return scope.page if hasattr(scope, "page") else scope


def find(scope, name, candidates, timeout=15000):
    """Return the first candidate (label, builder) whose locator has a visible match.

    scope is a Page or Frame. Candidates are tried in order on every poll, so the primary
    wins whenever it matches; a backup that wins is recorded as healed.
    """
    deadline = time.monotonic() + timeout / 1000
    while True:
        for i, (label, build) in enumerate(candidates):
            try:
                loc = build(scope).first
                if loc.is_visible():
                    if i > 0:
                        _record_heal(name, candidates[0][0], label)
                    return loc
            except PlaywrightError:
                pass
        if time.monotonic() > deadline:
            tried = ", ".join(label for label, _ in candidates)
            raise LocatorNotFound(f"{name}: nothing visible after {timeout // 1000}s (tried {tried})")
        _page_of(scope).wait_for_timeout(500)


def find_optional(scope, name, candidates, timeout=3000):
    try:
        return find(scope, name, candidates, timeout)
    except LocatorNotFound:
        return None


def _record_heal(name, primary, used):
    entry = {"locator": name, "primary": primary, "used": used, "at": datetime.now().isoformat(timespec="seconds")}
    if any(h["locator"] == name and h["used"] == used for h in HEALED):
        return
    HEALED.append(entry)
    print(f"[HEAL] {name}: primary '{primary}' not found, used '{used}'")
    HEAL_FILE.parent.mkdir(exist_ok=True)
    HEAL_FILE.write_text(json.dumps(HEALED, indent=2), encoding="utf-8")


# ---------------------------------------------------------------- read-only guard

# Any request that could change a route or task is aborted before it leaves the browser.
WRITE_PATH = re.compile(
    r"approve|reject|reassign|delegate|complete|promote|demote|abort|resume|restart|"
    r"delete|remove|modify|update|setstate|accept|signoff",
    I,
)


# Platform notices (e.g. "will be updated to R2027x") sit over the top bar search box and block the
# click. Hidden by a style added on every page load: nothing is clicked or dismissed on the platform.
HIDE_NOTICES_JS = """(() => { const add = () => { if (document.getElementById('qa-hide-notices')) return;
  const s = document.createElement('style'); s.id = 'qa-hide-notices';
  s.textContent = '.upgrade-prompt-alert, .alert-root.alert { display: none !important; }';
  (document.head || document.documentElement).appendChild(s); };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', add); else add(); })();"""


def hide_platform_notices(context):
    """Keep platform-wide notice banners from covering the top bar (all pages of the context)."""
    context.add_init_script(HIDE_NOTICES_JS)


def install_read_only_guard(context, blocked):
    def handler(route):
        req = route.request
        path = urlparse(req.url).path
        login_host = ".iam." in urlparse(req.url).netloc
        if not login_host and (
            req.method in ("PUT", "PATCH", "DELETE") or (req.method == "POST" and WRITE_PATH.search(path))
        ):
            blocked.append(f"{req.method} {req.url}")
            print(f"[READ-ONLY] blocked {req.method} {path}")
            route.abort()
            return
        route.continue_()

    context.route("**/*", handler)


# ---------------------------------------------------------------- login

USERNAME = [
    ("#username", lambda s: s.locator("#username")),
    ("input[name=username]", lambda s: s.locator("input[name=username]")),
    ("label Email/Username", lambda s: s.get_by_label(re.compile("e-?mail|user", I))),
]
CONTINUE = [
    ("submit Continue", lambda s: s.locator("input[type=submit][value='Continue']")),
    ("button Continue", lambda s: s.get_by_role("button", name=re.compile("^continue$", I))),
    ("input[type=submit]", lambda s: s.locator("input[type=submit]")),
]
PASSWORD = [
    ("#password", lambda s: s.locator("#password")),
    ("input[name=password]", lambda s: s.locator("input[name=password]")),
    ("input[type=password]", lambda s: s.locator("input[type=password]")),
]
LOG_IN = [
    ("submit Log in", lambda s: s.locator("input[type=submit][value*='Log' i]")),
    ("button Log in", lambda s: s.get_by_role("button", name=re.compile("log ?in|sign ?in", I))),
    ("input[type=submit]", lambda s: s.locator("input[type=submit]")),
]
COOKIES = [
    ("button Accept All", lambda s: s.get_by_role("button", name=re.compile("^accept all$", I))),
]
LOGIN_ERROR = [
    ("role=alert", lambda s: s.get_by_role("alert")),
    (".error", lambda s: s.locator(".error, .alert-error, .uwa-alert-error, .help-block")),
]


def _on_login_page(page):
    return ".iam." in urlparse(page.url).netloc


def ensure_logged_in(page, cfg):
    """Open the platform home. Log in on the 3DEXPERIENCE ID page only if the saved session has expired."""
    page.goto(cfg.base_url, wait_until="domcontentloaded", timeout=90000)
    base_host = cfg.platform_host
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        if _on_login_page(page) and page.locator("#username, input[name=username], input[type=password]").first.is_visible():
            break
        if urlparse(page.url).netloc == base_host:
            print("Session still valid, login skipped.")
            return
        page.wait_for_timeout(500)
    else:
        raise LoginError(f"Neither the login page nor the dashboard loaded (at {page.url}).")

    if not cfg.username or not cfg.password:
        raise LoginError("Login needed but THREEDX_USERNAME / THREEDX_PASSWORD are empty in .env.")

    print("Logging in on 3DEXPERIENCE ID ...")
    cookie = find_optional(page, "cookie banner Accept All", COOKIES)
    if cookie:
        cookie.click()

    # Step 1: username, Continue. Step 2: password, Log in.
    user_box = find_optional(page, "username box", USERNAME, timeout=5000)
    if user_box:
        user_box.fill(cfg.username)
        find(page, "Continue button", CONTINUE).click()
    find(page, "password box", PASSWORD, timeout=30000).fill(cfg.password)
    find(page, "Log in button", LOG_IN).click()

    try:
        page.wait_for_url(lambda u: urlparse(u).netloc == base_host, timeout=90000)
    except PlaywrightError:
        err = find_optional(page, "login error message", LOGIN_ERROR, timeout=2000)
        msg = err.inner_text().strip() if err else "still on the login page after 90s (check for MFA / captcha)"
        raise LoginError(f"Login failed: {msg}")
    print("Logged in.")


# ---------------------------------------------------------------- top bar search -> app

SEARCH_BOX = [
    ("#Search-dropdown-v3-input", lambda s: s.locator("#Search-dropdown-v3-input")),
    ("topbar search input", lambda s: s.locator("#topbar .search-input input, .topbar-search input, #topbar-search input")),
    ("role=searchbox", lambda s: s.get_by_role("searchbox")),
    ("placeholder Search", lambda s: s.get_by_placeholder(re.compile("search", I))),
    ("input[type=search]", lambda s: s.locator("input[type=search]")),
]
# Clicking the search dropdown opens a line editor; the search text is typed there.
# The tag panel's "Tag name" line editor also counts as visible and swallows the text, so it is excluded.
SEARCH_TEXT = [
    ("lineEditor-lite input", lambda s: s.locator("input.wux-controls-lineEditor-lite:visible:not([placeholder='Tag name' i])")),
]
# Suggestion shown under the line editor once the search text is typed.
SEARCH_SUGGESTION = [
    ("li#Route_0", lambda s: s.locator("li#Route_0")),
]
SEARCH_ICON = [
    ("topbar search icon", lambda s: s.locator(".topbar-search .fonticon-search, #topbar-search .fonticon-search, .search-input .fonticon-search")),
    ("button Search", lambda s: s.get_by_role("button", name=re.compile("^search$", I))),
    ("[title=Search]", lambda s: s.locator("[title='Search' i]")),
    (".fonticon-search", lambda s: s.locator("#topbar .fonticon-search, .fonticon-search")),
]


def app_result(app_name):
    exact = re.compile(rf"^\s*{re.escape(app_name)}\s*$", I)
    return [
        ("text app name", lambda s: s.get_by_text(exact)),
        ("[title=app name]", lambda s: s.locator(f"[title='{app_name}' i]")),
        ("link app name", lambda s: s.get_by_role("link", name=exact)),
        ("button app name", lambda s: s.get_by_role("button", name=exact)),
    ]


def open_app_via_search(page, cfg):
    """Type the search text in the dashboard top bar, click the search icon and open the app."""
    print(f"Searching '{cfg.search_text}' in the top bar ...")
    box = find(page, "top bar search box", SEARCH_BOX, timeout=60000)
    box.click()
    text_box = find_optional(page, "search text box", SEARCH_TEXT, timeout=10000)
    if text_box:
        _type_search_text(page, text_box, cfg.search_text)
    else:
        box.fill(cfg.search_text)
    suggestion = find_optional(page, "search suggestion", SEARCH_SUGGESTION, timeout=15000)
    if suggestion:
        suggestion.click()
    else:
        icon = find_optional(page, "top bar search icon", SEARCH_ICON, timeout=5000)
        if icon:
            icon.click()
        else:
            _record_heal("top bar search icon", SEARCH_ICON[0][0], "Enter key")
            page.keyboard.press("Enter")

    print(f"Opening '{cfg.app_name}' from the results ...")
    result = find_optional(page, f"{cfg.app_name} result", app_result(cfg.app_name), timeout=30000)
    if result:
        result.click()
    return wait_for_widget_frame(page, cfg.app_name)


def _type_search_text(page, text_box, text, attempts=3):
    """Type key by key (fill() does not fire the keystrokes that show the suggestion list).

    The line editor opens with an animation and takes focus late, so the first keys can be
    lost ("oute" instead of "Route"). Wait for focus, then read the value back and retype.
    """
    for attempt in range(1, attempts + 1):
        text_box.click()
        try:
            text_box.page.wait_for_function("el => document.activeElement === el", arg=text_box.element_handle(), timeout=5000)
        except PlaywrightError:
            pass
        page.wait_for_timeout(300)
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(text, delay=80)
        page.wait_for_timeout(300)
        try:
            typed = text_box.input_value()
        except PlaywrightError:
            typed = text
        if typed.strip().lower() == text.lower():
            return
        print(f"  search text came out as '{typed}', retyping ({attempt}/{attempts}) ...")
    print(f"  search text still not '{text}', filling it instead")
    text_box.fill(text)
    page.keyboard.press("End")


def wait_for_widget_frame(page, app_name, timeout=90):
    """The app runs inside an iframe; return that frame once it has content."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for frame in page.frames:
            if frame == page.main_frame:
                continue
            if re.search(r"ENXROUT|route", frame.url, I):
                try:
                    frame.wait_for_load_state("domcontentloaded", timeout=5000)
                    if frame.locator("body *").count() > 5:
                        return frame
                except PlaywrightError:
                    pass
        page.wait_for_timeout(1000)
    raise LocatorNotFound(f"{app_name} widget iframe did not load in {timeout}s")
