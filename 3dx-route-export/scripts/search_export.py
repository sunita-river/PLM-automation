"""Export every top bar search result for "Route" and open each record one by one.

Flow: platform home -> 3DEXPERIENCE ID login (only if the saved session expired) -> plain
dashboard (no app open) -> top bar search "Route" -> Enter -> results page ("<n> Results",
#search-nb-result) -> scroll the result grid until all <n> rows are read -> for each record:
right click the title -> Open with -> Route Management, read its properties, Tasks and Content,
close the app -> output/Route_Search_<date>.xlsx (Results / Details / Tasks / Content).

Progress is saved after every record (output/search_progress.json), so a stopped run
continues where it left off. Read only: writes are blocked by threedx.install_read_only_guard.

Usage:
  .venv\\Scripts\\python.exe scripts\\search_export.py                  full run (all results)
  .venv\\Scripts\\python.exe scripts\\search_export.py --max-records 3  open only the first 3 records
  .venv\\Scripts\\python.exe scripts\\search_export.py --list-only      save the result list, open nothing
  .venv\\Scripts\\python.exe scripts\\search_export.py --fresh          ignore saved progress
"""
import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timezone

from openpyxl import Workbook
from playwright.sync_api import Error as PlaywrightError, sync_playwright

from excel_writer import CONTENT_COLUMNS, _write_sheet, to_date
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from route_export import (CONTENT_ALIASES, CONTENT_TAB, READ_GRID_JS, TASKS_TAB, Capture, _flatten, _record_lists,
                          records_from_grid, records_from_json)
from threedx import (HEALED, PROJECT_DIR, SEARCH_BOX, Config, LocatorNotFound, LoginError, _type_search_text,
                     ensure_logged_in, find, find_optional, hide_platform_notices, install_read_only_guard)

I = re.IGNORECASE
PROFILE_DIR = PROJECT_DIR / ".auth" / "profile"
OUTPUT_DIR = PROJECT_DIR / "output"
PROGRESS_FILE = OUTPUT_DIR / "search_progress.json"
STAMP = datetime.now().strftime("%Y-%m-%d_%H%M")
DEBUG_DIR = PROJECT_DIR / "debug" / f"search_{STAMP}"
# Second start point for the search: the dashboard tab with Route Management (no ?ticket=, it is single use).
DASHBOARD_URL = os.getenv("THREEDX_DASHBOARD_URL") or (
    "https://r1132101268131-indw2-ifwe.3dexperience.3ds.com/"
    "#dashboard:c95cb8a9-2363-4b7f-88da-ae7e74da475b/tabId:ASA6xzyA08Hm--LyUWZK/app:ENXROUT_AP")

RESULT_COLUMNS = ["Row", "Title", "Name", "Type", "Description", "Maturity State", "Owner", "Revision",
                  "Project", "Creation Date", "Modification Date"]

# ---------------------------------------------------------------- locators

RESULT_COUNT = [
    ("#search-nb-result", lambda s: s.locator("#search-nb-result")),
    (".set-detail-view-count", lambda s: s.locator(".set-detail-view-count")),
]
RESULT_GRID = [
    ("search datagrid row", lambda s: s.locator(".wux-layouts-datagridview-row[row-id]")),
    ("search item link", lambda s: s.locator(".search_item_hyperlink")),
]
# Clicking the top bar search opens the Guided Search dialog; the text goes in its line editor.
# The "Tag name" line editor of the tag panel also matches a plain lineEditor locator and
# swallows the text, so it is excluded.
GUIDED_SEARCH_INPUT = [
    ("Guided Search line editor", lambda s: s.locator(
        "input.wux-controls-lineEditor-lite:visible:not([placeholder='Tag name' i])")),
]
SEARCH_INPUT = [
    ("#Search-dropdown-v3-input", lambda s: s.locator("#Search-dropdown-v3-input")),
]
RUN_SEARCH = [
    (".run_btn_search", lambda s: s.locator(".run_ctn_search:visible, .run_btn_search:visible")),
    ("[title=Search]", lambda s: s.locator("#topbar [title='Search']:visible")),
]
ROUTE_SUGGESTION = [
    ("li#Route_0", lambda s: s.locator("li#Route_0")),
    ("suggestion Route", lambda s: s.locator("li:visible").filter(has_text=re.compile(r"^\s*route\s*$", I))),
]
# The blue Search button at the bottom of the Guided Search dialog (not the "Search" tab at the top).
SEARCH_BUTTON = [
    ("button Search", lambda s: s.get_by_role("button", name=re.compile(r"^\s*search\s*$", I))),
    ("blue Search button", lambda s: s.locator("[class*=btn i]:visible, [class*=button i]:visible").filter(
        has_text=re.compile(r"^\s*search\s*$", I)).last),
]
OPEN_WITH = [
    ("menu Open with", lambda s: s.locator(".wux-menu:visible, .dropdown-menu:visible, [role=menu]:visible")
        .get_by_text(re.compile(r"^\s*open with\s*$", I))),
    ("menuitem Open with", lambda s: s.get_by_role("menuitem", name=re.compile("open with", I))),
    ("text Open with", lambda s: s.get_by_text(re.compile(r"^\s*open with\s*$", I))),
]
OPEN_WITH_ROUTE_MGMT = [
    ("menu Route Management", lambda s: s.locator(".wux-menu:visible, .dropdown-menu:visible, [role=menu]:visible")
        .get_by_text(re.compile(r"^\s*route management\s*$", I))),
    ("menuitem Route Management", lambda s: s.get_by_role("menuitem", name=re.compile("^route management$", I))),
    ("text Route Management", lambda s: s.get_by_text(re.compile(r"^\s*route management\s*$", I))),
]
CLOSE_APP = [
    ("app header Close", lambda s: s.locator(".widget-header [title='Close' i]:visible, .moduleHeader [title='Close' i]:visible")),
    ("[title=Close]", lambda s: s.locator("[title='Close' i]:visible")),
    ("close icon", lambda s: s.locator(".fonticon-close:visible, .wux-ui-3ds-close:visible")),
]

# Reads the rows of the virtual result grid that are in the DOM right now, keyed by row-id.
# Cells give the visible columns; the Type cell also carries data-ds6w* attributes with
# state, owner, revision, project and name, which fill columns scrolled out of view.
READ_ROWS_JS = r"""
() => {
  const txt = e => (e.innerText || e.textContent || '').trim().replace(/\s+/g, ' ');
  const headers = {};
  for (const h of document.querySelectorAll('.wux-layouts-datagridview-column-header-cell[column-id]'))
    headers[h.getAttribute('column-id')] = txt(h);
  const attr = {'data-ds6wstatus-nls': 'Maturity State', 'data-ds6wresponsible-nls': 'Owner',
                'data-ds6wgrevision-nls': 'Revision', 'data-ds6wproject-nls': 'Project',
                'data-ds6widentifier-nls': 'Name', 'data-ds6wmodified-nls': 'Modification Date',
                'data-ds6wtype-nls': 'Type'};
  const rows = {};
  for (const c of document.querySelectorAll('.wux-layouts-datagridview-cell[row-id][column-id]')) {
    const r = parseInt(c.getAttribute('row-id'), 10), name = headers[c.getAttribute('column-id')];
    if (r < 0 || !name) continue;
    const row = rows[r] = rows[r] || {};
    const value = txt(c);
    if (value) row[name] = value;
    const t = c.querySelector('[data-ds6wtype]');
    if (t) for (const [a, col] of Object.entries(attr)) {
      const v = (t.getAttribute(a) || '').trim();
      if (v && !row[col]) row[col] = v;
    }
  }
  return rows;
}
"""

# Same as READ_ROWS_JS but only the grid that is on screen (the hidden tabs keep their grids in the DOM).
VISIBLE_GRID_JS = READ_ROWS_JS.replace(
    "for (const h of document.querySelectorAll('.wux-layouts-datagridview-column-header-cell[column-id]'))",
    "for (const h of [...document.querySelectorAll('.wux-layouts-datagridview-column-header-cell[column-id]')].filter(e => e.offsetParent))",
).replace(
    "for (const c of document.querySelectorAll('.wux-layouts-datagridview-cell[row-id][column-id]')) {",
    "for (const c of [...document.querySelectorAll('.wux-layouts-datagridview-cell[row-id][column-id]')].filter(e => e.offsetParent)) {",
)

# Scrolls every vertical scroller of the result grid (left pinned + main) to y, or one screen down.
SCROLL_JS = r"""
(y) => {
  const s = [...document.querySelectorAll('.wux-scroller')]
    .filter(e => e.offsetHeight && e.scrollHeight > e.clientHeight + 20 && e.querySelector('[row-id]'));
  if (!s.length) return null;
  for (const el of s) el.scrollTop = (y === null) ? el.scrollTop + el.clientHeight * 0.8 : y;
  const el = s[0];
  return {top: el.scrollTop, max: el.scrollHeight - el.clientHeight};
}
"""

# Label / value pairs from whatever opened after a record click (property panels, forms, 2-column tables).
READ_DETAILS_JS = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const txt = e => (e.innerText || e.textContent || '').trim().replace(/\s+/g, ' ');
  const out = {};
  const put = (k, v) => { k = (k || '').replace(/[:*]\s*$/, '').trim(); v = (v || '').trim();
                          if (k && v && k.length < 60 && k !== v && !(k in out)) out[k] = v.slice(0, 2000); };
  for (const dt of document.querySelectorAll('dt')) if (vis(dt)) put(txt(dt), dt.nextElementSibling ? txt(dt.nextElementSibling) : '');
  for (const tr of document.querySelectorAll('tr')) {
    if (!vis(tr)) continue;
    const c = tr.querySelectorAll('th, td');
    if (c.length === 2) put(txt(c[0]), txt(c[1]));
  }
  for (const l of document.querySelectorAll('label, [class*=label i]')) {
    if (!vis(l) || l.closest('.wux-layouts-datagridview-cell')) continue;
    const next = l.nextElementSibling;
    if (next && vis(next)) put(txt(l), next.value || txt(next));
  }
  return out;
}
"""


# ---------------------------------------------------------------- steps

def wait_for_dashboard(page):
    """The top bar search is greyed out until the dashboard has loaded; clicking it earlier opens
    a stray menu and the text goes nowhere."""
    try:
        page.wait_for_load_state("networkidle", timeout=30000)
    except PlaywrightError:
        pass
    page.wait_for_function("""() => { const e = document.querySelector('#Search-dropdown-v3-input');
        return e && !e.disabled && !e.readOnly && e.offsetWidth > 0; }""", timeout=90000)
    page.wait_for_timeout(3000)


def run_search(page, cfg, settle=True):
    """One search attempt from the current dashboard; returns the result count element or None.
    settle=False when the dashboard is already loaded (searching again after a record)."""
    if settle:
        try:
            wait_for_dashboard(page)
        except PlaywrightError:
            print("  search box still greyed out after 90s, trying anyway")
    page.keyboard.press("Escape")  # close any menu left open
    find(page, "top bar search box", SEARCH_BOX, timeout=60000).click()
    page.wait_for_timeout(1500 if settle else 800)
    text_box = (find_optional(page, "Guided Search input", GUIDED_SEARCH_INPUT, timeout=8000)
                or find(page, "top bar search input", SEARCH_INPUT, timeout=15000))
    _type_search_text(page, text_box, cfg.search_text)

    # Enter only highlights the "Route" type in the suggestion list: pick it, then Enter runs the search.
    suggestion = find_optional(page, "Route type suggestion", ROUTE_SUGGESTION, timeout=5000)
    if suggestion:
        suggestion.click()
        page.wait_for_timeout(800)
    page.keyboard.press("Enter")
    count = find_optional(page, "result count", RESULT_COUNT, timeout=15000)
    if not count:
        button = find_optional(page, "Guided Search button", SEARCH_BUTTON, timeout=5000)
        if button:
            button.click()
        count = find_optional(page, "result count", RESULT_COUNT, timeout=45000)
    return count


def search_results(page, cfg):
    """Type the search text, press Enter and return the result count shown on the results page.

    First from the plain dashboard; if no results page appears, again from the Route dashboard
    tab (THREEDX_DASHBOARD_URL)."""
    starts = [f"https://{cfg.platform_host}/", DASHBOARD_URL]
    count = None
    for n, url in enumerate(starts, 1):
        print(f"Searching '{cfg.search_text}' in the top bar (attempt {n}/{len(starts)}) ...")
        page.goto(url, wait_until="domcontentloaded", timeout=90000)
        count = run_search(page, cfg)
        if count:
            break
        save_snapshot(page, f"search_attempt_{n}")
    if not count:
        raise LocatorNotFound("results page did not appear after searching from both dashboards")
    find(page, "result grid", RESULT_GRID, timeout=60000)
    total = 0
    for _ in range(20):  # the count is filled in after the first rows arrive
        total = int(re.sub(r"\D", "", count.inner_text()) or 0)
        if total:
            break
        page.wait_for_timeout(500)
    print(f"Results page shown: {total} results.")
    return total


def read_all_rows(page, total):
    """Scroll the virtual grid top to bottom and keep every row (the grid also loads more rows near the end)."""
    rows = {}
    page.evaluate(SCROLL_JS, 0)
    page.wait_for_timeout(1000)
    stuck = 0
    while len(rows) < total and stuck < 20:  # the last batches load slowly
        before = len(rows)
        rows.update({int(k): v for k, v in page.evaluate(READ_ROWS_JS).items()})
        pos = page.evaluate(SCROLL_JS, None)
        page.wait_for_timeout(700 if pos and pos["top"] < pos["max"] else 2000)
        stuck = stuck + 1 if len(rows) == before else 0
        if len(rows) // 200 != before // 200:
            print(f"  {len(rows)}/{total} rows read ...")
    rows.update({int(k): v for k, v in page.evaluate(READ_ROWS_JS).items()})
    page.evaluate(SCROLL_JS, 0)
    print(f"{len(rows)} of {total} result rows read.")
    if len(rows) < total:
        print("  The grid stopped loading before the count shown; the missing rows are listed in the summary.")
    return [{"Row": r + 1, **rows[r]} for r in sorted(rows)]


def scroll_to_row(page, row_id, title=""):
    """Bring one row into the DOM and on screen; returns its title link."""
    # :visible - the page also keeps a hidden copy of the grid from the search tab preview.
    link = page.locator(f".search_item_hyperlink[row-id='{row_id}']:visible").first
    height = page.evaluate("() => { const r = document.querySelector('.wux-layouts-datagridview-row[row-id]');"
                           " return r ? r.getBoundingClientRect().height || 26 : 26; }")
    target = max(0, row_id * height - 3 * height)
    # The grid loads rows in batches as it is scrolled, so deep rows are reached step by step.
    deadline = time.monotonic() + 180
    while time.monotonic() < deadline:
        if link.count() and link.is_visible():
            if title and link.inner_text().strip() != title:
                # Routes added since the list was read shift the rows: use the row with this title.
                same = page.locator(".search_item_hyperlink:visible").filter(
                    has_text=re.compile(rf"^\s*{re.escape(title)}\s*$"))
                if not same.count():
                    raise LocatorNotFound(f"row {row_id + 1} is now '{link.inner_text().strip()}', "
                                          f"'{title}' not found near it")
                link = same.first
            try:
                link.scroll_into_view_if_needed(timeout=10000)
                return link
            except PlaywrightError:
                # The grid redrew the row while scrolling ("not attached to the DOM"): look it up again.
                link = page.locator(f".search_item_hyperlink[row-id='{row_id}']:visible").first
                page.wait_for_timeout(800)
                continue
        pos = page.evaluate(SCROLL_JS, target)
        page.wait_for_timeout(600 if pos and pos["top"] >= target - height else 1200)
    raise LocatorNotFound(f"row {row_id + 1} did not come into view")


def results_visible(page):
    return page.locator(".search_item_hyperlink:visible").count() > 0


def back_to_results(page, cfg):
    """Opening a record in Route Management closes the search results and clears the top bar
    query, so search again on the dashboard that is already loaded (no reload); a full new
    search from the start page is only the fallback."""
    if results_visible(page):
        return
    try:
        if run_search(page, cfg, settle=False) and find_optional(page, "result grid", RESULT_GRID, timeout=20000):
            page.wait_for_timeout(1000)
            return
    except (PlaywrightError, LocatorNotFound):
        pass
    search_results(page, cfg)


def visible_grid_rows(frame):
    """All columns of the visible data grid in the frame (e.g. the Content tab), one dict per row."""
    rows = frame.evaluate(VISIBLE_GRID_JS)
    return [rows[k] for k in sorted(rows, key=int) if rows[k]]


def wait_for_new_route_frame(page, frames_before, timeout=90):
    """The Route Management app opened for the record runs in a new iframe; return it once it has content."""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for f in page.frames:
            if id(f) in frames_before or not re.search(r"ENXROUT|route", f.url, I):
                continue
            try:
                f.wait_for_load_state("domcontentloaded", timeout=5000)
                if f.locator("body *").count() > 5:
                    return f
            except PlaywrightError:
                pass
        page.wait_for_timeout(1000)
    raise LocatorNotFound(f"Route Management did not open in {timeout}s")


def open_record(page, capture, row, save_debug):
    """Right click one record title -> Open with -> Route Management, read the route, close the app."""
    link = scroll_to_row(page, row["Row"] - 1, row.get("Title", ""))
    frames_before = {id(f) for f in page.frames}
    mark = capture.mark()

    link.click(button="right")
    find(page, "context menu Open with", OPEN_WITH, timeout=10000).hover()
    find(page, "Open with Route Management", OPEN_WITH_ROUTE_MGMT, timeout=10000).click()
    frame = wait_for_new_route_frame(page, frames_before)
    frame.page.wait_for_timeout(3000)

    detail = {"Row": row["Row"], "Title": row.get("Title", ""), "Name": row.get("Name", ""), "Opened": frame.url[:300]}
    try:
        for k, v in frame.evaluate(READ_DETAILS_JS).items():
            detail.setdefault(k, v)
    except PlaywrightError:
        pass

    tasks, content = [], []
    tab = find_optional(frame, "Tasks tab", TASKS_TAB, timeout=10000)
    if tab:
        tab.click()
        page.wait_for_timeout(2000)
        tasks = all_task_fields(frame, capture.since(mark))
    tab = find_optional(frame, "Content tab", CONTENT_TAB, timeout=5000)
    if tab:
        tab.click()
        page.wait_for_timeout(2000)
        content = visible_grid_rows(frame)
    if save_debug:
        save_snapshot(page, f"record_{row['Row']:04d}")

    # Close the app; the results page is underneath (still "visible" to Playwright while covered,
    # so the check is on the app iframe instead).
    for _ in range(3):
        if not app_open(frame):
            break
        close = find_optional(page, "close Route Management", CLOSE_APP, timeout=3000)
        if close:
            close.click()
        else:
            page.keyboard.press("Escape")
        page.wait_for_timeout(1500)
    # The data is kept either way; the caller searches again if the results page did not come back.
    return detail, tasks, content


def all_task_fields(frame, items):
    """Every column of the Tasks tab, as shown in its grid; if the grid has no headers, every
    field of the task records the app loaded (JSON)."""
    grid = frame.evaluate(READ_GRID_JS)
    headers = [h for h in grid["headers"]]
    if headers and grid["rows"]:
        rows = [{h: c for h, c in zip(headers, cells) if h and c} for cells in grid["rows"]]
        return [r for r in rows if r]
    best = []
    for item in items:
        if re.search(r"task", item["url"], I):
            for lst in _record_lists(item["body"]):
                if len(lst) > len(best):
                    best = [_flatten(r) for r in lst]
    return best


def app_open(frame):
    try:
        return frame.frame_element().is_visible()
    except PlaywrightError:  # detached: the app was closed
        return False


def save_snapshot(page, name):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    try:
        page.screenshot(path=str(DEBUG_DIR / f"{name}.png"))
        for n, f in enumerate(page.frames):
            (DEBUG_DIR / f"{name}_frame_{n:02d}.html").write_text(f"<!-- {f.url} -->\n" + f.content(), encoding="utf-8")
    except PlaywrightError:
        pass


# ---------------------------------------------------------------- progress and output

def load_progress(total, fresh):
    empty = {"total": total, "rows": [], "details": {}, "tasks": {}, "content": {}}
    if fresh or not PROGRESS_FILE.exists():
        return empty
    data = {**empty, **json.loads(PROGRESS_FILE.read_text(encoding="utf-8"))}
    if data.get("total") != total:
        # Routes were added or removed since: the list is read again, records already opened are kept
        # and moved to their new row by Name (see remap_progress).
        print(f"Saved progress is for {data.get('total')} results, now {total}: reading the list again, "
              f"keeping the {len(data['details'])} records already opened.")
        data["total"], data["stale_rows"] = total, data["rows"]
        data["rows"] = []
    print(f"Continuing saved progress: {len(data['details'])} records already opened.")
    return data


def remap_progress(progress):
    """After the list was read again, move opened records to their new row number (matched by Name)."""
    old = {str(r["Row"]): r.get("Name") for r in progress.pop("stale_rows", [])}
    new = {r.get("Name"): str(r["Row"]) for r in progress["rows"] if r.get("Name")}
    for part in ("details", "tasks", "content"):
        moved = {}
        for k, v in progress[part].items():
            key = new.get(old.get(k))
            if key:
                if part == "details":
                    v = {**v, "Row": int(key)}
                moved[key] = v
        progress[part] = moved


def save_progress(progress):
    PROGRESS_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = PROGRESS_FILE.with_suffix(".tmp")
    tmp.write_text(json.dumps(progress, indent=1, ensure_ascii=False), encoding="utf-8")
    tmp.replace(PROGRESS_FILE)


THIN = Side(style="thin", color="BFBFBF")
# Task dates from the Route Management service are UTC: "2026-11-04T07:55:46.000Z" and also the
# ones without a zone, "2026-11-04 07:55:51" (completed 5 s after the task was originated).
UTC_TEXT = re.compile(r"^\s*\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(\.\d+)?Z?\s*$")
# 2026-11-06, 2026-11-06 07:55:47, 2026-11-04T07:55:46.000Z, 11/5/2026 12:36:05 PM
DATE_TEXT = re.compile(r"^\s*(\d{4}-\d{2}-\d{2}([ T]\d{2}:\d{2}(:\d{2})?(\.\d+)?Z?)?|\d{1,2}/\d{1,2}/\d{4}( \d{1,2}:\d{2}(:\d{2})?( [AP]M)?)?)\s*$", I)

# Task fields from the app (JSON keys) in reading order with readable headers; any other field
# follows these, and internal IDs go last.
TASK_LABELS = {
    "taskOrder": "Task Order", "name": "Task Name", "title": "Task Title", "type": "Task Type",
    "taskAssignee": "Assignee", "taskAssigneeUsername": "Assignee Username", "assigneeType": "Assignee Type",
    "assigneeTitle": "Assignee Title", "assigneeState": "Assignee State", "taskAction": "Action",
    "stateNLS": "Task State", "current": "Current State", "attrTaskStatus": "Task Status",
    "approvalStatus": "Approval Status", "taskDueDate": "Due Date", "taskActualCompletionDate": "Completed Date",
    "instructions": "Instructions", "comments": "Comments", "reviewersComments": "Reviewer Comments",
    "routeTaskPriority": "Priority", "dateOffsetFrom": "Due Date Offset From", "dueDateOffset": "Due Date Offset",
    "originator": "Originator", "originatorFullName": "Originator Username", "originated": "Originated Date",
    "modified": "Modified Date", "revision": "Revision",
}
# Task fields left out of the Tasks sheet: settings, user group data, candidates and internal IDs.
TASK_HIDDEN_FIELDS = {
    "taskRequirement", "needsOwnerReview", "allowDelegation", "assigneeSetDueDate", "parallelNodeProcessionRule",
    "userGroupAction", "userGroupInfo", "routeOwnerTask", "routeOwnerUGChoice", "templateTask", "isInitialTask",
    "taskEsignPolicy", "taskAcceptLink", "routeOwnerUGName", "routeOwnerUGUUID", "candidateAssignee",
    "candidateAssigneeId", "id", "identifier", "relId", "cestamp", "i3dx", "source", "relativePath",
    "originatorPhId", "assigneeId", "assigneePID", "ugUUID", "typeicon", "children",
}
_HIDDEN = {k.lower() for k in TASK_HIDDEN_FIELDS}


def readable_task(task):
    """Rename to readable headers in TASK_LABELS order; other fields next; TASK_HIDDEN_FIELDS left out."""
    out = {TASK_LABELS[k]: task[k] for k in TASK_LABELS if k in task}
    out.update({re.sub(r"(?<!^)(?=[A-Z])", " ", k).title(): v for k, v in task.items()
                if k not in TASK_LABELS and k.lower() not in _HIDDEN})
    return out


def group_by_route(ws):
    """One block per route: Route Title merged down the route's rows, alternate light shading
    per route and a thicker line under each block. (Excel tables cannot hold merged cells, so
    these sheets keep a plain filter instead.)"""
    shade = PatternFill("solid", fgColor="EAF1FB")
    thick = Side(style="medium", color="1F4E79")
    start = 2
    block = 0
    for r in range(2, ws.max_row + 2):
        if r <= ws.max_row and ws.cell(r, 1).value == ws.cell(start, 1).value:
            continue
        end = r - 1  # rows start..end belong to one route
        if end > start:
            ws.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)
        title = ws.cell(start, 2)
        title.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        title.font = Font(bold=True)
        for row in ws.iter_rows(min_row=start, max_row=end):
            for c in row:
                if block % 2:
                    c.fill = shade
                if c.row == end:
                    c.border = Border(left=THIN, right=THIN, top=c.border.top, bottom=thick)
        block += 1
        start = r
    ws.column_dimensions["B"].width = 40


def format_table(ws, as_table=True):
    """Excel table with banded rows and filters, thin borders, wrapped long text, widths capped at 60."""
    if as_table:
        ws.auto_filter.ref = None  # the table brings its own filter buttons
    last_col = get_column_letter(ws.max_column)
    if as_table and ws.max_row > 1:
        table = Table(displayName=f"t{ws.title}", ref=f"A1:{last_col}{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(table)
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
            if cell.row > 1 and isinstance(cell.value, str) and DATE_TEXT.match(cell.value):
                cell.value = to_date(cell.value) or cell.value  # date text in columns not named "... Date"
            if isinstance(cell.value, datetime):
                # Written as text (05-Nov-2026): previews such as the VS Code Excel viewer ignore date
                # formats and show the raw serial number (46331.63) for real date cells.
                cell.value = cell.value.strftime("%d-%b-%Y")
                cell.number_format = "@"
                cell.alignment = Alignment(vertical="top", horizontal="right")
                continue
            cell.alignment = Alignment(vertical="top", wrap_text=cell.row > 1 and len(str(cell.value or "")) > 60)
    ws.row_dimensions[1].height = 22
    for i in range(1, ws.max_column + 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = min(ws.column_dimensions[letter].width or 12, 60)


# Completed Date also comes as "11/5/2026 9:47:25 AM" - UTC as well (it matches the task's UTC Modified time).
UTC_US_TEXT = re.compile(r"^\s*\d{1,2}/\d{1,2}/\d{4} \d{1,2}:\d{2}:\d{2} [AP]M\s*$", I)


def utc_to_local(text):
    text = text.strip()
    if UTC_US_TEXT.match(text):
        t = datetime.strptime(text, "%m/%d/%Y %I:%M:%S %p")
    else:
        t = datetime.fromisoformat(text.replace("Z", "").replace(" ", "T")[:19])
    return t.replace(tzinfo=timezone.utc).astimezone().replace(tzinfo=None)


def local_task(task):
    """All date texts of one task (service data, UTC) as local datetimes."""
    return {k: utc_to_local(v) if isinstance(v, str) and (UTC_TEXT.match(v) or UTC_US_TEXT.match(v)) else v
            for k, v in task.items()}


def task_date(value):
    """Task date as local datetime: ISO times ending in Z are UTC, the others are already local."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, str) and (UTC_TEXT.match(value) or UTC_US_TEXT.match(value)):
        return utc_to_local(value)
    return to_date(value) if value not in (None, "", "None") else None


def days_between(start, end):
    return round((end - start).total_seconds() / 86400, 1)


def with_task_days(task):
    """Days Taken (Originated -> Completed) and Overdue (Completed, or today if open, after Due Date),
    placed right after Completed Date."""
    start, due, done = (task_date(task.get(k)) for k in ("Originated Date", "Due Date", "Completed Date"))
    extra = {"Days Taken": days_between(start, done) if start and done else None,
             "Overdue": ("Yes" if (done or datetime.now()) > due else "No") if due else None}
    out = {}
    for k, v in task.items():
        out[k] = v
        if k == "Completed Date":
            out.update(extra)
    return out if "Days Taken" in out else {**out, **extra}


ROUTE_DAY_COLUMNS = ["Route Status", "Route Start", "Route Completed", "Days to Complete", "Days Open So Far",
                     "Tasks", "Tasks Completed", "Tasks Overdue"]


def route_days(tasks):
    """Route duration from its tasks: start = earliest Originated, end = latest Completed once every
    task is complete. Open routes get the days open so far instead."""
    if not tasks:
        return {"Route Status": "No tasks"}
    starts = [d for d in (task_date(t.get("originated")) for t in tasks) if d]
    ends = [task_date(t.get("taskActualCompletionDate")) for t in tasks]
    done = [d for d in ends if d]
    complete = len(done) == len(tasks)
    start = min(starts) if starts else None
    end = max(done) if complete else None
    overdue = sum(1 for t in tasks if with_task_days({"Originated Date": t.get("originated"), "Due Date": t.get("taskDueDate"),
                                                      "Completed Date": t.get("taskActualCompletionDate")})["Overdue"] == "Yes")
    return {
        "Route Status": "Completed" if complete else "In progress",
        "Route Start": start,
        "Route Completed": end,
        "Days to Complete": days_between(start, end) if start and end else None,
        "Days Open So Far": days_between(start, datetime.now()) if start and not complete else None,
        "Tasks": len(tasks), "Tasks Completed": len(done), "Tasks Overdue": overdue,
    }


def write_search_workbook(path, progress):
    rows = progress["rows"]
    keys = sorted(progress["details"], key=int)
    details = [progress["details"][k] for k in keys]
    by_row = {str(r["Row"]): r for r in rows}

    def cell(v):  # lists / dicts from the app's JSON become readable text
        if isinstance(v, str) and UTC_TEXT.match(v):  # 2026-11-04T20:10:00.000Z is UTC: show the local date
            return utc_to_local(v)
        if isinstance(v, list):
            return ", ".join(cell(x) for x in v)
        if isinstance(v, dict):
            return json.dumps(v, ensure_ascii=False)
        return v

    def keyed(part):
        out = []
        for k in keys:
            for x in progress[part].get(k, []):
                if part == "tasks":
                    x = with_task_days(readable_task(local_task(x)))
                elif str(x.get("Type", "")).lower().endswith("task"):  # tasks picked up as content
                    continue
                out.append({"Row": int(k), "Route Title": by_row.get(k, {}).get("Title", ""),
                            "Route Name": by_row.get(k, {}).get("Name", ""),
                            **{f: cell(v) for f, v in x.items()}})
        return out

    details = [{**{f: cell(v) for f, v in d.items() if f != "Opened"}, **route_days(progress["tasks"].get(str(d["Row"]), []))}
               for d in details]

    def columns(first, records):
        return first + [c for c in dict.fromkeys(k for r in records for k in r) if c not in first]

    wb = Workbook()
    wb.active.title = "Results"
    sheets = [
        (wb.active, columns(RESULT_COLUMNS, rows), rows),
        (wb.create_sheet("Details"), columns(["Row", "Title", "Name"] + ROUTE_DAY_COLUMNS, details), details),
        (wb.create_sheet("Tasks"), columns(["Row", "Route Title", "Route Name"], keyed("tasks")), keyed("tasks")),
        (wb.create_sheet("Content"), columns(["Row", "Route Title"] + CONTENT_COLUMNS, keyed("content")), keyed("content")),
    ]
    for ws, cols, records in sheets:
        _write_sheet(ws, cols, records)
        if ws.title in ("Tasks", "Content"):
            format_table(ws, as_table=False)
            group_by_route(ws)
        else:
            format_table(ws)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--max-records", type=int, default=0, help="only open the first N records (testing)")
    ap.add_argument("--list-only", action="store_true", help="save the result list, do not open records")
    ap.add_argument("--fresh", action="store_true", help="ignore saved progress and start again")
    ap.add_argument("--headless", action="store_true", help="run without a visible browser")
    args = ap.parse_args()

    cfg = Config.from_env()
    report = {"failed": [], "blocked": []}
    for d in (PROFILE_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        headless = args.headless or cfg.headless
        # Headed: maximised window, page sized to the window (no fixed viewport hanging off screen).
        context = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), headless=headless,
            **({"viewport": {"width": 1600, "height": 950}} if headless else {"no_viewport": True, "args": ["--start-maximized"]}),
        )
        install_read_only_guard(context, report["blocked"])
        hide_platform_notices(context)
        page = context.pages[0] if context.pages else context.new_page()
        capture = Capture(page)
        progress = {"rows": [], "details": {}, "tasks": {}, "content": {}}
        try:
            ensure_logged_in(page, cfg)
            total = search_results(page, cfg)
            if not total:
                raise LocatorNotFound("the results page shows 0 results")
            progress = load_progress(total, args.fresh)
            if len(progress["rows"]) < total:
                progress["rows"] = read_all_rows(page, total)
                remap_progress(progress)
                save_progress(progress)
            rows = progress["rows"]

            todo = [] if args.list_only else rows[: args.max_records] if args.max_records else rows
            for n, row in enumerate(todo, 1):
                key = str(row["Row"])
                if key in progress["details"]:
                    continue
                label = row.get("Title") or row.get("Name") or f"row {row['Row']}"
                try:
                    detail, tasks, content = open_record(page, capture, row, save_debug=n <= 3)
                    progress["details"][key] = detail
                    progress["tasks"][key] = tasks
                    progress["content"][key] = content
                    save_progress(progress)
                    capture.items.clear()  # keep memory flat over 1000+ records
                    print(f"  [{n}/{len(todo)}] {label}: {len(detail) - 4} fields, {len(tasks)} tasks, {len(content)} content")
                    back_to_results(page, cfg)
                except (PlaywrightError, LocatorNotFound) as e:
                    msg = str(e).splitlines()[0]
                    report["failed"].append(f"row {row['Row']} {label}: {msg}")
                    print(f"  [{n}/{len(todo)}] {label}: FAILED {msg}")
                    save_snapshot(page, f"failed_{row['Row']:04d}")
                    page.keyboard.press("Escape")
                    close = find_optional(page, "close Route Management", CLOSE_APP, timeout=2000)
                    if close:
                        close.click()
                    back_to_results(page, cfg)
        except (LoginError, LocatorNotFound, PlaywrightError) as e:
            print(f"\nSTOPPED: {str(e).splitlines()[0]}")
            save_snapshot(page, "stopped")
            if not progress["rows"]:
                return 1
        finally:
            context.close()

    details = progress["details"]
    out = write_search_workbook(OUTPUT_DIR / f"Route_Search_{STAMP}.xlsx", progress)
    total = progress.get("total", len(progress["rows"]))
    missing = sorted(set(range(1, total + 1)) - {r["Row"] for r in progress["rows"]})

    print("\n================ Summary ================")
    print(f"Workbook        : {out}")
    print(f"Results shown   : {total}")
    print(f"Rows saved      : {len(progress['rows'])}" + (f" (missing rows: {missing[:20]}{' ...' if len(missing) > 20 else ''})" if missing else ""))
    print(f"Records opened  : {len(details)} of {len(progress['rows'])}")
    print(f"Tasks / Content : {sum(map(len, progress['tasks'].values()))} / {sum(map(len, progress['content'].values()))}")
    print(f"Healed locators : {len(HEALED)}" + (f" (see {OUTPUT_DIR / 'healed-locators.json'})" if HEALED else ""))
    print(f"Blocked writes  : {len(report['blocked'])}")
    print(f"Failed records  : {len(report['failed'])}")
    for f in report["failed"]:
        print(f"  - {f}")
    if len(details) < len(progress["rows"]) and not args.list_only and not args.max_records:
        print("Run the same command again to continue with the remaining records.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
