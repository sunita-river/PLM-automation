"""Tasks of N fully approved routes from the top bar "Route" search, with creation / modification
dates and time taken per task and per route - one search, no record opened.

Flow: platform home -> 3DEXPERIENCE ID login (only if the saved session expired) -> top bar
search "Route" (once) -> go down the results; routes still in process (search status) are skipped
-> for each completed route one read-only GET
  /enovia/resources/v1/modeler/routes/<id>?$include=tasks
(the same call the Route Management app makes when a route is opened) -> kept only when every
task is finished (approval tasks Approved, notifications Sent, comments given; none open or
rejected) -> stop at N routes (default 2) -> output/Route_Tasks_<date>.xlsx (one Tasks sheet).

Task columns follow the task box in the Route Management flowchart (Task Title, Status, Assignee,
State, Priority) plus Created, Due, Completed, Modified, Time Taken and Late. Status, Priority,
Due, Time Taken and Late are hidden in the Tasks sheet (still in the file; Unhide to see them).
Route Created and Route Completed (last task completed) sit next to Route Title, merged per route;
there is no separate Routes sheet. Route Completed stays empty while any task is open.

--all reads every route in the results, in process ones too (about 1 - 1.5 hours for ~1500 routes).
When PLM_SHEET_DRIVE_DIR is set in .env, the rows are saved as Route_Tasks_latest.csv in that Google
Drive folder and the sheet's Apps Script (docs/apps-script/plmRouteTasksImport.gs) imports them into
the Tasks tab. Otherwise, with PLM_SHEET_* set, the rows replace the contents of the shared Google Sheet tab
(sheet_push.py; route values repeated on every row there, no merged cells). The push goes through
the same Chrome profile, signed in to Google: the first time a window opens for the Google sign-in.

The route id comes from the search service response (resourceid of each result), so the results
page stays as it is: no right click, no Open with, no closing the app and no searching again.

Usage:
  .venv\\Scripts\\python.exe scripts\\route_tasks.py              2 approved routes
  .venv\\Scripts\\python.exe scripts\\route_tasks.py --routes 10  10 approved routes
  .venv\\Scripts\\python.exe scripts\\route_tasks.py --all        every route (weekly full run)
"""
import argparse
import json
import os
import re
import sys
from datetime import datetime
from urllib.parse import quote

from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from playwright.sync_api import Error as PlaywrightError, sync_playwright

from excel_writer import _write_sheet
from sheet_push import SheetPushError, drive_dir, push_via_browser, sheet_settings, write_drive_csv
from route_export import _flatten
from search_export import (OUTPUT_DIR, PROFILE_DIR, SCROLL_JS, STAMP, format_table, group_by_route,
                           save_snapshot, search_results, task_date)
from threedx import (HEALED, Config, LocatorNotFound, LoginError, ensure_logged_in, hide_platform_notices,
                     install_read_only_guard)

I = re.IGNORECASE
ROUTE_INCLUDE = "$include=tasks&$fields=all"
MAX_SCROLLS = 60  # grid pages to load while looking for approved routes
MAX_SCROLLS_ALL = 3000  # --all: scroll until every result is loaded (stops earlier when the grid stops growing)
STALL_SCROLLS = 8  # scrolls in a row without a new route = the grid has nothing more to load

TASK_COLUMNS = ["Row", "Route Title", "Route Created", "Route Completed", "Task Title", "Status", "Assignee", "State", "Priority",
                "Created", "Due", "Completed", "Modified", "Days Taken", "Time Taken", "Late"]
# Kept in the Tasks sheet but hidden (right click a column header > Unhide to see them again).
HIDDEN_TASK_COLUMNS = ["Status", "Priority", "Due", "Time Taken", "Late"]
# Route-level columns: one value per route, merged down the route's tasks like Route Title.
MERGED_ROUTE_COLUMNS = ["Route Created", "Route Completed"]


class SearchIds:
    """Routes in result order (id, name, title, status), read from the search service responses."""

    def __init__(self, context):
        self.items = []     # {"id", "name", "title", "status"} in result order
        self.space = None   # 3DSpace base URL, e.g. https://<tenant>-space.3dexperience.3ds.com/enovia
        context.on("response", self._on_response)

    def _on_response(self, resp):
        try:
            if not self.space:
                m = re.match(r"(https://[^/]+-space\.[^/]+/enovia)/resources/", resp.url)
                if m:
                    self.space = m.group(1)
            if "federated/search" not in resp.url:
                return
            for result in resp.json().get("results", []):
                self.add(result)
        except (PlaywrightError, ValueError, AttributeError):
            pass

    def add(self, result):
        attrs = {a.get("name"): a.get("value") for a in result.get("attributes", [])}
        rid = attrs.get("resourceid", "")
        status = attrs.get("ds6w:what/ds6w:status") or attrs.get("ds6w:status") or ""
        if not re.fullmatch(r"[0-9A-F]{32}", rid or "") or (status and not status.lower().startswith("route")):
            return
        if all(it["id"] != rid for it in self.items):
            self.items.append({"id": rid, "name": attrs.get("ds6w:identifier", ""),
                               "title": attrs.get("ds6w:label", ""), "status": status})


def maybe_complete(item):
    """Search status Route.Complete (or unknown): worth reading its tasks."""
    return not item["status"] or "complete" in item["status"].lower()


# GET from inside the dashboard page, with its 3DSpace session (the same way the page's own widgets call it).
GET_JSON_JS = r"""
async ([url, ctx]) => {
  const headers = {'Accept': 'application/json'};
  if (ctx) headers['SecurityContext'] = ctx;
  const r = await fetch(url, {method: 'GET', credentials: 'include', headers});
  const text = await r.text();
  return {status: r.status, text};
}
"""


def get_json(page, url, ctx=""):
    resp = page.evaluate(GET_JSON_JS, [url, ctx])
    if resp["status"] != 200:
        raise LocatorNotFound(f"3DSpace answered {resp['status']} for {url.split('?')[0]}")
    try:
        return json.loads(resp["text"])
    except ValueError:
        raise LocatorNotFound(f"3DSpace did not return JSON for {url.split('?')[0]} (session not valid?)")


def security_context(page, space):
    """ctx::<role>.<organization>.<collab space> from the user's preferred credentials;
    THREEDX_SECURITY_CONTEXT in .env overrides it."""
    if os.getenv("THREEDX_SECURITY_CONTEXT"):
        return os.getenv("THREEDX_SECURITY_CONTEXT").strip()
    data = get_json(page, f"{space}/resources/modeler/pno/person?current=true&select=preferredcredentials"
                          f"&select=collabspaces&xrequestedwith=xmlhttprequest")
    pc = data.get("preferredcredentials") or {}
    try:
        return f"ctx::{pc['role']['name']}.{pc['organization']['name']}.{pc['collabspace']['name']}"
    except (KeyError, TypeError):
        cs = data["collabspaces"][0]
        couple = cs["couples"][0]
        return f"ctx::{couple['role']['name']}.{couple['organization']['name']}.{cs['name']}"


def tenant_of(space):
    m = re.match(r"https://([a-z0-9]+)-", space, I)
    return m.group(1).upper() if m else ""


def fetch_route(page, space, ctx, route_id):
    """One GET: the route with its tasks."""
    url = (f"{space}/resources/v1/modeler/routes/{route_id}?{ROUTE_INCLUDE}&tenant={tenant_of(space)}"
           f"&SecurityContext={quote(ctx)}&$language=en&xrequestedwith=xmlhttprequest")
    return parse_route(get_json(page, url, ctx))


def clean(value):
    """Text Excel can hold (some route fields carry control characters)."""
    return ILLEGAL_CHARACTERS_RE.sub("", value) if isinstance(value, str) else value


def parse_route(body):
    """(route fields, raw task fields) from the routes/<id>?$include=tasks response."""
    data = (body.get("data") or [{}])[0]
    related = data.get("relateddata") or {}
    return _flatten(data), [_flatten(t) for t in related.get("tasks", [])]


def is_done(task):
    return str(task.get("current", "")).lower() == "complete" or bool(task.get("taskActualCompletionDate"))


def is_rejected(task):
    return str(task.get("approvalStatus", "")).lower() == "reject"


def all_approved(tasks):
    """Every task finished and none rejected (approvals Approved, notifications Sent, comments given)."""
    return bool(tasks) and all(is_done(t) and not is_rejected(t) for t in tasks)


def task_status(task):
    """The status line of the flowchart task box."""
    action = str(task.get("taskAction", "")).lower()
    approval = str(task.get("approvalStatus", "")).lower()
    if not is_done(task):
        return {"approve": "Awaiting Approval", "comment": "Awaiting Comment"}.get(action, "Awaiting Notification"
                                                                                 if "notify" in action else "Awaiting")
    if "notify" in action:
        return "Notification Sent"
    if action == "comment":
        return "Commented"
    return {"approve": "Approved", "reject": "Rejected", "abstain": "Abstained"}.get(approval, "Completed")


UUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", I)


def assignee(task):
    """Person name; for a user group task the assignee field holds the group's id, so its title is used."""
    for key in ("taskAssignee", "assigneeTitle", "routeOwnerUGName", "userGroupInfo", "taskAssigneeUsername"):
        value = clean(str(task.get(key) or "")).strip()
        if value and not UUID.match(value):
            return value
    return ""


def stamp(dt):
    """05-Nov-2026 15:17 (local time); text so every viewer shows it the same way."""
    return dt.strftime("%d-%b-%Y %H:%M") if dt else None


def duration(start, end):
    """(days with 2 decimals, "1d 4h 12m") between two datetimes."""
    if not (start and end):
        return None, None
    minutes = max(0, int((end - start).total_seconds() // 60))
    d, rest = divmod(minutes, 1440)
    h, m = divmod(rest, 60)
    return round(minutes / 1440, 2), (f"{d}d " if d else "") + f"{h}h {m}m"


def task_row(task):
    """One task: flowchart fields, its dates and the time it took (created -> completed)."""
    created, due = task_date(task.get("originated")), task_date(task.get("taskDueDate"))
    completed, modified = task_date(task.get("taskActualCompletionDate")), task_date(task.get("modified"))
    days, text = duration(created, completed)
    return {
        "Task Title": clean(task.get("title") or task.get("name")),
        "Status": task_status(task),
        "Assignee": assignee(task),
        "State": task.get("stateNLS") or ("Completed" if is_done(task) else "To Do"),
        "Priority": task.get("routeTaskPriority"),
        # "Due" not "Due Date": columns named "... Date" are written as date only, without the time.
        "Created": created, "Due": due, "Completed": completed, "Modified": modified,
        "Days Taken": days, "Time Taken": text,
        "Late": ("Yes" if completed > due else "No") if completed and due else None,
    }


def route_row(route, rows):
    """One route: its dates and the time from creation to the last task completed."""
    created = task_date(route.get("originated")) or min((r["Created"] for r in rows if r["Created"]), default=None)
    # Completed only when every task is: a route still in process has no completion date.
    done = rows and all(r["Completed"] for r in rows)
    completed = max(r["Completed"] for r in rows) if done else None
    days, text = duration(created, completed)
    longest = max((r for r in rows if r["Days Taken"] is not None), key=lambda r: r["Days Taken"], default=None)
    return {
        "Title": clean(route.get("title") or route.get("name")), "Name": route.get("name"),
        "Owner": route.get("owner"), "State": route.get("state"),
        "Created": created, "Started": min((r["Created"] for r in rows if r["Created"]), default=None),
        "Modified": task_date(route.get("modified")), "Completed": completed,
        "Days Taken": days, "Time Taken": text, "Tasks": len(rows),
        "Late Tasks": sum(r["Late"] == "Yes" for r in rows),
        "Longest Task": longest and longest["Task Title"], "Longest Task Days": longest and longest["Days Taken"],
    }


def as_text_dates(record):
    return {k: stamp(v) if isinstance(v, datetime) else v for k, v in record.items()}


def merge_per_route(ws, columns):
    """Merge each column down the rows of one route (rows with the same Row number)."""
    start = 2
    for r in range(2, ws.max_row + 2):
        if r <= ws.max_row and ws.cell(r, 1).value == ws.cell(start, 1).value:
            continue
        for col in columns:
            if r - 1 > start:
                ws.merge_cells(start_row=start, start_column=col, end_row=r - 1, end_column=col)
            ws.cell(start, col).alignment = Alignment(vertical="center", horizontal="right")
        start = r


def write_workbook(path, tasks):
    """One sheet, Tasks: the route's dates sit next to Route Title (no separate Routes sheet)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Tasks"
    _write_sheet(ws, TASK_COLUMNS, [as_text_dates(t) for t in tasks])
    format_table(ws, as_table=False)
    group_by_route(ws)
    merge_per_route(ws, [TASK_COLUMNS.index(c) + 1 for c in MERGED_ROUTE_COLUMNS])
    for col in HIDDEN_TASK_COLUMNS:
        ws.column_dimensions[get_column_letter(TASK_COLUMNS.index(col) + 1)].hidden = True
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def to_sheet(tasks, failed):
    """Hand the rows to the shared Google Sheet; returns the summary line.
    PLM_SHEET_DRIVE_DIR set: save Route_Tasks_latest.csv in the Google Drive folder; the sheet's Apps
    Script imports it (hourly, or menu PLM > Import latest). Otherwise: post to the web app."""
    if not tasks:
        return "nothing to write"
    try:
        if drive_dir():
            n = write_drive_csv(TASK_COLUMNS, tasks, drive_dir())
            return f"{n} rows saved to Google Drive ({drive_dir()}); the sheet imports them within the hour"
        if not sheet_settings():
            return "not set up (PLM_SHEET_DRIVE_DIR in .env); push later with --push-last"
        return f"{push_via_browser(TASK_COLUMNS, tasks, HIDDEN_TASK_COLUMNS, PROFILE_DIR)} rows written"
    except (SheetPushError, OSError) as e:
        failed.append(str(e))
        return f"FAILED {e} (push again with --push-last)"


def push_last():
    """Push the rows saved by the last run, without opening 3DEXPERIENCE."""
    Config.from_env()  # loads .env
    files = sorted(OUTPUT_DIR.glob("Route_Tasks_*.json"))
    if not files:
        print("No saved run found (output/Route_Tasks_*.json). Run the export first.")
        return 1
    tasks = json.loads(files[-1].read_text(encoding="utf-8"))
    failed = []
    print(f"{files[-1].name}: {len(tasks)} rows -> Google Sheet: {to_sheet(tasks, failed)}")
    return 1 if failed or not (drive_dir() or sheet_settings()) else 0


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--routes", type=int, default=2, help="how many fully approved routes (default 2)")
    ap.add_argument("--all", action="store_true",
                    help="every route in the results, in process ones too (full run; ignores --routes)")
    ap.add_argument("--no-sheet", action="store_true", help="do not push to the Google Sheet")
    ap.add_argument("--push-last", action="store_true",
                    help="only push the rows of the last run (output/Route_Tasks_*.json) to the Google Sheet")
    ap.add_argument("--headless", action="store_true", help="run without a visible browser")
    args = ap.parse_args()
    wanted = float("inf") if args.all else args.routes
    max_scrolls = MAX_SCROLLS_ALL if args.all else MAX_SCROLLS

    if args.push_last:
        return push_last()

    cfg = Config.from_env()
    failed, blocked, routes, tasks = [], [], [], []
    total, checked, skipped, tried = 0, 0, 0, set()
    for d in (PROFILE_DIR, OUTPUT_DIR):
        d.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        headless = args.headless or cfg.headless
        context = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), headless=headless,
            **({"viewport": {"width": 1600, "height": 950}} if headless else {"no_viewport": True, "args": ["--start-maximized"]}),
        )
        install_read_only_guard(context, blocked)
        hide_platform_notices(context)
        ids = SearchIds(context)
        page = context.pages[0] if context.pages else context.new_page()
        try:
            ensure_logged_in(page, cfg)
            total = search_results(page, cfg)
            if not total:
                raise LocatorNotFound("the results page shows 0 results")
            page.wait_for_timeout(1500)  # last search response
            if not ids.space:
                raise LocatorNotFound("3DSpace address not seen during the search")
            ctx = security_context(page, ids.space)
            print(f"Reading every route in the {total} results (security context {ctx}) ..." if args.all else
                  f"Looking for {args.routes} routes with every task approved (security context {ctx}) ...")

            tried, scrolls, stalled = set(), 0, 0
            while len(routes) < wanted:
                pending = [it for it in ids.items if it["id"] not in tried]
                if not pending:
                    if scrolls >= max_scrolls or len(ids.items) >= total or stalled >= STALL_SCROLLS:
                        break
                    before = len(ids.items)
                    pos = page.evaluate(SCROLL_JS, None)  # next page of results, same search
                    page.wait_for_timeout(1500)
                    scrolls += 1
                    stalled = 0 if len(ids.items) > before else stalled + 1
                    if pos and pos["top"] >= pos["max"] and not [it for it in ids.items if it["id"] not in tried]:
                        page.wait_for_timeout(2000)
                    continue
                item = pending[0]
                tried.add(item["id"])
                if not args.all and not maybe_complete(item):
                    skipped += 1
                    continue
                label = item["title"] or item["name"]
                checked += 1
                try:
                    route, raw = fetch_route(page, ids.space, ctx, item["id"])
                except (PlaywrightError, LocatorNotFound, ValueError) as e:
                    msg = str(e).splitlines()[0]
                    failed.append(f"{label}: {msg}")
                    print(f"  {label}: FAILED {msg}")
                    continue
                if not args.all and not all_approved(raw):
                    open_ = sum(not is_done(t) for t in raw)
                    print(f"  {label}: skipped ({open_} open, {sum(map(is_rejected, raw))} rejected)")
                    skipped += 1
                    continue
                n = len(routes) + 1
                rows = sorted((task_row(t) for t in raw), key=lambda r: (r["Created"] or datetime.min, r["Task Title"]))
                summary = route_row(route, rows)
                routes.append({"Row": n, **summary})
                tasks += [{"Row": n, "Route Title": summary["Title"], "Route Created": summary["Created"], "Route Completed": summary["Completed"], **r}
                          for r in rows]
                if args.all:
                    print(f"  [{n}/{total}] {label}: {len(rows)} tasks, "
                          + (f"took {summary['Time Taken']}" if summary["Completed"] else "in process"))
                else:
                    print(f"  [{n}/{args.routes}] {label}: {len(rows)} tasks, all approved, took {summary['Time Taken']}")
            if args.all and len(tried) < total:
                print(f"Only {len(tried)} of {total} results loaded in the grid; the rest were not read.")
            elif not args.all and len(routes) < args.routes:
                print(f"Only {len(routes)} approved routes found in the {len(tried)} results loaded.")
        except (LoginError, LocatorNotFound, PlaywrightError) as e:
            print(f"\nSTOPPED: {str(e).splitlines()[0]}")
            save_snapshot(page, "route_tasks_stopped")
            if not routes:
                return 1
        finally:
            context.close()

    out = write_workbook(OUTPUT_DIR / f"Route_Tasks_{STAMP}.xlsx", tasks)
    rows_file = OUTPUT_DIR / f"Route_Tasks_{STAMP}.json"
    rows_file.write_text(json.dumps([as_text_dates(t) for t in tasks], ensure_ascii=False), encoding="utf-8")
    sheet = "skipped (--no-sheet)" if args.no_sheet else to_sheet(tasks, failed)
    print("\n================ Summary ================")
    print(f"Workbook        : {out}")
    print(f"Google Sheet    : {sheet}")
    print(f"Results shown   : {total}")
    print(f"Routes exported : {len(routes)}" + ("" if args.all else " (all tasks approved)"))
    print(f"Routes skipped  : {skipped} (in process or not every task approved); routes read: {checked}")
    print(f"Tasks           : {len(tasks)}")
    print(f"Healed locators : {len(HEALED)}")
    print(f"Blocked writes  : {len(blocked)}")
    print(f"Failed routes   : {len(failed)}")
    for f in failed:
        print(f"  - {f}")
    if args.all:
        return 1 if failed or len(tried) < total else 0
    return 1 if failed or len(routes) < args.routes else 0


if __name__ == "__main__":
    sys.exit(main())
