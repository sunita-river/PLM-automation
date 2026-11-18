"""Export every route in 3DEXPERIENCE Route Management to Excel (Routes / Tasks / Content).

Flow: platform home -> 3DEXPERIENCE ID login (only if the saved session expired) ->
top bar search "Route" -> search icon -> Route Management -> filter All -> load every row ->
open each route for its tasks and content -> output/Routes_<date>.xlsx.

Read only: route/task changes are blocked at the network level (threedx.install_read_only_guard).

Usage:
  .venv\\Scripts\\python.exe scripts\\route_export.py              full export
  .venv\\Scripts\\python.exe scripts\\route_export.py --discover   login + search + open app, save page structure to debug/
  .venv\\Scripts\\python.exe scripts\\route_export.py --use-export also download the app's own Export file
  .venv\\Scripts\\python.exe scripts\\route_export.py --max-routes 5
"""
import argparse
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError, sync_playwright

from excel_writer import write_workbook
from threedx import (HEALED, PROJECT_DIR, Config, LocatorNotFound, LoginError, ensure_logged_in,
                     find, find_optional, hide_platform_notices, install_read_only_guard, open_app_via_search)

I = re.IGNORECASE
PROFILE_DIR = PROJECT_DIR / ".auth" / "profile"
OUTPUT_DIR = PROJECT_DIR / "output"
DOWNLOAD_DIR = PROJECT_DIR / "downloads"
STAMP = datetime.now().strftime("%Y-%m-%d_%H%M")
DEBUG_DIR = PROJECT_DIR / "debug" / STAMP

# Column header / JSON key aliases for each output column (compared lower-case, letters and digits only).
ROUTE_ALIASES = {
    "Route Name": ["routename", "name"],
    "Title": ["title"],
    "Description": ["description"],
    "State/Status": ["state", "status", "routestatus", "current", "maturitystate"],
    "Owner": ["owner", "ownerfullname", "ownername"],
    "Route Template": ["routetemplate", "template", "templatename"],
    "Created Date": ["created", "createddate", "creationdate", "originated"],
    "Modified Date": ["modified", "modifieddate", "modificationdate", "lastmodified"],
    "Due Date": ["duedate", "routeduedate", "scheduledcompletiondate", "due"],
    "Completion Date": ["completiondate", "actualcompletiondate", "completeddate", "completed"],
}
TASK_ALIASES = {
    "Task Name": ["taskname", "name", "title"],
    "Assignee": ["assignee", "assigneefullname", "assignedto", "assignees", "owner"],
    "Action": ["action", "routeaction", "tasktype"],
    "Task State": ["taskstate", "state", "status", "current"],
    "Due Date": ["duedate", "scheduledcompletiondate", "due"],
    "Completed Date": ["completeddate", "actualcompletiondate", "completiondate", "completed"],
    "Approval Status": ["approvalstatus", "approvalstatusvalue", "decision", "result"],
    "Comments": ["comments", "comment", "approvalcomments"],
}
CONTENT_ALIASES = {
    "Name": ["name", "title"],
    "Type": ["type", "objecttype"],
    "Revision": ["revision", "rev", "version"],
}

# ---------------------------------------------------------------- widget locators

FILTER_BUTTON = [
    ("filter dropdown", lambda f: f.locator("[class*=filter i] .dropdown, [class*=filter i] select, .wux-controls-combobox").filter(has_text=re.compile("route|active|all", I))),
    ("button My Routes/Active", lambda f: f.get_by_role("button", name=re.compile("my routes|active|owned|assigned", I))),
    ("text My Routes/Active", lambda f: f.get_by_text(re.compile(r"^\s*(my routes|active routes?|active)\s*$", I))),
]
FILTER_ALL = [
    ("option All Routes", lambda f: f.get_by_role("option", name=re.compile(r"^\s*all( routes)?\s*$", I))),
    ("menuitem All", lambda f: f.get_by_role("menuitem", name=re.compile(r"^\s*all( routes)?\s*$", I))),
    ("text All Routes", lambda f: f.get_by_text(re.compile(r"^\s*all( routes)?\s*$", I))),
]
EXPORT_BUTTON = [
    ("button Export", lambda f: f.get_by_role("button", name=re.compile("export|download|csv|excel", I))),
    ("[title*=Export]", lambda f: f.locator("[title*='export' i], [title*='download' i], [title*='csv' i], [title*='excel' i]")),
    ("text Export", lambda f: f.get_by_text(re.compile(r"^\s*(export|download).*$", I))),
]
TASKS_TAB = [
    ("tab Tasks", lambda f: f.get_by_role("tab", name=re.compile("task", I))),
    ("text Tasks", lambda f: f.get_by_text(re.compile(r"^\s*tasks?\s*(\(\d+\))?\s*$", I))),
]
CONTENT_TAB = [
    ("tab Content", lambda f: f.get_by_role("tab", name=re.compile("content|attachment", I))),
    ("text Content", lambda f: f.get_by_text(re.compile(r"^\s*(content|attachments?)\s*(\(\d+\))?\s*$", I))),
]
CLOSE_DETAILS = [
    ("button Close", lambda f: f.get_by_role("button", name=re.compile("^(close|back)$", I))),
    ("[title=Close]", lambda f: f.locator("[title='Close' i], [title='Back' i], .fonticon-close, .fonticon-left-open")),
]

# Reads the biggest visible table / ARIA grid / 3DS data grid in the frame as {headers, rows}.
READ_GRID_JS = r"""
() => {
  const vis = e => !!(e.offsetWidth || e.offsetHeight || e.getClientRects().length);
  const txt = e => (e.innerText || e.textContent || '').trim().replace(/\s+/g, ' ');
  const out = [];
  for (const t of document.querySelectorAll('table')) {
    if (!vis(t)) continue;
    const headers = [...t.querySelectorAll('thead th, tr:first-child th')].map(txt);
    const rows = [...t.querySelectorAll('tbody tr')].filter(vis).map(r => [...r.querySelectorAll('td')].map(txt));
    out.push({headers, rows});
  }
  for (const g of document.querySelectorAll('[role=grid], [role=treegrid], [class*=datagrid i]')) {
    if (!vis(g)) continue;
    const headers = [...g.querySelectorAll('[role=columnheader], [class*=header-cell i]')].filter(vis).map(txt);
    const rows = [...g.querySelectorAll('[role=row]')].filter(r => vis(r) && !r.querySelector('[role=columnheader]'))
      .map(r => [...r.querySelectorAll('[role=gridcell], [role=cell], [class*=cell i]')].map(txt));
    out.push({headers, rows});
  }
  out.sort((a, b) => b.rows.length - a.rows.length);
  return out[0] || {headers: [], rows: []};
}
"""


def norm(key):
    return re.sub(r"[^a-z0-9]", "", str(key).lower())


def map_record(record, aliases):
    """Pick each output column from a flat dict using the alias list (first alias that has a value)."""
    keyed = {norm(k): v for k, v in record.items()}
    row = {}
    for column, names in aliases.items():
        for n in names:
            v = keyed.get(n)
            if v not in (None, "", [], {}):
                row[column] = ", ".join(map(str, v)) if isinstance(v, list) else v
                break
    return row


# ---------------------------------------------------------------- network capture (JSON the widget loads)

class Capture:
    """Keeps every JSON response the app loads; the route list and route details usually arrive this way."""

    def __init__(self, page):
        self.items = []
        page.context.on("response", self._on_response)

    def _on_response(self, resp):
        try:
            if "json" not in (resp.headers.get("content-type") or ""):
                return
            if not re.search(r"route|task|dsrt|enxrout|content|attach", resp.url, I):
                return
            self.items.append({"url": resp.url, "method": resp.request.method, "body": resp.json()})
        except (PlaywrightError, ValueError):
            pass

    def mark(self):
        return len(self.items)

    def since(self, mark):
        return self.items[mark:]

    def save(self, folder):
        folder.mkdir(parents=True, exist_ok=True)
        for n, item in enumerate(self.items, 1):
            (folder / f"{n:03d}.json").write_text(json.dumps(item, indent=2, default=str), encoding="utf-8")


def _flatten(obj):
    """Merge 3DS-style nested fields (dataelements, attributes) into one flat dict."""
    flat = {}
    for k, v in obj.items():
        if isinstance(v, dict) and norm(k) in ("dataelements", "attributes", "attribute", "properties", "info"):
            flat.update({kk: vv for kk, vv in v.items() if not isinstance(vv, (dict, list))})
        elif not isinstance(v, (dict, list)) or (isinstance(v, list) and all(isinstance(x, str) for x in v)):
            flat[k] = v
    return flat


def _record_lists(body):
    """Yield every list of dicts inside a JSON body."""
    if isinstance(body, list):
        if body and all(isinstance(x, dict) for x in body):
            yield body
        for x in body:
            yield from _record_lists(x)
    elif isinstance(body, dict):
        for v in body.values():
            yield from _record_lists(v)


def records_from_json(items, aliases, url_pattern, min_fields=2):
    """Best list of records in the captured responses whose URL matches url_pattern."""
    best = []
    for item in items:
        if not re.search(url_pattern, item["url"], I):
            continue
        for lst in _record_lists(item["body"]):
            rows = [map_record(_flatten(r), aliases) for r in lst]
            rows = [r for r in rows if len(r) >= min_fields]
            if len(rows) > len(best):
                best = rows
    return best


def records_from_grid(frame, aliases):
    grid = frame.evaluate(READ_GRID_JS)
    headers = grid["headers"]
    if not headers:
        return []
    return [r for r in (map_record(dict(zip(headers, cells)), aliases) for cells in grid["rows"]) if r]


# ---------------------------------------------------------------- steps

def check_export(frame, page, use_export, report):
    btn = find_optional(frame, "Export button", EXPORT_BUTTON, timeout=5000)
    if not btn:
        report["export"] = "No Export / Download option found in the widget."
        print(report["export"])
        return
    label = (btn.get_attribute("title") or btn.inner_text() or "Export").strip()
    report["export"] = f"Export option found: '{label}'."
    print(report["export"])
    if not use_export:
        print("  Not downloaded. Run again with --use-export to save the app's own file to downloads/.")
        return
    try:
        with page.expect_download(timeout=60000) as dl:
            btn.click()
            fmt = find_optional(frame, "Export format", [
                ("menuitem Excel/CSV", lambda f: f.get_by_role("menuitem", name=re.compile("excel|xlsx|csv", I))),
                ("text Excel/CSV", lambda f: f.get_by_text(re.compile(r"excel|xlsx|csv", I))),
            ], timeout=3000)
            if fmt:
                fmt.click()
        target = DOWNLOAD_DIR / f"{STAMP}_{dl.value.suggested_filename}"
        dl.value.save_as(target)
        report["export"] += f" Downloaded to {target}."
        print(f"  Saved {target}")
    except PlaywrightError as e:
        report["export"] += f" Download failed: {e.message.splitlines()[0]}"


def show_all_routes(frame):
    btn = find_optional(frame, "route filter", FILTER_BUTTON, timeout=10000)
    if not btn:
        print("Route filter not found; keeping the default list (check the debug screenshot).")
        return False
    btn.click()
    find(frame, "filter option All", FILTER_ALL, timeout=10000).click()
    frame.page.wait_for_timeout(3000)
    print("Filter set to All.")
    return True


def load_every_row(frame):
    """Scroll the list until the row count stops growing (covers lazy loading and virtual scrolling)."""
    last, same = -1, 0
    while same < 3:
        count = frame.evaluate("() => document.querySelectorAll('[role=row], tbody tr').length")
        same = same + 1 if count == last else 0
        last = count
        frame.evaluate("""() => {
            for (const el of document.querySelectorAll('*'))
                if (el.scrollHeight > el.clientHeight + 20 && getComputedStyle(el).overflowY.match(/auto|scroll/))
                    el.scrollTop = el.scrollHeight;
            window.scrollTo(0, document.body.scrollHeight);
        }""")
        more = find_optional(frame, "Load more", [
            ("button Load more", lambda f: f.get_by_role("button", name=re.compile("load more|show more|next", I))),
        ], timeout=500)
        if more:
            more.click()
        frame.page.wait_for_timeout(1500)
    print(f"List fully loaded ({last} rows in the DOM).")


def route_details(frame, capture, route_name):
    """Open one route, read its Tasks and Content, close it again."""
    mark = capture.mark()
    frame.get_by_text(route_name, exact=True).first.click()
    frame.page.wait_for_timeout(2500)

    tab = find_optional(frame, "Tasks tab", TASKS_TAB, timeout=10000)
    if tab:
        tab.click()
        frame.page.wait_for_timeout(2000)
    tasks = records_from_json(capture.since(mark), TASK_ALIASES, r"task") or records_from_grid(frame, TASK_ALIASES)

    content = []
    tab = find_optional(frame, "Content tab", CONTENT_TAB, timeout=5000)
    if tab:
        tab.click()
        frame.page.wait_for_timeout(2000)
        content = (records_from_json(capture.since(mark), CONTENT_ALIASES, r"content|attach", min_fields=1)
                   or records_from_grid(frame, CONTENT_ALIASES))

    close = find_optional(frame, "close route details", CLOSE_DETAILS, timeout=2000)
    if close:
        close.click()
    else:
        frame.page.keyboard.press("Escape")
    frame.page.wait_for_timeout(1000)
    return tasks, content


def save_discovery(page, frame, capture):
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(DEBUG_DIR / "page.png"), full_page=True)
    for n, f in enumerate(page.frames):
        try:
            (DEBUG_DIR / f"frame_{n:02d}.html").write_text(f"<!-- {f.url} -->\n" + f.content(), encoding="utf-8")
        except PlaywrightError:
            pass
    if frame:
        controls = frame.evaluate("""() => [...document.querySelectorAll(
            'button, [role=button], [role=tab], [role=menuitem], select, input, [title]')]
            .filter(e => e.offsetWidth || e.offsetHeight)
            .map(e => [e.tagName, e.getAttribute('role') || '', e.getAttribute('title') || '',
                       (e.innerText || e.value || '').trim().slice(0, 60), e.className].join(' | '))""")
        (DEBUG_DIR / "widget_controls.txt").write_text("\n".join(controls), encoding="utf-8")
        (DEBUG_DIR / "widget_grid.json").write_text(json.dumps(frame.evaluate(READ_GRID_JS), indent=2), encoding="utf-8")
    capture.save(DEBUG_DIR / "network")
    print(f"Discovery saved to {DEBUG_DIR}")


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--discover", action="store_true", help="save page structure to debug/ and stop")
    ap.add_argument("--use-export", action="store_true", help="download the app's own Export file too")
    ap.add_argument("--max-routes", type=int, default=0, help="only open the first N routes (testing)")
    ap.add_argument("--headless", action="store_true", help="run without a visible browser")
    args = ap.parse_args()

    cfg = Config.from_env()
    report = {"export": "", "failed": [], "blocked": []}
    for d in (PROFILE_DIR, OUTPUT_DIR, DOWNLOAD_DIR):
        d.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            str(PROFILE_DIR), headless=args.headless or cfg.headless, accept_downloads=True,
            viewport={"width": 1600, "height": 950},
        )
        install_read_only_guard(context, report["blocked"])
        hide_platform_notices(context)
        page = context.pages[0] if context.pages else context.new_page()
        capture = Capture(page)
        frame = None
        try:
            ensure_logged_in(page, cfg)
            frame = open_app_via_search(page, cfg)
            print("Route Management loaded.")
            if args.discover:
                show_all_routes(frame)
                load_every_row(frame)
                save_discovery(page, frame, capture)
                return 0

            check_export(frame, page, args.use_export, report)
            mark = capture.mark()
            show_all_routes(frame)
            load_every_row(frame)

            routes = (records_from_json(capture.since(mark), ROUTE_ALIASES, r"route", min_fields=3)
                      or records_from_json(capture.items, ROUTE_ALIASES, r"route", min_fields=3)
                      or records_from_grid(frame, ROUTE_ALIASES))
            routes = list({r.get("Route Name"): r for r in routes if r.get("Route Name")}.values())
            if not routes:
                raise LocatorNotFound("No routes found in the list (run with --discover and check debug/).")
            print(f"{len(routes)} routes in the list.")

            tasks, content = [], []
            todo = routes[: args.max_routes] if args.max_routes else routes
            for n, r in enumerate(todo, 1):
                name = r["Route Name"]
                try:
                    t, c = route_details(frame, capture, name)
                    tasks += [{"Route Name": name, **x} for x in t]
                    content += [{"Route Name": name, **x} for x in c]
                    print(f"  [{n}/{len(todo)}] {name}: {len(t)} tasks, {len(c)} content")
                except (PlaywrightError, LocatorNotFound) as e:
                    msg = str(e).splitlines()[0]
                    report["failed"].append(f"{name}: {msg}")
                    print(f"  [{n}/{len(todo)}] {name}: FAILED {msg}")
                    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
                    page.screenshot(path=str(DEBUG_DIR / f"failed_{n:03d}.png"))
                    page.keyboard.press("Escape")

            out = write_workbook(OUTPUT_DIR / f"Routes_{STAMP}.xlsx", routes, tasks, content)
        except (LoginError, LocatorNotFound, PlaywrightError) as e:
            print(f"\nSTOPPED: {str(e).splitlines()[0]}")
            save_discovery(page, frame, capture)
            return 1
        finally:
            context.close()

    print("\n================ Summary ================")
    print(f"Workbook       : {out}")
    print(f"Routes         : {len(routes)}  ({len(todo)} opened)")
    print(f"Tasks          : {len(tasks)}")
    print(f"Content items  : {len(content)}")
    print(f"Export option  : {report['export']}")
    print(f"Healed locators: {len(HEALED)}" + (f" (see {OUTPUT_DIR / 'healed-locators.json'})" if HEALED else ""))
    print(f"Blocked writes : {len(report['blocked'])}")
    print(f"Failed rows    : {len(report['failed'])}")
    for f in report["failed"]:
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
