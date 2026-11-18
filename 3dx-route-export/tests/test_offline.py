"""Offline checks (no 3DEXPERIENCE connection): JSON mapping, grid reading, read-only guard, Excel format.

Run: .venv\\Scripts\\python.exe tests\\test_offline.py
"""
import sys
import tempfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

import route_export as rx
from excel_writer import write_workbook
from threedx import install_read_only_guard

RESULTS = []


def check(tc, name, ok, detail=""):
    RESULTS.append((tc, ok))
    print(f"{'PASS' if ok else 'FAIL'}  {tc}  {name}" + (f"  ({detail})" if detail and not ok else ""))


# 3DS-style REST payload: records under "data", fields under "dataelements".
SAMPLE_ROUTES = {"url": "https://x/resources/v1/modeler/dsrt/routes", "method": "GET", "body": {"data": [
    {"id": "1", "dataelements": {"name": "R-0001", "title": "ECO review", "description": "Bracket change",
                                 "state": "In Process", "owner": "Asha K", "routeTemplate": "ECO Std",
                                 "created": "2026-10-01T09:15:00Z", "modified": "2026-10-03T11:00:00Z",
                                 "dueDate": "2026-10-10", "completionDate": ""}},
    {"id": "2", "dataelements": {"name": "R-0002", "title": "Drawing release", "state": "Complete",
                                 "owner": "Ravi M", "created": "2026-09-20", "completionDate": "2026-09-25"}},
]}}
SAMPLE_TASKS = {"url": "https://x/resources/v1/modeler/dsrt/routes/1/tasks", "method": "GET", "body": {"data": [
    {"dataelements": {"title": "Approve drawing", "assignee": "Asha K", "routeAction": "Approve",
                      "state": "Assigned", "dueDate": "2026-10-08", "comments": ""}},
    {"dataelements": {"title": "Notify team", "assignee": "Ravi M", "routeAction": "Notify", "state": "Complete",
                      "completedDate": "2026-10-02", "approvalStatus": "None"}},
]}}


def test_json_mapping():
    routes = rx.records_from_json([SAMPLE_ROUTES], rx.ROUTE_ALIASES, r"route", min_fields=3)
    check("RTE-T01", "routes read from 3DS JSON (dataelements)", len(routes) == 2 and routes[0]["Route Name"] == "R-0001"
          and routes[0]["Route Template"] == "ECO Std" and routes[0]["State/Status"] == "In Process", routes)
    tasks = rx.records_from_json([SAMPLE_TASKS], rx.TASK_ALIASES, r"task")
    check("RTE-T02", "tasks read from 3DS JSON", len(tasks) == 2 and tasks[0]["Action"] == "Approve"
          and tasks[1]["Completed Date"] == "2026-10-02", tasks)
    none = rx.records_from_json([SAMPLE_TASKS], rx.ROUTE_ALIASES, r"nomatch")
    check("RTE-T03", "responses with other URLs ignored", none == [])


def test_grid_and_guard():
    html = """<table><thead><tr><th>Name</th><th>Title</th><th>State</th><th>Owner</th></tr></thead>
              <tbody><tr><td>R-0003</td><td>Spec</td><td>Active</td><td>Neha</td></tr>
                     <tr><td>R-0004</td><td>BOM</td><td>Complete</td><td>Om</td></tr></tbody></table>"""
    with sync_playwright() as p:
        b = p.chromium.launch(headless=True)
        page = b.new_page()
        page.set_content(html)
        rows = rx.records_from_grid(page.main_frame, rx.ROUTE_ALIASES)
        check("RTE-T04", "DOM table read and mapped by header", [r["Route Name"] for r in rows] == ["R-0003", "R-0004"]
              and rows[1]["Owner"] == "Om", rows)
        b.close()

    # Guard decision tested directly on fake requests (the handler is what the browser calls).
    class Req:
        def __init__(self, method, url):
            self.method, self.url = method, url

    class Route:
        def __init__(self, method, url):
            self.request, self.result = Req(method, url), None

        def abort(self):
            self.result = "aborted"

        def continue_(self):
            self.result = "sent"

    class Ctx:
        def route(self, _pattern, handler):
            self.handler = handler

    ctx, blocked = Ctx(), []
    install_read_only_guard(ctx, blocked)
    cases = [("GET", "https://t.3dexperience.3ds.com/resources/v1/modeler/dsrt/routes", "sent"),
             ("POST", "https://t.3dexperience.3ds.com/federated/search", "sent"),
             ("POST", "https://t.3dexperience.3ds.com/resources/v1/modeler/dsrt/tasks/9/approve", "aborted"),
             ("POST", "https://t.3dexperience.3ds.com/resources/route/reassign", "aborted"),
             ("PUT", "https://t.3dexperience.3ds.com/resources/v1/modeler/dsrt/routes/1", "aborted"),
             ("DELETE", "https://t.3dexperience.3ds.com/resources/v1/modeler/dsrt/routes/1", "aborted"),
             ("POST", "https://r1-eu1.iam.3dexperience.3ds.com/login", "sent")]
    got = []
    for method, url, want in cases:
        r = Route(method, url)
        ctx.handler(r)
        got.append(r.result == want)
    check("RTE-T05", "read-only guard blocks approve/reassign/PUT/DELETE, allows GET, search and login", all(got), got)


def test_excel():
    routes = [rx.map_record(rx._flatten(r), rx.ROUTE_ALIASES) for r in SAMPLE_ROUTES["body"]["data"]]
    tasks = [{"Route Name": "R-0001", **rx.map_record(rx._flatten(t), rx.TASK_ALIASES)} for t in SAMPLE_TASKS["body"]["data"]]
    content = [{"Route Name": "R-0001", "Name": "DRW-100", "Type": "Drawing", "Revision": "B"}]
    path = Path(tempfile.mkdtemp()) / "Routes_test.xlsx"
    write_workbook(path, routes, tasks, content)
    wb = load_workbook(path)
    ok_sheets = wb.sheetnames == ["Routes", "Tasks", "Content"]
    ws = wb["Routes"]
    fmt = all(wb[s]["A1"].font.b and wb[s].freeze_panes == "A2" and wb[s].auto_filter.ref for s in wb.sheetnames)
    created = ws["G2"]
    check("RTE-T06", "three sheets, bold headers, frozen row, filters", ok_sheets and fmt)
    check("RTE-T07", "dates stored as dates in DD-MMM-YYYY", isinstance(created.value, datetime)
          and created.number_format == "DD-MMM-YYYY" and created.value.date().isoformat() == "2026-10-01",
          f"{created.value!r} {created.number_format}")
    check("RTE-T08", "Route Name key on Tasks and Content", wb["Tasks"]["A2"].value == "R-0001"
          and wb["Content"]["A2"].value == "R-0001" and wb["Tasks"].max_row == 3)
    check("RTE-T09", "columns auto-sized", ws.column_dimensions["C"].width >= len("Bracket change"))


def test_route_tasks():
    """scripts/route_tasks.py: route id from the search response, route + tasks from one GET, approved-only
    filter, flowchart fields, dates and time taken."""
    import route_tasks as rt

    ids = rt.SearchIds.__new__(rt.SearchIds)
    ids.items = []
    attr = lambda n, v: {"format": "internal", "name": n, "type": "string", "value": v}
    ids.add({"attributes": [attr("resourceid", "iam:some.user"), attr("ds6w:label", "Some User")]})
    ids.add({"attributes": [attr("resourceid", "41661FF655AC2B006AE3F50F0002B792"),
                            attr("ds6w:identifier", "R-RMIN-0002810"), attr("ds6w:what/ds6w:status", "Route.In Process")]})
    ids.add({"attributes": [attr("resourceid", "41661FF655AC2B006AE3F50F0002B793"),
                            attr("ds6w:identifier", "R-RMIN-0002812"), attr("ds6w:what/ds6w:status", "Route.Complete")]})
    check("RTT-T01", "route ids and status from search response; people ignored; in-process routes not read",
          [i["name"] for i in ids.items] == ["R-RMIN-0002810", "R-RMIN-0002812"]
          and [rt.maybe_complete(i) for i in ids.items] == [False, True], ids.items)

    task = lambda **f: {"id": "T", "dataelements": f}
    common = dict(originated="2026-11-04T07:55:00.000Z", taskDueDate="2026-11-06T07:55:00.000Z",
                  modified="2026-11-06T08:00:00.000Z", routeTaskPriority="High")
    body = {"data": [{"id": "R1", "dataelements": {"name": "R-RMIN-0002823", "title": "RX02 HL and TL",
                                                   "originated": "11/3/2026 11:57:00 AM", "state": "Complete"},
                      "relateddata": {"tasks": [
        task(title="R2 - Integration", taskAction="Approve", approvalStatus="Approve", current="Complete",
             stateNLS="Completed", taskAssignee="subba reddy", taskActualCompletionDate="11/5/2026 9:57:00 AM", **common),
        task(title="Division Head - Electrical", taskAction="Notify Only", approvalStatus="None", current="Complete",
             stateNLS="Completed", taskActualCompletionDate="11/4/2026 7:55:00 AM", **common),
        task(title="R2 - Service", taskAction="Comment", approvalStatus="None", current="Complete",
             stateNLS="Completed", taskActualCompletionDate="11/11/2026 7:55:00 AM", **common),
    ]}}]}
    route, tasks = rt.parse_route(body)
    check("RTT-T02", "route and its tasks read from one response",
          route.get("name") == "R-RMIN-0002823" and len(tasks) == 3, (route, tasks))
    rows = [rt.task_row(t) for t in tasks]
    check("RTT-T03", "status line as in the flowchart",
          [r["Status"] for r in rows] == ["Approved", "Notification Sent", "Commented"], [r["Status"] for r in rows])
    open_task = {**tasks[0], "current": "Assigned", "taskActualCompletionDate": "", "approvalStatus": "None"}
    rejected = {**tasks[0], "approvalStatus": "Reject"}
    check("RTT-T04", "only routes with every task approved / done are kept",
          rt.all_approved(tasks) and not rt.all_approved(tasks + [open_task])
          and not rt.all_approved(tasks + [rejected]) and not rt.all_approved([]))
    check("RTT-T05", "task time taken = created -> completed; late when completed after due",
          rows[0]["Time Taken"] == "1d 2h 2m" and rows[0]["Days Taken"] == 1.08 and rows[0]["Late"] == "No"
          and rows[1]["Time Taken"] == "0h 0m" and rows[2]["Days Taken"] == 7.0 and rows[2]["Late"] == "Yes", rows)
    summary = rt.route_row(route, rows)
    check("RTT-T06", "route time taken = created -> last task completed; started = first task created",
          summary["Time Taken"] == "7d 19h 58m" and summary["Started"] == rows[0]["Created"]
          and summary["Late Tasks"] == 1 and summary["Longest Task"] == "R2 - Service", summary)
    check("RTT-T07", "dates written with time (dd-Mmm-yyyy hh:mm)",
          rt.stamp(datetime(2026, 11, 5, 15, 17)) == "05-Nov-2026 15:17")
    check("RTT-T08", "tenant read from the 3DSpace address",
          rt.tenant_of("https://r1132101268131-indw2-space.3dexperience.3ds.com/enovia") == "R1132101268131")

    path = Path(tempfile.mkdtemp()) / "Route_Tasks_test.xlsx"
    rt.write_workbook(path, [{"Row": 1, "Route Title": summary["Title"],
                              "Route Created": summary["Created"], "Route Completed": summary["Completed"], **r}
                             for r in rows])
    wb = load_workbook(path)
    ws = wb["Tasks"]
    hidden = {ws.cell(1, i).value for i in range(1, ws.max_column + 1)
              if ws.column_dimensions[get_column_letter(i)].hidden}
    check("RTT-T09", "Tasks sheet: Status, Priority, Due, Time Taken, Late hidden; data still written",
          hidden == {"Status", "Priority", "Due", "Time Taken", "Late"} and ws.cell(2, 6).value, hidden)
    heads = [ws.cell(1, i).value for i in range(1, ws.max_column + 1)]
    merged = {str(m) for m in ws.merged_cells.ranges}
    check("RTT-T10", "no Routes sheet, no Route Name; Route Created / Route Completed after Route Title, merged per route",
          wb.sheetnames == ["Tasks"] and heads[1:4] == ["Route Title", "Route Created", "Route Completed"]
          and "Route Name" not in heads
          and ws["C2"].value == rt.stamp(summary["Created"]) and ws["D2"].value == rt.stamp(summary["Completed"])
          and {f"C2:C{ws.max_row}", f"D2:D{ws.max_row}"} <= merged, (wb.sheetnames, heads, ws["D2"].value, merged))

    open_rows = [dict(r) for r in rows]
    open_rows[-1].update(Completed=None, **{"Days Taken": None, "Time Taken": None, "Late": None})
    open_summary = rt.route_row({"title": "Open route", "originated": "2026-11-01T10:00:00.000Z"}, open_rows)
    check("RTT-T11", "route in process (--all): Route Completed and route time taken left empty",
          open_summary["Completed"] is None and open_summary["Time Taken"] is None
          and summary["Completed"] is not None, open_summary)

    import sheet_push as sp
    records = [{"Row": i, "Route Title": "R", "Created": datetime(2026, 11, 5, 15, 17), "Late": None} for i in range(5)]
    bodies = sp.payloads("tok", "804083112", ["Row", "Route Title", "Created", "Late"], records, ["Late"], chunk=2,
                         updated="12-Nov-2026 10:00")
    sent = []
    written = sp.push(["Row", "Route Title", "Created", "Late"], records, ["Late"],
                      settings=("https://example.invalid/exec", "tok", "804083112"), post=lambda u, b: sent.append(b))
    check("RTT-T12", "sheet push: rows in parts, first clears (start 0), last hides columns; dates as text, empty for None",
          [b["start"] for b in bodies] == [0, 2, 4] and [b["last"] for b in bodies] == [False, False, True]
          and bodies[0]["rows"][0] == [0, "R", "05-Nov-2026 15:17", ""] and bodies[0]["gid"] == "804083112"
          and written == 5 and len(sent) == 1 and sent[0]["last"] and sent[0]["hidden"] == ["Late"]
          and sp.payloads("t", "g", ["A"], [], [])[0]["last"], bodies)

    class FakeResp:
        def __init__(self, status, url, text):
            self.status, self.url, self._text = status, url, text

        def text(self):
            return self._text

    class FakeRequest:
        def __init__(self, resp):
            self.resp, self.calls = resp, 0

        def post(self, url, **kw):
            self.calls += 1
            return self.resp

    def outcome(resp):
        req = FakeRequest(resp)
        try:
            return sp.browser_post(req)("https://example.invalid/exec", {"start": 0}), req.calls
        except sp.SheetPushError as e:
            return type(e).__name__ + (": token" if "token" in str(e) else ""), req.calls

    ok = outcome(FakeResp(200, "https://script.googleusercontent.com/macros/echo", '{"ok": true, "written": 3}'))
    unauth = outcome(FakeResp(401, "https://script.google.com/a/macros/x/exec", "Unauthorized"))
    login = outcome(FakeResp(200, "https://accounts.google.com/v3/signin/identifier", "<html>"))
    token = outcome(FakeResp(200, "https://script.googleusercontent.com/macros/echo", '{"ok": false, "error": "token does not match"}'))
    check("RTT-T13", "browser push: JSON ok answer returned; 401 / Google sign-in page = sign in needed (no retry); token error not retried",
          ok == ({"ok": True, "written": 3}, 1) and unauth == ("SignInNeeded", 1) and login == ("SignInNeeded", 1)
          and token == ("SheetPushError: token", 1), (ok, unauth, login, token))

    probe = {name: sp.signed_in(FakeRequest(resp), "https://example.invalid/exec") for name, resp in [
        ("401", FakeResp(401, "https://script.google.com/a/macros/x/exec", "Unauthorized")),
        ("jumpcloud", FakeResp(200, "https://accounts.google.com/samlredirect", "<html>")),
        ("token check", FakeResp(200, "https://script.googleusercontent.com/macros/echo", '{"ok": false, "error": "token does not match"}'))]}
    check("RTT-T14", "Google sign-in counted only when the web app itself answers (401 / SSO page = not yet)",
          probe == {"401": False, "jumpcloud": False, "token check": True}, probe)


    import csv as _csv
    with tempfile.TemporaryDirectory() as d:
        n = sp.write_drive_csv(["Row", "Route Title", "Created", "Late"], records[:3], d)
        with open(Path(d) / sp.DRIVE_CSV, encoding="utf-8", newline="") as f:
            got = list(_csv.reader(f))
        leftovers = sorted(p.name for p in Path(d).iterdir())
    check("RTT-T15", "Drive hand-over: Route_Tasks_latest.csv with header + rows, dates as text, no temp file left",
          n == 3 and got[0] == ["Row", "Route Title", "Created", "Late"] and got[1] == ["0", "R", "05-Nov-2026 15:17", ""]
          and len(got) == 4 and leftovers == [sp.DRIVE_CSV], (n, got[:2], leftovers))


if __name__ == "__main__":
    test_json_mapping()
    test_grid_and_guard()
    test_excel()
    test_route_tasks()
    failed = [tc for tc, ok in RESULTS if not ok]
    print(f"\n{len(RESULTS) - len(failed)}/{len(RESULTS)} passed" + (f"; failed: {', '.join(failed)}" if failed else ""))
    sys.exit(1 if failed else 0)
