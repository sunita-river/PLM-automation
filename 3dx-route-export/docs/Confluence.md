# PLM – 3DEXPERIENCE Route Export

## Purpose
Read-only export of every route in 3DEXPERIENCE Route Management (app ENXROUT_AP) to one Excel workbook for review and reporting.

## Flow
1. Open the 3DEXPERIENCE ID login page (`THREEDX_BASE_URL`).
2. If the saved browser session is still valid, login is skipped. Otherwise: accept cookies, enter the username, **Continue**, enter the password, **Log in**.
3. Type `Route` in the dashboard top bar search and click the search icon. The script waits for the search box to take focus, reads back what was typed and retypes it (up to 3 times) if letters were dropped.
4. Open **Route Management** from the results and wait for its iframe to load.
5. Report whether the app has an Export / Download option. It is downloaded only with `--use-export`.
6. Set the list filter to **All**, then scroll until no new rows appear.
7. Open each route, then read the **Tasks** and **Content** tabs.
8. Write `output/Routes_<date>.xlsx` and print the summary.

## Output
| Sheet   | One row per         | Columns |
|---------|---------------------|---------|
| Routes  | route               | Route Name, Title, Description, State/Status, Owner, Route Template, Created / Modified / Due / Completion Date |
| Tasks   | task                | Route Name, Task Name, Assignee, Action, Task State, Due Date, Completed Date, Approval Status, Comments |
| Content | attached item       | Route Name, Name, Type, Revision |

Headers are bold, the top row is frozen, filters are on, columns are auto-sized and dates show as DD-MMM-YYYY.

The summary reports route, task and content counts, the Export option, healed locators, blocked write requests and failed rows.

## Controls
- **Platform notices:** banners such as the R2027x update notice cover the top bar search box; they are hidden on every page load (`threedx.hide_platform_notices`). Nothing is clicked or dismissed on the platform.
- **Read only:** every PUT, PATCH or DELETE request is aborted in the browser, as is any POST to an approve, reject, reassign, complete, promote, demote, modify or update path. The summary lists anything that was blocked.
- **Credentials:** kept in `.env` on the run machine only and never printed. The session is reused from `.auth/profile`.
- **Self-healing locators:** every action has backup locators. If a backup is used, the run prints `[HEAL]` and records it in `output/healed-locators.json`.
- **Failures:** a route that cannot be opened is listed with its reason and a screenshot in `debug/<date>/`. The run then continues with the next route.

## Run
```
.venv\Scripts\python.exe scripts\route_export.py              full export
.venv\Scripts\python.exe scripts\route_export.py --discover   save page structure to debug/
.venv\Scripts\python.exe scripts\route_export.py --use-export also download the app's own export
.venv\Scripts\python.exe scripts\route_export.py --max-routes 5
.venv\Scripts\python.exe tests\test_offline.py                offline checks
```

## Test status (05-Nov-2026)
- **Offline checks:** RTE-T01 to RTE-T09 passed.
- **Live:** RTE-03 passed (the login page is detected and the run stops when no credentials are set). RTE-01, RTE-02 and RTE-04 to RTE-17 have not been run yet; they need the password in `.env`.
- **12-Nov-2026:** RTE-05 and RTE-06 passed (search, Route Management opened; update notice hidden). RTE-09 failed: the filter is not found and the list stays on My Routes (0), so the run stops. For route data use the Route Search or Route Tasks script.
- **Test cases:** `docs/PLM_Route_Export_Test_Cases.xlsx`.

---

# PLM – 3DEXPERIENCE Route Search Export

## Purpose
Read-only export of every route returned by the top bar search "Route" (1465 on 05-Nov-2026), with the tasks and content of each route, to one Excel workbook.

## Flow
1. Log in on 3DEXPERIENCE ID if the saved session has expired, then open the plain dashboard (no app open).
2. Wait until the top bar search is enabled, open Guided Search and type `Route`. What was typed is read back and retyped if letters were dropped.
3. Press Enter; if Enter only highlights the **Route** suggestion, pick it and run the search. If no results page appears, the search is repeated from the Route dashboard tab (`THREEDX_DASHBOARD_URL`).
4. Read the result count from `#search-nb-result` and scroll the result grid until every row is saved.
5. For each record: right click the title > **Open with** > **Route Management**, read the **Tasks** and **Content** tabs, close the app, and search again on the loaded dashboard to get back to the results.
6. Write `output/Route_Search_<date>.xlsx`.

## Output
| Sheet   | One row per   | Columns |
|---------|---------------|---------|
| Results | search result | Row, Title, Name, Type, Description, Maturity State, Owner, Revision, Project, Creation Date, Modification Date |
| Details | opened route  | Row, Title, Name, Opened |
| Tasks   | task          | Row, Route Title (merged over the route's tasks), Route Name, then every task field: Task Order, Task Name, Task Title, Assignee, Action, Task State, Approval Status, Due Date, Completed Date, Instructions, Comments, Priority, Originator, Originated Date, Modified Date ... Task settings (Requirement, Needs Owner Review, Allow Delegation, Assignee Sets Due Date, Parallel Rule, Template Task, Initial Task, E-sign Policy, Accept Link), user group and route owner UG fields, Candidate Assignee and internal IDs are left out |
| Content | content item  | Row, Route Title (merged), Route Name, every column of the Content grid |

Results and Details are Excel tables (banded rows, filters). Tasks and Content use a plain filter because they contain merged cells; each route is one block with alternate shading and a line under it. All dates are Excel dates shown as dd-mmm-yyyy. Creation Date is only filled for rows whose column the grid drew; Modification Date is filled for every row.

## Controls
- **Platform notices:** banners such as the R2027x update notice cover the top bar search box; they are hidden on every page load (`threedx.hide_platform_notices`). Nothing is clicked or dismissed on the platform.
- **Read only:** same network guard as the Route Export; the summary lists blocked writes.
- **Resume:** progress is saved after every record in `output/search_progress.json`. Running the same command again skips records already opened. If the result count has changed, the list is read again and opened records are kept, matched by Name.
- **Shifted rows:** a record is opened by its saved title; if routes were added and the title is not at its row, the record is listed as failed and the run continues.
- **Failures:** screenshot and page capture in `debug/search_<date>/`; the first three records are captured too.
- **Browser:** opens maximised and sized to the window.

## Run
```
.venv\Scripts\python.exe scripts\search_export.py                  all records (continues saved progress)
.venv\Scripts\python.exe scripts\search_export.py --max-records 5  first 5 records only
.venv\Scripts\python.exe scripts\search_export.py --list-only      result list only
.venv\Scripts\python.exe scripts\search_export.py --fresh          ignore saved progress
```
About 20–35 seconds per record, so a full run of 1465 records takes roughly 10–12 hours.

## Test status (05-Nov-2026)
- **Live:** RTS-01 to RTS-03, RTS-05 to RTS-10 and RTS-13 to RTS-15 passed. 1465 result rows saved; 6 records opened (79 tasks, 22 content items), 0 failed.
- **12-Nov-2026:** 1479 rows saved, 2 records opened (26 tasks, 4 content), 0 failed. RTS-16 (update notice over the search box) and RTS-17 (row redrawn while scrolling: looked up again) passed.
- **Not run yet:** RTS-04 (fallback start point), RTS-11 (count changed, records kept) and RTS-12 (shifted rows).
- **Test cases:** `docs/PLM_Route_Export_Test_Cases.xlsx`, sheet Route Search.

# PLM – 3DEXPERIENCE Route Tasks (approved routes, time taken)

## Purpose
Time taken by each task and each route, for routes where every task is approved. Default is 2 routes. Task columns follow the task box in the Route Management flowchart, plus creation, due, completion and modification dates.

## Flow
1. Log in on 3DEXPERIENCE ID if the saved session has expired, then open the plain dashboard.
2. Search `Route` in the top bar once and wait for the results page.
3. Go down the results. Routes still in process (status from the search service, e.g. `Route.In Process`) are skipped without being read. More results are loaded by scrolling the same list; the search is not repeated.
4. For each completed route, send one read-only GET from the dashboard page: `/enovia/resources/v1/modeler/routes/<id>?$include=tasks`. This is the call Route Management makes when a route is opened.
5. Keep the route only when every task is finished: approval tasks Approved, notification tasks Sent, comment tasks Commented; none open, none rejected.
6. Stop when N routes are kept (`--routes`, default 2) and write `output/Route_Tasks_<date>.xlsx` (one sheet, Tasks).
7. **Weekly full run (`--all`):** every route in the results, in-process ones included; no approval filter. The grid is scrolled until every result is loaded. About 1 - 1.5 hours for ~1,480 routes.
8. **Shared Google Sheet:** the rows replace the Tasks tab of `Route_Tasks_2026` (tab gid 804083112) through an Apps Script web app (`docs/apps-script/plmRouteTasks.gs`, token-checked). Route values are repeated on every row there (no merged cells) so filters work. The Excel file stays as the local copy.

## Output
| Sheet  | One row per | Columns |
|--------|-------------|---------|
| Tasks  | task        | Row, Route Title, Route Created, Route Completed (all three merged per route), Task Title, Status, Assignee, State, Priority, Created, Due, Completed, Modified, Days Taken, Time Taken, Late. No separate Routes sheet. Status, Priority, Due, Time Taken and Late are hidden (still in the file; select the columns either side, right click > Unhide) |

| Column | Meaning |
|--------|---------|
| Task Created | Date and time the task was assigned (`originated`) |
| Task Completed | Date and time the task was approved / commented / notified |
| Task Time Taken | Task Created to Task Completed, as days (2 decimals) and as `1d 4h 12m` |
| Late | Yes when the task was completed after its due date |
| Route Created | Date and time the route was created |
| Route Completed | Latest task completed; empty while any task is still open |

All dates are local time, written as `dd-Mmm-yyyy hh:mm`. Notification tasks complete when they are sent, so they show `0h 0m`.

## Controls
- **Platform notices:** banners such as the R2027x update notice cover the top bar search box; they are hidden on every page load (`threedx.hide_platform_notices`). Nothing is clicked or dismissed on the platform.
- **Read only:** only GET calls; the network guard is installed as in the other scripts and the summary lists blocked writes.
- **Security context:** taken from the user's preferred credentials; `THREEDX_SECURITY_CONTEXT` in `.env` overrides it.
- **Not enough approved routes:** the run stops after the loaded results are exhausted, writes the routes found and says how many.
- **Failures:** a route whose call fails is listed under Failed routes; the search for approved routes continues.

## Run
```
.venv\Scripts\python.exe scripts\route_tasks.py              2 approved routes
.venv\Scripts\python.exe scripts\route_tasks.py --routes 10  10 approved routes
.venv\Scripts\python.exe scripts\route_tasks.py --all        every route + Google Sheet (weekly)
.venv\Scripts\python.exe scripts\route_tasks.py --push-last  send the last run to the Google Sheet again
```
About 1 minute for 2 routes, most of it login and search.

## Test status (12-Nov-2026)
- **Live 12-Nov-2026:** RTT-01, RTT-02, RTT-08 and RTT-13 passed on the new layout. R-RMIN-0002844: 11 tasks, 0h 2m. RX02 TCKO ROUTE FOR DOUBLE PIPE CLIP: 11 tasks, 5d 3h 33m. 15 routes skipped; 0 failed, 0 blocked writes.
- **Live 11-Nov-2026:** RTT-01, RTT-02, RTT-05 to RTT-07 and RTT-09 to RTT-11 passed. R-RMIN-0002823: 17 tasks, 7d 18h 59m, 6 late. R-RMIN-0002812: 9 tasks, 10d 2h 58m, 1 late. 16 routes skipped (in process); 0 failed, 0 blocked writes.
- **Offline:** RTT-T01 to RTT-T10 passed (`tests\test_offline.py`, 19/19). 12-Nov-2026: Status, Priority, Due, Time Taken and Late hidden (RTT-T09); Routes sheet removed, Route Created / Route Completed moved into the Tasks sheet, Route Name removed (RTT-T10).
- **Offline 12-Nov-2026:** RTT-T11 (in-process route) and RTT-T12 (sheet push parts) passed, 21/21.
- **Not run yet:** RTT-03 (`--routes N`), RTT-04 (fewer approved routes than asked), RTT-12 (route call fails), RTT-15 and RTT-16 (Google Sheet, after the web app is deployed). RTT-14 (full run) in progress.
- **Test cases:** `docs/PLM_Route_Export_Test_Cases.xlsx`, sheet Route Tasks.
