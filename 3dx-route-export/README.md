# 3DEXPERIENCE Route Export

Read-only export of every route in the 3DEXPERIENCE Route Management app (ENXROUT_AP)
to an Excel workbook with three sheets: Routes, Tasks and Content.

## Folder layout

| Path          | Purpose                                          |
|---------------|--------------------------------------------------|
| `.venv/`      | Project virtual environment (Python 3.14)        |
| `scripts/`    | Automation scripts                               |
| `downloads/`  | Files from the app's own Export/CSV option       |
| `output/`     | Generated `.xlsx` workbooks                      |

## Setup (already done)

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m playwright install chromium
```

Check the setup:

```powershell
.venv\Scripts\python.exe scripts\check_setup.py
```

## Run

```powershell
.venv\Scripts\python.exe scripts\route_export.py            # full export
.venv\Scripts\python.exe scripts\route_export.py --discover # save page structure to debug/
.venv\Scripts\python.exe scripts\route_tasks.py            # time taken per task / route, 2 fully approved routes (--routes N)
.venv\Scripts\python.exe tests\test_offline.py              # offline checks
```

Settings and credentials are in `.env` (template: `.env.example`). Flow, controls and test
status: [docs/Confluence.md](docs/Confluence.md); test cases: `docs/PLM_Route_Export_Test_Cases.xlsx`.

## Rules

- Read only. Route/task write requests are blocked in the browser before they are sent.
- Credentials live only in `.env`; the session is reused from `.auth/profile`.
