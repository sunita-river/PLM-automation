"""Write the Routes / Tasks / Content workbook: bold headers, frozen top row, filters,
auto-sized columns and dates as DD-MMM-YYYY."""
from datetime import datetime

from dateutil import parser as date_parser
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

ROUTE_COLUMNS = ["Route Name", "Title", "Description", "State/Status", "Owner", "Route Template",
                 "Created Date", "Modified Date", "Due Date", "Completion Date"]
TASK_COLUMNS = ["Route Name", "Task Name", "Assignee", "Action", "Task State", "Due Date",
                "Completed Date", "Approval Status", "Comments"]
CONTENT_COLUMNS = ["Route Name", "Name", "Type", "Revision"]

DATE_FORMAT = "DD-MMM-YYYY"
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")


def to_date(value):
    """Return a datetime for anything that looks like a date, else None."""
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) and value > 10**11:  # epoch milliseconds
        return datetime.fromtimestamp(value / 1000)
    try:
        return date_parser.parse(str(value)).replace(tzinfo=None)
    except (ValueError, OverflowError):
        return None


def _write_sheet(ws, columns, rows):
    ws.append(columns)
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
    date_cols = {i for i, c in enumerate(columns) if "Date" in c}
    for row in rows:
        values = []
        for i, col in enumerate(columns):
            v = row.get(col, "")
            if i in date_cols:
                v = to_date(v) or (v if v else None)
            values.append(v)
        ws.append(values)
        for i in date_cols:
            cell = ws.cell(row=ws.max_row, column=i + 1)
            if isinstance(cell.value, datetime):
                cell.number_format = DATE_FORMAT
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(ws.max_row, 1)}"
    for i, col in enumerate(columns, start=1):
        longest = len(col)
        for (cell,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            text = "dd-mmm-yyyy" if isinstance(cell.value, datetime) else str(cell.value or "")
            longest = max(longest, min(len(text), 80))
        ws.column_dimensions[get_column_letter(i)].width = longest + 3


def write_workbook(path, routes, tasks, content):
    wb = Workbook()
    _write_sheet(wb.active, ROUTE_COLUMNS, routes)
    wb.active.title = "Routes"
    _write_sheet(wb.create_sheet("Tasks"), TASK_COLUMNS, tasks)
    _write_sheet(wb.create_sheet("Content"), CONTENT_COLUMNS, content)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
