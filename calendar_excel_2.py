from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
import calendar
from datetime import datetime
from io import BytesIO

def get_color(name):
    if name == "Brandon":
        return "FFFF00"  # Yellow
    elif name == "Tony":
        return "00FF00"  # Green
    elif name == "Erik":
        return "00B0F0"  # Blue
    else:
        return "FFFFFF"  # White

def generate_excel_calendar(year, month, schedule):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year}"

    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for col, day in enumerate(days, start=1):
        ws.cell(row=1, column=col).value = day
        ws.cell(row=1, column=col).font = Font(bold=True)
        ws.cell(row=1, column=col).alignment = Alignment(horizontal="center")

    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdayscalendar(year, month)
    row_start = 2

    for row_idx, week in enumerate(weeks, start=row_start):
        for col_idx, day in enumerate(week, start=1):
            if day == 0:
                continue
            cell = ws.cell(row=row_idx, column=col_idx)
            cell_value = f"{day}"
            date_obj = datetime(year, month, day).date()
            if date_obj in schedule:
                name_off = schedule[date_obj]
                on_names = [n for n in ["Brandon", "Tony", "Erik"] if n != name_off]
                cell_value += f"\n{name_off} OFF\n{on_names[0]} & {on_names[1]} ON"
                fill = PatternFill(start_color=get_color(name_off), end_color=get_color(name_off), fill_type="solid")
                cell.fill = fill
            cell.value = cell_value
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.font = Font(size=9)

    # Auto-adjust column widths
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = 18

    # Adjust row height
    for row in range(2, 2 + len(weeks)):
        ws.row_dimensions[row].height = 70

    excel_bytes = BytesIO()
    wb.save(excel_bytes)
    excel_bytes.seek(0)
    return excel_bytes# Excel calendar generation code placeholder
