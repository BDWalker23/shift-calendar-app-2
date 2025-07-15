import calendar
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.styles.borders import Border, Side
from openpyxl.utils import get_column_letter
from collections import Counter
import random
from datetime import timedelta

def generate_excel_calendar(year, month, schedule, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"

    # Add Month and Year title
    title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = title
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    header_row = 2
    row_start = header_row + 1

    # Headers
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for col, day in enumerate(days, start=1):
        ws.cell(row=header_row, column=col).value = day
        ws.cell(row=header_row, column=col).font = Font(bold=True)
        ws.cell(row=header_row, column=col).alignment = Alignment(horizontal="center")

    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdatescalendar(year, month)

    thin_border = Border(top=Side(style="thin"), bottom=Side(style="thin"))

    for week_idx, week in enumerate(month_days):
        for day_idx, day in enumerate(week):
            cell = ws.cell(row=row_start + week_idx, column=day_idx + 1)
            if day.month == month:
                off_name = schedule.get(day, "")
                on_names = [n for n in ["Brandon", "Tony", "Erik"] if n != off_name]

                cell_value = f"{day.day}\n"
                if off_name:
                    cell_value += f"{off_name} OFF\n{on_names[0]} & {on_names[1]} ON"

                cell.value = cell_value
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = thin_border  # ✅ Add border here

                if off_name:
                    fill_color = get_color(off_name)
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

                # Make day number bold and underlined
                if cell_value:
                    cell.font = Font(bold=False)
                    cell_value_lines = cell_value.split("\n")
                    if cell_value_lines:
                        day_str = cell_value_lines[0]
                        cell.value = f"{day_str}\n" + "\n".join(cell_value_lines[1:])

    for i in range(1, 8):  # Columns A–G (1–7)
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = 25

    wb.save(file_path)
