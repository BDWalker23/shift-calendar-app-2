import calendar
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from collections import Counter
import random
from datetime import timedelta

def get_color(name):
    if name == "Brandon":
        return "FFFF99"  # Yellow
    elif name == "Tony":
        return "CCFFCC"  # Green
    elif name == "Erik":
        return "99CCFF"  # Blue
    else:
        return "FFFFFF"  # Default white

def autopopulate_schedule(year, month, existing_schedule):
    all_names = ["Brandon", "Tony", "Erik"]
    schedule = existing_schedule.copy()

    # Count current off days
    off_count = Counter(schedule.values())
    while len(set(off_count.values())) > 1:
        # Normalize off counts to +/- 1
        most = off_count.most_common(1)[0][0]
        for day, name in schedule.items():
            if name == most:
                del schedule[day]
                off_count = Counter(schedule.values())
                break

    # Get all dates in the month
    cal = calendar.Calendar(firstweekday=6)
    all_days = [day for week in cal.monthdatescalendar(year, month) for day in week if day.month == month]

    # Fill in blanks
    for day in all_days:
        if day not in schedule:
            least_common = min(off_count, key=off_count.get)
            schedule[day] = least_common
            off_count[least_common] += 1

    # Ensure at least one full weekend off
    weekend_blocks = []
    for i in range(len(all_days) - 2):
        if all_days[i].weekday() == 4 and all_days[i+1].weekday() == 5 and all_days[i+2].weekday() == 6:
            weekend_blocks.append((all_days[i], all_days[i+1], all_days[i+2]))

    has_weekend_off = {name: False for name in all_names}
    for name in all_names:
        for fri, sat, sun in weekend_blocks:
            if all(schedule.get(d) == name for d in [fri, sat, sun]):
                has_weekend_off[name] = True
                break

    for name, has_off in has_weekend_off.items():
        if not has_off:
            for fri, sat, sun in weekend_blocks:
                if all(schedule.get(d) != name for d in [fri, sat, sun]):
                    schedule[fri] = name
                    schedule[sat] = name
                    schedule[sun] = name
                    break

    return schedule

def generate_excel_calendar(year, month, schedule, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"

    # Add Month and Year title
    title = f"{calendar.month_name[month]} {year} Shift Calendar"
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
                if off_name:
                    fill_color = get_color(off_name)
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

                # Make day number bold and underlined
                if cell_value:
                    cell.font = Font(bold=False)  # base font
                    cell_value_lines = cell_value.split("\n")
                    if cell_value_lines:
                        day_str = cell_value_lines[0]
                        cell.value = f"{day_str}\n" + "\n".join(cell_value_lines[1:])
                        
    # Set column widths for better layout
 from openpyxl.utils import get_column_letter

# Set fixed column widths for Sunday to Saturday (columns 1 to 7)
    for i in range(1, 8):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = 28  # Adjust as needed (25–30 is usually good)
        ws.column_dimensions[col_letter].width = 25  # adjust to 28 if you want even wider
    wb.save(file_path)
