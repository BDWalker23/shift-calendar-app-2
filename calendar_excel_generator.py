import calendar
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from collections import Counter
import random

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

    cal = calendar.Calendar(firstweekday=6)
    all_days = [d for week in cal.monthdatescalendar(year, month) for d in week if d.month == month]

    # Step 1: Find all possible Friday-Saturday-Sunday combos
    weekends = []
    for i in range(len(all_days) - 2):
        if all_days[i].weekday() == 4 and all_days[i+1].weekday() == 5 and all_days[i+2].weekday() == 6:
            weekends.append((all_days[i], all_days[i+1], all_days[i+2]))

    # Step 2: Assign each person to one unused weekend
    used_days = set(schedule.keys())
    assigned_weekends = set()

    for name in all_names:
        for fri, sat, sun in weekends:
            if fri not in used_days and sat not in used_days and sun not in used_days:
                schedule[fri] = name
                schedule[sat] = name
                schedule[sun] = name
                used_days.update([fri, sat, sun])
                assigned_weekends.add((fri, sat, sun))
                break

    # Step 3: Fill remaining unassigned days, balancing counts
    off_count = Counter(schedule.values())

    for day in all_days:
        if day not in schedule:
            # Choose the name with the fewest OFF days so far
            least = min(all_names, key=lambda n: off_count[n])
            schedule[day] = least
            off_count[least] += 1

    return schedule
def generate_excel_calendar(year, month, schedule, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"

    # Title
    title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = title
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    # Headers
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for col, day in enumerate(days, start=1):
        cell = ws.cell(row=2, column=col)
        cell.value = day
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Fill calendar days
    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdatescalendar(year, month)

    thin_border = Border(
        top=Side(style="thin"),
        bottom=Side(style="thin"),
        left=Side(style="thin"),
        right=Side(style="thin")
    )

    for week_idx, week in enumerate(month_days):
        for day_idx, day in enumerate(week):
            cell = ws.cell(row=3 + week_idx, column=1 + day_idx)
            if day.month == month:
                off_name = schedule.get(day, "")
                on_names = [n for n in ["Brandon", "Tony", "Erik"] if n != off_name]

                lines = [f"{day.day}"]
                if off_name:
                    lines.append(f"{off_name} OFF")
                    lines.append(f"{on_names[0]} & {on_names[1]} ON")

                cell.value = "\n".join(lines)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = thin_border

                if off_name:
                    fill_color = get_color(off_name)
                    cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

                # Bold + underline the date number only
                if lines:
                    date_font = Font(bold=True, underline="single")
                    cell.font = date_font
            else:
                cell.border = thin_border

    # Set column width
    for col_idx in range(1, 8):
        col_letter = chr(64 + col_idx)
        ws.column_dimensions[col_letter].width = 25

    wb.save(file_path)# This is a placeholder so Streamlit doesn’t throw an error when trying to import.
# Replace this file content with actual function calls in your project.
