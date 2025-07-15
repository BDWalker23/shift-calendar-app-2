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

    all_days = [day for week in cal.monthdatescalendar(year, month) for day in week if day.month == month]
    locked_days = set(schedule.keys())

    # Step 1 – Assign full Fri-Sat-Sun weekend OFF per person
    weekend_blocks = []
    for week in cal.monthdatescalendar(year, month):
        fri, sat, sun = week[calendar.FRIDAY], week[calendar.SATURDAY], week[calendar.SUNDAY]
        if all(d.month == month for d in [fri, sat, sun]):
            weekend_blocks.append((fri, sat, sun))

    random.shuffle(weekend_blocks)
    assigned_weekends = set()
    for name in all_names:
        for fri, sat, sun in weekend_blocks:
            if (fri, sat, sun) not in assigned_weekends:
                if all(d not in locked_days for d in [fri, sat, sun]):
                    schedule[fri] = name
                    schedule[sat] = name
                    schedule[sun] = name
                    locked_days.update([fri, sat, sun])
                    assigned_weekends.add((fri, sat, sun))
                    break

    # Step 2 – Evenly distribute remaining OFF days (±1)
    off_count = Counter(schedule.values())
    unfilled = [d for d in all_days if d not in schedule]

    base = len(all_days) // len(all_names)
    extras = len(all_days) % len(all_names)
    target = {name: base for name in all_names}
    for i in range(extras):
        target[all_names[i]] += 1

    random.shuffle(unfilled)
    for day in unfilled:
        sorted_names = sorted(all_names, key=lambda n: off_count[n])
        for name in sorted_names:
            if off_count[name] < target[name]:
                schedule[day] = name
                off_count[name] += 1
                break

    return schedule

def generate_excel_calendar(year, month, schedule, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"

    # Title row
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    cell = ws.cell(row=1, column=1)
    cell.value = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"
    cell.font = Font(size=16, bold=True)
    cell.alignment = Alignment(horizontal="center")

    # Day headers
    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for i, day in enumerate(days, start=1):
        header = ws.cell(row=2, column=i)
        header.value = day
        header.font = Font(bold=True)
        header.alignment = Alignment(horizontal="center")

    cal = calendar.Calendar(firstweekday=6)
    weeks = cal.monthdatescalendar(year, month)

    border = Border(
        top=Side(style="thin"), bottom=Side(style="thin"),
        left=Side(style="thin"), right=Side(style="thin")
    )

    for i, week in enumerate(weeks):
        for j, d in enumerate(week):
            cell = ws.cell(row=3 + i, column=1 + j)
            cell.border = border
            if d.month == month:
                off = schedule.get(d, "")
                on = [n for n in ["Brandon", "Tony", "Erik"] if n != off]
                lines = [f"{d.day}"]
                if off:
                    lines.append(f"{off} OFF")
                    lines.append(f"{on[0]} & {on[1]} ON")
                cell.value = "\n".join(lines)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if off:
                    fill = PatternFill(start_color=get_color(off), end_color=get_color(off), fill_type="solid")
                    cell.fill = fill
                cell.font = Font(bold=True, underline="single")

    for i in range(1, 8):
        ws.column_dimensions[chr(64 + i)].width = 25

    wb.save(file_path)
