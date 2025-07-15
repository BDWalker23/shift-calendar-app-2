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
        return "FFFFFF"

def autopopulate_schedule(year, month, existing_schedule):
    all_names = ["Brandon", "Tony", "Erik"]
    schedule = existing_schedule.copy()
    cal = calendar.Calendar(firstweekday=6)
    all_days = [day for week in cal.monthdatescalendar(year, month) for day in week if day.month == month]
    locked_days = set(schedule.keys())

    # Step 1: Assign 1 full weekend OFF to each person (Fri/Sat/Sun)
    weekends = []
    for week in cal.monthdatescalendar(year, month):
        fri, sat, sun = week[calendar.FRIDAY], week[calendar.SATURDAY], week[calendar.SUNDAY]
        if all(d.month == month and d not in locked_days for d in [fri, sat, sun]):
            weekends.append((fri, sat, sun))

    used_days = set()
    used_names = set()
    for fri, sat, sun in weekends:
        available = [n for n in all_names if n not in used_names]
        if not available:
            break
        pick = available[0]
        schedule[fri] = pick
        schedule[sat] = pick
        schedule[sun] = pick
        used_names.add(pick)
        used_days.update([fri, sat, sun])

    locked_days.update(used_days)

    # Step 2: Fill in rest of unassigned OFF days evenly
    unfilled = [d for d in all_days if d not in schedule]
    off_count = Counter(schedule.values())

    base = len(all_days) // len(all_names)
    extras = len(all_days) % len(all_names)
    targets = {n: base for n in all_names}
    for i in range(extras):
        targets[all_names[i]] += 1

    random.shuffle(unfilled)
    for d in unfilled:
        sorted_names = sorted(all_names, key=lambda n: off_count[n])
        for n in sorted_names:
            if off_count[n] < targets[n]:
                schedule[d] = n
                off_count[n] += 1
                break

    return schedule

def generate_excel_calendar(year, month, schedule, file_path):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=7)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"{calendar.month_name[month]} {year} Shift Calendar – N969PW"
    title_cell.font = Font(size=16, bold=True)
    title_cell.alignment = Alignment(horizontal="center")

    days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for col, name in enumerate(days, 1):
        cell = ws.cell(row=2, column=col)
        cell.value = name
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    cal = calendar.Calendar(firstweekday=6)
    month_days = cal.monthdatescalendar(year, month)
    thin_border = Border(top=Side(style="thin"), bottom=Side(style="thin"),
                         left=Side(style="thin"), right=Side(style="thin"))

    for r, week in enumerate(month_days, start=3):
        for c, day in enumerate(week, start=1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if day.month == month:
                off = schedule.get(day, "")
                on = [n for n in ["Brandon", "Tony", "Erik"] if n != off]
                lines = [f"{day.day}"]
                if off:
                    lines.append(f"{off} OFF")
                    lines.append(f"{on[0]} & {on[1]} ON")
                cell.value = "\n".join(lines)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                if off:
                    color = get_color(off)
                    cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                cell.font = Font(bold=True, underline="single")

    for i in range(1, 8):
        ws.column_dimensions[chr(64 + i)].width = 25

    wb.save(file_path)
