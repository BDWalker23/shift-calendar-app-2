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
    from collections import Counter
    from datetime import date
    import calendar
    import random

    all_names = ["Brandon", "Tony", "Erik"]
    schedule = existing_schedule.copy()
    cal = calendar.Calendar(firstweekday=6)

    # 1. Get all valid days in the month
    all_days = [day for week in cal.monthdatescalendar(year, month) for day in week if day.month == month]

    # 2. Lock in manually selected OFF days
    locked_days = set(schedule.keys())

    # 3. Assign one full Fri–Sat–Sun OFF block to each person
    weekends = []
    for week in cal.monthdatescalendar(year, month):
        fri = week[calendar.FRIDAY]
        sat = week[calendar.SATURDAY]
        sun = week[calendar.SUNDAY]
        if all(d.month == month and d not in locked_days for d in [fri, sat, sun]):
            weekends.append((fri, sat, sun))

    used_days = set()
    assigned = {}

    for name in all_names:
        for fri, sat, sun in weekends:
            if not any(d in used_days or d in locked_days for d in [fri, sat, sun]):
                schedule[fri] = name
                schedule[sat] = name
                schedule[sun] = name
                used_days.update([fri, sat, sun])
                assigned[name] = (fri, sat, sun)
                break

    locked_days.update(used_days)

    # 4. Fill remaining unassigned days
    off_count = Counter(schedule.values())
    unfilled = [d for d in all_days if d not in schedule]

    # Calculate target OFF days per person (balanced ±1)
    base_off = len(all_days) // len(all_names)
    extras = len(all_days) % len(all_names)
    target_off = {name: base_off for name in all_names}
    for i in range(extras):
        target_off[all_names[i]] += 1

    # Assign remaining OFF days based on current count
    random.shuffle(unfilled)
    for day in unfilled:
        sorted_names = sorted(all_names, key=lambda n: off_count[n])
        for name in sorted_names:
            if off_count[name] < target_off[name]:
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
