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
    from collections import Counter
    import calendar
    import random

    all_names = ["Brandon", "Tony", "Erik"]
    schedule = existing_schedule.copy()
    cal = calendar.Calendar(firstweekday=6)

    # Get all valid days in the month
    all_days = [day for week in cal.monthdatescalendar(year, month) for day in week if day.month == month]

    # Lock in manually selected OFF days
    locked_days = set(schedule.keys())
    weekend_blocks = []

    for week in cal.monthdatescalendar(year, month):
        fri = week[calendar.FRIDAY]
        sat = week[calendar.SATURDAY]
        sun = week[calendar.SUNDAY]
        if all(d.month == month for d in [fri, sat, sun]):
            weekend_blocks.append((fri, sat, sun))

    assigned_weekend_names = set()

    # If someone selected only 1–2 days of a weekend, treat it as their full weekend off
    for name in all_names:
        for fri, sat, sun in weekend_blocks:
            selected_days = [d for d in [fri, sat, sun] if schedule.get(d) == name]
            if len(selected_days) >= 1 and len(selected_days) < 3:
                for d in [fri, sat, sun]:
                    if d not in locked_days:
                        schedule[d] = name
                        locked_days.add(d)
                assigned_weekend_names.add(name)
                break

    # Now try to give remaining names one full weekend off
    for name in all_names:
        if name in assigned_weekend_names:
            continue
        for fri, sat, sun in weekend_blocks:
            if all(d not in locked_days for d in [fri, sat, sun]):
                schedule[fri] = name
                schedule[sat] = name
                schedule[sun] = name
                locked_days.update([fri, sat, sun])
                assigned_weekend_names.add(name)
                break

    # Fill remaining unassigned days
    off_count = Counter(schedule.values())
    unfilled = [d for d in all_days if d not in schedule]

    # Calculate target OFF days per person (balanced ±1)
    base_off = len(all_days) // len(all_names)
    extras = len(all_days) % len(all_names)
    target_off = {name: base_off for name in all_names}
    for i in range(extras):
        target_off[all_names[i]] += 1

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
