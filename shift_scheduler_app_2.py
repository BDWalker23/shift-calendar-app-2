import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from io import BytesIO
from datetime import datetime
import calendar

st.set_page_config(layout="wide")

st.title("Monthly Shift Calendar App 2")

# Define names and colors
names = ["Tony", "Erik", "Brandon"]
colors = {
    "Tony": "90EE90",      # Light Green
    "Erik": "ADD8E6",      # Light Blue
    "Brandon": "FFFF99"    # Light Yellow
}

# Month and year selection
year = st.selectbox("Select Year", list(range(2025, 2031)), index=0)
month = st.selectbox("Select Month", list(calendar.month_name)[1:], index=datetime.now().month - 1)

# Generate days in selected month
num_days = calendar.monthrange(year, list(calendar.month_name).index(month))[1]
dates = [datetime(year, list(calendar.month_name).index(month), day) for day in range(1, num_days + 1)]

# Initialize schedule dictionary
schedule = {}

# Editable calendar input
st.markdown("### Assign OFF days")
cols = st.columns(7)
for date in dates:
    col = cols[date.weekday() % 7]
    with col:
        selected = st.selectbox(f"{date.strftime('%a %d')}", ["", *names], key=str(date))
        if selected:
            schedule[date] = selected

# Generate Excel calendar
def generate_excel(schedule, month, year):
    wb = Workbook()
    ws = wb.active
    ws.title = f"{month}_{year}"

    # Set up header row with weekdays
    days_of_week = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    for col_num, day in enumerate(days_of_week, 1):
        ws.cell(row=1, column=col_num, value=day)
        ws.column_dimensions[chr(64 + col_num)].width = 20

    cal = calendar.Calendar(firstweekday=6)
    month_dates = cal.monthdatescalendar(year, list(calendar.month_name).index(month))

    for row_num, week in enumerate(month_dates, 2):
        for col_num, day in enumerate(week, 1):
            if day.month == list(calendar.month_name).index(month):
                cell = ws.cell(row=row_num, column=col_num)
                off_name = schedule.get(day, "")
                on_names = [n for n in names if n != off_name]
                cell_value = f"{day.day}"
                if off_name:
                    cell_value += f"{day.day}\n{name} OFF\n{on_names} ON"
{off_name} OFF
{on_names[0]} & {on_names[1]} ON"
                cell.value = cell_value
                cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="top")
                cell.font = Font(name='Calibri', size=11)
                if off_name in colors:
                    fill = PatternFill(start_color=colors[off_name], end_color=colors[off_name], fill_type="solid")
                    cell.fill = fill
                # Bold and underline the day number
                cell.font = Font(bold=True, underline="single", name='Calibri', size=11)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# Button to export
if st.button("Generate Calendar Excel"):
    output = generate_excel(schedule, month, year)
    st.download_button(
        label="📥 Download Calendar Excel File",
        data=output,
        file_name=f"shift_calendar_{month}_{year}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
