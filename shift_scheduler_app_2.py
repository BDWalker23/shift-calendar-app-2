import streamlit as st
from datetime import date
from collections import defaultdict
import calendar
import random
from calendar_excel_generator import generate_excel_calendar

st.set_page_config(layout="wide")

st.title("📆 Shift Scheduler (Version 2) – N969PW")

names = ["Brandon", "Tony", "Erik"]
today = date.today()

selected_year = st.selectbox("Select Year", list(range(2025, 2031)), index=0)
selected_month = st.selectbox("Select Month", list(calendar.month_name)[1:], index=today.month - 1)

def evenly_distribute(schedule, days_in_month):
    name_counts = {name: list(schedule.values()).count(name) for name in names}
    target = days_in_month // len(names)
    extras = days_in_month % len(names)

    ideal_counts = {name: target for name in names}
    for i in range(extras):
        ideal_counts[names[i]] += 1

    date_list = list(schedule.keys())
    random.shuffle(date_list)

    for dt in date_list:
        current = schedule[dt]
        for name in names:
            if name_counts[name] < ideal_counts[name]:
                name_counts[current] -= 1
                name_counts[name] += 1
                schedule[dt] = name
                break

    return schedule

def assign_weekend(schedule, year, month):
    weekends_given = {name: False for name in names}
    cal = calendar.Calendar()
    month_days = list(cal.itermonthdates(year, month))

    # Build list of valid Fri–Sat–Sun weekend blocks
    weekend_blocks = []
    for i in range(len(month_days) - 2):
        fri, sat, sun = month_days[i:i+3]
        if (fri.weekday() == 4 and sat.weekday() == 5 and sun.weekday() == 6 and
            fri.month == month and sat.month == month and sun.month == month):
            weekend_blocks.append((fri, sat, sun))

    # Assign one full weekend OFF per person
    for name in names:
        for fri, sat, sun in weekend_blocks:
            # Only assign if all 3 days are unassigned
            if all(d not in schedule for d in [fri, sat, sun]):
                schedule[fri] = name
                schedule[sat] = name
                schedule[sun] = name
                weekends_given[name] = True
                break

    return schedule

st.markdown("### Select OFF days for each person")

schedule = defaultdict(str)
month_range = calendar.monthrange(selected_year, list(calendar.month_name).index(selected_month))[1]

cols = st.columns(7)
for day in range(1, month_range + 1):
    col = cols[(day - 1) % 7]
    with col:
        dt = date(selected_year, list(calendar.month_name).index(selected_month), day)
        selected = st.selectbox(f"{day}", [""] + names, key=str(dt))
        if selected:
            schedule[dt] = selected

even = st.checkbox("✅ Evenly distribute other unselected OFF days")
assign_weekend_pref = st.checkbox("✅ Ensure everyone gets one weekend OFF")

if st.button("📥 Generate Calendar"):
    try:
        output_path = f"/tmp/{selected_month}_{selected_year}_calendar.xlsx"
        final_schedule = autopopulate_schedule(
            selected_year,
            list(calendar.month_name).index(selected_month),
            dict(schedule)
        )
        generate_excel_calendar(
            selected_year,
            list(calendar.month_name).index(selected_month),
            final_schedule,
            output_path
        )

        with open(output_path, "rb") as f:
            st.success("✅ Excel calendar generated!")
            st.download_button(
                label="📥 Download Excel Calendar",
                data=f,
                file_name=f"{selected_month}_{selected_year}_Shift_Calendar.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        st.error(f"Error generating calendar: {e}")
