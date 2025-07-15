import streamlit as st
from datetime import date
from collections import defaultdict
import calendar
import random
from calendar_excel_generator import generate_excel_calendar

st.set_page_config(layout="wide")

st.title("📆 Shift Scheduler (Version 2)")

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
    month_days = cal.itermonthdates(year, month)

    for week in calendar.monthcalendar(year, month):
        fri = week[calendar.FRIDAY]
        sat = week[calendar.SATURDAY]
        sun = week[calendar.SUNDAY]

        if fri and sat and sun:
            available_names = [name for name, given in weekends_given.items() if not given]
            if available_names:
                chosen = available_names[0]
                for d in [fri, sat, sun]:
                    try:
                        schedule[date(year, month, d)] = chosen
                    except:
                        continue
                weekends_given[chosen] = True
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
        all_days = [date(selected_year, list(calendar.month_name).index(selected_month), d)
                    for d in range(1, month_range + 1)]

        unfilled = [d for d in all_days if d not in schedule]

        for d in unfilled:
            schedule[d] = random.choice(names)

        if assign_weekend_pref:
            schedule = assign_weekend(schedule, selected_year, list(calendar.month_name).index(selected_month))

        if even:
            schedule = evenly_distribute(schedule, month_range)

        excel_bytes = generate_excel_calendar(selected_year, list(calendar.month_name).index(selected_month), schedule)

        st.success("✅ Excel calendar generated!")
        st.download_button(
            label="📥 Download Excel Calendar",
            data=excel_bytes,
            file_name=f"{selected_month}_{selected_year}_Shift_Calendar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Error generating calendar: {e}")
