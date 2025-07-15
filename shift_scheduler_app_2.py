
import streamlit as st
import calendar
from datetime import date, datetime, timedelta
from calendar_pdf_2 import CalendarPDF, get_color

st.set_page_config(layout="wide")
st.title("Shift Calendar App v2")

# Allow user to pick any month/year from 2025 to 2030
st.sidebar.header("Calendar Settings")
year = int(st.selectbox("Select Year", options=years))
month = int(months.index(st.selectbox("Select Month", options=months)) + 1)

names = ["Brandon", "Tony", "Erik"]

# Get calendar layout
def get_default_schedule(y, m):
    start_date = date(y, m, 1)
    end_date = date(y, m, calendar.monthrange(y, m)[1])
    return {start_date + timedelta(days=i): "" for i in range((end_date - start_date).days + 1)}

shift_data = get_default_schedule(year, month)

# Render a calendar grid layout using columns
st.subheader(f"Select Off Days for {calendar.month_name[month]} {year}")
days_of_week = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
calendar.setfirstweekday(0)
weeks = calendar.Calendar().monthdatescalendar(year, month)

for week in weeks:
    cols = st.columns(7)
    for i, day in enumerate(week):
        if day.month == month:
            shift_data[day] = cols[i].selectbox(
                f"{day.strftime('%a %d')}", [""] + names, key=str(day)
            )
        else:
            cols[i].markdown(" ")

# Generate calendar-style PDF
if st.button("Generate PDF"):
    pdf = CalendarPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.draw_calendar(year, month, shift_data)
    filename = f"shift_calendar_{calendar.month_name[month].lower()}_{year}_v2.pdf"
    pdf.output(filename)
    with open(filename, "rb") as f:
        st.download_button("📥 Download Calendar PDF", f, file_name=filename)
