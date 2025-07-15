
import streamlit as st
import calendar
from datetime import date, datetime, timedelta
from calendar_pdf_2 import CalendarPDF, get_color
from autopopulate_schedule_2 import autopopulate_schedule

st.set_page_config(layout="wide")
st.title("Shift Calendar App v2")

years = list(range(2025, 2031))
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Allow user to pick any month/year from 2025 to 2030
st.sidebar.header("Calendar Settings")
selected_year = st.selectbox("Select Year", options=years)
selected_month_name = st.selectbox("Select Month", options=months)
year = int(selected_year)
month = int(months.index(selected_month_name)) + 1

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
    try:
        from autopopulate_schedule_2 import autopopulate_schedule  # make sure this file is also in your repo

        full_schedule = autopopulate_schedule(year, month, schedule)

        pdf = CalendarPDF()
        pdf.draw_calendar(year, month, full_schedule)

        output_filename = f"shift_calendar_{calendar.month_name[month].lower()}_{year}_v2.pdf"
        pdf.output(output_filename)

        with open(output_filename, "rb") as f:
            st.download_button("Download PDF", f, file_name=output_filename)

    except Exception as e:
        st.error(f"Error generating PDF: {e}")
