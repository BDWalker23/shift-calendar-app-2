# Streamlit app code placeholder
import streamlit as st
import calendar
from datetime import date
from calendar_excel_generator import generate_excel_calendar

st.set_page_config(layout="wide")
st.title("Shift Calendar App v2")

years = list(range(2025, 2031))
months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

# Sidebar controls
st.sidebar.header("Calendar Settings")
selected_year = st.sidebar.selectbox("Select Year", options=years)
selected_month_name = st.sidebar.selectbox("Select Month", options=months)
year = int(selected_year)
month = months.index(selected_month_name) + 1
names = ["Brandon", "Tony", "Erik"]

# User input section
st.header("Assign OFF Days")
shift_data = {}
weeks = calendar.Calendar(firstweekday=6).monthdatescalendar(year, month)

for week in weeks:
    cols = st.columns(7)
    for i, day in enumerate(week):
        if day.month == month:
            shift_data[day] = cols[i].selectbox(
                f"{day.strftime('%a %d')}", [""] + names, key=str(day)
            )
        else:
            cols[i].markdown(" ")

# Generate calendar PDF
if st.button("Generate PDF"):
    try:
        output_filename = generate_excel_calendar(year, month, shift_data)
        with open(output_filename, "rb") as f:
            st.download_button("Download PDF", f, file_name=output_filename)
    except Exception as e:
        st.error(f"Error generating calendar: {e}")
