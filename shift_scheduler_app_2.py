
import streamlit as st
from datetime import date, timedelta
import calendar
from calendar_pdf_2 import CalendarPDF, get_color

st.set_page_config(layout="wide")
st.title("August 2025 Shift Calendar (v2)")

names = ["Brandon", "Tony", "Erik"]
year = 2025
month = 8

# Build a default empty schedule for the month
def get_default_schedule():
    start_date = date(year, month, 1)
    end_date = date(year, month, calendar.monthrange(year, month)[1])
    return {start_date + timedelta(days=i): "" for i in range((end_date - start_date).days + 1)}

# Calendar-style selector
st.subheader("Assign Off Days")
shift_data = get_default_schedule()
cols = st.columns(7)
for d in shift_data:
    idx = d.weekday() if calendar.firstweekday() == 0 else (d.weekday() + 1) % 7
    with cols[idx]:
        shift_data[d] = st.selectbox(f"{d.strftime('%a %d')}", [""] + names, key=str(d))

# Generate calendar-style PDF when button is pressed
if st.button("Generate PDF"):
    pdf = CalendarPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()
    pdf.draw_calendar(year, month, shift_data)
    pdf_path = "shift_calendar_august_2025_v2.pdf"
    pdf.output(pdf_path)
    with open(pdf_path, "rb") as f:
        st.download_button("📥 Download Calendar PDF", f, file_name=pdf_path)
