
from fpdf import FPDF
import calendar
from datetime import date, timedelta
from collections import Counter
import random

def get_color(name):
    if name == "Brandon":
        return (0, 102, 204)
    elif name == "Tony":
        return (200, 30, 30)
    elif name == "Erik":
        return (0, 153, 76)
    else:
        return (0, 0, 0)

class CalendarPDF(FPDF):
    def __init__(self, orientation="L", unit="mm", format="A4"):
        super().__init__(orientation, unit, format)
        self.set_auto_page_break(auto=True, margin=10)

    def draw_calendar(self, year, month, schedule):
        self.add_page()
        self.set_font("Arial", "B", 16)
        title = f"{calendar.month_name[month]} {year} Shift Schedule (Color-Coded Off Days)"
        self.set_text_color(0)
        self.cell(0, 10, title, ln=True, align="C")
        self.ln(5)

        self.set_font("Arial", "B", 10)
        cell_width = 40
        cell_height = 30  # Tripled height
        days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
        for day in days:
            self.cell(cell_width, 10, day, border=1, align="C")
        self.ln()

        cal = calendar.Calendar(firstweekday=6)
        weeks = cal.monthdatescalendar(year, month)

        for week in weeks:
            for day in week:
                x = self.get_x()
                y = self.get_y()
                self.rect(x, y, cell_width, cell_height)
                if day.month == month:
                    self.set_xy(x + 1, y + 1)
                    self.set_font("Arial", "B", 9)
                    self.set_text_color(0)
                    self.cell(cell_width - 2, 5, str(day.day), ln=1)

                    if day in schedule:
                        name = schedule[day]
                        others = [n for n in ["Brandon", "Tony", "Erik"] if n != name]
                        r, g, b = get_color(name)
                        self.set_font("Arial", "", 9)
                        self.set_xy(x + 1, y + 7)
                        self.set_text_color(r, g, b)
                        self.cell(cell_width - 2, 5, f"{name} OFF", ln=1)
                        self.set_xy(x + 1, y + 13)
                        self.set_text_color(0)
                        self.cell(cell_width - 2, 5, f"{others[0]} & {others[1]} ON", ln=1)
                self.set_xy(x + cell_width, y)
            self.ln()
