
from fpdf import FPDF
import calendar

def get_color(name):
    if name == "Brandon":
        return (0, 0, 255)
    elif name == "Tony":
        return (200, 0, 0)
    elif name == "Erik":
        return (0, 130, 0)
    else:
        return (0, 0, 0)

class CalendarPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, "August 2025 - Shift Calendar", ln=True, align="C")

    def draw_calendar(self, year, month, schedule):
        self.set_font("Arial", "B", 10)
        days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
        cell_w = 40
        cell_h = 25

        for day in days:
            self.cell(cell_w, 10, day, border=1, align="C", fill=True)
        self.ln()

        cal = calendar.Calendar(firstweekday=6)
        weeks = cal.monthdatescalendar(year, month)

        for week in weeks:
            for date in week:
                x = self.get_x()
                y = self.get_y()
                if date.month == month:
                    name = schedule.get(date, "")
                    r, g, b = get_color(name)
                    self.set_draw_color(0)
                    self.set_fill_color(255, 255, 255)
                    self.rect(x, y, cell_w, cell_h)
                    self.set_xy(x + 2, y + 2)
                    self.set_text_color(0)
                    self.cell(cell_w - 4, 5, str(date.day), ln=1)
                    self.set_text_color(r, g, b)
                    self.set_xy(x + 2, y + 10)
                    self.set_font("Arial", "", 10)
                    self.multi_cell(cell_w - 4, 5, name)
                    self.set_xy(x + cell_w, y)
                else:
                    self.cell(cell_w, cell_h, "", border=1)
            self.ln()
