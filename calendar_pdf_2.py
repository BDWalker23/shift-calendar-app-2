
from fpdf import FPDF
import calendar

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
    def header(self):
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, "Shift Calendar", ln=True, align="C")
        self.ln(5)

    def draw_calendar(self, year, month, schedule):
        self.set_font("Arial", "B", 10)
        self.set_fill_color(220, 220, 220)

        days = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        cell_w = 40
        cell_h = 20

        # Draw weekday headers
        for day in days:
            self.cell(cell_w, 10, day, border=1, align="C", fill=True)
        self.ln()

        cal = calendar.Calendar(firstweekday=0)
        weeks = cal.monthdatescalendar(year, month)

        for week in weeks:
            for date in week:
                x = self.get_x()
                y = self.get_y()
                self.rect(x, y, cell_w, cell_h)
                if date.month == month:
                    name = schedule.get(date, "")
                    r, g, b = get_color(name)
                    self.set_xy(x + 2, y + 2)
                    self.set_font("Arial", "", 9)
                    self.set_text_color(0)
                    self.cell(cell_w - 4, 5, str(date.day), ln=1)
                    if name:
                        self.set_xy(x + 2, y + 8)
                        self.set_text_color(r, g, b)
                        self.multi_cell(cell_w - 4, 5, name)
                else:
                    self.set_xy(x + 2, y + 2)
                    self.set_font("Arial", "", 9)
                    self.set_text_color(150)
                    self.cell(cell_w - 4, 5, str(date.day), ln=1)
                self.set_xy(x + cell_w, y)
            self.ln()
