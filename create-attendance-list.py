import datetime
import sys
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

def load_perslist(file_name, sheetname=None):
    wb = load_workbook(file_name)
    ws = wb[sheetname] if sheetname else wb.active
    persons = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            persons.append((row[0], row[1]))
    return persons

def create_team_sheet(wb, sheetname, persons, start_date, end_date):
    from openpyxl.utils import get_column_letter
    from openpyxl.styles import Alignment, PatternFill
    from openpyxl.formatting.rule import CellIsRule
    import datetime

    ws = wb.create_sheet(title=sheetname)

    center_align = Alignment(horizontal="center", vertical="center")
    rotated_align = Alignment(horizontal="left", vertical="bottom", text_rotation=90)

    saturday_fill = PatternFill(start_color="3399ff", end_color="3399ff", fill_type="solid")        # Saturday / Gameday bright blue
    excused_fill = PatternFill(start_color="ffff00", end_color="ffff00", fill_type="solid")         # "e" excused = kind of yellow
    excusedholiday_fill = PatternFill(start_color="ff8000", end_color="ff8000", fill_type="solid")  # "f" holiday = bright kind of yellow
    unexcused_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")       # "u" unexcused = red
    attendpractice_fill = PatternFill(start_color="6fdc6f", end_color="6fdc6f", fill_type="solid")  # "a" attend / present = green
    gameattend_fill = PatternFill(start_color="99ccff", end_color="99ccff", fill_type="solid")      # "s" Game attend = bright bright blue

    practice_days = [0, 2]  # Monday and Wednesday
    gameday = 5             # Saturday

    # ------------------------------------
    # Headers
    # ------------------------------------
    ws["A1"] = sheetname
    ws["A2"] = "Trainings bis heute:"
    ws["A3"] = "Vorname"
    ws["B3"] = "Nachname"

    # ------------------------------------
    # Dates (Mon, Wed, Sat) from column C
    # ------------------------------------
    termine = []
    col = 3  # Column C
    current = start_date
    while current <= end_date:
        if current.weekday() in practice_days or current.weekday() == gameday:
            zelle = ws.cell(row=3, column=col, value=current)
            zelle.number_format = "DD.MM.YYYY"
            zelle.alignment = rotated_align
            termine.append((col, current))
            if current.weekday() == gameday:
                ws.cell(row=3, column=col).fill = saturday_fill
            col += 1
        current += datetime.timedelta(days=1)

    # practice dates
    practice_dates = [c for c, d in termine if d.weekday() in practice_days]
    game_dates = [c for c, d in termine if d.weekday() == gameday]

    # Count possible practice days until today (for B2)
    ws["B2"] = f'=SUMPRODUCT(--(WEEKDAY(C3:N3,2)=1), --(C3:N3 <= TODAY()))'
    #ws["B2"] = f'=SUMPRODUCT(--(WEEKDAY(C3:N3,2)=&practice_day), --(C3:N3 <= TODAY()))'

    # ------------------------------------
    # Additional columns for quota + games
    # ------------------------------------
    practicesquote_col = col
    game_col = col + 1
    ws.cell(row=3, column=practicesquote_col, value="Präsenz (%)").alignment=rotated_align
    ws.cell(row=3, column=game_col, value="Anz. Spiele").alignment=rotated_align

    # ------------------------------------
    # Participant data from row 4
    # ------------------------------------
    for row_idx, (name, vorname) in enumerate(persons, start=4):
        ws.cell(row=row_idx, column=1, value=name)
        ws.cell(row=row_idx, column=2, value=vorname)

        # Conditional formatting for each date cell
        for col_num, _ in termine:
            cell = ws.cell(row=row_idx, column=col_num).alignment=center_align
            col_letter = get_column_letter(col_num)
            cell_range = f"{col_letter}{row_idx}"

            ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"u"'], fill=unexcused_fill))
            ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"e"'], fill=excused_fill))
            ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"a"'], fill=attendpractice_fill))
            ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"s"'], fill=gameattend_fill))
            ws.conditional_formatting.add(cell_range, CellIsRule(operator='equal', formula=['"f"'], fill=excusedholiday_fill))

        # Training area for this line
        if practice_dates:
            first_col_t = get_column_letter(practice_dates[0])
            last_col_t = get_column_letter(practice_dates[-1])
            practices_range = f"{first_col_t}{row_idx}:{last_col_t}{row_idx}"
            quote_formula = f'=IFERROR(COUNTIF({practices_range}, "a") / $B$2*100, "")'
            ws.cell(row=row_idx, column=practicesquote_col, value=quote_formula)
        else:
            ws.cell(row=row_idx, column=practicesquote_col, value="")

        # Game area for this line
        if game_dates:
            first_col_s = get_column_letter(game_dates[0])
            last_col_s = get_column_letter(game_dates[-1])
            game_range = f"{first_col_s}{row_idx}:{last_col_s}{row_idx}"
            game_formula = f'=COUNTIF({game_range}, "s")'
            ws.cell(row=row_idx, column=game_col, value=game_formula)
        else:
            ws.cell(row=row_idx, column=game_col, value="")

    # Total attendance
    total_start_row = len(persons) +6
    ws.cell(row=total_start_row, column=1, value="TOTAL:")
    num_attendance_formula = f'=COUNTIF(C4:C19,"a")'
    ws.cell(row=total_start_row, column=3, value=num_attendance_formula)

    # ------------------------------------
    # Insert legend below the list
    # ------------------------------------
    legend_start_row = len(persons) + 8
    ws.cell(row=legend_start_row, column=1, value="Legende:")
    ws.cell(row=legend_start_row + 1, column=1, value='a = anwesend',).fill=attendpractice_fill
    ws.cell(row=legend_start_row + 2, column=1, value='e = entschuldigt abwesend').fill=excused_fill
    ws.cell(row=legend_start_row + 3, column=1, value='f = ferien abwesend').fill=excusedholiday_fill
    ws.cell(row=legend_start_row + 4, column=1, value='s = Spielteilnahme').fill=gameattend_fill
    ws.cell(row=legend_start_row + 5, column=1, value='u = unentschuldigt abwesend').fill=unexcused_fill

### main 
# Input Verification - Usage message
def usage():
    print("Usage: python create-attendance-list.py <start date> <end date>")
    print("Example: python create-attendance-list.py 2025-07-01 2025-09-30")
    sys.exit(1)

# Check: Did the user pass exactly two arguments?
if len(sys.argv) != 3:
    print("❌ Error: Exactly two date arguments must be specified.")
    usage()

# Try to parse the dates
try:
    start_date = datetime.strptime(sys.argv[1], "%Y-%m-%d").date()
    end_date = datetime.strptime(sys.argv[2], "%Y-%m-%d").date()
except ValueError:
    print("❌ Error: Please enter the date values in the format YYYY-MM-DD.")
    usage()

# Logic check: Start date must be before or equal to end date
if start_date > end_date:
    print("❌ Error: Start date must not be after the end date.")
    usage()

persons_teamDa = load_perslist("teamDa.xlsx")
persons_teamDb = load_perslist("teamDb.xlsx")

wb = Workbook()
# Remove the default “Sheet”
std = wb.active
wb.remove(std)

create_team_sheet(wb, "Team Da", persons_teamDa, start_date, end_date)
create_team_sheet(wb, "Team Db", persons_teamDb, start_date, end_date)

wb.save("attendance-list.xlsx")
