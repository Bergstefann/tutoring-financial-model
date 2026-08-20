"""Build the tutoring financial model workbook from the anonymised extract.

Reads data/lessons_anonymized.csv (produced by extract.py) and writes
workbook/tutoring-financial-model.xlsx. Every derived figure is an Excel
formula, not a computed-in-Python value -- this script only writes labels,
raw inputs, and formula strings. Formula correctness is verified separately
by scripts/recalc.py (which forces Excel to recalculate and checks for
errors) -- openpyxl does not evaluate formulas itself.
"""

import csv
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
DATA_CSV = ROOT / "data" / "lessons_anonymized.csv"
OUT_PATH = ROOT / "workbook" / "tutoring-financial-model.xlsx"

AS_AT_DATE = "2026-08-20"
WIND_UP_DATE = "2026-09-18"

# ---------------------------------------------------------------- styles --
BLUE = Font(color="0000FF")
BOLD_BLUE = Font(color="0000FF", bold=True)
BLACK = Font(color="000000")
GREEN = Font(color="008000")
BOLD = Font(bold=True)
BOLD_GREEN = Font(color="008000", bold=True)
TITLE_FONT = Font(bold=True, size=14)
SECTION_FONT = Font(bold=True, size=11, color="1F4E78")
NOTE_FONT = Font(italic=True, size=9, color="666666")
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
YELLOW_FILL = PatternFill("solid", fgColor="FFFF00")
GREY_FILL = PatternFill("solid", fgColor="F2F2F2")
PLACEHOLDER_FILL = PatternFill("solid", fgColor="FFD966")

CURRENCY = '$#,##0;($#,##0)'
CURRENCY2 = '$#,##0.00;($#,##0.00)'
PCT1 = '0.0%'
NUM1 = '0.0'
NUM2 = '0.00'
INT_FMT = '0'
DATE_FMT = 'yyyy-mm-dd'

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header_row(ws, row, first_col, last_col):
    for col in range(first_col, last_col + 1):
        c = ws.cell(row=row, column=col)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def sc(ws, addr, value, font=None, fill=None, number_format=None, bold=False,
       align=None, border=False, wrap=False):
    c = ws[addr]
    c.value = value
    if font is not None:
        c.font = font
    elif bold:
        c.font = BOLD
    if fill is not None:
        c.fill = fill
    if number_format is not None:
        c.number_format = number_format
    if align is not None:
        c.alignment = align
    elif wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    if border:
        c.border = BOX
    return c


# ------------------------------------------------------------- load data --
def load_lessons():
    with open(DATA_CSV, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    for r in rows:
        r["duration_minutes"] = int(r["duration_minutes"])
        r["rate"] = float(r["rate"])
    return rows


# ---------------------------------------------------------------- sheets --
def build_readme(wb):
    ws = wb.active
    ws.title = "README"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 90

    sc(ws, "A1", "Tutoring Business Financial Model", font=TITLE_FONT)
    sc(ws, "A3", "As at:", bold=True)
    sc(ws, "B3", AS_AT_DATE)
    sc(ws, "A4", "Business wound up:", bold=True)
    sc(ws, "B4", f"{WIND_UP_DATE} (end of school term)")
    sc(ws, "A5", "Student identities:", bold=True)
    sc(ws, "B5",
       "Pseudonymised. Every student appears as \"Student NN\"; no real name, "
       "parent name, email address, or address appears anywhere in this "
       "workbook. The mapping from pseudonym to real identity is kept in a "
       "file that is not part of this repository.", wrap=True)

    r = 7
    sc(ws, f"A{r}", "What this model does", font=SECTION_FONT)
    r += 1
    sc(ws, f"A{r}",
       "A driver-based financial model of a private tutoring business, built "
       "from actual lesson-schedule and billing data for one school term. It "
       "derives revenue, cost-to-serve, capacity utilisation, and a "
       "provisioned P&L from that data, then explores scenarios and rate/"
       "utilisation sensitivity.", wrap=True)
    ws.merge_cells(f"A{r}:B{r}")
    ws.row_dimensions[r].height = 45
    r += 2

    sc(ws, f"A{r}", "Colour key", font=SECTION_FONT)
    r += 1
    rows = [
        ("Blue text", "Hardcoded input — a number typed in, not calculated"),
        ("Black text", "Formula — calculated from other cells"),
        ("Green text", "Cross-sheet link — formula that pulls from another sheet"),
        ("Yellow fill", "Key assumption — worth checking before trusting the output"),
        ("Orange fill", "Placeholder — not supplied; a stand-in value is used so the "
                         "model runs, but it needs confirming"),
    ]
    for label, desc in rows:
        sc(ws, f"A{r}", label, font=BOLD)
        sc(ws, f"B{r}", desc, wrap=True)
        r += 1
    r += 1

    sc(ws, f"A{r}", "Sheet order", font=SECTION_FONT)
    r += 1
    sheet_list = [
        "Assumptions", "Data", "Revenue by student", "Cost to serve",
        "Capacity", "P&L and provisioning", "Scenarios", "Sensitivity", "Charts",
    ]
    for name in sheet_list:
        sc(ws, f"A{r}", f"• {name}")
        r += 1
    r += 1

    sc(ws, f"A{r}", "What isn't in this model", font=SECTION_FONT)
    r += 1
    sc(ws, f"A{r}",
       "No suburb or address field exists anywhere in the source data, so "
       "travel cost/time is modelled as a flat weekly figure allocated "
       "across students by lesson share, not as a per-suburb distance "
       "calculation. There is no suburb-level chart for the same reason. "
       "The maximum teachable hours/week figure on the Assumptions sheet is "
       "a placeholder (orange fill) pending confirmation.", wrap=True)
    ws.merge_cells(f"A{r}:B{r}")
    ws.row_dimensions[r].height = 60

    return ws


ASSUMPTIONS_SPEC = [
    ("section", "Rates & Lesson Structure"),
    ("row", "rate_per_lesson", "Rate per lesson", 40.0, "$/lesson", False, CURRENCY,
     "From data: flat $40 per 30-minute lesson across all 37 students "
     "(Portfolio Lesson Schedule extract, Term 1 2026).", False),
    ("row", "duration_min", "Standard lesson duration", 30, "min", False, INT_FMT,
     "Confirmed by business owner, 2026-08-20 — not present in source data.", False),
    ("row", "hourly_rate", "Implied hourly rate", "={rate_per_lesson}/({duration_min}/60)",
     "$/hr", True, CURRENCY,
     "Derived: rate per lesson ÷ (duration ÷ 60).", False),

    ("section", "Statutory & Provisioning"),
    ("row", "super_rate", "Superannuation provision rate", 0.125, "%", False, PCT1,
     "Statutory (ATO) rate, FY2025-26.", True),
    ("row", "tax_rate", "Income tax provisioning rate", 0.20, "%", False, PCT1,
     "ESTIMATE for provisioning purposes only. Tutoring is being wound up "
     "mid financial year; the real liability depends on total income across "
     "the whole year, which isn't modelled here. Not a derived or lodged "
     "figure.", True),

    ("section", "Working Pattern"),
    ("row", "terms_per_year", "Terms per year", 4, "terms", False, INT_FMT,
     "Owner estimate.", False),
    ("row", "weeks_per_term", "Weeks per term", 9, "weeks", False, INT_FMT,
     "Owner estimate — matches the 9 weekly blocks in the extracted schedule.", False),
    ("row", "working_weeks_per_year", "Working weeks per year",
     "={terms_per_year}*{weeks_per_term}", "weeks", True, INT_FMT,
     "Derived: terms per year × weeks per term.", False),
    ("row", "working_days_per_week", "Working days per week", 5, "days", False, INT_FMT,
     "From data: 5 populated weekdays per week in the extract.", False),
    ("row", "max_lessons_per_week", "Max lessons per week (scheduling ceiling)", 37,
     "lessons", False, INT_FMT,
     "Owner confirmed, 2026-08-20 — equals current scheduled lesson volume. "
     "The business is already scheduled at its capacity ceiling; the gap to "
     "lessons actually delivered is show-up rate, not spare capacity.", True),
    ("row", "max_teachable_hours", "Max teachable hours per week",
     "={max_lessons_per_week}*{duration_min}/60", "hrs", True, NUM1,
     "Derived: max lessons per week × duration ÷ 60. This is a ceiling on "
     "lesson-delivery time; travel time is separate, additive time spent "
     "outside it (see Capacity sheet), not subtracted from it.", False),

    ("section", "Travel & Vehicle"),
    ("row", "travel_time_per_day", "Travel time per day", 1.5, "hrs", False, NUM1,
     "Owner estimate — flat figure. No per-suburb location data exists "
     "in the source, so travel can't be modelled by distance.", False),
    ("row", "travel_time_per_week", "Travel time per week",
     "={travel_time_per_day}*{working_days_per_week}", "hrs", True, NUM1,
     "Derived: travel time per day × working days per week.", False),
    ("row", "insurance_per_month", "Vehicle insurance", 150.0, "$/mo", False, CURRENCY,
     "Owner estimate.", False),
    ("row", "fuel_price", "Fuel price", 1.80, "$/L", False, CURRENCY2,
     "Owner estimate (\"generally\").", False),
    ("row", "tank_size", "Fuel tank size", 40.0, "L", False, NUM1,
     "Owner estimate.", False),
    ("row", "refill_weeks", "Tank refill frequency", 1.5, "weeks", False, NUM1,
     "Owner estimate.", False),
    ("row", "insurance_per_week", "Insurance cost per working week",
     "={insurance_per_month}*12/{working_weeks_per_year}", "$/wk", True, CURRENCY,
     "Derived: monthly insurance × 12 ÷ working weeks per year.", False),
    ("row", "fuel_per_week", "Fuel cost per week",
     "={tank_size}*{fuel_price}/{refill_weeks}", "$/wk", True, CURRENCY,
     "Derived: tank size × fuel price ÷ refill frequency.", False),
    ("row", "vehicle_cost_per_week", "Vehicle running cost per week",
     "={insurance_per_week}+{fuel_per_week}", "$/wk", True, CURRENCY,
     "Derived: insurance/week + fuel/week. Modelled flat (not $/km) because "
     "no distance data exists to drive a per-km figure.", True),

    ("section", "Scenario Parameters"),
    ("row", "scenario_switch", "Scenario switch", 1, "1/2/3", False, INT_FMT,
     "1 = Base (actual), 2 = More students at current rate, 3 = Higher rate, "
     "fewer students. Drives the \"Selected scenario\" callout on the "
     "Scenarios sheet.", True),
    ("row", "scenario2_volume_mult", "Scenario 2: lesson volume multiplier", 1.30, "×",
     False, NUM2, "Illustrative — adjust as needed. 30% more lesson volume "
     "at the current rate.", False),
    ("row", "scenario3_rate_mult", "Scenario 3: rate multiplier", 1.25, "×", False, NUM2,
     "Illustrative — adjust as needed. A 25% rate rise.", False),
    ("row", "scenario3_volume_mult", "Scenario 3: lesson volume multiplier", 0.85, "×",
     False, NUM2, "Illustrative — adjust as needed. 15% fewer lessons at "
     "the higher rate.", False),
]


def build_assumptions(wb):
    ws = wb.create_sheet("Assumptions")
    ws.sheet_view.showGridLines = False
    widths = {"A": 40, "B": 14, "C": 10, "D": 62}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    sc(ws, "A1", "Assumptions", font=TITLE_FONT)
    sc(ws, "A2", "Every input the model uses lives here, in its own labelled "
                 "cell. Formulas elsewhere reference these cells; nothing "
                 "downstream hardcodes a number.", font=NOTE_FONT, wrap=True)

    row = 4
    sc(ws, f"A{row}", "Label", bold=True)
    sc(ws, f"B{row}", "Value", bold=True)
    sc(ws, f"C{row}", "Units", bold=True)
    sc(ws, f"D{row}", "Source / basis", bold=True)
    style_header_row(ws, row, 1, 4)
    row += 1

    refs = {}
    for item in ASSUMPTIONS_SPEC:
        if item[0] == "section":
            row += 1
            sc(ws, f"A{row}", item[1], font=SECTION_FONT, fill=GREY_FILL)
            ws.merge_cells(f"A{row}:D{row}")
            row += 1
            continue

        _, key, label, value, units, is_formula, number_format, source, flag = item
        sc(ws, f"A{row}", label, border=True)
        value_addr = f"B{row}"
        if is_formula:
            formula = value
            for ref_key, ref_addr in refs.items():
                formula = formula.replace("{" + ref_key + "}", ref_addr)
            sc(ws, value_addr, formula, font=BLACK, number_format=number_format, border=True)
        else:
            fill = PLACEHOLDER_FILL if flag == "placeholder" else (
                YELLOW_FILL if flag else None)
            sc(ws, value_addr, value, font=BOLD_BLUE, number_format=number_format,
               fill=fill, border=True)
        sc(ws, f"C{row}", units, font=NOTE_FONT, border=True)
        sc(ws, f"D{row}", source, font=NOTE_FONT, wrap=True, border=True)
        refs[key] = f"$B${row}"
        row += 1

    ws.freeze_panes = "A5"
    return ws, refs


def build_data(wb, lessons):
    ws = wb.create_sheet("Data")
    ws.sheet_view.showGridLines = False
    widths = {"A": 14, "B": 14, "C": 12, "D": 16, "E": 12}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    sc(ws, "A1", "Data", font=TITLE_FONT)
    ws.merge_cells("A1:E1")
    sc(ws, "A2", "Anonymised lesson schedule, one row per scheduled lesson "
                 "slot. Extracted from the Portfolio Lesson Schedule sheet "
                 "(Term 1 2026, 9 weeks); not hand-edited. Source: "
                 "scripts/extract.py.", font=NOTE_FONT, wrap=True)
    ws.merge_cells("A2:E2")

    header_row = 4
    headers = ["Student", "Lesson Date", "Attended", "Duration (min)", "Rate ($)"]
    for i, h in enumerate(headers, start=1):
        sc(ws, f"{get_column_letter(i)}{header_row}", h)
    style_header_row(ws, header_row, 1, 5)

    first_data_row = header_row + 1
    for i, rec in enumerate(lessons):
        r = first_data_row + i
        sc(ws, f"A{r}", rec["student_pseudonym"], font=BLUE)
        sc(ws, f"B{r}", rec["lesson_date"], font=BLUE, number_format=DATE_FMT)
        sc(ws, f"C{r}", rec["attended"], font=BLUE, align=Alignment(horizontal="center"))
        sc(ws, f"D{r}", rec["duration_minutes"], font=BLUE, number_format=INT_FMT)
        sc(ws, f"E{r}", rec["rate"], font=BLUE, number_format=CURRENCY)
    last_data_row = first_data_row + len(lessons) - 1

    ws.freeze_panes = f"A{first_data_row}"
    return ws, first_data_row, last_data_row


def build_revenue_by_student(wb, students, data_range):
    ws = wb.create_sheet("Revenue by student")
    ws.sheet_view.showGridLines = False
    widths = {"A": 14, "B": 14, "C": 14, "D": 14, "E": 14, "F": 16, "G": 14, "H": 12, "I": 8}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    sc(ws, "A1", "Revenue by student", font=TITLE_FONT)
    ws.merge_cells("A1:I1")

    header_row = 3
    headers = ["Student", "Lessons scheduled", "Lessons delivered", "Hours delivered",
               "Gross revenue ($)", "Effective hourly rate ($/hr)", "Share of revenue",
               "Tie-break (rank helper)", "Rank"]
    for i, h in enumerate(headers, start=1):
        sc(ws, f"{get_column_letter(i)}{header_row}", h)
    style_header_row(ws, header_row, 1, 9)

    first_row = header_row + 1
    last_row = first_row + len(students) - 1
    total_row = last_row + 1

    dA, dC, dD, dE = (f"Data!${c}${data_range[0]}:${c}${data_range[1]}" for c in "ACDE")

    for i, student in enumerate(students):
        r = first_row + i
        sc(ws, f"A{r}", student, font=BLUE, border=True)
        sc(ws, f"B{r}", f'=COUNTIFS({dA},A{r})', font=BLACK, number_format=INT_FMT, border=True)
        sc(ws, f"C{r}", f'=COUNTIFS({dA},A{r},{dC},"Y")', font=BLACK, number_format=INT_FMT, border=True)
        sc(ws, f"D{r}", f'=SUMIFS({dD},{dA},A{r},{dC},"Y")/60', font=BLACK, number_format=NUM2, border=True)
        sc(ws, f"E{r}", f'=SUMIFS({dE},{dA},A{r},{dC},"Y")', font=BLACK, number_format=CURRENCY, border=True)
        sc(ws, f"F{r}", f'=IFERROR(E{r}/D{r},0)', font=BLACK, number_format=CURRENCY, border=True)
        sc(ws, f"G{r}", f'=IFERROR(E{r}/$E${total_row},0)', font=BLACK, number_format=PCT1, border=True)
        sc(ws, f"H{r}", f'=E{r}+(ROW()-{first_row})/100000', font=NOTE_FONT, number_format=NUM2, border=True)
        sc(ws, f"I{r}", f'=RANK(H{r},$H${first_row}:$H${last_row},0)', font=BLACK, number_format=INT_FMT, border=True)

    sc(ws, f"A{total_row}", "TOTAL", bold=True, border=True)
    for col in "BCDE":
        sc(ws, f"{col}{total_row}", f'=SUM({col}{first_row}:{col}{last_row})',
           font=BOLD, number_format=CURRENCY if col == "E" else (NUM2 if col == "D" else INT_FMT), border=True)
    sc(ws, f"F{total_row}", f'=IFERROR(E{total_row}/D{total_row},0)', font=BOLD, number_format=CURRENCY, border=True)
    sc(ws, f"G{total_row}", f'=SUM(G{first_row}:G{last_row})', font=BOLD, number_format=PCT1, border=True)

    ws.freeze_panes = f"A{first_row}"
    refs = {
        "first_row": first_row, "last_row": last_row, "total_row": total_row,
        "lessons_scheduled_total": f"'Revenue by student'!$B${total_row}",
        "lessons_delivered_total": f"'Revenue by student'!$C${total_row}",
        "hours_delivered_total": f"'Revenue by student'!$D${total_row}",
        "gross_revenue_total": f"'Revenue by student'!$E${total_row}",
    }
    return ws, refs


def build_cost_to_serve(wb, students, rbs_refs, a_refs, date_min, date_max):
    ws = wb.create_sheet("Cost to serve")
    ws.sheet_view.showGridLines = False
    widths = {"A": 14, "B": 16, "C": 16, "D": 18, "E": 18, "F": 14, "G": 14, "H": 14, "I": 20, "J": 12, "K": 8}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    sc(ws, "A1", "Cost to serve", font=TITLE_FONT)
    ws.merge_cells("A1:K1")
    sc(ws, "A2", f"Period covered: {date_min} to {date_max} (9 weeks, one term). "
                 "No suburb/location field exists anywhere in the source data, so "
                 "travel cost and time are modelled as a flat weekly total "
                 "(Assumptions sheet) allocated across students by their share "
                 "of lessons delivered — not by per-suburb distance.",
       font=NOTE_FONT, wrap=True)
    ws.merge_cells("A2:K2")
    ws.row_dimensions[2].height = 30

    sc(ws, "A4", "Total vehicle running cost (period)", bold=True)
    sc(ws, "B4", f"=Assumptions!{a_refs['vehicle_cost_per_week']}*Assumptions!{a_refs['weeks_per_term']}",
       font=GREEN, number_format=CURRENCY)
    sc(ws, "A5", "Total travel time (period, hours)", bold=True)
    sc(ws, "B5", f"=Assumptions!{a_refs['travel_time_per_week']}*Assumptions!{a_refs['weeks_per_term']}",
       font=GREEN, number_format=NUM2)

    header_row = 7
    headers = ["Student", "Lessons delivered", "Share of lessons delivered",
               "Allocated travel cost ($)", "Allocated travel time (hrs)",
               "Gross revenue ($)", "Contribution ($)", "Committed hours",
               "Contribution per hour committed ($/hr)", "Tie-break", "Rank"]
    for i, h in enumerate(headers, start=1):
        sc(ws, f"{get_column_letter(i)}{header_row}", h)
    style_header_row(ws, header_row, 1, 11)

    first_row = header_row + 1
    last_row = first_row + len(students) - 1
    total_row = last_row + 1
    rbs_first, rbs_last = rbs_refs["first_row"], rbs_refs["last_row"]

    for i, student in enumerate(students):
        r = first_row + i
        rbs_r = rbs_first + i
        sc(ws, f"A{r}", f"='Revenue by student'!A{rbs_r}", font=GREEN, border=True)
        sc(ws, f"B{r}", f"='Revenue by student'!C{rbs_r}", font=GREEN, number_format=INT_FMT, border=True)
        sc(ws, f"C{r}", f'=IFERROR(B{r}/{rbs_refs["lessons_delivered_total"]},0)', font=GREEN, number_format=PCT1, border=True)
        sc(ws, f"D{r}", f'=C{r}*$B$4', font=BLACK, number_format=CURRENCY, border=True)
        sc(ws, f"E{r}", f'=C{r}*$B$5', font=BLACK, number_format=NUM2, border=True)
        sc(ws, f"F{r}", f"='Revenue by student'!E{rbs_r}", font=GREEN, number_format=CURRENCY, border=True)
        sc(ws, f"G{r}", f'=F{r}-D{r}', font=BLACK, number_format=CURRENCY, border=True)
        sc(ws, f"H{r}", f"='Revenue by student'!D{rbs_r}+E{r}", font=GREEN, number_format=NUM2, border=True)
        sc(ws, f"I{r}", f'=IFERROR(G{r}/H{r},0)', font=BLACK, number_format=CURRENCY2, border=True)
        sc(ws, f"J{r}", f'=G{r}+(ROW()-{first_row})/100000', font=NOTE_FONT, number_format=NUM2, border=True)
        sc(ws, f"K{r}", f'=RANK(J{r},$J${first_row}:$J${last_row},0)', font=BLACK, number_format=INT_FMT, border=True)

    sc(ws, f"A{total_row}", "TOTAL", bold=True, border=True)
    for col in "BDEFGH":
        fmt = CURRENCY if col in "DFG" else NUM2
        sc(ws, f"{col}{total_row}", f'=SUM({col}{first_row}:{col}{last_row})', font=BOLD, number_format=fmt, border=True)
    sc(ws, f"I{total_row}", f'=IFERROR(G{total_row}/H{total_row},0)', font=BOLD, number_format=CURRENCY2, border=True)

    ws.freeze_panes = f"A{first_row}"
    refs = {
        "first_row": first_row, "last_row": last_row, "total_row": total_row,
        "travel_cost_total": f"'Cost to serve'!$D${total_row}",
        "contribution_total": f"'Cost to serve'!$G${total_row}",
    }
    return ws, refs


def build_capacity(wb, a_refs, rbs_refs, cts_refs):
    ws = wb.create_sheet("Capacity")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 60

    sc(ws, "A1", "Capacity", font=TITLE_FONT)
    ws.merge_cells("A1:C1")
    sc(ws, "A2", "Max lessons/week is a hard ceiling on lesson-delivery time "
                 "only. Travel time is separate, additive time spent outside "
                 "it (owner already teaches at the scheduling ceiling), so it "
                 "is not subtracted here — see Cost to serve for how travel "
                 "time is treated as unpaid capacity consumed per student.",
       font=NOTE_FONT, wrap=True)
    ws.merge_cells("A2:C2")
    ws.row_dimensions[2].height = 30

    A = lambda k: f"Assumptions!{a_refs[k]}"
    RBS_TOT = rbs_refs["gross_revenue_total"]

    rows = [
        ("max_lessons", "Max lessons per week (scheduling ceiling)", f"={A('max_lessons_per_week')}", INT_FMT, GREEN, ""),
        ("max_hours", "Max teachable hours per week", f"={A('max_teachable_hours')}", NUM1, GREEN, ""),
        ("travel_hours", "Travel time per week (hrs, additive)", f"={A('travel_time_per_week')}", NUM1, GREEN, "Not subtracted from teaching ceiling — see note above."),
        ("total_commitment", "Total weekly time commitment at full capacity", "=B5+B6", NUM1, BLACK, "Teaching ceiling + travel time — the real total workload if fully booked."),
        (None, None, None, None, None, None),
        ("actual_hours", "Actual hours delivered/week (avg)", f"='Revenue by student'!$D${rbs_refs['total_row']}/{A('weeks_per_term')}", NUM2, GREEN, "Averaged over the 9-week extract."),
        ("actual_lessons", "Actual lessons delivered/week (avg)", f"='Revenue by student'!$C${rbs_refs['total_row']}/{A('weeks_per_term')}", NUM2, GREEN, ""),
        ("utilisation", "Utilisation (actual hrs ÷ max teachable hrs)", "=IFERROR(B9/B5,0)", PCT1, BLACK, "How close to the scheduling ceiling the business ran."),
        (None, None, None, None, None, None),
        ("scheduled_lessons", "Scheduled lessons/week (avg)", f"='Revenue by student'!$B${rbs_refs['total_row']}/{A('weeks_per_term')}", NUM2, GREEN, "Already equal to the max — see note below."),
        ("showup_rate", "Show-up rate (delivered ÷ scheduled)", "=IFERROR(B10/B13,0)", PCT1, BLACK, "Scheduled already sits at the max lessons/week ceiling, so this equals utilisation above: the constraint here is no-shows/cancellations, not unfilled slots."),
        (None, None, None, None, None, None),
        ("ceiling_revenue", "Ceiling annual revenue (if every slot delivered)", f"=B5*{A('working_weeks_per_year')}*{A('hourly_rate')}", CURRENCY, GREEN, "Max teachable hrs × working weeks/yr × hourly rate."),
        ("actual_pace", "Actual annualised revenue pace", f"={RBS_TOT}/{A('weeks_per_term')}*{A('working_weeks_per_year')}", CURRENCY, GREEN, "Extrapolated run-rate, NOT an actual full-year figure — the business wound up ~18 Sept 2026."),
    ]

    r = 4
    refs = {}
    for key, label, formula, fmt, font, note in rows:
        if key is None:
            r += 1
            continue
        sc(ws, f"A{r}", label, border=True)
        sc(ws, f"B{r}", formula, font=font, number_format=fmt, border=True)
        sc(ws, f"C{r}", note, font=NOTE_FONT, wrap=True, border=True)
        refs[key] = f"$B${r}"
        r += 1

    return ws, refs


def build_pl(wb, a_refs, rbs_refs, cts_refs):
    ws = wb.create_sheet("P&L and provisioning")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 55

    sc(ws, "A1", "P&L and provisioning", font=TITLE_FONT)
    ws.merge_cells("A1:C1")

    A = lambda k: f"Assumptions!{a_refs[k]}"

    sc(ws, "A3", "Period actuals (9 weeks, from the extract)", font=SECTION_FONT, fill=GREY_FILL)
    ws.merge_cells("A3:C3")
    labels = [
        ("Gross revenue", f"='Revenue by student'!${'E'}${rbs_refs['total_row']}", GREEN, ""),
        ("Less: allocated travel cost", f"=-'Cost to serve'!${'D'}${cts_refs['total_row']}", GREEN, ""),
        ("Net operating income", "=B4+B5", BLACK, ""),
        ("Super provision", f"=B6*{A('super_rate')}", GREEN, ""),
        ("Taxable income after super", "=B6-B7", BLACK, ""),
        ("Tax provision", f"=B8*{A('tax_rate')}", GREEN, "ESTIMATE only — see Assumptions sheet note."),
        ("Net income (period)", "=B6-B7-B9", BLACK, ""),
    ]
    r = 4
    for label, formula, font, note in labels:
        sc(ws, f"A{r}", label, border=True, bold=(label in ("Net operating income", "Net income (period)")))
        sc(ws, f"B{r}", formula, font=font, number_format=CURRENCY, border=True,
           bold=(label in ("Net operating income", "Net income (period)")))
        sc(ws, f"C{r}", note, font=NOTE_FONT, wrap=True, border=True)
        r += 1

    r += 1
    sc(ws, f"A{r}", "Annualised run-rate (extrapolated — NOT an actual full-year "
                     "figure; the business wound up ~18 Sept 2026)", font=SECTION_FONT, fill=GREY_FILL)
    ws.merge_cells(f"A{r}:C{r}")
    r += 1
    sc(ws, f"A{r}", "Annualisation factor (working weeks/yr ÷ weeks/term)")
    sc(ws, f"B{r}", f"={A('working_weeks_per_year')}/{A('weeks_per_term')}", font=GREEN, number_format=NUM2)
    factor_row = r
    r += 1
    ann_start = r
    for label in ["Gross revenue", "Less: allocated travel cost", "Net operating income",
                  "Super provision", "Taxable income after super", "Tax provision",
                  "Net income (annualised)"]:
        period_row = 4 + (r - ann_start)
        sc(ws, f"A{r}", label, border=True, bold=(label in ("Net operating income", "Net income (annualised)")))
        sc(ws, f"B{r}", f"=B{period_row}*$B${factor_row}", font=BLACK, number_format=CURRENCY, border=True,
           bold=(label in ("Net operating income", "Net income (annualised)")))
        r += 1

    return ws


def build_scenarios(wb, a_refs, capacity_refs):
    ws = wb.create_sheet("Scenarios")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    for col in "BCD":
        ws.column_dimensions[col].width = 20

    sc(ws, "A1", "Scenarios", font=TITLE_FONT)
    ws.merge_cells("A1:D1")
    sc(ws, "A2", "All three columns are live formulas driven by the "
                 "multipliers on the Assumptions sheet — not separate "
                 "copied models. Vehicle/travel cost is held constant across "
                 "scenarios: there's no distance-based driver in the data to "
                 "scale it with lesson volume.", font=NOTE_FONT, wrap=True)
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 30

    header_row = 4
    sc(ws, f"A{header_row}", "")
    sc(ws, f"B{header_row}", "1. Base (actual)")
    sc(ws, f"C{header_row}", "2. More students, current rate")
    sc(ws, f"D{header_row}", "3. Higher rate, fewer students")
    style_header_row(ws, header_row, 1, 4)

    A = lambda k: f"Assumptions!{a_refs[k]}"
    r = header_row + 1
    lessons_row = r
    sc(ws, f"A{r}", "Lessons/week (avg)", border=True)
    sc(ws, f"B{r}", f"=Capacity!{capacity_refs['actual_lessons']}", font=GREEN, number_format=NUM2, border=True)
    sc(ws, f"C{r}", f"=B{r}*{A('scenario2_volume_mult')}", font=GREEN, number_format=NUM2, border=True)
    sc(ws, f"D{r}", f"=B{r}*{A('scenario3_volume_mult')}", font=GREEN, number_format=NUM2, border=True)

    r += 1
    rate_row = r
    sc(ws, f"A{r}", "Rate per lesson ($)", border=True)
    sc(ws, f"B{r}", f"={A('rate_per_lesson')}", font=GREEN, number_format=CURRENCY, border=True)
    sc(ws, f"C{r}", f"={A('rate_per_lesson')}", font=GREEN, number_format=CURRENCY, border=True)
    sc(ws, f"D{r}", f"={A('rate_per_lesson')}*{A('scenario3_rate_mult')}", font=GREEN, number_format=CURRENCY, border=True)

    r += 1
    rev_row = r
    sc(ws, f"A{r}", "Annualised revenue ($)", border=True, bold=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={col}{lessons_row}*{A('working_weeks_per_year')}*{col}{rate_row}",
           font=BLACK, number_format=CURRENCY, border=True, bold=True)

    r += 1
    cost_row = r
    sc(ws, f"A{r}", "Annualised vehicle/travel cost ($)", border=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={A('vehicle_cost_per_week')}*{A('working_weeks_per_year')}",
           font=GREEN, number_format=CURRENCY, border=True)

    r += 1
    noi_row = r
    sc(ws, f"A{r}", "Net operating income ($)", border=True, bold=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={col}{rev_row}-{col}{cost_row}", font=BLACK, number_format=CURRENCY, border=True, bold=True)

    r += 1
    super_row = r
    sc(ws, f"A{r}", "Super provision ($)", border=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={col}{noi_row}*{A('super_rate')}", font=GREEN, number_format=CURRENCY, border=True)

    r += 1
    taxable_row = r
    sc(ws, f"A{r}", "Taxable income after super ($)", border=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={col}{noi_row}-{col}{super_row}", font=BLACK, number_format=CURRENCY, border=True)

    r += 1
    tax_row = r
    sc(ws, f"A{r}", "Tax provision ($)", border=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={col}{taxable_row}*{A('tax_rate')}", font=GREEN, number_format=CURRENCY, border=True)

    r += 1
    net_row = r
    sc(ws, f"A{r}", "Net income (annualised, $)", border=True, bold=True)
    for col in "BCD":
        sc(ws, f"{col}{r}", f"={col}{noi_row}-{col}{super_row}-{col}{tax_row}",
           font=BLACK, number_format=CURRENCY, border=True, bold=True)

    r += 2
    sc(ws, f"A{r}", "Selected scenario (Assumptions!scenario_switch)", bold=True)
    sc(ws, f"B{r}", f'=CHOOSE({A("scenario_switch")},B{header_row},C{header_row},D{header_row})',
       font=GREEN, border=True)
    r += 1
    sc(ws, f"A{r}", "Selected scenario net income ($)", bold=True)
    sc(ws, f"B{r}", f'=CHOOSE({A("scenario_switch")},B{net_row},C{net_row},D{net_row})',
       font=GREEN, number_format=CURRENCY, border=True, bold=True)

    return ws


def build_sensitivity(wb, a_refs, pl_annual_revenue_row, pl_annual_cost_row):
    ws = wb.create_sheet("Sensitivity")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 16
    for col in "CDEFG":
        ws.column_dimensions[col].width = 14

    sc(ws, "A1", "Sensitivity: net income by rate × utilisation", font=TITLE_FONT)
    ws.merge_cells("A1:G1")
    sc(ws, "A2", "Explicit formula grid (not a native Excel Data Table). "
                 "Each cell = (annualised base revenue × row volume "
                 "multiplier × column rate multiplier − annualised "
                 "vehicle/travel cost) × (1 − super rate) × "
                 "(1 − tax rate). Cost is held flat, same simplification "
                 "as Scenarios.", font=NOTE_FONT, wrap=True)
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 30

    sc(ws, "A4", "Base annualised revenue ($)", bold=True)
    sc(ws, "B4", f"='P&L and provisioning'!$B${pl_annual_revenue_row}", font=GREEN, number_format=CURRENCY)
    sc(ws, "A5", "Base annualised vehicle/travel cost ($)", bold=True)
    sc(ws, "B5", f"=ABS('P&L and provisioning'!$B${pl_annual_cost_row})", font=GREEN, number_format=CURRENCY)

    A = lambda k: f"Assumptions!{a_refs[k]}"
    row_mults = [0.8, 0.9, 1.0, 1.1, 1.2]
    col_mults = [0.9, 0.95, 1.0, 1.05, 1.1]

    grid_header_row = 7
    sc(ws, f"A{grid_header_row}", "Volume ↓  /  Rate →", bold=True, border=True)
    for j, cm in enumerate(col_mults):
        col = get_column_letter(3 + j)
        sc(ws, f"{col}{grid_header_row}", cm, font=BOLD_BLUE, number_format=PCT1, border=True,
           align=Alignment(horizontal="center"))
    style_header_row(ws, grid_header_row, 1, 2)

    for i, rm in enumerate(row_mults):
        r = grid_header_row + 1 + i
        sc(ws, f"B{r}", rm, font=BOLD_BLUE, number_format=PCT1, border=True)
        for j, cm in enumerate(col_mults):
            col = get_column_letter(3 + j)
            col_header_addr = f"{col}${grid_header_row}"
            formula = (
                f"=(($B$4*$B{r}*{col_header_addr})-$B$5)"
                f"*(1-{A('super_rate')})*(1-{A('tax_rate')})"
            )
            sc(ws, f"{col}{r}", formula, font=GREEN, number_format=CURRENCY, border=True)
    ws.merge_cells(f"A{grid_header_row+1}:A{grid_header_row+len(row_mults)}")
    sc(ws, f"A{grid_header_row+1}", "Lesson volume\nmultiplier (row)\nRate multiplier\n(column) →",
       font=NOTE_FONT, wrap=True, align=Alignment(wrap_text=True, vertical="center"))

    return ws


def add_bar_chart(ws, anchor, title, cats_ref, data_ref, y_title, x_title=""):
    chart = BarChart()
    chart.type = "col"
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = x_title
    chart.width = 22
    chart.height = 10
    chart.add_data(data_ref, titles_from_data=False)
    chart.set_categories(cats_ref)
    chart.legend = None
    chart.y_axis.scaling.min = 0
    ws.add_chart(chart, anchor)


def build_charts(wb, rbs_ws, rbs_refs, cts_ws, cts_refs, capacity_ws, capacity_refs):
    ws = wb.create_sheet("Charts")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 16

    sc(ws, "A1", "Charts", font=TITLE_FONT)
    sc(ws, "A2", "Sorted helper tables backing each chart are below/beside "
                 "it — built with INDEX/MATCH against a rank column on "
                 "the source sheet, not a native sort.", font=NOTE_FONT, wrap=True)

    # --- Chart 1: revenue concentration by student (sorted desc) ---
    sc(ws, "A4", "Revenue concentration by student (sorted, high→low)", font=SECTION_FONT)
    sc(ws, "A5", "Student", bold=True)
    sc(ws, "B5", "Gross revenue ($)", bold=True)
    n = rbs_refs["last_row"] - rbs_refs["first_row"] + 1
    rbs_first, rbs_last = rbs_refs["first_row"], rbs_refs["last_row"]
    for i in range(n):
        r = 6 + i
        rank_target = i + 1
        sc(ws, f"A{r}",
           f"=INDEX('Revenue by student'!$A${rbs_first}:$A${rbs_last},"
           f"MATCH({rank_target},'Revenue by student'!$I${rbs_first}:$I${rbs_last},0))",
           font=GREEN, number_format=INT_FMT)
        sc(ws, f"B{r}",
           f"=INDEX('Revenue by student'!$E${rbs_first}:$E${rbs_last},"
           f"MATCH({rank_target},'Revenue by student'!$I${rbs_first}:$I${rbs_last},0))",
           font=GREEN, number_format=CURRENCY)
    last_sorted_row = 6 + n - 1
    add_bar_chart(
        ws, "D4", "Revenue concentration by student",
        Reference(ws, min_col=1, min_row=6, max_row=last_sorted_row),
        Reference(ws, min_col=2, min_row=6, max_row=last_sorted_row),
        "Gross revenue ($)",
    )

    # --- Chart 2: contribution per hour committed, by student (sorted desc) ---
    base2 = last_sorted_row + 3
    sc(ws, f"A{base2}", "Contribution per hour committed, by student (sorted, high→low)", font=SECTION_FONT)
    sc(ws, f"A{base2+1}", "Student", bold=True)
    sc(ws, f"B{base2+1}", "Contribution/hr ($)", bold=True)
    cts_first, cts_last = cts_refs["first_row"], cts_refs["last_row"]
    m = cts_last - cts_first + 1
    for i in range(m):
        r = base2 + 2 + i
        rank_target = i + 1
        sc(ws, f"A{r}",
           f"=INDEX('Cost to serve'!$A${cts_first}:$A${cts_last},"
           f"MATCH({rank_target},'Cost to serve'!$K${cts_first}:$K${cts_last},0))",
           font=GREEN, number_format=INT_FMT)
        sc(ws, f"B{r}",
           f"=INDEX('Cost to serve'!$I${cts_first}:$I${cts_last},"
           f"MATCH({rank_target},'Cost to serve'!$K${cts_first}:$K${cts_last},0))",
           font=GREEN, number_format=CURRENCY2)
    last2 = base2 + 2 + m - 1
    add_bar_chart(
        ws, f"D{base2}", "Contribution per hour committed, by student",
        Reference(ws, min_col=1, min_row=base2 + 2, max_row=last2),
        Reference(ws, min_col=2, min_row=base2 + 2, max_row=last2),
        "Contribution per hour ($/hr)",
    )

    # --- Chart 3: utilisation vs ceiling ---
    base3 = last2 + 3
    sc(ws, f"A{base3}", "Utilisation vs capacity ceiling (hrs/week)", font=SECTION_FONT)
    sc(ws, f"A{base3+1}", "Actual hours/week (avg)", border=True)
    sc(ws, f"B{base3+1}", f"=Capacity!{capacity_refs['actual_hours']}", font=GREEN, number_format=NUM2, border=True)
    sc(ws, f"A{base3+2}", "Max teachable hours/week (ceiling)", border=True)
    sc(ws, f"B{base3+2}", f"=Capacity!{capacity_refs['max_hours']}", font=GREEN, number_format=NUM2, border=True)
    add_bar_chart(
        ws, f"D{base3}", "Utilisation vs capacity ceiling",
        Reference(ws, min_col=1, min_row=base3 + 1, max_row=base3 + 2),
        Reference(ws, min_col=2, min_row=base3 + 1, max_row=base3 + 2),
        "Hours/week",
    )

    return ws


def main():
    lessons = load_lessons()
    students = sorted({r["student_pseudonym"] for r in lessons})
    date_min = min(r["lesson_date"] for r in lessons)
    date_max = max(r["lesson_date"] for r in lessons)

    wb = openpyxl.Workbook()
    build_readme(wb)
    a_ws, a_refs = build_assumptions(wb)
    data_ws, data_first, data_last = build_data(wb, lessons)
    rbs_ws, rbs_refs = build_revenue_by_student(wb, students, (data_first, data_last))
    cts_ws, cts_refs = build_cost_to_serve(wb, students, rbs_refs, a_refs, date_min, date_max)
    capacity_ws, capacity_refs = build_capacity(wb, a_refs, rbs_refs, cts_refs)
    pl_ws = build_pl(wb, a_refs, rbs_refs, cts_refs)
    scenarios_ws = build_scenarios(wb, a_refs, capacity_refs)
    # P&L annualised block: header row, factor row, then a 7-row waterfall ->
    # Gross revenue is the first row of that waterfall, cost is the second.
    sensitivity_ws = build_sensitivity(wb, a_refs, pl_annual_revenue_row=14, pl_annual_cost_row=15)
    build_charts(wb, rbs_ws, rbs_refs, cts_ws, cts_refs, capacity_ws, capacity_refs)

    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    print(f"Wrote {OUT_PATH}")
    print(f"Students: {len(students)}  Data rows: {data_last - data_first + 1}")


if __name__ == "__main__":
    main()
