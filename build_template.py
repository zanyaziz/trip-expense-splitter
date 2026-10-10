#!/usr/bin/env python3
"""Build the Trip Expense Splitter Excel template.

Sheets:
  People   - traveler names (up to 12) + usage instructions
  Expenses - expense log with Paid By dropdown + per-person TRUE/FALSE split flags
  Balances - per-person totals: paid, owed, net
  Settle Up - greedy minimal-transfer settlement computed with formulas
"""
import copy
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

MAX_PEOPLE = 12
PEOPLE_FIRST_ROW = 10          # People!A10:A21
EXP_FIRST_ROW = 2
EXP_LAST_ROW = 101
EXP_FLAG_FIRST_COL = 5          # E
EXP_FLAG_LAST_COL = EXP_FLAG_FIRST_COL + MAX_PEOPLE - 1  # P
SHARE_COL = EXP_FLAG_LAST_COL + 1  # Q

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E78")
SUBTITLE_FONT = Font(bold=True, size=11, color="1F4E78")
MONEY_FMT = '"$"#,##0.00'
THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)


def style_header_row(ws, ncols):
    for col in range(1, ncols + 1):
        c = ws.cell(row=1, column=col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER
    ws.row_dimensions[1].height = 30


def build():
    wb = Workbook()

    # ---------------------------------------------------------------- People
    ws = wb.active
    ws.title = "People"
    ws["A1"] = "Trip Expense Splitter"
    ws["A1"].font = TITLE_FONT
    instructions = [
        "A free, no-signup alternative to Splitwise. Enter each trip expense once,",
        "mark who paid and who shares it, and the sheets compute who owes whom.",
        "",
        "HOW TO USE (5 steps):",
        "1. Type the traveler names below (up to 12). The other sheets update automatically.",
        "2. On the Expenses sheet, add one row per expense: date, description, who paid,",
        "   amount, then put TRUE in each traveler's column who shares that expense.",
        "3. Share Each is computed for you: Amount ÷ number of TRUEs (0 if nobody shares).",
        "4. Check the Balances sheet: Net = Total Paid − Total Owed. Positive means",
        "   the group owes that person; negative means they owe the group.",
        "5. Open the Settle Up sheet for the shortest list of 'X pays Y $Z' transfers",
        "   that clears every debt.",
        "",
        "Tips: leave unused traveler rows as-is; the sheets ignore people with no",
        "expenses. Amounts in any one currency per trip. Delete sample rows on the",
        "Expenses sheet before your trip (or leave them — formulas skip blank rows).",
    ]
    for i, line in enumerate(instructions):
        ws.cell(row=2 + i, column=1, value=line)
    ws.cell(row=9, column=1, value="Traveler name (edit me)").font = SUBTITLE_FONT
    for i in range(MAX_PEOPLE):
        c = ws.cell(row=PEOPLE_FIRST_ROW + i, column=1, value=f"Person {i + 1}")
        c.border = THIN_BORDER
        c.alignment = Alignment(horizontal="left")
    ws.column_dimensions["A"].width = 70
    for r in range(2, 9 + len(instructions)):
        ws.row_dimensions[r].height = 15

    # Named range for the traveler list (used by dropdowns)
    wb.defined_names["Travelers"] = DefinedName(
        "Travelers", attr_text=f"People!$A${PEOPLE_FIRST_ROW}:$A${PEOPLE_FIRST_ROW + MAX_PEOPLE - 1}"
    )

    # --------------------------------------------------------------- Expenses
    ex = wb.create_sheet("Expenses")
    headers = ["Date", "Description", "Paid By", "Amount"]
    for col, h in enumerate(headers, start=1):
        ex.cell(row=1, column=col, value=h)
    for i in range(MAX_PEOPLE):
        # Header mirrors the name typed on the People sheet
        ex.cell(row=1, column=EXP_FLAG_FIRST_COL + i,
                value=f"=People!$A${PEOPLE_FIRST_ROW + i}")
    ex.cell(row=1, column=SHARE_COL, value="Share Each")
    style_header_row(ex, SHARE_COL)

    # Paid By dropdown -> traveler names
    dv_paid = DataValidation(type="list", formula1="Travelers", allow_blank=True,
                             showErrorMessage=True)
    dv_paid.error = "Pick a traveler from the People sheet."
    dv_paid.prompt = "Who paid for this expense?"
    dv_paid.showInputMessage = True
    ex.add_data_validation(dv_paid)
    dv_paid.add(f"C{EXP_FIRST_ROW}:C{EXP_LAST_ROW}")

    # Split flags -> TRUE / FALSE dropdown
    dv_flag = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    dv_flag.prompt = "TRUE = this traveler shares the expense."
    dv_flag.showInputMessage = True
    ex.add_data_validation(dv_flag)
    flag_range = (f"{get_column_letter(EXP_FLAG_FIRST_COL)}{EXP_FIRST_ROW}:"
                  f"{get_column_letter(EXP_FLAG_LAST_COL)}{EXP_LAST_ROW}")
    dv_flag.add(flag_range)

    for r in range(EXP_FIRST_ROW, EXP_LAST_ROW + 1):
        first_flag = f"{get_column_letter(EXP_FLAG_FIRST_COL)}{r}"
        last_flag = f"{get_column_letter(EXP_FLAG_LAST_COL)}{r}"
        # Unquoted TRUE criterion matches the boolean TRUE the dropdown writes,
        # in Excel, Google Sheets, and LibreOffice alike.
        ex.cell(row=r, column=SHARE_COL,
                value=(f"=IFERROR(D{r}/COUNTIF({first_flag}:{last_flag},TRUE),0)"))
        ex.cell(row=r, column=SHARE_COL).number_format = MONEY_FMT
        ex.cell(row=r, column=4).number_format = MONEY_FMT
        ex.cell(row=r, column=1).number_format = "MM/DD/YYYY"
        for col in range(1, SHARE_COL + 1):
            ex.cell(row=r, column=col).border = THIN_BORDER

    ex.column_dimensions["A"].width = 12
    ex.column_dimensions["B"].width = 32
    ex.column_dimensions["C"].width = 16
    ex.column_dimensions["D"].width = 12
    for i in range(MAX_PEOPLE):
        ex.column_dimensions[get_column_letter(EXP_FLAG_FIRST_COL + i)].width = 13
    ex.column_dimensions[get_column_letter(SHARE_COL)].width = 12
    ex.freeze_panes = "A2"
    ex.auto_filter.ref = f"A1:{get_column_letter(SHARE_COL)}{EXP_LAST_ROW}"
    ex.sheet_properties.pageSetUpPr.fitToPage = True

    # --------------------------------------------------------------- Balances
    bal = wb.create_sheet("Balances")
    for col, h in enumerate(["Person", "Total Paid", "Total Share Owed", "Net (Paid − Owed)"], start=1):
        bal.cell(row=1, column=col, value=h)
    style_header_row(bal, 4)
    for i in range(MAX_PEOPLE):
        r = 2 + i
        prow = PEOPLE_FIRST_ROW + i
        flag_col = get_column_letter(EXP_FLAG_FIRST_COL + i)
        bal.cell(row=r, column=1, value=f"=People!$A${prow}")
        # Total paid by this person
        bal.cell(row=r, column=2,
                 value=(f"=SUMIF(Expenses!$C${EXP_FIRST_ROW}:$C${EXP_LAST_ROW},"
                        f"$A{r},Expenses!$D${EXP_FIRST_ROW}:$D${EXP_LAST_ROW})"))
        # Total share owed: sum of Share Each on rows where this person's flag is TRUE
        # (the dropdown writes boolean TRUE; the unquoted criterion matches it
        # in Excel, Google Sheets, and LibreOffice)
        bal.cell(row=r, column=3,
                 value=(f"=SUMIF(Expenses!${flag_col}${EXP_FIRST_ROW}:${flag_col}${EXP_LAST_ROW},"
                        f"TRUE,Expenses!${get_column_letter(SHARE_COL)}${EXP_FIRST_ROW}:"
                        f"${get_column_letter(SHARE_COL)}${EXP_LAST_ROW})"))
        bal.cell(row=r, column=4, value=f"=B{r}-C{r}")
        for col in range(1, 5):
            c = bal.cell(row=r, column=col)
            c.border = THIN_BORDER
            if col > 1:
                c.number_format = MONEY_FMT
    # Totals / sanity check row
    tr = 2 + MAX_PEOPLE + 1
    bal.cell(row=tr, column=1, value="CHECK (Net must sum to 0)").font = SUBTITLE_FONT
    bal.cell(row=tr, column=4, value=f"=SUM(D2:D{1 + MAX_PEOPLE})")
    bal.cell(row=tr, column=4).number_format = MONEY_FMT
    for w, col in zip([22, 14, 18, 18], "ABCD"):
        bal.column_dimensions[col].width = w

    # -------------------------------------------------------------- Settle Up
    st = wb.create_sheet("Settle Up")
    st["A1"] = "Who pays whom (fewest transfers)"
    st["A1"].font = TITLE_FONT
    st["A2"] = ("Greedy settlement: each step matches the person who owes the most with the "
                "person who is owed the most. Clearing one person per step means at most "
                f"{MAX_PEOPLE - 1} transfers for {MAX_PEOPLE} travelers.")
    st["A2"].alignment = Alignment(wrap_text=True)
    st.merge_cells("A2:R2")
    for col, h in enumerate(["Step", "From (pays)", "To (receives)", "Amount"], start=1):
        c = st.cell(row=4, column=col, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = THIN_BORDER

    HDR = 4
    first_step_row = HDR + 1            # rows 5..15 hold steps 1..11
    WORK_NAME_COL = 6                  # F
    WORK_START_COL = 7                 # G = starting net; H..R = after steps 1..11
    n_steps = MAX_PEOPLE - 1

    # Working-balance header
    st.cell(row=HDR, column=WORK_NAME_COL, value="Person").font = HEADER_FONT
    st.cell(row=HDR, column=WORK_NAME_COL).fill = HEADER_FILL
    st.cell(row=HDR, column=WORK_START_COL, value="Net balance").font = HEADER_FONT
    st.cell(row=HDR, column=WORK_START_COL).fill = HEADER_FILL
    for k in range(1, n_steps + 1):
        c = st.cell(row=HDR, column=WORK_START_COL + k, value=f"After step {k}")
        c.font = Font(bold=True, size=9)
        c.alignment = Alignment(horizontal="center")

    person_rows = []
    for i in range(MAX_PEOPLE):
        r = first_step_row + i
        person_rows.append(r)
        st.cell(row=r, column=WORK_NAME_COL, value=f"=Balances!$A${2 + i}")
        st.cell(row=r, column=WORK_START_COL, value=f"=ROUND(Balances!$D${2 + i},2)")
        st.cell(row=r, column=WORK_START_COL).number_format = MONEY_FMT
        st.cell(row=r, column=WORK_NAME_COL).border = THIN_BORDER

    name_range = f"${get_column_letter(WORK_NAME_COL)}${first_step_row}:${get_column_letter(WORK_NAME_COL)}${first_step_row + MAX_PEOPLE - 1}"

    for k in range(1, n_steps + 1):
        srow = HDR + k                      # transfer-table row for step k
        wcol_prev = WORK_START_COL + k - 1   # working column BEFORE step k
        wcol = WORK_START_COL + k            # working column AFTER step k
        wrange = (f"${get_column_letter(wcol_prev)}${first_step_row}:"
                  f"${get_column_letter(wcol_prev)}${first_step_row + MAX_PEOPLE - 1}")
        st.cell(row=srow, column=1, value=k)
        # Amount: min(what the biggest debtor owes, what the biggest creditor is owed)
        st.cell(row=srow, column=4,
                value=(f"=IF(AND(MIN({wrange})<-0.005,MAX({wrange})>0.005),"
                       f"MIN(-MIN({wrange}),MAX({wrange})),0)"))
        st.cell(row=srow, column=4).number_format = MONEY_FMT
        # From = biggest debtor, To = biggest creditor
        st.cell(row=srow, column=2,
                value=(f"=IF($D${srow}>0.005,INDEX({name_range},"
                       f"MATCH(MIN({wrange}),{wrange},0)),\"\")"))
        st.cell(row=srow, column=3,
                value=(f"=IF($D${srow}>0.005,INDEX({name_range},"
                       f"MATCH(MAX({wrange}),{wrange},0)),\"\")"))
        for col in (1, 2, 3, 4):
            st.cell(row=srow, column=col).border = THIN_BORDER
        # Working balances after this step
        pcol = get_column_letter(wcol_prev)
        for i, prow in enumerate(person_rows):
            fcell = f"${get_column_letter(WORK_NAME_COL)}${prow}"
            st.cell(row=prow, column=wcol,
                    value=(f"=ROUND(${pcol}${prow}"
                           f"+IF($B${srow}={fcell},$D${srow},0)"
                           f"-IF($C${srow}={fcell},$D${srow},0),2)"))
            st.cell(row=prow, column=wcol).number_format = MONEY_FMT

    # Final check row: remaining unsettled total should be 0
    chk_row = first_step_row + MAX_PEOPLE + 1
    final_col = get_column_letter(WORK_START_COL + n_steps)
    st.cell(row=chk_row, column=1, value="Remaining unsettled (must be 0):").font = SUBTITLE_FONT
    st.merge_cells(start_row=chk_row, start_column=1, end_row=chk_row, end_column=3)
    st.cell(row=chk_row, column=4,
            value=(f"=SUMPRODUCT(ABS(${final_col}${first_step_row}:${final_col}${first_step_row + MAX_PEOPLE - 1}))"))
    st.cell(row=chk_row, column=4).number_format = MONEY_FMT

    for col, w in zip("ABCD", [8, 18, 18, 12]):
        st.column_dimensions[col].width = w
    st.column_dimensions["F"].width = 16
    for c in range(WORK_START_COL, WORK_START_COL + n_steps + 1):
        st.column_dimensions[get_column_letter(c)].width = 11
    st.freeze_panes = "A5"

    # Readability: light banding on Expenses flag area is skipped to keep file simple.
    return wb


if __name__ == "__main__":
    wb = build()
    out = "/home/hatch/workspace/trip-expense-splitter/trip-expense-splitter.xlsx"
    wb.save(out)
    print("saved", out)
