"""Force a full recalculation of the workbook via Excel COM automation and
report any formula errors. openpyxl never evaluates formulas -- this is the
only way to prove they actually compute, not just that they parse.

Requires Microsoft Excel installed (pywin32 for COM access). Saves the
workbook back with calculated values cached, so opening it normally in
Excel shows correct numbers without needing a manual recalc.
"""

import sys
from pathlib import Path

import win32com.client as win32

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "workbook" / "tutoring-financial-model.xlsx"

ERROR_VALUES = {
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!", "#SPILL!",
}


def main():
    if not WORKBOOK.exists():
        raise SystemExit(f"Workbook not found: {WORKBOOK}")

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    errors = []
    try:
        wb = excel.Workbooks.Open(str(WORKBOOK))
        excel.CalculateFullRebuild()

        for sheet in wb.Sheets:
            used = sheet.UsedRange
            values = used.Value
            if values is None:
                continue
            # Value is a scalar for a 1x1 range, else a tuple of tuples
            rows = values if isinstance(values, tuple) else ((values,),)
            row0 = used.Row
            col0 = used.Column
            for i, row in enumerate(rows):
                if not isinstance(row, tuple):
                    row = (row,)
                for j, val in enumerate(row):
                    if isinstance(val, str) and val in ERROR_VALUES:
                        cell = sheet.Cells(row0 + i, col0 + j)
                        errors.append((sheet.Name, cell.Address, val, cell.Formula))

        wb.Save()
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    if errors:
        print(f"FOUND {len(errors)} formula error(s):")
        for sheet_name, addr, val, formula in errors:
            print(f"  {sheet_name}!{addr}: {val}   [{formula}]")
        sys.exit(1)
    else:
        print("Recalculated clean: zero formula errors found across all sheets.")


if __name__ == "__main__":
    main()
