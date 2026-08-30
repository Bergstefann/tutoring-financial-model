"""Export PNG screenshots of key sheets for the README.

Uses Excel COM automation to export each sheet's used range to a one-page
PDF (fit-to-page), then rasterises that PDF page to PNG with PyMuPDF. More
reliable than the CopyPicture/paste-into-chart clipboard trick, which
produced blank images.

Requires Microsoft Excel installed (pywin32) and pymupdf.
"""

from pathlib import Path

import pymupdf
import win32com.client as win32
from PIL import Image, ImageChops

ROOT = Path(__file__).resolve().parent.parent
WORKBOOK = ROOT / "workbook" / "tutoring-financial-model.xlsx"
OUT_DIR = ROOT / "docs" / "screenshots"
TMP_DIR = ROOT / "data" / "tmp_pdf"

XL_TYPE_PDF = 0
XL_LANDSCAPE = 2
XL_PORTRAIT = 1

# (sheet name, output filename, landscape?)
SHEETS = [
    ("Assumptions", "assumptions.png", False),
    ("Revenue by student", "revenue-by-student.png", True),
    ("Cost to serve", "cost-to-serve.png", True),
    ("Capacity", "capacity.png", False),
    ("P&L and provisioning", "pl.png", False),
    ("Scenarios", "scenarios.png", True),
    ("Sensitivity", "sensitivity.png", True),
]

DPI = 200


def export_sheet(ws, out_png, landscape):
    ps = ws.PageSetup
    ps.Orientation = XL_LANDSCAPE if landscape else XL_PORTRAIT
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.LeftMargin = ps.RightMargin = ps.TopMargin = ps.BottomMargin = 20
    ps.CenterHorizontally = True

    TMP_DIR.mkdir(parents=True, exist_ok=True)
    pdf_path = TMP_DIR / (out_png.stem + ".pdf")
    ws.ExportAsFixedFormat(Type=XL_TYPE_PDF, Filename=str(pdf_path))

    doc = pymupdf.open(str(pdf_path))
    page = doc[0]
    zoom = DPI / 72
    pix = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
    pix.save(str(out_png))
    doc.close()
    pdf_path.unlink()
    crop_whitespace(out_png)


def crop_whitespace(png_path, padding=24):
    im = Image.open(png_path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    diff = ImageChops.difference(im, bg)
    bbox = diff.getbbox()
    if bbox is None:
        return
    left, top, right, bottom = bbox
    left = max(0, left - padding)
    top = max(0, top - padding)
    right = min(im.width, right + padding)
    bottom = min(im.height, bottom + padding)
    im.crop((left, top, right, bottom)).save(png_path)


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(WORKBOOK))
        for sheet_name, filename, landscape in SHEETS:
            ws = wb.Sheets(sheet_name)
            out_png = OUT_DIR / filename
            export_sheet(ws, out_png, landscape)
            print(f"Wrote {out_png.relative_to(ROOT)}")
        wb.Close(SaveChanges=False)
    finally:
        excel.Quit()

    if TMP_DIR.exists():
        TMP_DIR.rmdir()


if __name__ == "__main__":
    main()
