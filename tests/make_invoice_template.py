# -*- coding: utf-8 -*-
"""Builds a placeholder KBB commercial-invoice template that matches the
cell layout documented in the kbb-shipment-invoice skill, so the Excel
building logic (insert_rows gotchas, merge fix, wrap_text, formulas) can
be exercised end-to-end. This is NOT Elihay's real template — swap
assets/kbb_invoice_template.xlsx for the real file before real use.
"""
import os
import openpyxl
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment

HERE = os.path.dirname(os.path.abspath(__file__))
FIX = os.path.join(HERE, "fixtures")
os.makedirs(FIX, exist_ok=True)

THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Invoice"

    ws["B1"] = "COMMERCIAL INVOICE"
    ws["B1"].font = Font(bold=True, size=16)

    ws["G3"] = "Reference No:"
    ws["G5"] = "Ship Date:"
    ws["G15"] = "Gross weight:"
    ws["G22"] = "Defective unit return:"

    headers = ["Part number", "Description", "QTY", "Unit Price $", "Total $", "Country Of Origin"]
    for i, h in enumerate(headers):
        c = ws.cell(row=31, column=3 + i, value=h)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDDDDD")
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")

    # one example row + 20 formatted rows (32-51)
    for row in range(32, 52):
        for col in range(3, 9):  # C..H
            cell = ws.cell(row=row, column=col)
            cell.border = BORDER
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(horizontal="center")
            if col == 8:  # H = Country Of Origin default
                cell.value = "China"
            if col == 5:  # E = numeric qty
                cell.number_format = "0"
            if col in (6, 7):  # F,G = currency
                cell.number_format = "$#,##0.00"
    ws["C32"] = "EXAMPLE-PART"
    ws["D32"] = "Example description"
    ws["E32"] = 1
    ws["F32"] = 10.0
    ws["G32"] = "=E32*F32"

    ws["F52"] = "TOTAL"
    ws["F52"].font = Font(bold=True)
    ws["G52"] = "=SUM(G32:G51)"
    ws["G52"].font = Font(bold=True)
    ws["G52"].number_format = "$#,##0.00"

    ws["C53"] = "DIMS:"
    ws["C53"].font = Font(bold=True)
    ws["D53"] = ""

    ws.merge_cells("B54:I54")
    ws["B54"] = ("We hereby certify that the information on this invoice is "
                 "true and correct and that the contents of this shipment "
                 "are as stated above.")
    ws["B54"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["B54"].font = Font(italic=True, size=9)

    ws.page_setup.scale = 45
    ws.page_setup.orientation = "landscape"

    path = os.path.join(FIX, "COMMERCIAL_INVOICE_TEMPLATE_placeholder.xlsx")
    wb.save(path)
    print("wrote", path)


if __name__ == "__main__":
    build()
