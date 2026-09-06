# -*- coding: utf-8 -*-
import os
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm

HERE = os.path.dirname(os.path.abspath(__file__))
FIX = os.path.join(HERE, "fixtures")
os.makedirs(FIX, exist_ok=True)

styles = getSampleStyleSheet()


def make_invoice_pdf(filename, shipment_number, date, hawb, gross_weight, items):
    """items: list of (order, item_no, part_number, description, country, qty, unit_price)"""
    path = os.path.join(FIX, filename)
    doc = SimpleDocTemplate(path, pagesize=A4)
    story = []

    story.append(Paragraph(f"Shipment Number: {shipment_number}", styles["Normal"]))
    story.append(Paragraph(f"Date: {date}", styles["Normal"]))
    story.append(Spacer(1, 12))

    table_data = [["Order", "Item", "Part Number", "Description", "Qty", "Unit Price", "Total Price"]]
    grand_total = 0.0
    for order, item_no, part, desc, country, qty, unit_price in items:
        total_price = qty * unit_price
        grand_total += total_price
        table_data.append([
            str(order), str(item_no), part, f"{desc} {country}",
            f"{qty:.3f}", f"{unit_price:.2f}", f"{total_price:.2f}",
        ])

    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
    ]))
    story.append(t)
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Grand Total: {grand_total:.2f}", styles["Normal"]))
    story.append(Spacer(1, 24))
    story.append(Paragraph("Shipping Advice Only", styles["Heading4"]))
    story.append(Paragraph(f"HAWB {hawb}", styles["Normal"]))
    story.append(Paragraph(f"Gross Weight {gross_weight} KG", styles["Normal"]))
    story.append(Paragraph(f"Net Weight {gross_weight - 0.5:.2f} KG", styles["Normal"]))
    story.append(Paragraph(f"Total Value {grand_total:.2f}", styles["Normal"]))

    doc.build(story)
    print("wrote", path, "grand_total=", grand_total)


if __name__ == "__main__":
    make_invoice_pdf(
        "Proforma_1_KBB_5001.pdf", "SHP-5001", "01/09/2026", "KBB 62127", 8.4,
        items=[
            (1, 1, "HB661-1001", "iPhone Battery", "CN", 1, 45.00),
            (1, 2, "UV661-2002", "iPhone Camera Module", "CN", 1, 32.50),
            (1, 3, "ZM661-3003", "iPhone Screen", "VN", 1, 120.00),
        ],
    )
    make_invoice_pdf(
        "Proforma_2_KBB_5002.pdf", "SHP-5002", "02/09/2026", "KBB 62128", 5.1,
        items=[
            (1, 1, "HB661-1001", "iPhone Battery", "CN", 1, 45.00),  # same part, same price -> merges
            (1, 2, "PA661-4004", "iPad Charging Port", "CN", 1, 18.75),
        ],
    )
