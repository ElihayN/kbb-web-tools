# -*- coding: utf-8 -*-
import os
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

HERE = os.path.dirname(os.path.abspath(__file__))
FIX = os.path.join(HERE, "fixtures")
os.makedirs(FIX, exist_ok=True)


def make_shipment_pdf(filename, header, items):
    path = os.path.join(FIX, filename)
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    # page 1: header block
    c.setFont("Helvetica", 11)
    y = height - 60
    c.drawString(50, y, "KBB / DCS Packing List")
    y -= 30
    for label, value in header.items():
        c.drawString(50, y, f"{label}: {value}")
        y -= 18
    c.showPage()

    # page 2+: item table, ~25 rows per page
    rows_per_page = 25
    for i in range(0, len(items), rows_per_page):
        c.setFont("Helvetica", 10)
        y = height - 60
        c.drawString(50, y, "Single Packed Boxes")
        y -= 24
        c.drawString(50, y, "Part Number   Return Order Number   Purchase Order Number   Country Of Origin")
        y -= 18
        for part, ret_ord, po, country in items[i:i + rows_per_page]:
            c.drawString(50, y, f"{part} {ret_ord} {po} {country}")
            y -= 16
        c.showPage()

    c.save()
    print("wrote", path)


if __name__ == "__main__":
    shipment_a_items = [
        ("HB661-001", "HAL1000001", "PO90001", "China"),
        ("HB661-002", "HAL1000002", "PO90001", "China"),
        ("UV661-003", "HAL1000003", "PO90002", "Vietnam"),
    ]
    make_shipment_pdf(
        "PackingList_KBB_1001.pdf",
        {
            "Date": "01/09/2026",
            "Delivery Address": "1 Apple Park Way, Cupertino, CA",
            "Ship From": "KBB Israel Ltd, Tel Aviv",
            "Shipper": "KBB Logistics",
            "Tracking Number": "KBB-1001",
            "Contact": "Dana Cohen",
            "Number of Single Packed Parts": "3",
            "Number of Overpacks": "1",
            "Total number of boxes": "1",
            "Length": "40",
            "Width": "30",
            "Height": "20",
            "Total Weight": "5.4",
            "Customer Notes": "RMA return batch",
        },
        shipment_a_items,
    )

    shipment_b_items = [
        ("HB661-002", "HAL2000001", "PO90101", "Vietnam"),  # same part, conflicting country
        ("ZM661-010", "HAL2000002", "PO90102", "China"),
        ("ZM661-010", "HAL2000003", "PO90102", "China"),
        ("PA661-020", "HAL2000004", "PO90103", "China"),
    ]
    make_shipment_pdf(
        "PackingList_KBB_1002.pdf",
        {
            "Date": "02/09/2026",
            "Delivery Address": "1 Apple Park Way, Cupertino, CA",
            "Ship From": "KBB Israel Ltd, Tel Aviv",
            "Shipper": "KBB Logistics",
            "Tracking Number": "KBB-1002",
            "Contact": "Dana Cohen",
            "Number of Single Packed Parts": "4",
            "Number of Overpacks": "1",
            "Total number of boxes": "2",
            "Length": "45",
            "Width": "32",
            "Height": "22",
            "Total Weight": "7.1",
            "Customer Notes": "RMA return batch 2",
        },
        shipment_b_items,
    )
