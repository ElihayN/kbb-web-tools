# -*- coding: utf-8 -*-
"""
kbb-packing-list-consolidation — merges several KBB/DCS Packing List PDFs
(RMA return shipments to Apple) into one unified Excel packing list.

Deterministic re-implementation of the corresponding Claude skill. No AI,
no vision reading — pure pdfplumber + regex, with the exact verification
asserts the skill calls "load-bearing": row counts must reconcile with the
stated header totals before anything gets consolidated.
"""
import os
import re
import statistics
from collections import defaultdict, Counter

import pdfplumber
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter

from . import office

ITEM_LINE_RE = re.compile(r"^(\S+)\s+(HAL\S+)\s+(\S+)\s+(.+)$")

HEADER_FIELD_ALIASES = {
    "date": "date",
    "delivery address": "delivery_address",
    "ship from": "ship_from",
    "shipper": "shipper",
    "tracking number": "tracking_number",
    "contact": "contact",
    "number of single packed parts": "n_single_packed",
    "number of overpacks": "n_overpacks",
    "total number of boxes": "total_boxes",
    "length": "length",
    "width": "width",
    "height": "height",
    "total weight": "total_weight",
    "customer notes": "customer_notes",
}

THIN = Side(style="thin", color="B0B4BD")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F2A44")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial")
CENTER = Alignment(horizontal="center", vertical="center")


class PackingListError(RuntimeError):
    pass


def _parse_header_fields(text: str) -> dict:
    fields = {}
    for line in text.splitlines():
        m = re.match(r"^\s*([A-Za-z][A-Za-z /]+?)\s*:\s*(.+?)\s*$", line)
        if not m:
            continue
        label = m.group(1).strip().lower()
        value = m.group(2).strip()
        key = HEADER_FIELD_ALIASES.get(label)
        if key:
            fields[key] = value
    return fields


def _to_int(value, default=None):
    if value is None:
        return default
    m = re.search(r"-?\d+", str(value).replace(",", ""))
    return int(m.group()) if m else default


def _to_float(value, default=None):
    if value is None:
        return default
    m = re.search(r"-?\d+(\.\d+)?", str(value).replace(",", ""))
    return float(m.group()) if m else default


def _extract_shipment(pdf_path: str) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            raise PackingListError(f"{os.path.basename(pdf_path)}: קובץ PDF ריק.")
        header_text = pdf.pages[0].extract_text() or ""
        header = _parse_header_fields(header_text)

        items = []
        for page in pdf.pages[1:]:
            text = page.extract_text() or ""
            for line in text.splitlines():
                m = ITEM_LINE_RE.match(line.strip())
                if not m:
                    continue
                part_number, return_order, po_number, country = m.groups()
                items.append({
                    "part_number": part_number,
                    "return_order": return_order,
                    "po_number": po_number,
                    "country": country.strip(),
                })

    expected = _to_int(header.get("n_single_packed"))
    tracking = header.get("tracking_number") or os.path.splitext(os.path.basename(pdf_path))[0]

    if expected is None:
        raise PackingListError(
            f"{os.path.basename(pdf_path)}: לא הצלחתי לזהות את השדה "
            f"'Number of Single Packed Parts' בעמוד הראשון — לא ניתן לאמת "
            f"את מספר השורות שחולצו. בדוק שהקובץ תואם לפורמט הצפוי."
        )
    if len(items) != expected:
        raise PackingListError(
            f"משלוח {tracking}: חולצו {len(items)} שורות פריט, אך העמוד הראשון "
            f"מצהיר על {expected} ('Number of Single Packed Parts'). "
            f"לא ממשיכים לאיחוד עד שהפער הזה יובן — ייתכן שינוי בפורמט ה-PDF "
            f"שדורש עדכון בביטוי הרגולרי של החילוץ."
        )

    header["tracking_number"] = tracking
    header["n_single_packed"] = expected
    header["_source_file"] = os.path.basename(pdf_path)
    return {"header": header, "items": items}


def consolidate(pdf_paths, work_dir):
    """Steps 1-3 of the skill: extract every shipment, verify, consolidate.
    Returns a dict ready for _build_excel().
    """
    shipments = [_extract_shipment(p) for p in pdf_paths]

    by_part = defaultdict(lambda: {"qty": 0, "countries": Counter(), "sources": []})
    for shp in shipments:
        tn = shp["header"]["tracking_number"]
        for item in shp["items"]:
            entry = by_part[item["part_number"]]
            entry["qty"] += 1
            if item["country"]:
                entry["countries"][item["country"]] += 1
            entry["sources"].append(tn)

    conflicts = []
    consolidated = []
    for part, entry in sorted(by_part.items()):
        countries = entry["countries"]
        if countries:
            top_country, top_n = countries.most_common(1)[0]
            if len(countries) > 1:
                conflicts.append({
                    "part_number": part,
                    "countries": dict(countries),
                    "chosen": top_country,
                })
        else:
            top_country = ""
        consolidated.append({
            "part_number": part,
            "quantity": entry["qty"],
            "country": top_country,
            "sources": sorted(set(entry["sources"])),
        })

    total_expected = sum(s["header"]["n_single_packed"] for s in shipments)
    total_merged = sum(c["quantity"] for c in consolidated)
    if total_merged != total_expected:
        raise PackingListError(
            f"בדיקת סכימה נכשלה: סה\"כ יחידות מאוחדות ({total_merged}) "
            f"אינו תואם לסכום ה-'Number of Single Packed Parts' מכל המשלוחים "
            f"({total_expected}). לא ממשיכים — ייתכן חוסר עקביות בחילוץ."
        )

    return {
        "shipments": shipments,
        "consolidated": consolidated,
        "conflicts": conflicts,
        "total_quantity": total_merged,
    }


def _clear_and_set_widths(ws, widths: dict, last_col: int):
    """Step 5 gotcha fix: never touch column_dimensions incrementally.
    Clear everything, then set each column individually and explicitly so
    openpyxl can't leave overlapping <col> ranges that LibreOffice's PDF
    export then misreads."""
    ws.column_dimensions.clear()
    for col in range(1, last_col + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = widths.get(col, 16)


def _style_row(ws, row, ncols, header=False, center=True):
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.border = BORDER
        cell.font = HEADER_FONT if header else BODY_FONT
        if header:
            cell.fill = HEADER_FILL
        if center:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=header)
    if header:
        ws.row_dimensions[row].height = 26


def build_excel(data: dict, work_dir: str) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Packing List מאוחד"
    ws.sheet_view.rightToLeft = True

    shipments = data["shipments"]
    consolidated = data["consolidated"]

    r = 1
    ws.cell(row=r, column=1, value="Packing List מאוחד — KBB").font = Font(name="Arial", bold=True, size=14)
    r += 2

    # --- general info block --------------------------------------------
    first = shipments[0]["header"]
    general_fields = [
        ("Delivery Address", first.get("delivery_address", "")),
        ("Ship From", first.get("ship_from", "")),
        ("Shipper", first.get("shipper", "")),
        ("Contact", first.get("contact", "")),
        ("Tracking Numbers", ", ".join(s["header"]["tracking_number"] for s in shipments)),
    ]
    for label, value in general_fields:
        ws.cell(row=r, column=1, value=label).font = Font(name="Arial", bold=True)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        value_cell = ws.cell(row=r, column=2, value=value)
        value_cell.font = BODY_FONT
        # long addresses/tracking lists would otherwise get visually
        # clipped by the neighboring label cell (same gotcha as the KBB
        # invoice template's header fields) — wrap + give the row room.
        value_cell.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1

    items_start_placeholder_row = None  # filled in later once we know it
    summary_row_start = r
    ws.cell(row=r, column=1, value="שורות פריט ייחודיות").font = Font(name="Arial", bold=True)
    unique_items_cell = f"B{r}"
    r += 1
    ws.cell(row=r, column=1, value='סה"כ כמות').font = Font(name="Arial", bold=True)
    total_qty_cell = f"B{r}"
    r += 1
    ws.cell(row=r, column=1, value='סה"כ משקל (ק"ג)').font = Font(name="Arial", bold=True)
    total_weight_cell = f"B{r}"
    r += 2

    # --- per-shipment breakdown table -----------------------------------
    ws.cell(row=r, column=1, value="פירוט לפי משלוח").font = Font(name="Arial", bold=True, size=12)
    r += 1
    breakdown_header_row = r
    headers = ["Tracking Number", "Date", "Boxes", "Weight (kg)", "Length (cm)", "Width (cm)", "Height (cm)"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=r, column=c, value=h)
    _style_row(ws, r, len(headers), header=True)
    r += 1
    breakdown_first_data_row = r
    for shp in shipments:
        h = shp["header"]
        ws.cell(row=r, column=1, value=h.get("tracking_number", ""))
        ws.cell(row=r, column=2, value=h.get("date", ""))
        ws.cell(row=r, column=3, value=_to_int(h.get("total_boxes"), 0))
        ws.cell(row=r, column=4, value=_to_float(h.get("total_weight"), 0))
        ws.cell(row=r, column=5, value=_to_float(h.get("length"), None))
        ws.cell(row=r, column=6, value=_to_float(h.get("width"), None))
        ws.cell(row=r, column=7, value=_to_float(h.get("height"), None))
        _style_row(ws, r, len(headers))
        r += 1
    breakdown_last_data_row = r - 1
    ws.cell(row=r, column=1, value="TOTAL")
    ws.cell(row=r, column=3, value=f"=SUM(C{breakdown_first_data_row}:C{breakdown_last_data_row})")
    ws.cell(row=r, column=4, value=f"=SUM(D{breakdown_first_data_row}:D{breakdown_last_data_row})")
    ws.cell(row=r, column=5, value="-")
    ws.cell(row=r, column=6, value="-")
    ws.cell(row=r, column=7, value="-")
    _style_row(ws, r, len(headers), header=True)
    breakdown_total_row = r
    r += 2

    # now that we know the breakdown's weight total cell, wire the summary
    ws[total_weight_cell] = f"=D{breakdown_total_row}"
    ws[total_weight_cell].font = BODY_FONT

    # --- consolidated items table ----------------------------------------
    ws.cell(row=r, column=1, value="פריטים מאוחדים").font = Font(name="Arial", bold=True, size=12)
    r += 1
    items_header_row = r
    item_headers = ["Part Number", "Quantity", "Country Of Origin"]
    for c, h in enumerate(item_headers, start=1):
        ws.cell(row=r, column=c, value=h)
    _style_row(ws, r, len(item_headers), header=True)
    r += 1
    items_first_data_row = r
    for item in consolidated:
        ws.cell(row=r, column=1, value=item["part_number"])
        ws.cell(row=r, column=2, value=item["quantity"])
        ws.cell(row=r, column=3, value=item["country"])
        if len(item["sources"]) > 1 or True:
            ws.cell(row=r, column=1).comment = Comment(
                "מקור: " + ", ".join(item["sources"]), "kbb-packing-list-consolidation"
            )
        _style_row(ws, r, len(item_headers))
        r += 1
    items_last_data_row = r - 1
    ws.cell(row=r, column=1, value="TOTAL")
    ws.cell(row=r, column=2, value=f"=SUM(B{items_first_data_row}:B{items_last_data_row})")
    ws.cell(row=r, column=3, value="")
    _style_row(ws, r, len(item_headers), header=True)
    items_total_row = r

    # summary formulas that depend on the items table
    ws[unique_items_cell] = f"=COUNTA(A{items_first_data_row}:A{items_last_data_row})"
    ws[unique_items_cell].font = BODY_FONT
    ws[total_qty_cell] = f"=B{items_total_row}"
    ws[total_qty_cell].font = BODY_FONT

    # --- formatting fixes --------------------------------------------------
    last_col = max(len(headers), len(item_headers))
    # column C is shared between the breakdown table's "Boxes" and the
    # items table's "Country Of Origin" header — size it for the wider one.
    widths = {1: 24, 2: 16, 3: 20, 4: 14, 5: 14, 6: 14, 7: 14}
    _clear_and_set_widths(ws, widths, last_col)

    wb.calculation.fullCalcOnLoad = True

    # print setup, ready for the (separate, post-approval) PDF export step
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{items_header_row}:{items_header_row}"

    path = os.path.join(work_dir, "Packing_List_מאוחד.xlsx")
    wb.save(path)
    return path


def process(pdf_paths, work_dir):
    os.makedirs(work_dir, exist_ok=True)
    data = consolidate(pdf_paths, work_dir)
    xlsx_path = build_excel(data, work_dir)
    xlsx_path = office.recalc_xlsx(xlsx_path, work_dir)
    return {
        "xlsx_path": xlsx_path,
        "unique_items": len(data["consolidated"]),
        "total_quantity": data["total_quantity"],
        "conflicts": data["conflicts"],
        "n_shipments": len(data["shipments"]),
    }


def export_pdf(xlsx_path: str, work_dir: str) -> str:
    """Step 8: PDF export only after the user has reviewed/approved the
    Excel — page setup was already baked in by build_excel(), so this is
    just the conversion + (caller-side) visual sanity check."""
    os.makedirs(work_dir, exist_ok=True)
    return office.to_pdf(xlsx_path, work_dir)
