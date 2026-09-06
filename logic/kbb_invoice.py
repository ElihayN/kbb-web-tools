# -*- coding: utf-8 -*-
"""
kbb-shipment-invoice — consolidates several Apple/DCS proforma invoice PDFs
into the fixed KBB commercial-invoice Excel template.

Deterministic re-implementation of the corresponding Claude skill.

IMPORTANT — read before trusting this in production:
The original skill read each source PDF with Claude's vision-capable Read
tool, which tolerates real-world layout noise (page breaks splitting a
description across two pages, odd spacing, OCR artifacts) far better than
a fixed regex/table parser can. The extraction logic below (see
`_extract_invoice_tables` / `_ITEM ROW PARSING` for the exact spot to
tune) is a best-effort port based on the skill's written description of
the PDF layout — it has NOT been validated against a real Apple/DCS
invoice PDF, because none was available while building this tool.

The one safety net kept from the original skill is load-bearing and is
NOT optional: every invoice's extracted line items must sum to that
invoice's own stated Grand Total, and the full merged total must match
the sum of every source invoice's Grand Total. If the parser is
mis-tuned for a real file's layout, this assertion is what stands between
that and a wrong commercial invoice reaching a customs shipment — it
raises loudly instead of silently emitting bad numbers. Before relying on
this for a real shipment, run it once against a real invoice batch and
manually double check the drafted Excel line-by-line against the source
PDFs. If the totals assert cleanly, the extraction can be trusted; if it
raises, the regex needs a tuning pass against that real file (send it to
be recalibrated) rather than being forced through.
"""
import os
import re
from collections import defaultdict, Counter

import pdfplumber
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from . import office

DEFAULT_COUNTRY = "CN"

# --------------------------------------------------------------------------
# ITEM ROW PARSING — the spot to tune once a real sample invoice is on hand.
# Expected columns per the skill: Order, Item, Part Number/Harmonised code,
# Description/Country of Origin/Customer Reference (stacked), Qty, Unit
# Price, Total Price.
# --------------------------------------------------------------------------
PART_NUMBER_RE = re.compile(r"^[A-Z]{2,4}\d{2,4}-[A-Z0-9]+$")
MONEY_RE = re.compile(r"^-?\d[\d,]*\.\d{2}$")
GRAND_TOTAL_RE = re.compile(r"Grand\s*Total\s*:?\s*([\d,]+\.\d{2})", re.IGNORECASE)
SHIPMENT_NUMBER_RE = re.compile(r"Shipment\s*Number\s*:?\s*([A-Za-z0-9][A-Za-z0-9\-]*)", re.IGNORECASE)
HAWB_RE = re.compile(r"HAWB\s*:?\s*([A-Za-z]*\s*\d+)", re.IGNORECASE)
GROSS_WEIGHT_RE = re.compile(r"Gross\s*Weight\s*:?\s*([\d.,]+)", re.IGNORECASE)


class InvoiceError(RuntimeError):
    pass


def _money(s):
    return float(str(s).replace(",", "").strip())


def _extract_rows_from_tables(pdf) -> list:
    """Primary strategy: use pdfplumber's table detector, which copes with
    real ruled invoice tables far better than manual text-line regexes."""
    rows = []
    for page in pdf.pages:
        for table in page.extract_tables() or []:
            for raw_row in table:
                cells = [(_c or "").strip() for _c in raw_row]
                if len(cells) < 5:
                    continue
                # find a cell that looks like a part number and two
                # money-shaped cells to its right (unit price, total)
                money_idx = [i for i, c in enumerate(cells) if MONEY_RE.match(c)]
                if len(money_idx) < 2:
                    continue
                unit_idx, total_idx = money_idx[-2], money_idx[-1]
                part = None
                for c in cells[:unit_idx]:
                    if PART_NUMBER_RE.match(c):
                        part = c
                        break
                if not part:
                    continue
                # drop purely-numeric cells (Order, Item, Qty columns) so
                # the description text doesn't get polluted with them
                description = " ".join(
                    c for c in cells[:unit_idx]
                    if c and c != part and not re.match(r"^-?\d+(\.\d+)?$", c)
                ).strip()
                country = None
                m = re.search(r"\b([A-Z]{2})\b", description)
                if m:
                    country = m.group(1)
                rows.append({
                    "part_number": part,
                    "description": description,
                    "unit_price": _money(cells[unit_idx]),
                    "country": country or DEFAULT_COUNTRY,
                })
    return rows


def _extract_rows_from_text(pdf) -> list:
    """Fallback strategy when the PDF has no ruled table pdfplumber can
    detect: pair a part-number/price line with the description line that
    follows it."""
    rows = []
    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = [l.strip() for l in text.splitlines() if l.strip()]
        for i, line in enumerate(lines):
            tokens = line.split()
            money_tokens = [t for t in tokens if MONEY_RE.match(t)]
            part_tokens = [t for t in tokens if PART_NUMBER_RE.match(t)]
            if not part_tokens or len(money_tokens) < 2:
                continue
            part = part_tokens[0]
            unit_price = _money(money_tokens[-2])
            description = ""
            country = DEFAULT_COUNTRY
            if i + 1 < len(lines):
                nxt = lines[i + 1]
                cm = re.search(r"\b([A-Z]{2})\b", nxt)
                if cm:
                    country = cm.group(1)
                description = nxt
            rows.append({
                "part_number": part,
                "description": description,
                "unit_price": unit_price,
                "country": country,
            })
    return rows


def _extract_invoice(pdf_path: str) -> dict:
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)
        rows = _extract_rows_from_tables(pdf)
        if not rows:
            rows = _extract_rows_from_text(pdf)

    gt_match = GRAND_TOTAL_RE.search(full_text)
    if not gt_match:
        raise InvoiceError(
            f"{os.path.basename(pdf_path)}: לא נמצא 'Grand Total' בקובץ — "
            f"לא ניתן לאמת את סכום השורות שחולצו, ולכן לא ממשיכים."
        )
    grand_total = _money(gt_match.group(1))

    if not rows:
        raise InvoiceError(
            f"{os.path.basename(pdf_path)}: לא זוהתה אף שורת פריט. פורמט "
            f"החשבונית הזו כנראה שונה מהצפוי — יש לכייל את החילוץ מול קובץ "
            f"זה לפני שממשיכים."
        )

    extracted_sum = round(sum(r["unit_price"] for r in rows), 2)
    if abs(extracted_sum - grand_total) > 0.01:
        raise InvoiceError(
            f"{os.path.basename(pdf_path)}: בדיקת סכימה נכשלה — סכום השורות "
            f"שחולצו ({extracted_sum:.2f}) אינו תואם ל-Grand Total המוצהר "
            f"בקובץ ({grand_total:.2f}). ההפרש ({extracted_sum - grand_total:.2f}) "
            f"מעיד על שגיאת חילוץ — לא ממשיכים כדי לא לשלוח חשבונית שגויה. "
            f"יש לכייל את ביטויי הרגולרי מול הפורמט המדויק של קובץ זה."
        )

    shp_match = SHIPMENT_NUMBER_RE.search(full_text)
    hawb_match = HAWB_RE.search(full_text)
    gw_match = GROSS_WEIGHT_RE.search(full_text)

    return {
        "source_file": os.path.basename(pdf_path),
        "rows": rows,
        "grand_total": grand_total,
        "shipment_number": shp_match.group(1).strip() if shp_match else os.path.splitext(os.path.basename(pdf_path))[0],
        "hawb": hawb_match.group(1).strip() if hawb_match else "",
        "gross_weight": _money(gw_match.group(1)) if gw_match else None,
    }


def consolidate(pdf_paths):
    invoices = [_extract_invoice(p) for p in pdf_paths]

    by_part = defaultdict(lambda: {"qty": 0, "description": None, "prices": [], "countries": Counter()})
    for inv in invoices:
        for row in inv["rows"]:
            entry = by_part[row["part_number"]]
            entry["qty"] += 1
            entry["prices"].append(row["unit_price"])
            entry["countries"][row["country"] or DEFAULT_COUNTRY] += 1
            if entry["description"] is None:
                entry["description"] = row["description"]

    price_conflicts = []
    merged = []
    for part, entry in sorted(by_part.items()):
        prices = entry["prices"]
        unit_price = prices[0]
        if any(abs(p - unit_price) > 0.005 for p in prices):
            price_conflicts.append({"part_number": part, "prices": prices})
        country = entry["countries"].most_common(1)[0][0]
        merged.append({
            "part_number": part,
            "description": entry["description"] or "",
            "qty": entry["qty"],
            "unit_price": unit_price,
            "country": country,
        })

    total_expected = round(sum(inv["grand_total"] for inv in invoices), 2)
    total_merged = round(sum(m["qty"] * m["unit_price"] for m in merged), 2)
    if abs(total_merged - total_expected) > 0.02:
        raise InvoiceError(
            f"בדיקת סכימה כוללת נכשלה: סה\"כ מאוחד ({total_merged:.2f}) "
            f"אינו תואם לסכום ה-Grand Total מכל החשבוניות ({total_expected:.2f})."
        )

    return {
        "invoices": invoices,
        "merged": merged,
        "price_conflicts": price_conflicts,
        "total": total_merged,
    }


# --------------------------------------------------------------------- xlsx
ITEM_HEADER_ROW = 31
FIRST_ITEM_ROW = 32
TEMPLATE_LAST_ITEM_ROW = 51  # 20 pre-formatted rows: 32..51
TEMPLATE_TOTAL_ROW = 52
TEMPLATE_CERT_ROW = 54
TEMPLATE_CERT_RANGE = "B54:I54"

COL_PART = "C"
COL_DESC = "D"
COL_QTY = "E"
COL_UNIT_PRICE = "F"
COL_TOTAL = "G"
COL_COUNTRY = "H"


def build_invoice_excel(data: dict, template_path: str, work_dir: str,
                         reference_no: str, gross_weight: str, defective_return: str,
                         ship_date: str, dims: str) -> str:
    if not os.path.exists(template_path):
        raise InvoiceError(
            "תבנית ה-Excel הקבועה של KBB (COMMERCIAL INVOICE MANUAL RMA "
            "KBB.xlsx) לא נמצאה. יש להעתיק את קובץ התבנית האמיתי שלכם "
            f"לנתיב: {template_path} לפני הפעלת הכלי."
        )

    wb = openpyxl.load_workbook(template_path)
    ws = wb["Invoice"] if "Invoice" in wb.sheetnames else wb.active

    merged = data["merged"]
    n_items = len(merged)
    n_extra = max(0, n_items - (TEMPLATE_LAST_ITEM_ROW - FIRST_ITEM_ROW + 1))

    # clear the template's example row unless it's going to be overwritten
    # anyway (it always will be, since we rewrite every item row below).

    if n_extra > 0:
        ws.insert_rows(TEMPLATE_LAST_ITEM_ROW + 1, amount=n_extra)
        # openpyxl does not copy formatting into newly inserted rows —
        # copy every style attribute from the template's row 32.
        template_row = FIRST_ITEM_ROW
        for new_row in range(TEMPLATE_LAST_ITEM_ROW + 1, TEMPLATE_LAST_ITEM_ROW + 1 + n_extra):
            for col_letter in (COL_PART, COL_DESC, COL_QTY, COL_UNIT_PRICE, COL_TOTAL, COL_COUNTRY):
                src = ws[f"{col_letter}{template_row}"]
                dst = ws[f"{col_letter}{new_row}"]
                dst.font = src.font.copy()
                dst.border = src.border.copy()
                dst.fill = src.fill.copy()
                dst.number_format = src.number_format
                dst.alignment = src.alignment.copy()

        # openpyxl does not shift merged ranges on insert_rows — the
        # certification-note merge's row is now a brand-new blank row
        # (the real text moved down to cert_row, computed below), so the
        # stale "B54:I54" entry no longer matches real cell state at all.
        # ws.unmerge_cells() assumes the range's non-anchor cells are live
        # MergedCell placeholders it can delete, but the freshly-inserted
        # rows never got such placeholders — it raises KeyError. Dropping
        # the stale range directly (no cell cleanup needed, there's
        # nothing real there to clean up) is the version that actually
        # works.
        ws.merged_cells.remove(TEMPLATE_CERT_RANGE)

    last_item_row = TEMPLATE_LAST_ITEM_ROW + n_extra
    total_row = TEMPLATE_TOTAL_ROW + n_extra
    cert_row = TEMPLATE_CERT_ROW + n_extra

    for i, item in enumerate(merged):
        row = FIRST_ITEM_ROW + i
        ws[f"{COL_PART}{row}"] = item["part_number"]
        ws[f"{COL_DESC}{row}"] = item["description"]
        ws[f"{COL_QTY}{row}"] = item["qty"]
        ws[f"{COL_UNIT_PRICE}{row}"] = item["unit_price"]
        ws[f"{COL_TOTAL}{row}"] = f"=E{row}*F{row}"
        ws[f"{COL_COUNTRY}{row}"] = item["country"]

    # clear any leftover template rows beyond our data but still inside the
    # original 20-row block (e.g. we only had 5 merged items)
    for row in range(FIRST_ITEM_ROW + n_items, last_item_row + 1):
        for col_letter in (COL_PART, COL_DESC, COL_QTY, COL_UNIT_PRICE, COL_TOTAL, COL_COUNTRY):
            ws[f"{col_letter}{row}"] = None

    ws[f"{COL_TOTAL}{total_row}"] = f"=SUM({COL_TOTAL}{FIRST_ITEM_ROW}:{COL_TOTAL}{last_item_row})"

    if n_extra > 0:
        ws.merge_cells(f"B{cert_row}:I{cert_row}")
        # unmerge/rewrite resets style to default — re-copy from a sibling
        # item row so the certification note keeps its font/border/fill.
        cert_cell = ws[f"B{cert_row}"]
        template_style_src = ws[f"{COL_PART}{FIRST_ITEM_ROW}"]
        cert_cell.font = template_style_src.font.copy()

    # header fields
    ws["H3"] = reference_no
    ws["H3"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = max(ws.row_dimensions[3].height or 15, 30)

    ws["H5"] = ship_date
    ws["H15"] = gross_weight

    defective_lines = [l for l in re.split(r",\s*", defective_return) if l]
    half = (len(defective_lines) + 1) // 2
    ws["H22"] = ", ".join(defective_lines[:half])
    ws["H23"] = ", ".join(defective_lines[half:])
    for ref in ("H22", "H23"):
        ws[ref].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[22].height = max(ws.row_dimensions[22].height or 15, 30)
    ws.row_dimensions[23].height = max(ws.row_dimensions[23].height or 15, 30)

    # DIMS lives at D53 in the template; insert_rows() already shifted its
    # *position* down by n_extra along with every other row after the
    # insertion point (only styles/merges need manual fixing, not values),
    # so the row to write to is simply cert_row - 1.
    dims_row = cert_row - 1
    ws[f"D{dims_row}"] = dims

    # print setup — the template ships hard-coded to ~45% scale for the
    # original ~20-item layout; with more items that squashes everything
    # onto one illegible page.
    ws.page_setup.scale = None
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{ITEM_HEADER_ROW}:{ITEM_HEADER_ROW}"

    wb.calculation.fullCalcOnLoad = True

    out_path = os.path.join(work_dir, "KBB_Commercial_Invoice.xlsx")
    wb.save(out_path)
    return out_path


def process(pdf_paths, template_path, work_dir, reference_no=None, gross_weight=None,
            defective_return=None, ship_date=None, dims="- x - x - cm"):
    import datetime as dt
    os.makedirs(work_dir, exist_ok=True)
    data = consolidate(pdf_paths)

    if reference_no is None:
        reference_no = ", ".join(sorted({inv["shipment_number"] for inv in data["invoices"]}))
    if gross_weight is None:
        weights = [inv["gross_weight"] for inv in data["invoices"] if inv["gross_weight"] is not None]
        gross_weight = f"{sum(weights):.2f} KG" if weights else ""
    if defective_return is None:
        defective_return = ", ".join(inv["hawb"] for inv in data["invoices"] if inv["hawb"])
    if ship_date is None:
        ship_date = dt.date.today().strftime("%d/%m/%Y")

    xlsx_path = build_invoice_excel(
        data, template_path, work_dir,
        reference_no=reference_no, gross_weight=gross_weight,
        defective_return=defective_return, ship_date=ship_date, dims=dims,
    )
    xlsx_path = office.recalc_xlsx(xlsx_path, work_dir)

    return {
        "xlsx_path": xlsx_path,
        "unique_items": len(data["merged"]),
        "total_value": data["total"],
        "price_conflicts": data["price_conflicts"],
        "n_invoices": len(data["invoices"]),
        "reference_no": reference_no,
        "gross_weight": gross_weight,
        "defective_return": defective_return,
    }


def export_pdf(xlsx_path: str, work_dir: str) -> str:
    os.makedirs(work_dir, exist_ok=True)
    return office.to_pdf(xlsx_path, work_dir)
