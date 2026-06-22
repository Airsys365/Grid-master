"""
pdf_export.py — генерация PDF-оффера для клиента (English / Eesti).
"""

from datetime import datetime
from typing import List, Dict

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.colors import HexColor, black, white, grey
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    _HAS_REPORTLAB = True
except ImportError:
    _HAS_REPORTLAB = False


BRAND_GREEN = HexColor("#2E8B57") if _HAS_REPORTLAB else None
LIGHT_GREY  = HexColor("#f5f5f5") if _HAS_REPORTLAB else None

# ---------------------------------------------------------------------------
# Переводы
# ---------------------------------------------------------------------------

_STRINGS = {
    "en": {
        "offer_no":    "Offer No.:",
        "date":        "Date:",
        "client":      "Client:",
        "order_items": "Order Items",
        "col_no":      "#",
        "col_pos":     "Position",
        "col_size":    "Size (mm)",
        "col_qty":     "Qty",
        "col_note":    "Notes",
        "total_price": "TOTAL PRICE:",
        "vat_note":    "Price excl. VAT. Offer valid for 30 days.",
        "prepared":    "Offer prepared:",
    },
    "et": {
        "offer_no":    "Pakkumine nr:",
        "date":        "Kuupäev:",
        "client":      "Klient:",
        "order_items": "Tellimuse koosseis",
        "col_no":      "#",
        "col_pos":     "Positsioon",
        "col_size":    "Suurus (mm)",
        "col_qty":     "Kogus",
        "col_note":    "Märkused",
        "total_price": "KOGUMAKSUMUS:",
        "vat_note":    "Hind ilma käibemaksuta. Pakkumine kehtib 30 päeva.",
        "prepared":    "Pakkumine koostatud:",
    },
}


# ---------------------------------------------------------------------------
# Главная функция
# ---------------------------------------------------------------------------

def export_offer_pdf(
    output_path: str,
    client_name: str,
    parts: list,
    total_cost: float,
    margin_pct: float,
    offer_number: str = "",
    company_name: str = "Grid Master",
    company_contact: str = "",
    notes: str = "",
    language: str = "en",       # "en" | "et"
) -> str:
    if not _HAS_REPORTLAB:
        raise RuntimeError("Install: pip install reportlab")

    t = _STRINGS.get(language, _STRINGS["en"])
    client_price = round(total_cost * (1 + margin_pct / 100), 2)
    date_str = datetime.now().strftime("%d.%m.%Y")
    if not offer_number:
        offer_number = datetime.now().strftime("OFF-%Y%m%d-%H%M")

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20*mm, leftMargin=20*mm,
        topMargin=20*mm,   bottomMargin=20*mm,
    )

    style_normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=10, leading=14)
    style_small  = ParagraphStyle("small",  fontName="Helvetica", fontSize=8,  leading=12, textColor=grey)
    style_h1     = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=BRAND_GREEN)
    style_h2     = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, leading=18)

    story = []

    # Шапка
    story.append(Paragraph(company_name, style_h1))
    if company_contact:
        story.append(Paragraph(company_contact, style_small))
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=BRAND_GREEN))
    story.append(Spacer(1, 4*mm))

    # Реквизиты
    meta = [
        [t["offer_no"], offer_number, t["date"], date_str],
        [t["client"],   client_name,  "",        ""],
    ]
    meta_table = Table(meta, colWidths=[35*mm, 70*mm, 20*mm, 45*mm])
    meta_table.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 0), (-1, -1), 10),
        ("FONTNAME",      (0, 0), (0, -1),  "Helvetica-Bold"),
        ("FONTNAME",      (2, 0), (2, -1),  "Helvetica-Bold"),
        ("VALIGN",        (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 6*mm))

    # Таблица позиций
    story.append(Paragraph(t["order_items"], style_h2))
    story.append(Spacer(1, 3*mm))

    table_data = [[t["col_no"], t["col_pos"], t["col_size"], t["col_qty"], t["col_note"]]]
    for i, p in enumerate(parts, 1):
        l2 = p.get("length2", 0) or 0
        size = (f"{p.get('length','')} / {l2} x {p.get('width','')}"
                if l2 else f"{p.get('length','')} x {p.get('width','')}")
        note = str(p.get("notes", ""))
        # Strip Cyrillic from notes for PDF safety
        note = "".join(c for c in note if ord(c) < 128)
        table_data.append([str(i), str(p.get("position", "")), size,
                           str(p.get("quantity", 1)), note])

    parts_table = Table(table_data, colWidths=[12*mm, 55*mm, 40*mm, 18*mm, 45*mm], repeatRows=1)
    parts_table.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, 0),  BRAND_GREEN),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  white),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  9),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        ("FONTNAME",      (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, 1), (-1, -1), 9),
        ("ALIGN",         (3, 1), (3, -1),  "CENTER"),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [white, LIGHT_GREY]),
        ("GRID",          (0, 0), (-1, -1), 0.3, grey),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
    ]))
    story.append(parts_table)
    story.append(Spacer(1, 8*mm))

    # Итоговая цена
    story.append(HRFlowable(width="100%", thickness=0.5, color=grey))
    story.append(Spacer(1, 4*mm))

    price_table = Table(
        [[t["total_price"], f"{client_price:,.2f} EUR".replace(",", " ")]],
        colWidths=[120*mm, 50*mm]
    )
    price_table.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (0, 0),   "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (0, 0),   13),
        ("FONTNAME",      (1, 0), (1, 0),   "Helvetica-Bold"),
        ("FONTSIZE",      (1, 0), (1, 0),   18),
        ("TEXTCOLOR",     (1, 0), (1, 0),   BRAND_GREEN),
        ("ALIGN",         (1, 0), (1, 0),   "RIGHT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
    ]))
    story.append(price_table)
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph(t["vat_note"], style_small))

    if notes:
        safe_notes = "".join(c for c in notes if ord(c) < 128)
        if safe_notes.strip():
            story.append(Spacer(1, 5*mm))
            story.append(Paragraph(safe_notes, style_normal))

    story.append(Spacer(1, 10*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=grey))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"{t['prepared']} {date_str} | {company_name}", style_small))

    doc.build(story)
    return output_path
