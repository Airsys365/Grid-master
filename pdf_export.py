"""
pdf_export.py — генерация PDF-оффера для клиента.

Показывает итоговую цену без расшифровки себестоимости.
"""

from datetime import datetime
from pathlib import Path
from typing import Dict, Optional

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib.colors import HexColor, black, white, grey
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    _HAS_REPORTLAB = True
except ImportError:
    _HAS_REPORTLAB = False


BRAND_GREEN  = HexColor("#2E8B57") if _HAS_REPORTLAB else None
BRAND_DARK   = HexColor("#1a1a2e") if _HAS_REPORTLAB else None
LIGHT_GREY   = HexColor("#f5f5f5") if _HAS_REPORTLAB else None


def export_offer_pdf(
    output_path: str,
    client_name: str,
    parts: list,           # список dict из ai_recognizer / ручного ввода
    total_cost: float,     # себестоимость
    margin_pct: float,     # % маржи
    offer_number: str = "",
    company_name: str = "Grid Master",
    company_contact: str = "",
    notes: str = "",
) -> str:
    """
    Генерирует PDF-оффер. Возвращает путь к файлу.
    Клиент видит только итоговую цену (без себестоимости).
    """
    if not _HAS_REPORTLAB:
        raise RuntimeError("Установите пакет: pip install reportlab")

    client_price = round(total_cost * (1 + margin_pct / 100), 2)
    date_str = datetime.now().strftime("%d.%m.%Y")

    if not offer_number:
        offer_number = datetime.now().strftime("OFF-%Y%m%d-%H%M")

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        rightMargin=20*mm,
        leftMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm,
    )

    styles = getSampleStyleSheet()
    style_normal = ParagraphStyle("normal", fontName="Helvetica", fontSize=10, leading=14)
    style_small  = ParagraphStyle("small",  fontName="Helvetica", fontSize=8,  leading=12, textColor=grey)
    style_h1     = ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=20, leading=24, textColor=BRAND_GREEN)
    style_h2     = ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, leading=18)
    style_right  = ParagraphStyle("right", fontName="Helvetica", fontSize=10, alignment=TA_RIGHT)
    style_price  = ParagraphStyle("price", fontName="Helvetica-Bold", fontSize=22, leading=28,
                                  textColor=BRAND_GREEN, alignment=TA_RIGHT)

    story = []

    # --- Шапка ---
    story.append(Paragraph(company_name, style_h1))
    if company_contact:
        story.append(Paragraph(company_contact, style_small))
    story.append(Spacer(1, 6*mm))
    story.append(HRFlowable(width="100%", thickness=1, color=BRAND_GREEN))
    story.append(Spacer(1, 4*mm))

    # --- Реквизиты ---
    meta = [
        ["Предложение №:", offer_number,  "Дата:", date_str],
        ["Клиент:",        client_name,   "",      ""],
    ]
    meta_table = Table(meta, colWidths=[35*mm, 70*mm, 20*mm, 45*mm])
    meta_table.setStyle(TableStyle([
        ("FONTNAME",  (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE",  (0, 0), (-1, -1), 10),
        ("FONTNAME",  (0, 0), (0, -1),  "Helvetica-Bold"),
        ("FONTNAME",  (2, 0), (2, -1),  "Helvetica-Bold"),
        ("VALIGN",    (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 6*mm))

    # --- Таблица позиций ---
    story.append(Paragraph("Состав заказа", style_h2))
    story.append(Spacer(1, 3*mm))

    table_data = [["№", "Позиция", "Размер (мм)", "Кол-во", "Примечание"]]
    for i, p in enumerate(parts, 1):
        pos   = str(p.get("position", ""))
        dims  = f"{p.get('length', '')} × {p.get('width', '')}"
        qty   = str(p.get("quantity", 1))
        note  = str(p.get("notes", ""))
        table_data.append([str(i), pos, dims, qty, note])

    col_widths = [12*mm, 55*mm, 40*mm, 18*mm, 45*mm]
    parts_table = Table(table_data, colWidths=col_widths, repeatRows=1)
    parts_table.setStyle(TableStyle([
        # Заголовок
        ("BACKGROUND",    (0, 0), (-1, 0),  BRAND_GREEN),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  white),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, 0),  9),
        ("ALIGN",         (0, 0), (-1, 0),  "CENTER"),
        # Данные
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

    # --- Итоговая цена ---
    story.append(HRFlowable(width="100%", thickness=0.5, color=grey))
    story.append(Spacer(1, 4*mm))

    price_data = [
        ["ИТОГОВАЯ ЦЕНА:", f"{client_price:,.2f} €".replace(",", " ")],
    ]
    price_table = Table(price_data, colWidths=[120*mm, 50*mm])
    price_table.setStyle(TableStyle([
        ("FONTNAME",      (0, 0), (0, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (0, 0),  13),
        ("FONTNAME",      (1, 0), (1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (1, 0), (1, 0),  18),
        ("TEXTCOLOR",     (1, 0), (1, 0),  BRAND_GREEN),
        ("ALIGN",         (1, 0), (1, 0),  "RIGHT"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
    ]))
    story.append(price_table)
    story.append(Spacer(1, 2*mm))
    story.append(Paragraph("Цена указана без НДС. Срок действия предложения — 30 дней.", style_small))

    # --- Примечания ---
    if notes:
        story.append(Spacer(1, 5*mm))
        story.append(Paragraph("Примечания:", style_h2))
        story.append(Paragraph(notes, style_normal))

    # --- Подпись ---
    story.append(Spacer(1, 10*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=grey))
    story.append(Spacer(1, 3*mm))
    story.append(Paragraph(f"Предложение подготовлено: {date_str} | {company_name}", style_small))

    doc.build(story)
    return output_path
