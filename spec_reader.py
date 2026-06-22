"""
spec_reader.py — чтение спецификации заказчика из Word/Excel/PDF.

Возвращает список позиций с полем quantity из документа заказчика.
"""

import re
from typing import List, Dict, Optional
from pathlib import Path

try:
    import docx as _docx
    _HAS_DOCX = True
except ImportError:
    _HAS_DOCX = False

try:
    import pandas as pd
    _HAS_PANDAS = True
except ImportError:
    _HAS_PANDAS = False

try:
    import fitz
    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False


# ---------------------------------------------------------------------------
# Главная функция
# ---------------------------------------------------------------------------

def read_spec(path: str) -> List[Dict]:
    """
    Читает спецификацию заказчика из файла.
    Возвращает список: [{position, quantity, weight_kg, mat_article, source}]
    """
    ext = Path(path).suffix.lower()
    if ext in (".docx", ".doc"):
        return _read_word(path)
    elif ext in (".xlsx", ".xls"):
        return _read_excel(path)
    elif ext == ".pdf":
        return _read_pdf_spec(path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {ext}")


# ---------------------------------------------------------------------------
# Word (.docx)
# ---------------------------------------------------------------------------

def _read_word(path: str) -> List[Dict]:
    if not _HAS_DOCX:
        raise RuntimeError("Установите: pip install python-docx")

    doc = _docx.Document(path)
    result = []

    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            if len(cells) < 5:
                continue

            # Ищем позицию — колонка с обозначением типа K.XX.XX.XX.XXX
            pos, pos_idx = _find_position_with_idx(cells)
            if not pos:
                continue

            # Берём только ячейки ПОСЛЕ колонки с позицией
            after = cells[pos_idx + 1:]
            qty = _find_quantity(after)
            weight = _find_weight(after)
            article = _find_article(after)

            result.append({
                "position":    pos,
                "quantity":    qty,
                "weight_kg":   weight,
                "mat_article": article,
                "source":      Path(path).name,
            })

    return result


# ---------------------------------------------------------------------------
# Excel (.xlsx)
# ---------------------------------------------------------------------------

def _read_excel(path: str) -> List[Dict]:
    if not _HAS_PANDAS:
        raise RuntimeError("Установите: pip install pandas openpyxl")

    result = []
    xf = pd.ExcelFile(path)
    for sheet in xf.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str).fillna("")
        for _, row in df.iterrows():
            cells = [str(c).strip() for c in row.tolist()]
            pos = _find_position(cells)
            if not pos:
                continue
            result.append({
                "position":    pos,
                "quantity":    _find_quantity(cells),
                "weight_kg":   _find_weight(cells),
                "mat_article": _find_article(cells),
                "source":      Path(path).name,
            })
    return result


# ---------------------------------------------------------------------------
# PDF спецификация (текстовый слой)
# ---------------------------------------------------------------------------

def _read_pdf_spec(path: str) -> List[Dict]:
    if not _HAS_FITZ:
        raise RuntimeError("Установите: pip install pymupdf")

    doc = fitz.open(path)
    result = []
    for page in doc:
        text = page.get_text()
        for line in text.splitlines():
            parts = line.split()
            pos = _find_position(parts)
            if pos:
                result.append({
                    "position":    pos,
                    "quantity":    _find_quantity(parts),
                    "weight_kg":   _find_weight(parts),
                    "mat_article": _find_article(parts),
                    "source":      Path(path).name,
                })
    return result


# ---------------------------------------------------------------------------
# Сравнение с AI-распознанными позициями
# ---------------------------------------------------------------------------

def compare(ai_parts: List[Dict], spec_parts: List[Dict]) -> List[Dict]:
    """
    Сравнивает список от ИИ со спецификацией заказчика.
    Возвращает merged-список с полем 'diff'.
    """
    spec_by_pos = {}
    for sp in spec_parts:
        key = _normalize_pos(sp["position"])
        spec_by_pos[key] = sp

    result = []
    matched_keys = set()

    for ai in ai_parts:
        key = _normalize_pos(ai.get("position", ""))
        sp = spec_by_pos.get(key)
        merged = dict(ai)

        if sp:
            matched_keys.add(key)
            diffs = []
            # Количество
            ai_qty = int(ai.get("quantity", 1))
            sp_qty = int(sp.get("quantity", 1))
            if ai_qty != sp_qty:
                diffs.append(f"кол-во: ИИ={ai_qty}, заказчик={sp_qty}")
                merged["quantity_conflict"] = {"ai": ai_qty, "spec": sp_qty}
            # Предлагаем взять количество из спеки (оно точнее)
            merged["quantity_spec"]  = sp_qty
            merged["weight_kg_spec"] = sp.get("weight_kg")
            merged["diff"] = diffs
            merged["in_spec"] = True
        else:
            merged["diff"] = ["нет в спецификации заказчика"]
            merged["in_spec"] = False

        result.append(merged)

    # Позиции из спеки которых нет у ИИ
    for key, sp in spec_by_pos.items():
        if key not in matched_keys:
            entry = {
                "position":      sp["position"],
                "length":        0,
                "width":         0,
                "quantity":      sp["quantity"],
                "notes":         "",
                "confidence":    "low",
                "diff":          ["нет в чертежах ИИ"],
                "in_spec":       True,
                "only_in_spec":  True,
                "mat_article":   sp.get("mat_article"),
                "weight_kg_spec": sp.get("weight_kg"),
            }
            result.append(entry)

    return result


def apply_spec_quantities(parts: List[Dict]) -> List[Dict]:
    """Применяет количества из спецификации (quantity_spec) к позициям."""
    for p in parts:
        if "quantity_spec" in p:
            p["quantity"] = p["quantity_spec"]
    return parts


# ---------------------------------------------------------------------------
# Парсеры полей
# ---------------------------------------------------------------------------

_POS_RE = re.compile(r'[A-Z]\.\d{2}\.\d{2}\.\d{2}\.\d{3}(?:-\d+)?', re.IGNORECASE)

def _find_position(cells: List[str]) -> Optional[str]:
    pos, _ = _find_position_with_idx(cells)
    return pos

def _find_position_with_idx(cells: List[str]):
    for i, c in enumerate(cells):
        m = _POS_RE.search(c)
        if m:
            return m.group().upper(), i
    return None, -1

def _normalize_pos(pos: str) -> str:
    return pos.upper().strip()

def _find_quantity(cells: List[str]) -> int:
    """Ищет количество (целое 1-99) — берём последнее подходящее число в строке."""
    candidates = []
    for c in cells:
        s = c.strip().replace(",", ".")
        # Пропускаем ячейки с обозначением позиции
        if _POS_RE.search(s):
            continue
        try:
            v = float(s)
            if v == int(v) and 1 <= int(v) <= 99 and "." not in s:
                candidates.append(int(v))
        except Exception:
            pass
    # Берём последнее найденное (в спеках количество обычно в конце)
    return candidates[-1] if candidates else 1

def _find_weight(cells: List[str]) -> Optional[float]:
    """Ищет дробное число (вес кг) — содержит запятую или точку."""
    for c in cells:
        s = c.strip().replace(",", ".")
        try:
            v = float(s)
            if v != int(v) and 0.1 <= v <= 9999:
                return round(v, 3)
        except Exception:
            pass
    return None

def _find_article(cells: List[str]) -> Optional[str]:
    for c in cells:
        if "33x11" in c or "P 33" in c or "DIN24537" in c or re.search(r'P\s+\d+x\d+', c):
            return c.strip()
    return None
