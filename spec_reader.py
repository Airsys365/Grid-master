"""
spec_reader.py — чтение спецификации заказчика из Word/Excel/PDF через Claude AI.

Возвращает список позиций с полем quantity из документа заказчика.
"""

import json
import re
from typing import List, Dict, Optional
from pathlib import Path

try:
    import anthropic
    _HAS_ANTHROPIC = True
except ImportError:
    _HAS_ANTHROPIC = False

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


_SPEC_SYSTEM = """Ты — инженер, читающий спецификации металлоконструкций.
Твоя задача: извлечь из таблицы все позиции изделий с их количеством.

Правила:
- position — обозначение позиции, обычно формат K.XX.XX.XX.XXX или аналогичный
- quantity — количество штук (целое число, обычно 1–50). Не путай с весом, номером строки, размером
- weight_kg — масса одной единицы в кг (дробное число), если указана
- mat_article — артикул решётки если указан (например "P 33x11/30x3")
- Отвечай ТОЛЬКО валидным JSON-массивом. Никакого текста до или после.
"""

_SPEC_PROMPT = """Извлеки все позиции из этой спецификации.

Верни JSON-массив:
[
  {
    "position": "K.09.57.01.101",
    "quantity": 2,
    "weight_kg": 14.75,
    "mat_article": "P 33x11/30x3"
  }
]

Спецификация:
"""


# ---------------------------------------------------------------------------
# Главная функция
# ---------------------------------------------------------------------------

def read_spec(path: str, api_key: str = "") -> List[Dict]:
    """
    Читает спецификацию заказчика из файла через Claude AI.
    Возвращает список: [{position, quantity, weight_kg, mat_article, source}]
    """
    ext = Path(path).suffix.lower()
    if ext in (".docx", ".doc"):
        text = _extract_word_text(path)
    elif ext in (".xlsx", ".xls"):
        text = _extract_excel_text(path)
    elif ext == ".pdf":
        text = _extract_pdf_text(path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {ext}")

    if not text.strip():
        return []

    parts = _ask_claude(text, api_key)
    for p in parts:
        p["source"] = Path(path).name
    return parts


# ---------------------------------------------------------------------------
# Извлечение текста из файлов
# ---------------------------------------------------------------------------

def _extract_word_text(path: str) -> str:
    if not _HAS_DOCX:
        raise RuntimeError("Установите: pip install python-docx")
    doc = _docx.Document(path)
    lines = []
    for table in doc.tables:
        for row in table.rows:
            cells = [c.text.strip() for c in row.cells]
            lines.append(" | ".join(cells))
    return "\n".join(lines)


def _extract_excel_text(path: str) -> str:
    if not _HAS_PANDAS:
        raise RuntimeError("Установите: pip install pandas openpyxl")
    lines = []
    xf = pd.ExcelFile(path)
    for sheet in xf.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str).fillna("")
        for _, row in df.iterrows():
            cells = [str(c).strip() for c in row.tolist()]
            lines.append(" | ".join(cells))
    return "\n".join(lines)


def _extract_pdf_text(path: str) -> str:
    if not _HAS_FITZ:
        raise RuntimeError("Установите: pip install pymupdf")
    doc = fitz.open(path)
    return "\n".join(page.get_text() for page in doc)


# ---------------------------------------------------------------------------
# Claude AI
# ---------------------------------------------------------------------------

def _ask_claude(text: str, api_key: str) -> List[Dict]:
    if not _HAS_ANTHROPIC:
        raise RuntimeError("Установите: pip install anthropic")
    if not api_key:
        raise RuntimeError("API ключ не указан")

    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        system=_SPEC_SYSTEM,
        messages=[{"role": "user", "content": _SPEC_PROMPT + text}],
    )
    return _parse_response(response.content[0].text)


def _parse_response(text: str) -> List[Dict]:
    match = re.search(r'\[.*\]', text, re.DOTALL)
    if not match:
        return []
    try:
        data = json.loads(match.group())
        if not isinstance(data, list):
            return []
        return [_normalize(p) for p in data if isinstance(p, dict)]
    except json.JSONDecodeError:
        return []


def _normalize(p: Dict) -> Dict:
    qty = _int_or_default(p.get("quantity"), 1)
    if qty < 1 or qty > 99:
        qty = 1
    return {
        "position":    str(p.get("position") or "").upper().strip(),
        "quantity":    qty,
        "weight_kg":   _float_or_none(p.get("weight_kg")),
        "mat_article": p.get("mat_article") or None,
        "source":      "",
    }


def _int_or_default(v, default: int) -> int:
    try:
        return int(float(str(v).replace(",", ".")))
    except Exception:
        return default


def _float_or_none(v) -> Optional[float]:
    try:
        return round(float(str(v).replace(",", ".")), 3)
    except Exception:
        return None


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
        key = sp["position"].upper().strip()
        spec_by_pos[key] = sp

    result = []
    matched_keys = set()

    for ai in ai_parts:
        key = ai.get("position", "").upper().strip()
        sp = spec_by_pos.get(key)
        merged = dict(ai)

        if sp:
            matched_keys.add(key)
            diffs = []
            ai_qty = int(ai.get("quantity", 1))
            sp_qty = int(sp.get("quantity", 1))
            if ai_qty != sp_qty:
                diffs.append(f"кол-во: ИИ={ai_qty}, заказчик={sp_qty}")
                merged["quantity_conflict"] = {"ai": ai_qty, "spec": sp_qty}
            merged["quantity_spec"]  = sp_qty
            merged["weight_kg_spec"] = sp.get("weight_kg")
            merged["diff"] = diffs
            merged["in_spec"] = True
        else:
            merged["diff"] = ["нет в спецификации заказчика"]
            merged["in_spec"] = False

        result.append(merged)

    for key, sp in spec_by_pos.items():
        if key not in matched_keys:
            result.append({
                "position":       sp["position"],
                "length":         0,
                "width":          0,
                "quantity":       sp["quantity"],
                "notes":          "",
                "confidence":     "low",
                "diff":           ["нет в чертежах ИИ"],
                "in_spec":        True,
                "only_in_spec":   True,
                "mat_article":    sp.get("mat_article"),
                "weight_kg_spec": sp.get("weight_kg"),
            })

    return result


def apply_spec_quantities(parts: List[Dict]) -> List[Dict]:
    """Применяет количества из спецификации (quantity_spec) к позициям."""
    for p in parts:
        if "quantity_spec" in p:
            p["quantity"] = p["quantity_spec"]
    return parts
