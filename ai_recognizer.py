"""
ai_recognizer.py — распознавание PDF-чертежей через Claude API.

Возвращает список позиций (Part-совместимых dict) с полем confidence.
"""

import base64
import json
import re
from pathlib import Path
from typing import List, Dict

try:
    import anthropic
    _HAS_ANTHROPIC = True
except ImportError:
    _HAS_ANTHROPIC = False

try:
    import fitz  # PyMuPDF
    _HAS_FITZ = True
except ImportError:
    _HAS_FITZ = False


SYSTEM_PROMPT = """Ты — инженер-технолог на производстве металлических решётчатых настилов и грязезащитных решёток.
Твоя задача: извлечь из чертежей все позиции изделий и вернуть их в виде JSON.

Правила:
- Извлекай ТОЛЬКО то, что явно указано. Если поле неясно — ставь null и confidence="low".
- Размеры всегда в мм. Если на чертеже другие единицы — конвертируй.
- position — номер/обозначение позиции с чертежа (например "K.09.57.01.101" или "Porirest 1").
- length — большая параллельная сторона вдоль полос мата, мм. width — сторона поперёк полос, мм.
- length2 — вторая параллельная сторона если деталь трапеция (две разные длины вдоль полос). Если прямоугольник — null.
- quantity — количество штук (если не указано — 1). Количество ВСЕГДА от 1 до 99. Если видишь число больше 99 — это не количество (скорее всего размер, номер позиции или код), ставь quantity=1.
- mat_article — артикул решётки если явно указан (например "P 33x11/30x3"), иначе null.
- coating — покрытие: "Zn" (цинк), "RST" (нержавейка), null если неясно.
- frame_thickness_mm — толщина листа обрамления в мм если указана.
- cutout_radius_mm — радиус кругового или секторного выреза в мм если явно указан на чертеже (например "R150", "⌀300" → radius=150). Если выреза нет — null.
- cutout_arc_fraction — доля окружности выреза: 1.0 = полный круг, 0.5 = полукруг, 0.25 = четверть. Если вырез круглый целый — 1.0. Если выреза нет — null.
- notes — любые важные замечания (зеркальное отражение, особые условия, неясные вырезы без размеров и т.д.).
- confidence — "high" если всё чётко, "medium" если есть сомнения, "low" если угадываешь.

Отвечай ТОЛЬКО валидным JSON — массивом объектов. Никакого текста до или после JSON.
"""

EXTRACTION_PROMPT = """Проанализируй этот чертёж и извлеки все позиции изделий.

Верни JSON-массив в точно таком формате:
[
  {
    "position": "K.09.57.01.101",
    "length": 1300,
    "length2": 800,
    "width": 680,
    "quantity": 1,
    "mat_article": "P 33x11/30x3",
    "coating": "Zn",
    "frame_thickness_mm": 3,
    "cutout_radius_mm": 150,
    "cutout_arc_fraction": 0.25,
    "notes": "вырез четверть круга R150 в углу; зеркальное отражение — K.09.57.01.101-01",
    "confidence": "high"
  }
]
"""


# ---------------------------------------------------------------------------
# Основная функция
# ---------------------------------------------------------------------------

def recognize_pdf(pdf_path: str, api_key: str) -> List[Dict]:
    """
    Распознаёт PDF чертёж через Claude Vision.
    Возвращает список позиций или пустой список при ошибке.
    """
    if not _HAS_ANTHROPIC:
        raise RuntimeError("Установите пакет: pip install anthropic")
    if not api_key:
        raise RuntimeError("API ключ не указан. Введите ключ Anthropic на главном экране.")

    pages = _pdf_to_images(pdf_path)
    if not pages:
        if not _HAS_FITZ:
            raise RuntimeError("Установите PyMuPDF: pip install pymupdf")
        # Fallback: попробуем как текст через PyMuPDF
        return _recognize_text_pdf(pdf_path, api_key)

    client = anthropic.Anthropic(api_key=api_key)
    all_parts = []

    for i, img_b64 in enumerate(pages):
        content = [
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": img_b64,
                },
            },
            {
                "type": "text",
                "text": EXTRACTION_PROMPT,
            },
        ]

        response = client.messages.create(
            model="claude-opus-4-8",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=[{"role": "user", "content": content}],
        )

        parts = _parse_response(response.content[0].text)
        for p in parts:
            p["source_page"] = i + 1
            p["source_file"] = Path(pdf_path).name
        all_parts.extend(parts)

    return _deduplicate(all_parts)


def recognize_text(text: str, filename: str, api_key: str) -> List[Dict]:
    """Распознаёт уже извлечённый текст (для PDF без изображений)."""
    if not _HAS_ANTHROPIC:
        raise RuntimeError("Установите пакет: pip install anthropic")

    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{
            "role": "user",
            "content": f"Текст извлечён из файла {filename}:\n\n{text}\n\n{EXTRACTION_PROMPT}"
        }],
    )
    parts = _parse_response(response.content[0].text)
    for p in parts:
        p["source_file"] = filename
    return parts


# ---------------------------------------------------------------------------
# Вспомогательные
# ---------------------------------------------------------------------------

def _pdf_to_images(pdf_path: str) -> List[str]:
    """Конвертирует страницы PDF в base64-PNG. Возвращает [] если нет PyMuPDF."""
    if not _HAS_FITZ:
        return []
    try:
        doc = fitz.open(pdf_path)
        images = []
        for page in doc:
            pix = page.get_pixmap(dpi=150)
            png_bytes = pix.tobytes("png")
            images.append(base64.standard_b64encode(png_bytes).decode())
        return images
    except Exception:
        return []


def _recognize_text_pdf(pdf_path: str, api_key: str) -> List[Dict]:
    """Fallback: извлекаем текст и отправляем как текст."""
    if not _HAS_FITZ:
        return []
    try:
        doc = fitz.open(pdf_path)
        text = "\n".join(page.get_text() for page in doc)
        return recognize_text(text, Path(pdf_path).name, api_key)
    except Exception:
        return []


def _parse_response(text: str) -> List[Dict]:
    """Извлекает JSON из ответа Claude."""
    # Ищем JSON-массив в тексте
    match = re.search(r'\[.*\]', text, re.DOTALL)
    if not match:
        return []
    try:
        data = json.loads(match.group())
        if not isinstance(data, list):
            return []
        return [_normalize_part(p) for p in data if isinstance(p, dict)]
    except json.JSONDecodeError:
        return []


def _normalize_part(p: Dict) -> Dict:
    """Приводит поля к стандартным типам."""
    # Радиус: ⌀ означает диаметр → делим на 2
    raw_r = p.get("cutout_radius_mm")
    radius = _int_or_zero(raw_r) if raw_r is not None else 0

    raw_arc = p.get("cutout_arc_fraction")
    try:
        arc = float(raw_arc) if raw_arc is not None else 1.0
        arc = max(0.0, min(1.0, arc))
    except Exception:
        arc = 1.0

    return {
        "position":           str(p.get("position") or ""),
        "length":             _int_or_zero(p.get("length")),
        "length2":            _int_or_zero(p.get("length2")) if p.get("length2") else 0,
        "width":              _int_or_zero(p.get("width")),
        "quantity":           _safe_quantity(p.get("quantity")),
        "mat_article":        p.get("mat_article"),
        "coating":            p.get("coating"),
        "frame_thickness_mm": _int_or_zero(p.get("frame_thickness_mm")),
        "radius":             radius,
        "radius_part":        arc if radius > 0 else 1.0,
        "notes":              str(p.get("notes") or ""),
        "confidence":         str(p.get("confidence") or "medium"),
        "source_file":        str(p.get("source_file") or ""),
        "source_page":        p.get("source_page", 1),
    }


def _deduplicate(parts: List[Dict]) -> List[Dict]:
    """Убирает дубли по position+length+width если несколько страниц."""
    seen = set()
    result = []
    for p in parts:
        key = (p["position"], p["length"], p["width"])
        if key not in seen:
            seen.add(key)
            result.append(p)
    return result


def _safe_quantity(v) -> int:
    n = _int_or_zero(v)
    if n < 1 or n > 99:
        return 1
    return n


def _int_or_zero(v) -> int:
    try:
        return int(float(str(v).replace(",", ".")))
    except Exception:
        return 0
