"""
gridmaster_core.py — чистый модуль расчёта себестоимости решётчатых настилов.

Не импортирует tkinter и не обращается к UI. Принимает Order, возвращает OrderResult.
Формулы идентичны Grid_Master_V50.py — только отцеплены от виджетов.
"""

import math
from dataclasses import dataclass, field
from typing import List, Tuple, Dict, Optional


# ---------------------------------------------------------------------------
# Вспомогательные функции
# ---------------------------------------------------------------------------

def parse_num(value, default: float = 0.0) -> float:
    try:
        return float(str(value).strip().replace(",", "."))
    except Exception:
        return float(default)


def weight_per_m(table: List[Tuple[str, float]], type_name: str) -> float:
    """Вес кг/м для типа из справочника весов."""
    for t, w in table:
        if str(t) == str(type_name):
            return float(w)
    return 0.0


# ---------------------------------------------------------------------------
# Структуры данных
# ---------------------------------------------------------------------------

@dataclass
class Part:
    """Одна деталь/позиция заказа."""
    position: str
    length: float           # мм, большая параллельная сторона (вдоль полос)
    width: float            # мм, связующая полоса (поперёк)
    quantity: int
    angle: str              # длина перфоуголка, мм (или "0"/"")
    bumper: float           # длина прямого кикплейта, мм
    work: float             # нормо-часы
    radius: float = 0.0
    radius_part: float = 1.0
    subtract_bumper: bool = True
    bumper_length: float = 0.0
    length2: float = 0.0    # мм, вторая параллельная сторона (трапеция; 0 = прямоугольник)


@dataclass
class MatSpec:
    """Артикул мата и его параметры."""
    article: str
    price_per_m2: float     # €/м² — из таблицы или введено вручную
    weight_per_m2: float    # кг/м²


@dataclass
class Order:
    """Полные входные данные для расчёта одного заказа."""
    parts: List[Part]

    # Мат
    mat: MatSpec
    mat_length_mm: int = 6100   # размер стандартного листа
    mat_width_mm: int = 1000
    manual_mats_count: float = 0.0  # 0 = авторасчёт по деталям
    manual_mats_k: bool = False     # True = без +15%, False = +15%

    # Обрамление
    frame_type: str = ""
    frame_price: float = 0.0       # €/кг
    frame_coef: float = 1.0
    frame_weights: List[Tuple[str, float]] = field(default_factory=list)

    # Перфоуголок
    angle_type: str = ""
    angle_price: float = 0.0       # €/шт (1 шт = 6 м)
    angle_weights: List[Tuple[str, float]] = field(default_factory=list)

    # Отбойник
    bumper_type: str = ""
    bumper_price: float = 0.0      # €/кг
    bumper_weights: List[Tuple[str, float]] = field(default_factory=list)

    # Работа
    work_price: float = 0.0        # €/час
    work_coef: float = 1.0

    # Цинкование
    zinc_price: float = 0.0        # €/кг


@dataclass
class OrderResult:
    """Результат расчёта."""
    mat_cost: float
    frame_cost: float
    angle_cost: float
    bumper_cost: float
    work_cost: float
    zinc_cost: float
    total_cost: float

    mats_count: float           # кол-во листов (без коэфф.)
    mats_count_with_k: float    # с коэфф. +15% или 1.0

    frame_length_m: float
    angle_length_m: float
    angle_piece_count: float
    bumper_length_m: float

    weight_mats: float
    weight_frame: float
    weight_angle: float
    weight_bumper: float
    total_weight: float
    total_weight_with_zinc: float

    cost_per_kg: float


# ---------------------------------------------------------------------------
# Расчётные функции
# ---------------------------------------------------------------------------

def _bumper_lengths(bumper_mm: float, radius: float, radius_part: float = 1.0) -> Tuple[float, float]:
    """(прямая_часть_мм, итого_мм)"""
    bumper_mm = parse_num(bumper_mm)
    radius = parse_num(radius)
    radius_part = parse_num(radius_part) or 1.0
    arc = 2 * math.pi * radius * radius_part if radius > 0 else 0.0
    return bumper_mm, bumper_mm + arc


def _frame_length_m(parts: List[Part]) -> float:
    total = 0.0
    for p in parts:
        linear_mm, total_bumper_mm = _bumper_lengths(p.bumper, p.radius, p.radius_part)
        if p.length2 > 0:
            # Трапеция: три стороны = L1 + L2 + W (четвёртая — скошенная, считается вручную)
            per_mm = p.length + p.length2 + p.width
        else:
            per_mm = (2 * p.width + p.length) if p.width < 1000 else (2 * p.width)
        # Вычитаем прямую часть кикплейта если указано
        if p.bumper and p.subtract_bumper:
            per_mm -= linear_mm
        # Дуга радиуса всегда уходит в кикплейт, не в обрамление
        arc_mm = total_bumper_mm - linear_mm
        if arc_mm > 0:
            per_mm -= arc_mm
        total += max(0.0, round(per_mm / 1000.0, 2)) * p.quantity
    return round(total, 2)


def _angle_length_m(parts: List[Part]) -> float:
    total = 0.0
    for p in parts:
        v = str(p.angle).strip()
        if not v or v == "0":
            continue
        try:
            total += (float(v) / 1000.0) * p.quantity
        except ValueError:
            pass
    return round(total, 2)


def _bumper_length_m(parts: List[Part]) -> float:
    """Суммарная длина кикплейта: прямая часть + дуги радиусов."""
    total = 0.0
    for p in parts:
        _, total_mm = _bumper_lengths(p.bumper, p.radius, p.radius_part)
        total += (total_mm / 1000.0) * p.quantity
    return round(total, 2)


def _rect_area_m2(p: Part) -> float:
    """Площадь по прямоугольнику для стоимости мата (скосы и вырезы — в отход)."""
    return (p.length * p.width) / 1_000_000.0


def _net_area_m2(p: Part) -> float:
    """Реальная площадь для веса: трапеция минус вырез радиуса."""
    if p.length2 > 0:
        base = ((p.length + p.length2) / 2.0 * p.width) / 1_000_000.0
    else:
        base = _rect_area_m2(p)
    cutout = (math.pi * p.radius ** 2 * p.radius_part) / 1_000_000.0 if p.radius > 0 else 0.0
    return max(0.0, base - cutout)


def _mats_count_auto(parts: List[Part], mat_length_mm: int, mat_width_mm: int) -> float:
    """Авторасчёт количества листов по деталям (без +15%). Трапеция — по прямоугольнику."""
    if mat_length_mm <= 0 or mat_width_mm <= 0:
        return 0.0
    total_area = sum(_rect_area_m2(p) * p.quantity for p in parts)
    sheet_area = (mat_length_mm * mat_width_mm) / 1_000_000.0
    return round(total_area / sheet_area, 1) if sheet_area > 0 else 0.0


def calculate(order: Order) -> OrderResult:
    """Главная функция расчёта. Принимает Order, возвращает OrderResult."""

    parts = order.parts

    # --- Кол-во матов ---
    if order.manual_mats_count > 0:
        mats_count = round(order.manual_mats_count, 1)
    else:
        mats_count = _mats_count_auto(parts, order.mat_length_mm, order.mat_width_mm)

    k = 1.0 if order.manual_mats_k else 1.15
    mats_count_with_k = round(mats_count * k, 3)

    # --- Себестоимость матов ---
    L, W = order.mat_length_mm, order.mat_width_mm
    area_m2 = (L * W) / 1_000_000.0 if L > 0 and W > 0 else 0.0
    mat_cost = round(mats_count_with_k * area_m2 * order.mat.price_per_m2, 2)

    # --- Обрамление ---
    frame_length_m = _frame_length_m(parts)
    wpm_frame = weight_per_m(order.frame_weights, order.frame_type)
    weight_frame = round(frame_length_m * wpm_frame * order.frame_coef, 2)
    frame_cost = round(weight_frame * order.frame_price, 2)

    # --- Перфоуголок ---
    angle_length_m = _angle_length_m(parts)
    angle_piece_count = math.ceil((angle_length_m / 6.0) * 10) / 10.0
    wpm_angle = weight_per_m(order.angle_weights, order.angle_type)
    weight_angle = round(angle_length_m * wpm_angle, 2)
    angle_cost = round(angle_piece_count * order.angle_price, 2)

    # --- Отбойник ---
    has_bumper = any(parse_num(p.bumper) > 0 for p in parts)
    bumper_length_m = _bumper_length_m(parts) if has_bumper else 0.0
    wpm_bumper = weight_per_m(order.bumper_weights, order.bumper_type)
    weight_bumper = round(bumper_length_m * wpm_bumper, 2) if has_bumper else 0.0
    bumper_cost = round(bumper_length_m * wpm_bumper * order.bumper_price, 2) if has_bumper else 0.0

    # --- Работа ---
    hours = round(sum(p.work * p.quantity for p in parts), 2)
    work_cost = round(hours * order.work_price * order.work_coef, 2)

    # --- Веса ---
    total_area_parts = sum(_net_area_m2(p) * p.quantity for p in parts)
    weight_mats = round(total_area_parts * order.mat.weight_per_m2, 2)
    total_weight = round(weight_mats + weight_frame + weight_angle + weight_bumper, 2)
    total_weight_with_zinc = round(total_weight * 1.1, 2)

    # --- Цинкование ---
    zinc_cost = round(order.zinc_price * total_weight_with_zinc, 2)

    total_cost = round(mat_cost + frame_cost + angle_cost + bumper_cost + work_cost + zinc_cost, 2)
    cost_per_kg = round(total_cost / total_weight, 2) if total_weight > 0 else 0.0

    return OrderResult(
        mat_cost=mat_cost,
        frame_cost=frame_cost,
        angle_cost=angle_cost,
        bumper_cost=bumper_cost,
        work_cost=work_cost,
        zinc_cost=zinc_cost,
        total_cost=total_cost,
        mats_count=mats_count,
        mats_count_with_k=mats_count_with_k,
        frame_length_m=frame_length_m,
        angle_length_m=angle_length_m,
        angle_piece_count=angle_piece_count,
        bumper_length_m=bumper_length_m,
        weight_mats=weight_mats,
        weight_frame=weight_frame,
        weight_angle=weight_angle,
        weight_bumper=weight_bumper,
        total_weight=total_weight,
        total_weight_with_zinc=total_weight_with_zinc,
        cost_per_kg=cost_per_kg,
    )


# ---------------------------------------------------------------------------
# Cutting / nesting
# ---------------------------------------------------------------------------

def cutting_build_columns(items: list, mat_width: int, kerf: int = 5) -> list:
    """
    Pack individual pieces into columns along the mat's width axis.
    Each item: {"position", "length", "width"}.
    Width = dimension across strips (≤ mat_width=1000mm).
    Length = dimension along strips (≤ mat_length=6000mm, no rotation).
    Returns list of columns: {"width_x": int, "height_used": int, "items": [...]}.
    """
    remaining = sorted(items, key=lambda d: (d["length"], d["width"]), reverse=True)
    columns: list = []
    for part in remaining:
        best_col, best_delta = None, None
        for col in columns:
            need_h = part["width"] if col["height_used"] == 0 else kerf + part["width"]
            if col["height_used"] + need_h > mat_width:
                continue
            new_wx = max(col["width_x"], part["length"])
            delta = new_wx - col["width_x"]
            if best_delta is None or delta < best_delta:
                best_delta, best_col = delta, col
        if best_col is None:
            columns.append({"width_x": part["length"], "height_used": part["width"], "items": [part]})
        else:
            best_col["width_x"] = max(best_col["width_x"], part["length"])
            best_col["items"].append(part)
            best_col["height_used"] += kerf + part["width"]
    for col in columns:
        col["items"].sort(key=lambda d: d["width"], reverse=True)
    return columns


def cutting_place_mats(columns: list, mat_length: int, mat_width: int, kerf: int = 5) -> list:
    """
    Pack columns into mats (bins).
    Returns list of mats, each mat is a list of placed rects:
      {"position", "length", "width", "x", "y"}.
    x is along mat_length (6000mm axis), y is along mat_width (1000mm axis).
    """
    cols = sorted(columns, key=lambda c: c["width_x"], reverse=True)
    mats: list = []
    while cols:
        x = 0
        placed: list = []
        i = 0
        while i < len(cols):
            col = cols[i]
            need_x = col["width_x"] if x == 0 else kerf + col["width_x"]
            if x + need_x <= mat_length:
                y = 0
                for j, d in enumerate(col["items"]):
                    if j > 0:
                        y += kerf
                    placed.append({"position": d["position"], "length": d["length"],
                                   "width": d["width"],
                                   "x": x if x == 0 else x + kerf, "y": y})
                    y += d["width"]
                x += need_x
                cols.pop(i)
                continue
            i += 1
        if not placed and cols:
            col = cols.pop(0)
            y = 0
            for j, d in enumerate(col["items"]):
                if j > 0:
                    y += kerf
                placed.append({"position": d["position"], "length": d["length"],
                               "width": d["width"], "x": 0, "y": y})
                y += d["width"]
        mats.append(placed)
    return mats


def cutting_mat_fraction(rects: list, mat_length: int) -> float:
    """Fraction of mat length used (0.0–1.0), snapped to 0.1 steps."""
    if not rects:
        return 0.0
    used = max(d["x"] + d["length"] for d in rects)
    ratio = used / mat_length if mat_length > 0 else 0.0
    if ratio >= 0.9:
        return 1.0
    if ratio < 0.1:
        return 0.0
    return round(ratio, 1)


def cutting_run(parts: List[Part], mat_length_mm: int = 6100, mat_width_mm: int = 1000,
                kerf: int = 5) -> dict:
    """
    Run nesting for all parts (expanded by quantity).
    Returns {
        "mats": list of mats (each = list of placed rects),
        "mat_count": float (fractional mats needed),
        "skipped": int (parts that didn't fit),
    }.
    """
    items = []
    skipped = 0
    for p in parts:
        try:
            l, w, q = int(p.length), int(p.width), int(p.quantity)
        except Exception:
            continue
        if l <= 0 or w <= 0:
            continue
        if l > mat_length_mm or w > mat_width_mm:
            skipped += q
            continue
        for _ in range(max(q, 0)):
            items.append({"position": p.position, "length": l, "width": w})

    if not items:
        return {"mats": [], "mat_count": 0.0, "skipped": skipped}

    columns = cutting_build_columns(items, mat_width_mm, kerf)
    mats = cutting_place_mats(columns, mat_length_mm, mat_width_mm, kerf)
    total = round(sum(cutting_mat_fraction(m, mat_length_mm) for m in mats), 1)
    return {"mats": mats, "mat_count": total, "skipped": skipped}
