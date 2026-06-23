"""
steps_core.py — pure calculation logic for stair steps.
Uses shared manager_config.json data (mat_data, angle_weights, bumper_weights).
Adds only side_data (Боковины) specific to stairs.
"""

from typing import Dict, Optional


def _to_float(v, default: float = 0.0) -> float:
    if v is None:
        return default
    try:
        return float(str(v).strip().replace(" ", "").replace(",", "."))
    except Exception:
        return default


def _find_side(cfg: Dict, name: str) -> Optional[Dict]:
    for s in cfg.get("side_data", []):
        if str(s.get("name", "")).strip() == str(name).strip():
            return s
    return None


def _find_mat(cfg: Dict, article: str) -> Optional[Dict]:
    for m in cfg.get("mat_data", []):
        if str(m.get("article", "")).strip() == str(article).strip():
            return m
    return None


def _find_weight(table: list, name: str) -> Optional[float]:
    """Lookup weight (kg/m) from angle_weights or bumper_weights list."""
    for item in table:
        n, w = item[0], item[1]
        if str(n).strip() == str(name).strip():
            return float(w)
    return None


def calc_steps(
    length_mm: float,
    quantity: int,
    side_name: str,
    mat_article: str,
    angle_name: Optional[str],
    kickplate_name: Optional[str],
    work_hours: float,
    cfg: Dict,
) -> Dict:
    """
    Calculate step cost. Returns result dict with full breakdown.

    Prices are read from cfg["steps_prices"]:
      zinc_eur_kg, angle_eur_m, kickplate_eur_kg, work_eur_h

    Returns dict with keys:
      ok, error, breakdown, totals
    """
    prices = cfg.get("steps_prices", {})
    zinc_price   = _to_float(prices.get("zinc_eur_kg", 0))
    angle_price  = _to_float(prices.get("angle_eur_m", 0))
    kick_price   = _to_float(prices.get("kickplate_eur_kg", 0))
    work_rate    = _to_float(prices.get("work_eur_h", 0))

    breakdown = {}
    errors = []

    # ── Боковины ──────────────────────────────────────────────────────────────
    side = _find_side(cfg, side_name)
    if side is None:
        return {"ok": False, "error": f"Боковина не найдена: {side_name}"}
    side_width_mm   = _to_float(side.get("width_mm", 0))
    side_weight_1pc = _to_float(side.get("weight_kg", 0))
    side_price_1pc  = _to_float(side.get("price_eur", 0))

    side_pcs_total    = 2 * quantity
    side_weight_total = side_weight_1pc * side_pcs_total
    side_price_total  = side_price_1pc  * side_pcs_total
    breakdown["sides"] = {
        "pcs": side_pcs_total,
        "weight_kg": round(side_weight_total, 3),
        "cost_eur": round(side_price_total, 2),
    }

    # ── Решётка ───────────────────────────────────────────────────────────────
    mat = _find_mat(cfg, mat_article)
    if mat is None:
        return {"ok": False, "error": f"Мат не найден: {mat_article}"}
    mat_w_m2 = _to_float(mat.get("weight_per_m2", 0))
    mat_p_m2 = _to_float(mat.get("price_per_m2", 0))

    grid_area_m2    = (length_mm * side_width_mm) / 1_000_000.0
    grid_weight_1pc = grid_area_m2 * mat_w_m2
    grid_price_1pc  = grid_area_m2 * mat_p_m2
    breakdown["grid"] = {
        "area_m2_per_unit": round(grid_area_m2, 4),
        "weight_kg": round(grid_weight_1pc * quantity, 3),
        "cost_eur":  round(grid_price_1pc  * quantity, 2),
    }

    # ── Перфоуголок ───────────────────────────────────────────────────────────
    angle_weight_total = 0.0
    angle_price_total  = 0.0
    angle_len_total_m  = 0.0
    if angle_name and angle_name != "Нет":
        w_per_m = _find_weight(cfg.get("angle_weights", []), angle_name)
        if w_per_m is None:
            errors.append(f"Перфоуголок не найден: {angle_name}")
        else:
            angle_len_unit_m    = length_mm / 1000.0
            angle_len_total_m   = angle_len_unit_m * quantity
            angle_weight_total  = w_per_m * angle_len_total_m
            angle_price_total   = angle_len_total_m * angle_price
            breakdown["angle"] = {
                "len_m": round(angle_len_total_m, 3),
                "weight_kg": round(angle_weight_total, 3),
                "cost_eur":  round(angle_price_total, 2),
            }

    # ── Кикплейт ──────────────────────────────────────────────────────────────
    kick_weight_total = 0.0
    kick_price_total  = 0.0
    if kickplate_name and kickplate_name != "Нет":
        w_per_m = _find_weight(cfg.get("bumper_weights", []), kickplate_name)
        if w_per_m is None:
            errors.append(f"Кикплейт не найден: {kickplate_name}")
        else:
            kick_len_m       = length_mm / 1000.0
            kick_weight_unit = w_per_m * kick_len_m
            kick_weight_total = kick_weight_unit * quantity
            kick_price_total  = kick_weight_total * kick_price
            breakdown["kickplate"] = {
                "weight_kg": round(kick_weight_total, 3),
                "cost_eur":  round(kick_price_total, 2),
            }

    # ── Работа ────────────────────────────────────────────────────────────────
    work_cost_total = work_hours * work_rate * quantity
    breakdown["work"] = {
        "hours_per_unit": work_hours,
        "hours_total": work_hours * quantity,
        "cost_eur": round(work_cost_total, 2),
    }

    # ── Итоги ─────────────────────────────────────────────────────────────────
    total_weight = (
        side_weight_total
        + grid_weight_1pc * quantity
        + angle_weight_total
        + kick_weight_total
    )
    total_weight_zinc = total_weight * 1.1
    material_cost = (
        side_price_total
        + grid_price_1pc * quantity
        + angle_price_total
        + kick_price_total
    )
    zinc_cost = total_weight_zinc * zinc_price if zinc_price > 0 else 0.0
    total_cost = material_cost + work_cost_total + zinc_cost

    if zinc_price > 0:
        breakdown["zinc"] = {
            "weight_with_zinc_kg": round(total_weight_zinc, 3),
            "cost_eur": round(zinc_cost, 2),
        }

    totals = {
        "total_weight_kg":      round(total_weight, 3),
        "total_weight_zinc_kg": round(total_weight_zinc, 3),
        "material_cost_eur":    round(material_cost, 2),
        "work_cost_eur":        round(work_cost_total, 2),
        "zinc_cost_eur":        round(zinc_cost, 2),
        "total_cost_eur":       round(total_cost, 2),
        "unit_cost_eur":        round(total_cost / quantity, 2) if quantity else 0,
    }

    return {"ok": True, "breakdown": breakdown, "totals": totals, "warnings": errors}
