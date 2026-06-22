"""
price_manager.py — загрузка цен из Excel, хранение в config, ручные переопределения.

Логика при старте:
  1. Нет config → просим Excel-файлы → грузим → сохраняем
  2. Есть config → сравниваем mtime Excel с сохранённым → если новее, перегружаем
  3. Ручной ввод цены → сохраняется с флагом manual_override=True
     При перезагрузке Excel manual_override-цены НЕ трогаем
"""

import json
import os
import time
from typing import Dict, List, Tuple, Optional
import pandas as pd

CONFIG_FILE = "manager_config.json"

# Ключи цен в конфиге
PRICE_KEYS = ["mat", "frame", "angle", "bumper", "work", "zinc"]


def _read_excel_safe(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    return df.fillna("")


# ---------------------------------------------------------------------------
# Загрузка данных из Excel-файлов
# ---------------------------------------------------------------------------

def load_mat_data(path: str) -> List[Dict]:
    df = _read_excel_safe(path)
    result = []
    for _, row in df.iterrows():
        article = str(row.get("Артикул", "")).strip()
        if not article:
            continue
        result.append({
            "article":      article,
            "cell":         str(row.get("Ячейка", "")),
            "price_per_m2": _to_float(row.get("Цена м²", row.get("Цена м2", 0))),
            "weight_per_m2": _to_float(row.get("Вес м²",  row.get("Вес м2",  0))),
            "mat_length_mm": int(_to_float(row.get("Длина мата", 6100)) or 6100),
        })
    return result


def load_weight_table(path: str) -> List[Tuple[str, float]]:
    """Загружает таблицу (Размер, Вес) → [(name, kg_per_m), ...]"""
    df = _read_excel_safe(path)
    result = []
    # Колонки могут называться по-разному
    name_col   = _find_col(df, ["Размер", "Название", "Имя", "Name"])
    weight_col = _find_col(df, ["Вес", "Weight", "Ширина"])
    if name_col is None or weight_col is None:
        return result
    for _, row in df.iterrows():
        name = str(row[name_col]).strip()
        w    = _to_float(row[weight_col])
        if name:
            result.append((name, w))
    return result


def load_strip_types(path: str) -> List[Dict]:
    """Загружает типы полос (Название, Ширина, Толщина)."""
    df = _read_excel_safe(path)
    result = []
    for _, row in df.iterrows():
        name = str(row.get("Название", "")).strip()
        if name:
            result.append({
                "name":      name,
                "width_mm":  _to_float(row.get("Ширина",   0)),
                "thick_mm":  _to_float(row.get("Толщина",  0)),
            })
    return result


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def _default_config() -> Dict:
    return {
        "excel_files": {
            "mats":         "",   # путь к файлу матов
            "frame_weight": "",   # веса обрамления
            "angle_weight": "",   # веса перфоуголка
            "bumper_weight": "",  # веса отбойника
            "strip_types":  "",   # типы полос
        },
        "excel_mtime": {},        # path → mtime при последней загрузке
        "mat_data":    [],
        "frame_weights": [],
        "angle_weights": [],
        "bumper_weights": [],
        "strip_types":  [],
        "prices": {
            # каждая цена: {"value": float, "manual_override": bool}
            "mat":    {"value": 0.0, "manual_override": False},
            "frame":  {"value": 0.0, "manual_override": False},
            "angle":  {"value": 0.0, "manual_override": False},
            "bumper": {"value": 0.0, "manual_override": False},
            "work":   {"value": 0.0, "manual_override": False},
            "zinc":   {"value": 0.0, "manual_override": False},
        },
        "last_mat_article": "",
        "last_frame_type":  "",
        "last_angle_type":  "",
        "last_bumper_type": "",
        "margin_pct": 35.0,
    }


def load_config() -> Dict:
    if not os.path.exists(CONFIG_FILE):
        return _default_config()
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        # Добавляем недостающие ключи если config старый
        default = _default_config()
        for k, v in default.items():
            if k not in cfg:
                cfg[k] = v
        return cfg
    except Exception:
        return _default_config()


def save_config(cfg: Dict):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def excel_files_set(cfg: Dict) -> bool:
    """Все обязательные Excel-файлы указаны?"""
    files = cfg.get("excel_files", {})
    required = ["mats", "frame_weight", "angle_weight", "bumper_weight"]
    return all(files.get(k, "") for k in required)


def _mtime(path: str) -> float:
    try:
        return os.path.getmtime(path)
    except Exception:
        return 0.0


def reload_excel_if_needed(cfg: Dict) -> Tuple[Dict, bool]:
    """
    Проверяет mtime всех Excel. Если хоть один изменился — перегружает.
    Возвращает (обновлённый_cfg, был_ли_reload).
    """
    files = cfg.get("excel_files", {})
    saved_mtime = cfg.get("excel_mtime", {})
    needs_reload = False

    for key, path in files.items():
        if not path or not os.path.exists(path):
            continue
        if _mtime(path) != saved_mtime.get(path, 0):
            needs_reload = True
            break

    if not needs_reload:
        return cfg, False

    return reload_excel(cfg), True


def reload_excel(cfg: Dict) -> Dict:
    """Принудительно перегружает все Excel. manual_override-цены не трогает."""
    files = cfg.get("excel_files", {})

    if files.get("mats") and os.path.exists(files["mats"]):
        cfg["mat_data"] = load_mat_data(files["mats"])

    if files.get("frame_weight") and os.path.exists(files["frame_weight"]):
        cfg["frame_weights"] = load_weight_table(files["frame_weight"])

    if files.get("angle_weight") and os.path.exists(files["angle_weight"]):
        cfg["angle_weights"] = load_weight_table(files["angle_weight"])

    if files.get("bumper_weight") and os.path.exists(files["bumper_weight"]):
        cfg["bumper_weights"] = load_weight_table(files["bumper_weight"])

    if files.get("strip_types") and os.path.exists(files["strip_types"]):
        cfg["strip_types"] = load_strip_types(files["strip_types"])

    # Обновляем mtime
    mtime_map = {}
    for path in files.values():
        if path and os.path.exists(path):
            mtime_map[path] = _mtime(path)
    cfg["excel_mtime"] = mtime_map

    # Обновляем цену мата из таблицы (если нет manual_override)
    _sync_mat_price(cfg)

    return cfg


def _sync_mat_price(cfg: Dict):
    """Подтягивает цену выбранного мата из таблицы если нет ручного ввода."""
    if cfg["prices"]["mat"]["manual_override"]:
        return
    article = cfg.get("last_mat_article", "")
    for m in cfg.get("mat_data", []):
        if m["article"] == article:
            cfg["prices"]["mat"]["value"] = m["price_per_m2"]
            return


# ---------------------------------------------------------------------------
# Ручной ввод цены
# ---------------------------------------------------------------------------

def set_price(cfg: Dict, key: str, value: float, manual: bool = True):
    """Устанавливает цену. manual=True → помечается как ручной ввод."""
    cfg["prices"][key] = {"value": value, "manual_override": manual}


def get_price(cfg: Dict, key: str) -> float:
    return cfg["prices"].get(key, {}).get("value", 0.0)


def get_mat_spec(cfg: Dict, article: str):
    """Возвращает dict мата по артикулу или None."""
    for m in cfg.get("mat_data", []):
        if m["article"] == article:
            return m
    return None


# ---------------------------------------------------------------------------
# Вспомогательные
# ---------------------------------------------------------------------------

def _to_float(v, default: float = 0.0) -> float:
    try:
        return float(str(v).strip().replace(",", "."))
    except Exception:
        return default


def _find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    return None
