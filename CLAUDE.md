# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

Desktop app for production managers at a steel grating factory. Managers receive customer PDF drawings, recognize grating positions via Claude Vision, verify and edit them, calculate cost, and generate a price offer PDF and/or Excel export. A separate Steps Master module handles stair tread calculations. All processing is local — no cloud storage, no server.

## Running the app

```bash
python manager_app.py
```

Quick launch on Windows: `run.bat` (double-click).

Dependencies:
```bash
pip install customtkinter anthropic pymupdf reportlab pandas openpyxl python-docx matplotlib
```

No build step, no test suite. Verify by running the app directly.

## Architecture

Seven modules, clear separation:

| File | Role |
|---|---|
| `manager_app.py` | All UI (CustomTkinter). 4 screens: history → upload → review → result. |
| `gridmaster_core.py` | Pure calculation engine — no UI imports. `calculate(Order) → OrderResult`. Also `cutting_run()` and `calc_work_hours()`. |
| `ai_recognizer.py` | Claude Vision: PDF → list of part dicts. Uses `claude-opus-4-8` at 150 DPI. Extracts dimensions, frame/kickplate analysis per edge. |
| `spec_reader.py` | Reads customer spec (Word/Excel/PDF), sends text to Claude for parsing, compares with AI-recognized parts. |
| `pdf_export.py` | Generates client offer PDF via reportlab. English/Estonian only (Cyrillic is stripped). |
| `price_manager.py` | Loads material data from Excel, manages `manager_config.json`. |
| `steps_core.py` | Pure calculation engine for stair steps. Uses shared config data. |
| `steps_app.py` | Steps Master UI as CTkToplevel, opened from history screen. |

## Data flow

```
PDF drawings → ai_recognizer (Claude Vision) → list of part dicts
                  ↓ (frame/bumper lengths extracted per edge)
Customer spec → spec_reader (Claude text) → quantities/positions
                                    ↓
              calc_work_hours() auto-fills labour hours per part
                                    ↓
                    manager reviews & edits in UI
                                    ↓
                  gridmaster_core.calculate(Order) → OrderResult
                                    ↓
    pdf_export (offer PDF)  /  Excel export  /  cutting_run (mat nesting)
```

## Key domain concepts

**Mat** — standard steel grating sheet, 6000×1000 mm (or 6100×1000). Source material.

**Part fields:**
- `length` — larger dimension **along** mat's 6000 mm axis (strips direction). Never rotated.
- `width` — dimension across strips (≤ 1000 mm).
- `length2` — second parallel side if the part is a trapezoid (0 = rectangle).
- `radius` / `radius_part` — circular cutout: radius in mm, fraction of circle (1.0=full, 0.5=half, 0.25=quarter).
- `work` — labour hours for this part (auto-calculated by `calc_work_hours`, editable).
- `frame_override` — if > 0, replaces auto-calculated frame perimeter (mm). Set by AI or manually.
- `bumper_override` — if > 0, replaces auto-calculated kickplate length (mm). Set by AI or manually.
- `frame_analysis` — AI reasoning text explaining which edge type is on which side.

**Cost components:** mat + frame (обрамление) + perforated angle (перфоуголок) + kickplate (кикплейт/отбойник) + labour + zinc.

**Frame vs kickplate rule:** Arc from a radius cutout always goes to kickplate, never to frame. Linear kickplate length is subtracted from frame length if `subtract_bumper=True`. If `frame_override` > 0, the formula is bypassed entirely.

**Mat cost** uses rectangle area (`length × width`) ignoring cutouts — offcuts are waste. **Weight** uses real area: trapezoid `(length+length2)/2 × width` minus cutout `π·r²·arc_fraction`.

## AI frame/kickplate analysis

`ai_recognizer.py` instructs Claude to analyze leader lines and labels on each edge and determine:
- which edges have standard frame (обрамление)
- which edges have kickplate/bumper (кикплейт/отбойник) — identified by width callouts, heavier profile, or explicit type labels (e.g. "Tun D", "тип B", "kickplate")
- the total length of each type in mm

Claude returns `frame_mm`, `bumper_mm`, and `frame_analysis` (reasoning text). These populate `frame_override` and `bumper_override` on the Part. In the UI:
- Part row shows **🤖** when AI filled in edge analysis
- Part edit dialog shows a read-only "🤖 Анализ обрамления" block with the reasoning
- Manager can verify and override values manually if needed

## Labour hours auto-calculation

`gridmaster_core.calc_work_hours(part, norms) -> float` computes hours per part:

```
hours = (straight_perimeter_m × t_straight + arc_m × t_arc + n_diagonal × t_diagonal)
        × complexity_factor
complexity = max(1.0, complexity_k / sqrt(area_m2))
```

Norms are stored in `manager_config.json` under `labour_norms` and edited in the ⚙ Settings dialog. Auto-applied after AI recognition; recalculated for all parts via **↻ Трудозатраты** button on the review screen.

Default norms (placeholder — must be calibrated from real measurements):

| Key | Default | Meaning |
|---|---|---|
| `t_straight_h_per_m` | 0.083 | h/m straight cut |
| `t_diagonal_h_per_cut` | 0.25 | h per diagonal cut (trapezoid) |
| `t_arc_h_per_m` | 0.33 | h/m arc cut (radius) |
| `complexity_k` | 0.30 | complexity coefficient constant |
| `min_hours_per_part` | 0.15 | minimum hours per part |

## Config and first run

Config lives in `manager_config.json` (auto-created). On first launch, the app asks for Excel files:
- Mats: columns `Артикул`, `Цена м²`, `Вес м²`, `Длина мата`
- Frame/angle/bumper weight tables: columns `Размер`, `Вес`
- Sides (Боковины) for Steps: columns `Размер`, `Ширина`, `Вес`, `Цена`

**Price memory:** Prices for frame/angle/bumper types are remembered per type in `type_prices` config key. Selecting a type auto-fills the last entered price for that type. Mat price auto-fills from the Excel table when the article is changed.

Steps-specific prices (zinc, angle €/m, kickplate €/kg, work €/h) are stored under `steps_prices` key.

After loading Excel, `_refresh_material_combos()` updates all dropdowns. API key and labour norms are stored in config, edited via the ⚙ Settings button (available on upload and review screens).

## Result screen

Displays:
- Cost breakdown (mat, frame, angle, bumper, labour, zinc, total)
- Composition: positions count, parts count, **mat sheet count (+15%)**, **perforated angle pieces (3m each)**
- Weights
- Client price with configurable margin %

Exports:
- **PDF offer** (`pdf_export.export_offer_pdf`, language en/et)
- **Excel for warehouse** — two sheets: «Позиции» (all parts with dimensions and labour hours) and «Сводка» (cost, material quantities, weights, client price). Filename: `ClientName_YYYY-MM-DD.xlsx`

## Steps Master module

Opened via **🪜 Ступеньки** button on the history screen. Opens as `CTkToplevel` (separate window, same process).

**Shared data from manager_config.json:** `mat_data`, `angle_weights`, `bumper_weights` — no duplicate Excel loading.

**Steps-specific:** `side_data` (Боковины) — loaded via "Боковины (Excel)" button in Steps window or via main Excel settings dialog.

**Calculation** (`steps_core.calc_steps`):
- Боковины × 2 per step: weight and price from side table
- Grid (решётка): `length_mm × side_width_mm / 1e6 × mat_price_m2`
- Angle (перфоуголок): `length_m × angle_price_eur_per_m`
- Kickplate: `length_m × w_per_m × kickplate_price_eur_per_kg`
- Work: `hours × rate × quantity`
- Zinc: `total_weight × 1.1 × zinc_price`

Result includes full breakdown and Excel export (`Ступени_YYYY-MM-DD.xlsx`).

## Cutting/nesting module

`gridmaster_core.cutting_run(parts, mat_length_mm, mat_width_mm)` runs a first-fit decreasing column packer:
1. Parts are expanded by quantity into individual rectangles.
2. `cutting_build_columns` packs pieces into columns along the width (1000 mm) axis, kerf=5 mm.
3. `cutting_place_mats` bins columns into 6000 mm mats.
4. `cutting_mat_fraction` snaps each mat's used-length ratio to 0.1 steps (≥0.9 → 1.0).

Result: `{"mats": [...], "mat_count": float, "skipped": int}`. UI in `manager_app._cutting_dialog` / `_cutting_preview`.

## PDF offer

`pdf_export.export_offer_pdf(..., language="en"|"et")`. Uses Helvetica — no Cyrillic. All notes have non-ASCII stripped before rendering. Price displayed with space-separated thousands: `1 234.56 EUR`.

## Branch

Active development branch: `claude/serene-heisenberg-7wufi5`
