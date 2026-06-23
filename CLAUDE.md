# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

Desktop app for production managers at a steel grating factory. Managers receive customer PDF drawings, need to recognize grating positions (via Claude Vision), verify them, calculate cost, and generate a price offer PDF. All processing is local — no cloud storage, no server.

## Running the app

```bash
python manager_app.py
```

Dependencies:
```bash
pip install customtkinter anthropic pymupdf reportlab pandas openpyxl python-docx matplotlib
```

No build step, no test suite. Verify by running the app directly.

## Architecture

Five modules, clear separation:

| File | Role |
|---|---|
| `manager_app.py` | All UI (CustomTkinter). 4 screens: history → upload → review → result. |
| `gridmaster_core.py` | Pure calculation engine — no UI imports. `calculate(Order) → OrderResult`. Also contains cutting/nesting logic (`cutting_run`). |
| `ai_recognizer.py` | Claude Vision: PDF → list of part dicts. Uses `claude-opus-4-8` at 150 DPI. |
| `spec_reader.py` | Reads customer spec (Word/Excel/PDF), sends text to Claude for parsing, compares with AI-recognized parts. |
| `pdf_export.py` | Generates client offer PDF via reportlab. English/Estonian only (Cyrillic is stripped). |
| `price_manager.py` | Loads material data from Excel, manages `manager_config.json`. |

## Data flow

```
PDF drawings → ai_recognizer (Claude Vision) → list of part dicts
Customer spec → spec_reader (Claude text) → quantities/positions
                                    ↓
                    manager reviews & edits in UI
                                    ↓
                  gridmaster_core.calculate(Order) → OrderResult
                                    ↓
              pdf_export (offer PDF)  /  cutting_run (mat nesting for warehouse)
```

## Key domain concepts

**Mat** — standard steel grating sheet, 6000×1000 mm (or 6100×1000). Source material.

**Part fields:**
- `length` — larger dimension **along** mat's 6000 mm axis (strips direction). Never rotated.
- `width` — dimension across strips (≤ 1000 mm).
- `length2` — second parallel side if the part is a trapezoid (0 = rectangle).
- `radius` / `radius_part` — circular cutout: radius in mm, fraction of circle (1.0=full, 0.5=half, 0.25=quarter).

**Cost components:** mat + frame (обрамление) + perforated angle (перфоуголок) + kickplate (кикплейт/отбойник) + labour + zinc.

**Frame vs kickplate rule:** Arc from a radius cutout always goes to kickplate, never to frame. Linear kickplate length is subtracted from frame length if `subtract_bumper=True`.

**Mat cost** uses rectangle area (`length × width`) ignoring cutouts — offcuts are waste. **Weight** uses real area: trapezoid `(length+length2)/2 × width` minus cutout `π·r²·arc_fraction`.

## Config and first run

Config lives in `manager_config.json` (auto-created). On first launch, the app asks for four Excel files:
- Mats: columns `Артикул`, `Цена м²`, `Вес м²`, `Длина мата`
- Frame/angle/bumper weight tables: columns `Размер`, `Вес`

After loading Excel, `_refresh_material_combos()` updates all dropdowns. API key is stored in config, entered via the ⚙ Settings button (not on the main screen).

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

Active development branch: `claude/trusting-tesla-f7axg8`
