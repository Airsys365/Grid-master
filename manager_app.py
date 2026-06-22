"""
manager_app.py — прога менеджера Grid Master.

Запуск: python manager_app.py
"""

import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List, Dict, Optional
import customtkinter as ctk

import gridmaster_core as core
import price_manager as pm
import ai_recognizer as ai
import pdf_export as pdf_exp

APP_TITLE   = "Grid Master — Менеджер"
ORDERS_FILE = "orders_history.json"
API_KEY_ENV = "ANTHROPIC_API_KEY"


# ---------------------------------------------------------------------------
# Вспомогательные
# ---------------------------------------------------------------------------

def _load_orders() -> List[Dict]:
    if not os.path.exists(ORDERS_FILE):
        return []
    try:
        with open(ORDERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _save_orders(orders: List[Dict]):
    with open(ORDERS_FILE, "w", encoding="utf-8") as f:
        json.dump(orders, f, ensure_ascii=False, indent=2)


def _float(s, default=0.0) -> float:
    try:
        return float(str(s).strip().replace(",", "."))
    except Exception:
        return default


# ---------------------------------------------------------------------------
# Главное окно
# ---------------------------------------------------------------------------

class ManagerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x720")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.cfg = pm.load_config()
        self._check_excel_on_start()

        self.orders = _load_orders()
        self.api_key = os.environ.get(API_KEY_ENV, self.cfg.get("api_key", ""))

        # Текущий заказ (редактируется на экранах 2-4)
        self.current_parts: List[Dict] = []
        self.current_client = ""

        self._build_ui()
        self._show_screen("history")

    # -----------------------------------------------------------------------
    # Старт: проверка Excel
    # -----------------------------------------------------------------------

    def _check_excel_on_start(self):
        if not pm.excel_files_set(self.cfg):
            self.after(300, self._ask_excel_files)
        else:
            self.cfg, reloaded = pm.reload_excel_if_needed(self.cfg)
            if reloaded:
                pm.save_config(self.cfg)

    def _ask_excel_files(self):
        win = ctk.CTkToplevel(self)
        win.title("Настройка — файлы данных")
        win.geometry("560x380")
        win.grab_set()

        ctk.CTkLabel(win, text="Укажите Excel-файлы с данными", font=("Segoe UI", 14, "bold")).pack(pady=14)

        fields = {
            "mats":          "Файл матов (Артикул, Ячейка, Цена м², Вес м²):",
            "frame_weight":  "Веса обрамления (Размер, Вес):",
            "angle_weight":  "Веса перфоуголка (Размер, Вес):",
            "bumper_weight": "Веса отбойника (Размер, Вес):",
            "strip_types":   "Типы полос/Название, Ширина, Толщина (необязательно):",
        }
        entries = {}
        for key, label in fields.items():
            row = ctk.CTkFrame(win)
            row.pack(fill="x", padx=16, pady=3)
            ctk.CTkLabel(row, text=label, anchor="w", width=340).pack(side="left")
            e = ctk.CTkEntry(row, width=100)
            e.insert(0, self.cfg["excel_files"].get(key, ""))
            e.pack(side="left", padx=4)
            entries[key] = e
            ctk.CTkButton(row, text="...", width=32,
                          command=lambda k=key, entry=e: self._browse_excel(k, entry)).pack(side="left")

        def save():
            for k, e in entries.items():
                self.cfg["excel_files"][k] = e.get().strip()
            self.cfg = pm.reload_excel(self.cfg)
            pm.save_config(self.cfg)
            win.destroy()
            messagebox.showinfo("Готово", "Данные загружены успешно!")

        ctk.CTkButton(win, text="Сохранить и загрузить", command=save).pack(pady=12)

    def _browse_excel(self, key: str, entry: ctk.CTkEntry):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path:
            entry.delete(0, "end")
            entry.insert(0, path)

    # -----------------------------------------------------------------------
    # Структура UI
    # -----------------------------------------------------------------------

    def _build_ui(self):
        self.container = ctk.CTkFrame(self)
        self.container.pack(fill="both", expand=True)
        self.screens: Dict[str, ctk.CTkFrame] = {}
        for name in ("history", "upload", "review", "result"):
            frame = ctk.CTkFrame(self.container)
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.screens[name] = frame

        self._build_history_screen()
        self._build_upload_screen()
        self._build_review_screen()
        self._build_result_screen()

    def _show_screen(self, name: str):
        self.screens[name].lift()

    # -----------------------------------------------------------------------
    # Экран 1 — История
    # -----------------------------------------------------------------------

    def _build_history_screen(self):
        s = self.screens["history"]

        top = ctk.CTkFrame(s, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=16)
        ctk.CTkLabel(top, text="Grid Master — Менеджер", font=("Segoe UI", 20, "bold")).pack(side="left")
        ctk.CTkButton(top, text="+ Новый заказ", width=140, command=self._new_order).pack(side="right")
        ctk.CTkButton(top, text="⚙ Настройки", width=110,
                      fg_color="transparent", border_width=1,
                      command=self._ask_excel_files).pack(side="right", padx=8)

        # API key field
        api_row = ctk.CTkFrame(s, fg_color="transparent")
        api_row.pack(fill="x", padx=20, pady=(0, 8))
        ctk.CTkLabel(api_row, text="API ключ Claude:", width=130, anchor="w").pack(side="left")
        self._api_entry = ctk.CTkEntry(api_row, width=340, show="*")
        self._api_entry.insert(0, self.api_key)
        self._api_entry.pack(side="left", padx=6)
        ctk.CTkButton(api_row, text="Сохранить", width=90,
                      command=self._save_api_key).pack(side="left")

        # Список заказов
        list_frame = ctk.CTkScrollableFrame(s, label_text="История заказов")
        list_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        self._history_frame = list_frame
        self._refresh_history()

    def _refresh_history(self):
        for w in self._history_frame.winfo_children():
            w.destroy()
        if not self.orders:
            ctk.CTkLabel(self._history_frame, text="Заказов пока нет. Нажмите «+ Новый заказ».",
                         text_color="grey").pack(pady=40)
            return
        for order in reversed(self.orders):
            self._history_row(order)

    def _history_row(self, order: Dict):
        row = ctk.CTkFrame(self._history_frame, fg_color=("grey90", "grey20"), corner_radius=8)
        row.pack(fill="x", pady=3, padx=4)
        ctk.CTkLabel(row, text=order.get("date", ""), width=90).pack(side="left", padx=10)
        ctk.CTkLabel(row, text=order.get("client", ""), width=220, anchor="w").pack(side="left")
        price = order.get("client_price", 0)
        ctk.CTkLabel(row, text=f"{price:,.2f} €".replace(",", " "),
                     font=("Segoe UI", 12, "bold"), width=120).pack(side="left")
        ctk.CTkButton(row, text="Открыть ▸", width=90,
                      command=lambda o=order: self._open_order(o)).pack(side="right", padx=8, pady=6)

    def _save_api_key(self):
        self.api_key = self._api_entry.get().strip()
        self.cfg["api_key"] = self.api_key
        pm.save_config(self.cfg)
        messagebox.showinfo("Сохранено", "API ключ сохранён.")

    def _new_order(self):
        self.current_parts = []
        self.current_client = ""
        self._client_entry.delete(0, "end")
        self._files_label.configure(text="Файлы не выбраны")
        self._pdf_paths = []
        self._show_screen("upload")

    def _open_order(self, order: Dict):
        self.current_parts  = order.get("parts", [])
        self.current_client = order.get("client", "")
        self._populate_review()
        self._show_screen("review")

    # -----------------------------------------------------------------------
    # Экран 2 — Загрузка чертежей
    # -----------------------------------------------------------------------

    def _build_upload_screen(self):
        s = self.screens["upload"]
        self._pdf_paths: List[str] = []

        ctk.CTkButton(s, text="← Назад", width=80, fg_color="transparent",
                      command=lambda: self._show_screen("history")).pack(anchor="w", padx=16, pady=12)

        ctk.CTkLabel(s, text="Новый заказ", font=("Segoe UI", 18, "bold")).pack(pady=(0, 12))

        client_row = ctk.CTkFrame(s, fg_color="transparent")
        client_row.pack(pady=6)
        ctk.CTkLabel(client_row, text="Клиент:", width=70).pack(side="left")
        self._client_entry = ctk.CTkEntry(client_row, width=300, placeholder_text="Название компании / имя")
        self._client_entry.pack(side="left", padx=8)

        # Drop zone
        drop = ctk.CTkFrame(s, width=500, height=160, fg_color=("grey85", "grey25"),
                            corner_radius=16, border_width=2, border_color=("grey70", "grey40"))
        drop.pack(pady=16)
        drop.pack_propagate(False)
        ctk.CTkLabel(drop, text="📄  Перетащи PDF сюда\nили нажми для выбора",
                     font=("Segoe UI", 13), text_color="grey").pack(expand=True)
        drop.bind("<Button-1>", lambda e: self._choose_files())

        self._files_label = ctk.CTkLabel(s, text="Файлы не выбраны", text_color="grey")
        self._files_label.pack()

        btn_row = ctk.CTkFrame(s, fg_color="transparent")
        btn_row.pack(pady=16)
        ctk.CTkButton(btn_row, text="Выбрать файлы", width=160,
                      fg_color="transparent", border_width=1,
                      command=self._choose_files).pack(side="left", padx=8)
        self._recog_btn = ctk.CTkButton(btn_row, text="Распознать →", width=160,
                                         command=self._start_recognition)
        self._recog_btn.pack(side="left", padx=8)

        self._recog_status = ctk.CTkLabel(s, text="", text_color="grey")
        self._recog_status.pack()

    def _choose_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите чертежи",
            filetypes=[("PDF", "*.pdf"), ("Все файлы", "*.*")]
        )
        if paths:
            self._pdf_paths = list(paths)
            names = ", ".join(os.path.basename(p) for p in paths)
            self._files_label.configure(text=names, text_color="white")

    def _start_recognition(self):
        if not self._pdf_paths:
            messagebox.showwarning("Нет файлов", "Сначала выберите PDF-чертежи.")
            return
        # Берём ключ прямо из поля — не требуем нажатия «Сохранить»
        self.api_key = self._api_entry.get().strip()
        if not self.api_key:
            messagebox.showwarning("Нет API ключа", "Введите API ключ Claude на главном экране.")
            return
        self.cfg["api_key"] = self.api_key
        pm.save_config(self.cfg)
        self.current_client = self._client_entry.get().strip()
        self._recog_btn.configure(state="disabled", text="Распознаю...")
        self._recog_status.configure(text="")

        def work():
            all_parts = []
            errors = []
            for path in self._pdf_paths:
                try:
                    self.after(0, lambda p=path: self._recog_status.configure(
                        text=f"Обрабатываю: {os.path.basename(p)}...", text_color="grey"))
                    parts = ai.recognize_pdf(path, self.api_key)
                    all_parts.extend(parts)
                except Exception as e:
                    import traceback
                    errors.append(f"{os.path.basename(path)}: {traceback.format_exc()}")
            self.current_parts = all_parts
            if errors:
                err_text = "\n\n".join(errors)
                self.after(0, lambda t=err_text: messagebox.showerror("Ошибка распознавания", t))
            self.after(0, self._on_recognition_done)

        threading.Thread(target=work, daemon=True).start()

    def _on_recognition_done(self):
        self._recog_btn.configure(state="normal", text="Распознать →")
        if not self.current_parts:
            self._recog_status.configure(
                text="Не удалось распознать позиции. Добавьте вручную.", text_color="orange")
            self._populate_review()
        else:
            n = len(self.current_parts)
            self._recog_status.configure(text=f"Распознано {n} поз.", text_color="green")
        self._populate_review()
        self._show_screen("review")

    # -----------------------------------------------------------------------
    # Экран 3 — Проверка позиций и параметры
    # -----------------------------------------------------------------------

    def _build_review_screen(self):
        s = self.screens["review"]

        top = ctk.CTkFrame(s, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=10)
        ctk.CTkButton(top, text="← Назад", width=80, fg_color="transparent",
                      command=lambda: self._show_screen("upload")).pack(side="left")
        ctk.CTkLabel(top, text="Проверка позиций", font=("Segoe UI", 16, "bold")).pack(side="left", padx=12)
        ctk.CTkButton(top, text="Рассчитать →", width=140,
                      fg_color=("#2E8B57", "#2E8B50"),
                      command=self._calculate).pack(side="right")

        # Основная область: слева позиции, справа параметры
        main = ctk.CTkFrame(s, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        main.columnconfigure(0, weight=3)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        # --- Левая колонка: позиции ---
        left = ctk.CTkFrame(main)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        parts_top = ctk.CTkFrame(left, fg_color="transparent")
        parts_top.pack(fill="x", padx=8, pady=6)
        ctk.CTkLabel(parts_top, text="Позиции", font=("Segoe UI", 12, "bold")).pack(side="left")
        ctk.CTkButton(parts_top, text="+ Добавить", width=90,
                      command=self._add_part_manual).pack(side="right")

        self._parts_scroll = ctk.CTkScrollableFrame(left)
        self._parts_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # --- Правая колонка: материалы и цены ---
        right = ctk.CTkScrollableFrame(main, label_text="Материалы и цены")
        right.grid(row=0, column=1, sticky="nsew")
        self._params_frame = right
        self._build_params_panel(right)

    def _build_params_panel(self, parent):
        def row(label, widget_fn):
            r = ctk.CTkFrame(parent, fg_color="transparent")
            r.pack(fill="x", padx=8, pady=3)
            ctk.CTkLabel(r, text=label, width=120, anchor="w").pack(side="left")
            widget_fn(r)

        # Материалы
        ctk.CTkLabel(parent, text="Материалы", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=8, pady=(8,2))

        mat_articles = [m["article"] for m in self.cfg.get("mat_data", [])]
        self._mat_cb = ctk.CTkComboBox(parent, values=mat_articles or ["—"], width=180,
                                        command=self._on_mat_changed)
        last = self.cfg.get("last_mat_article", "")
        if last in mat_articles:
            self._mat_cb.set(last)
        elif mat_articles:
            self._mat_cb.set(mat_articles[0])

        def mat_row(p):
            r = ctk.CTkFrame(p, fg_color="transparent")
            r.pack(fill="x", padx=8, pady=3)
            ctk.CTkLabel(r, text="Мат:", width=120, anchor="w").pack(side="left")
            self._mat_cb.pack(side="left")
        mat_row(parent)

        frame_types = [t for t, _ in self.cfg.get("frame_weights", [])]
        self._frame_cb = ctk.CTkComboBox(parent, values=frame_types or ["—"], width=180)
        last_f = self.cfg.get("last_frame_type", "")
        if last_f in frame_types:
            self._frame_cb.set(last_f)
        elif frame_types:
            self._frame_cb.set(frame_types[0])
        row("Обрамление:", lambda p: self._frame_cb.pack(side="left"))

        angle_types = [t for t, _ in self.cfg.get("angle_weights", [])]
        self._angle_cb = ctk.CTkComboBox(parent, values=angle_types or ["—"], width=180)
        last_a = self.cfg.get("last_angle_type", "")
        if last_a in angle_types:
            self._angle_cb.set(last_a)
        elif angle_types:
            self._angle_cb.set(angle_types[0])
        row("Перфоуголок:", lambda p: self._angle_cb.pack(side="left"))

        bumper_types = [t for t, _ in self.cfg.get("bumper_weights", [])]
        self._bumper_cb = ctk.CTkComboBox(parent, values=bumper_types or ["—"], width=180)
        last_b = self.cfg.get("last_bumper_type", "")
        if last_b in bumper_types:
            self._bumper_cb.set(last_b)
        elif bumper_types:
            self._bumper_cb.set(bumper_types[0])
        row("Отбойник:", lambda p: self._bumper_cb.pack(side="left"))

        # Цены
        ctk.CTkLabel(parent, text="Цены (редактируемые)", font=("Segoe UI", 11, "bold")).pack(
            anchor="w", padx=8, pady=(12, 2))

        self._price_entries: Dict[str, ctk.CTkEntry] = {}
        price_labels = {
            "mat":    "Мат €/м²:",
            "frame":  "Обрамление €/кг:",
            "angle":  "Перфоуголок €/шт:",
            "bumper": "Отбойник €/кг:",
            "work":   "Работа €/час:",
            "zinc":   "Цинк €/кг:",
        }
        for key, label in price_labels.items():
            r = ctk.CTkFrame(parent, fg_color="transparent")
            r.pack(fill="x", padx=8, pady=2)
            ctk.CTkLabel(r, text=label, width=140, anchor="w").pack(side="left")
            e = ctk.CTkEntry(r, width=90)
            val = pm.get_price(self.cfg, key)
            e.insert(0, f"{val:.2f}" if val else "")
            e.pack(side="left")
            e.bind("<FocusOut>", lambda ev, k=key: self._on_price_edited(k))
            self._price_entries[key] = e

        # Коэффициенты
        ctk.CTkLabel(parent, text="Прочее", font=("Segoe UI", 11, "bold")).pack(
            anchor="w", padx=8, pady=(12, 2))

        r = ctk.CTkFrame(parent, fg_color="transparent")
        r.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(r, text="Коэф. обрамления:", width=140, anchor="w").pack(side="left")
        self._frame_coef_entry = ctk.CTkEntry(r, width=60)
        self._frame_coef_entry.insert(0, "1.0")
        self._frame_coef_entry.pack(side="left")

        r2 = ctk.CTkFrame(parent, fg_color="transparent")
        r2.pack(fill="x", padx=8, pady=2)
        ctk.CTkLabel(r2, text="Коэф. работы:", width=140, anchor="w").pack(side="left")
        self._work_coef_entry = ctk.CTkEntry(r2, width=60)
        self._work_coef_entry.insert(0, "1.0")
        self._work_coef_entry.pack(side="left")

    def _on_mat_changed(self, article: str):
        spec = pm.get_mat_spec(self.cfg, article)
        if spec and not self.cfg["prices"]["mat"]["manual_override"]:
            e = self._price_entries.get("mat")
            if e:
                e.delete(0, "end")
                e.insert(0, f"{spec['price_per_m2']:.2f}")

    def _on_price_edited(self, key: str):
        e = self._price_entries.get(key)
        if not e:
            return
        val = _float(e.get())
        pm.set_price(self.cfg, key, val, manual=True)

    def _populate_review(self):
        """Заполняет список позиций из self.current_parts."""
        for w in self._parts_scroll.winfo_children():
            w.destroy()
        for i, part in enumerate(self.current_parts):
            self._part_row(i, part)

    def _part_row(self, idx: int, part: Dict):
        conf = part.get("confidence", "high")
        bg = ("grey90", "grey22") if conf == "high" else ("yellow", "#4a4000") if conf == "medium" else ("orange", "#5a2000")
        row = ctk.CTkFrame(self._parts_scroll, fg_color=bg, corner_radius=6)
        row.pack(fill="x", pady=2, padx=2)

        icon = "✅" if conf == "high" else "⚠️"
        ctk.CTkLabel(row, text=icon, width=24).pack(side="left", padx=4)
        ctk.CTkLabel(row, text=str(part.get("position", "")), width=160, anchor="w").pack(side="left")
        ctk.CTkLabel(row, text=f"{part.get('length',0)} × {part.get('width',0)} мм", width=130).pack(side="left")
        ctk.CTkLabel(row, text=f"× {part.get('quantity',1)}", width=40).pack(side="left")

        ctk.CTkButton(row, text="✏", width=30, fg_color="transparent",
                      command=lambda i=idx: self._edit_part(i)).pack(side="right", padx=4)
        ctk.CTkButton(row, text="✕", width=30, fg_color="transparent", text_color="red",
                      command=lambda i=idx: self._delete_part(i)).pack(side="right")

        if part.get("notes"):
            ctk.CTkLabel(row, text=part["notes"], text_color="grey", font=("Segoe UI", 9)).pack(
                side="left", padx=8)

    def _add_part_manual(self):
        part = {"position": "", "length": 0, "width": 0, "quantity": 1,
                "notes": "", "confidence": "high"}
        self.current_parts.append(part)
        self._edit_part(len(self.current_parts) - 1)

    def _edit_part(self, idx: int):
        part = self.current_parts[idx]
        win = ctk.CTkToplevel(self)
        win.title("Редактировать позицию")
        win.geometry("380x320")
        win.grab_set()

        fields = [
            ("Позиция/обозначение", "position", str(part.get("position", ""))),
            ("Длина (мм)",          "length",   str(part.get("length", ""))),
            ("Ширина (мм)",         "width",    str(part.get("width", ""))),
            ("Количество",          "quantity", str(part.get("quantity", 1))),
            ("Радиус (мм)",         "radius",   str(part.get("radius", ""))),
            ("Доля окружности",     "radius_part", str(part.get("radius_part", "1.0"))),
            ("Отбойник (мм)",       "bumper",   str(part.get("bumper", ""))),
            ("Работа (часов)",      "work",     str(part.get("work", ""))),
            ("Примечание",          "notes",    str(part.get("notes", ""))),
        ]
        entries = {}
        for label, key, val in fields:
            r = ctk.CTkFrame(win, fg_color="transparent")
            r.pack(fill="x", padx=16, pady=3)
            ctk.CTkLabel(r, text=label, width=160, anchor="w").pack(side="left")
            e = ctk.CTkEntry(r, width=160)
            e.insert(0, val)
            e.pack(side="left")
            entries[key] = e

        def save():
            for key, e in entries.items():
                v = e.get().strip()
                if key in ("length", "width", "quantity"):
                    part[key] = int(_float(v) or 0)
                elif key in ("radius", "bumper", "work"):
                    part[key] = _float(v)
                elif key == "radius_part":
                    part[key] = _float(v) or 1.0
                else:
                    part[key] = v
            part["confidence"] = "high"
            self.current_parts[idx] = part
            self._populate_review()
            win.destroy()

        ctk.CTkButton(win, text="Сохранить", command=save).pack(pady=12)

    def _delete_part(self, idx: int):
        self.current_parts.pop(idx)
        self._populate_review()

    # -----------------------------------------------------------------------
    # Расчёт
    # -----------------------------------------------------------------------

    def _build_order(self) -> core.Order:
        parts = []
        for p in self.current_parts:
            parts.append(core.Part(
                position=str(p.get("position", "")),
                length=float(p.get("length", 0)),
                width=float(p.get("width", 0)),
                quantity=int(p.get("quantity", 1)),
                angle=str(p.get("angle", "0")),
                bumper=float(p.get("bumper", 0)),
                work=float(p.get("work", 0)),
                radius=float(p.get("radius", 0)),
                radius_part=float(p.get("radius_part", 1.0)),
                subtract_bumper=bool(p.get("subtract_bumper", True)),
                bumper_length=float(p.get("bumper_length", 0)),
            ))

        article = self._mat_cb.get()
        spec = pm.get_mat_spec(self.cfg, article)
        mat = core.MatSpec(
            article=article,
            price_per_m2=_float(self._price_entries["mat"].get()),
            weight_per_m2=spec["weight_per_m2"] if spec else 0.0,
        )

        return core.Order(
            parts=parts,
            mat=mat,
            frame_type=self._frame_cb.get(),
            frame_price=_float(self._price_entries["frame"].get()),
            frame_coef=_float(self._frame_coef_entry.get()) or 1.0,
            frame_weights=self.cfg.get("frame_weights", []),
            angle_type=self._angle_cb.get(),
            angle_price=_float(self._price_entries["angle"].get()),
            angle_weights=self.cfg.get("angle_weights", []),
            bumper_type=self._bumper_cb.get(),
            bumper_price=_float(self._price_entries["bumper"].get()),
            bumper_weights=self.cfg.get("bumper_weights", []),
            work_price=_float(self._price_entries["work"].get()),
            work_coef=_float(self._work_coef_entry.get()) or 1.0,
            zinc_price=_float(self._price_entries["zinc"].get()),
        )

    def _calculate(self):
        if not self.current_parts:
            messagebox.showwarning("Нет позиций", "Добавьте хотя бы одну позицию.")
            return
        try:
            order = self._build_order()
            self._last_result = core.calculate(order)
            self._show_result()
        except Exception as e:
            messagebox.showerror("Ошибка расчёта", str(e))

    # -----------------------------------------------------------------------
    # Экран 4 — Результат
    # -----------------------------------------------------------------------

    def _build_result_screen(self):
        s = self.screens["result"]

        top = ctk.CTkFrame(s, fg_color="transparent")
        top.pack(fill="x", padx=16, pady=10)
        ctk.CTkButton(top, text="← Назад", width=80, fg_color="transparent",
                      command=lambda: self._show_screen("review")).pack(side="left")
        ctk.CTkLabel(top, text="Результат", font=("Segoe UI", 16, "bold")).pack(side="left", padx=12)

        # Левая часть — таблица себестоимости
        body = ctk.CTkFrame(s, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16)
        body.columnconfigure(0, weight=2)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ctk.CTkLabel(left, text="Себестоимость", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=8)

        self._cost_labels: Dict[str, ctk.CTkLabel] = {}
        rows_cfg = [
            ("mat",    "Маты (+15%):"),
            ("frame",  "Обрамление:"),
            ("angle",  "Перфоуголок:"),
            ("bumper", "Отбойник:"),
            ("work",   "Работа:"),
            ("zinc",   "Цинкование:"),
            ("total",  "ИТОГО себест.:"),
        ]
        for key, label in rows_cfg:
            r = ctk.CTkFrame(left, fg_color="transparent")
            r.pack(fill="x", padx=12, pady=2)
            font = ("Segoe UI", 11, "bold") if key == "total" else ("Segoe UI", 11)
            ctk.CTkLabel(r, text=label, width=160, anchor="w", font=font).pack(side="left")
            lbl = ctk.CTkLabel(r, text="—", width=100, anchor="e", font=font)
            lbl.pack(side="right")
            self._cost_labels[key] = lbl

        # Веса
        ctk.CTkLabel(left, text="Веса", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=(14,4))
        self._weight_labels: Dict[str, ctk.CTkLabel] = {}
        for key, label in [("w_mats","Решётка, кг:"), ("w_frame","Обрамление, кг:"),
                            ("w_total","Общий вес, кг:"), ("w_zinc","С цинком (+10%), кг:")]:
            r = ctk.CTkFrame(left, fg_color="transparent")
            r.pack(fill="x", padx=12, pady=2)
            ctk.CTkLabel(r, text=label, width=160, anchor="w").pack(side="left")
            lbl = ctk.CTkLabel(r, text="—", width=100, anchor="e")
            lbl.pack(side="right")
            self._weight_labels[key] = lbl

        # Правая часть — маржа и цена клиенту
        right = ctk.CTkFrame(body)
        right.grid(row=0, column=1, sticky="nsew")

        ctk.CTkLabel(right, text="Цена клиенту", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=12, pady=8)

        ctk.CTkLabel(right, text="Маржа %:", anchor="w").pack(anchor="w", padx=12)
        self._margin_var = ctk.StringVar(value=str(self.cfg.get("margin_pct", 35)))
        self._margin_entry = ctk.CTkEntry(right, width=100, textvariable=self._margin_var)
        self._margin_entry.pack(anchor="w", padx=12, pady=4)
        self._margin_var.trace_add("write", lambda *_: self._update_client_price())

        self._client_price_label = ctk.CTkLabel(right, text="— €",
            font=("Segoe UI", 28, "bold"), text_color=("#2E8B57", "#3CB371"))
        self._client_price_label.pack(pady=20)

        self._cost_per_kg_label = ctk.CTkLabel(right, text="", text_color="grey")
        self._cost_per_kg_label.pack()

        # Кнопки
        btns = ctk.CTkFrame(right, fg_color="transparent")
        btns.pack(pady=20)
        ctk.CTkButton(btns, text="📄 PDF оффер", width=140,
                      fg_color=("#2E8B57", "#2E8B50"),
                      command=self._export_pdf).pack(pady=6)
        ctk.CTkButton(btns, text="💾 Сохранить заказ", width=140,
                      command=self._save_order).pack(pady=6)

    def _show_result(self):
        r = self._last_result
        self._cost_labels["mat"].configure(text=f"{r.mat_cost:.2f} €")
        self._cost_labels["frame"].configure(text=f"{r.frame_cost:.2f} €")
        self._cost_labels["angle"].configure(text=f"{r.angle_cost:.2f} €")
        self._cost_labels["bumper"].configure(text=f"{r.bumper_cost:.2f} €")
        self._cost_labels["work"].configure(text=f"{r.work_cost:.2f} €")
        self._cost_labels["zinc"].configure(text=f"{r.zinc_cost:.2f} €")
        self._cost_labels["total"].configure(text=f"{r.total_cost:.2f} €")

        self._weight_labels["w_mats"].configure(text=f"{r.weight_mats:.2f}")
        self._weight_labels["w_frame"].configure(text=f"{r.weight_frame:.2f}")
        self._weight_labels["w_total"].configure(text=f"{r.total_weight:.2f}")
        self._weight_labels["w_zinc"].configure(text=f"{r.total_weight_with_zinc:.2f}")
        self._cost_per_kg_label.configure(text=f"Себест. €/кг: {r.cost_per_kg:.2f}")

        self._update_client_price()
        self._show_screen("result")

    def _update_client_price(self):
        if not hasattr(self, "_last_result"):
            return
        margin = _float(self._margin_var.get()) if self._margin_var.get() else 0.0
        client_price = round(self._last_result.total_cost * (1 + margin / 100), 2)
        self._client_price_label.configure(text=f"{client_price:,.2f} €".replace(",", " "))

    def _export_pdf(self):
        if not hasattr(self, "_last_result"):
            return
        margin = _float(self._margin_var.get()) or 0.0
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"Оффер_{self.current_client or 'клиент'}.pdf",
        )
        if not path:
            return
        try:
            pdf_exp.export_offer_pdf(
                output_path=path,
                client_name=self.current_client,
                parts=self.current_parts,
                total_cost=self._last_result.total_cost,
                margin_pct=margin,
                company_name=self.cfg.get("company_name", "Grid Master"),
                company_contact=self.cfg.get("company_contact", ""),
            )
            messagebox.showinfo("Готово", f"PDF сохранён:\n{path}")
        except Exception as e:
            messagebox.showerror("Ошибка PDF", str(e))

    def _save_order(self):
        if not hasattr(self, "_last_result"):
            return
        from datetime import datetime
        margin = _float(self._margin_var.get()) or 0.0
        client_price = round(self._last_result.total_cost * (1 + margin / 100), 2)
        order = {
            "date":         datetime.now().strftime("%d.%m.%Y"),
            "client":       self.current_client,
            "total_cost":   self._last_result.total_cost,
            "margin_pct":   margin,
            "client_price": client_price,
            "parts":        self.current_parts,
        }
        self.orders.append(order)
        _save_orders(self.orders)
        self.cfg["margin_pct"] = margin
        pm.save_config(self.cfg)
        self._refresh_history()
        messagebox.showinfo("Сохранено", "Заказ добавлен в историю.")


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = ManagerApp()
    app.mainloop()
