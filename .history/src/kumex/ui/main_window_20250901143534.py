"""
Главное окно GUI (минимальный функционал):
- Поле "Месяц (YYYY-MM)"
- Поле "Папка PDF" + кнопка "Выбрать…"
- Кнопки: "Сканировать PDF", "Сформировать отчёт", "Применить"
- Сохранение простого конфига при выходе (state/kumex_config.json)
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
import time
import re
from kumex.io.pdf_reader import read_pdf_text
from kumex.io.file_ops import load_json, save_json


class MainWindow(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master: tk.Tk = master
        self.year_var = tk.StringVar()
        self.month_num_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Готово")
        # Настройки склада Kumex (минимум: два материала) — пока без логики


        
        self.pdf_files = []          # список путей найденных PDF
        self.material_rows = []      # сюда позже положим строки из PDF-парсера


        # --- пути и состояние ---
        # <корень проекта> = ../../../.. от этого файла
        self.project_root = Path(__file__).resolve().parents[4]
        self.state_dir = self.project_root / "state"
        self.config_path = self.state_dir / "kumex_config.json"
        # kumex_stock.json → загрузка материалов
        self.stock_path = self.state_dir / "kumex_stock.json"
        stock_data = load_json(self.stock_path, default={"materials": {}})

            # Оставляем в интерфейсе только два материала
        allowed = {"POM Valge", "POM Must"}
        mats = stock_data.get("materials", {})
        filtered = {k: v for k, v in mats.items() if k in allowed}
        # если JSON содержит лишние материалы — очистим и сохраним
        if filtered != mats:
            stock_data["materials"] = filtered
            save_json(self.stock_path, stock_data)
        # если вообще пусто — создадим по умолчанию
        if not stock_data["materials"]:
            stock_data["materials"] = {
                "POM Valge": {"enabled": True, "stock_m3": 0.0, "remain_m3": 0.0},
                "POM Must":  {"enabled": True, "stock_m3": 0.0, "remain_m3": 0.0}
            }
            save_json(self.stock_path, stock_data)


        # строим Tkinter-переменные для GUI на основе JSON
        self.materials_cfg = {}
        for name, rec in stock_data["materials"].items():
            self.materials_cfg[name] = {
                "enabled": tk.BooleanVar(value=bool(rec.get("enabled", True))),
                "stock_m3": tk.StringVar(value=f'{float(rec.get("stock_m3", 0.0)):0.2f}'),
                "remain_m3": tk.StringVar(value=f'{float(rec.get("remain_m3", 0.0)):0.2f}'),
            }

        # Итоги конвертации в м² (по текущему месяцу) — для КАЖДОГО материала из JSON
        self.conv_totals = {name: tk.StringVar(value="0.00") for name in self.materials_cfg.keys()}

        # --- переменные GUI ---
        self.month_var = tk.StringVar()
        self.pdf_dir_var = tk.StringVar()
        self.make_default_var = tk.BooleanVar(value=False)  # чекбокс "сделать по умолчанию"

        # --- загрузка конфига / дефолтов ---
        self._load_defaults()

        # --- построение интерфейса ---
        self._build_ui()

        # --- обработчик закрытия: сохранить конфиг ---
        self.master.protocol("WM_DELETE_WINDOW", self._on_exit)

    # ---------------- UI ----------------

    def _build_ui(self):
        self.master.minsize(720, 360)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=16, pady=16)

        # Разделитель «-»
        ttk.Label(container, text="-").grid(row=1, column=1, padx=(64, 0), sticky="w")

        # Ряд 1: Месяц (YYYY-MM)
        ttk.Label(container, text="Date :").grid(row=1, column=0, sticky="w")
        month_entry = ttk.Entry(container, textvariable=self.month_var, width=12)
        month_entry.grid(row=1, column=1, sticky="w")
        
        # Месяцы 01..12
        months = [f"{m:02d}" for m in range(1, 13)]
        cb_month = ttk.Combobox(container, textvariable=self.month_num_var, values=months, width=4, state="readonly")
        cb_month.grid(row=1, column=1, padx=(80, 0), sticky="w")
        cb_month.bind("<<ComboboxSelected>>", self._on_date_change)
        
        # Годы: текущий-10 … текущий+1
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 10, current_year + 2)]
        cb_year = ttk.Combobox(container, textvariable=self.year_var, values=years, width=6, state="readonly")
        cb_year.grid(row=1, column=1, sticky="w")
        cb_year.bind("<<ComboboxSelected>>", self._on_date_change)

        # Ряд 2: Папка PDF
        ttk.Label(container, text="Папка PDF:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        pdf_entry = ttk.Entry(container, textvariable=self.pdf_dir_var, width=60)
        pdf_entry.grid(row=2, column=1, sticky="we", pady=(8, 0))
        choose_btn = ttk.Button(container, text="Выбрать…", command=self._choose_pdf_dir)
        choose_btn.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=(8, 0))

        # Ряд 3: чекбокс "сделать по умолчанию"
        default_cb = ttk.Checkbutton(
            container, text="Сделать по умолчанию", variable=self.make_default_var
        )
        default_cb.grid(row=3, column=1, sticky="w", pady=(4, 12))

        # Ряд 5: Два списка (PDF и материалы)
        lists_frame = ttk.Frame(container)
        lists_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(16, 0))
        container.rowconfigure(5, weight=1)
        container.columnconfigure(1, weight=1)
        lists_frame.columnconfigure(0, weight=1)
        lists_frame.columnconfigure(1, weight=1)
        lists_frame.rowconfigure(1, weight=1)

        # Левая панель — PDF
        pdf_group = ttk.LabelFrame(lists_frame, text="PDF-файлы (фильтр по месяцу)")
        pdf_group.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        pdf_group.columnconfigure(0, weight=1)
        pdf_group.rowconfigure(1, weight=1)

        self.pdf_count_lbl = ttk.Label(pdf_group, text="Найдено: 0")
        self.pdf_count_lbl.grid(row=0, column=0, sticky="w", padx=8, pady=(4, 0))

        self.pdf_list = tk.Listbox(pdf_group, height=12)
        pdf_scroll = ttk.Scrollbar(pdf_group, orient="vertical", command=self.pdf_list.yview)
        self.pdf_list.configure(yscrollcommand=pdf_scroll.set)
        self.pdf_list.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=8)
        pdf_scroll.grid(row=1, column=1, sticky="ns", pady=8)

        # Правая панель — Материалы из PDF (позже заполним реальным парсером)
        mat_group = ttk.LabelFrame(lists_frame, text="Перечень заказного материала (по выбранному месяцу)")
        mat_group.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        mat_group.columnconfigure(0, weight=1)
        mat_group.rowconfigure(1, weight=1)

        self.mat_count_lbl = ttk.Label(mat_group, text="Позиций: 0")
        self.mat_count_lbl.grid(row=0, column=0, sticky="w", padx=8, pady=(4, 0))

        # Таблица материалов
        cols = ("desc", "qty", "po", "date")
        self.mat_tree = ttk.Treeview(mat_group, columns=cols, show="headings", height=12)
        self.mat_tree.heading("desc", text="Material / Size")
        self.mat_tree.heading("qty", text="Qty")
        self.mat_tree.heading("po", text="PO No.")
        self.mat_tree.heading("date", text="Date")

        # ширины/выравнивание по минималке
        self.mat_tree.column("desc", width=220, anchor="w")
        self.mat_tree.column("qty", width=60, anchor="e")
        self.mat_tree.column("po", width=100, anchor="center")
        self.mat_tree.column("date", width=100, anchor="center")

        mat_scroll = ttk.Scrollbar(mat_group, orient="vertical", command=self.mat_tree.yview)
        self.mat_tree.configure(yscrollcommand=mat_scroll.set)
        self.mat_tree.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=8)
        mat_scroll.grid(row=1, column=1, sticky="ns", pady=8)
        # --- Нижняя панель: слева — склад Kumex (ввод), справа — конвертация в м² (вывод) ---
        bottom_frame = ttk.Frame(container)
        bottom_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(8, 0))

        bottom_frame.columnconfigure(0, weight=0)  # левая узкая колонка
        bottom_frame.columnconfigure(1, weight=1)  # правая широкая колонка

        # Левая группа (под левым списком): настройки склада Kumex
        cfg_group = ttk.LabelFrame(bottom_frame, text="Склад Kumex — материалы для учёта (м²)")
        cfg_group.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        for c in range(4):
            cfg_group.columnconfigure(c, weight=0)

        # Заголовки
        ttk.Label(cfg_group, text="Материал").grid(row=0, column=0, padx=8, pady=(6, 2), sticky="w")
        ttk.Label(cfg_group, text="Склад, м²").grid(row=0, column=2, padx=8, pady=(6, 2), sticky="w")
        ttk.Label(cfg_group, text="Остаток, м²").grid(row=0, column=3, padx=8, pady=(6, 2), sticky="w")

        # Ряды по материалам
        # Ряды по материалам ИЗ JSON (с сохранением флага enabled)
        _row = 1
        for name, cfg in self.materials_cfg.items():
            ttk.Label(cfg_group, text=name).grid(row=_row, column=0, padx=8, pady=2, sticky="w")
            e_stock = ttk.Entry(cfg_group, textvariable=cfg["stock_m3"], width=10)
            e_stock.grid(row=_row, column=2, padx=8, pady=2, sticky="w")

            e_rem = ttk.Entry(cfg_group, textvariable=cfg["remain_m3"], width=10, state="readonly")
            e_rem.grid(row=_row, column=3, padx=8, pady=2, sticky="w")

            _row += 1


        conv_group = ttk.LabelFrame(bottom_frame, text="Конвертация (по выбранному месяцу)")
        conv_group.grid(row=0, column=1, sticky="nsew")
        # 0: имя материала, 1: значение, 2: "м²", 3: спейсер для растяжения
        conv_group.columnconfigure(0, weight=0)
        conv_group.columnconfigure(1, weight=0)
        conv_group.columnconfigure(2, weight=0)
        conv_group.columnconfigure(3, weight=1)


        # Строки итогов (только вывод) — динамически для всех материалов
        _row_conv = 0
        for name in self.conv_totals.keys():
            pad_top = (6 if _row_conv == 0 else 2, 2)

            ttk.Label(conv_group, text=f"{name}:").grid(
                row=_row_conv, column=0, padx=8, pady=pad_top, sticky="w"
            )
            # значение ближе к "м²": уменьшаем отступ справа у Entry
            ttk.Entry(
                conv_group, textvariable=self.conv_totals[name],
                width=10, state="readonly"
            ).grid(
                row=_row_conv, column=1, padx=(8, 2), pady=pad_top, sticky="w"
            )
            # "м²" почти вплотную к полю
            ttk.Label(conv_group, text="м²").grid(
                row=_row_conv, column=2, padx=(0, 8), pady=pad_top, sticky="w"
            )

            _row_conv += 1



        
        # Кнопка "Рассчитать" внизу окна
        calc_btn = ttk.Button(container, text="Рассчитать", command=self._apply_stub)
        calc_btn.grid(row=7, column=0, columnspan=3, sticky="e", padx=16, pady=(12, 8))

        # Статус-бар
        
        status = ttk.Label(self, textvariable=self.status_var, anchor="w")
        status.pack(fill="x", padx=16, pady=(4, 8))


        # колонки
        container.columnconfigure(1, weight=1)
        
        self._scan_pdfs()

    # ------------- Загрузка/сохранение конфигурации -------------

    def _load_defaults(self):
        cfg = load_json(self.config_path, default={})
        current_month = datetime.now().strftime("%Y-%m")
        # всегда системный месяц
        target = current_month
        yy, mm = target.split("-")
        self.month_var.set(target)
        self.year_var.set(yy)
        self.month_num_var.set(mm)


        # Папка PDF: по умолчанию data/input_pdf
        default_pdf_dir = (self.project_root / "data" / "input_pdf").as_posix()
        self.pdf_dir_var.set(cfg.get("pdf_dir", default_pdf_dir))

    def _save_config(self):
        self.state_dir.mkdir(parents=True, exist_ok=True)
        data = {
            "pdf_dir": self.pdf_dir_var.get().strip()
        }
        save_json(self.config_path, data)
        # ---------------- Служебные обработчики ----------------

    def _choose_pdf_dir(self):
        initial = self.pdf_dir_var.get() or (self.project_root / "data" / "input_pdf").as_posix()
        chosen = filedialog.askdirectory(initialdir=initial, title="Выбрать папку с PDF")
        if chosen:
            self.pdf_dir_var.set(Path(chosen).as_posix())
            # Если стоит чекбокс — сразу записываем как дефолт (минимальная логика)
            if self.make_default_var.get():
                self._save_config()
                messagebox.showinfo("Kumex", "Папка сохранена как путь по умолчанию.")

    def _scan_pdfs(self):
    
        folder = Path(self.pdf_dir_var.get()).expanduser()
        if not folder.exists():
            messagebox.showerror("Папка недоступна", f"Папка не существует:\n{folder}")
            return

        # читаем год/месяц из комбобоксов
        yy_str = self.year_var.get().strip()
        mm_str = self.month_num_var.get().strip()
        if not (yy_str.isdigit() and mm_str.isdigit()):
            messagebox.showerror("Неверная дата", f"Ожидается формат YYYY-MM, получено: {yy_str}-{mm_str}")
            return
        yy = int(yy_str)
        mm = int(mm_str)

        # границы месяца (локальное время)
        start_ts = time.mktime((yy, mm, 1, 0, 0, 0, 0, 0, -1))
        if mm == 12:
            end_ts = time.mktime((yy + 1, 1, 1, 0, 0, 0, 0, 0, -1))
        else:
            end_ts = time.mktime((yy, mm + 1, 1, 0, 0, 0, 0, 0, -1))

        # сбор PDF в папке
        pdfs = []
        for p in folder.glob("*.pdf"):
            try:
                mtime = p.stat().st_mtime
            except Exception:
                continue
            if start_ts <= mtime < end_ts:
                pdfs.append(p)

        # обновляем левый список
        self.pdf_files = sorted(pdfs, key=lambda x: x.name.lower())
        self.pdf_list.delete(0, tk.END)
        for p in self.pdf_files:
            self.pdf_list.insert(tk.END, p.name)
        self.pdf_count_lbl.config(text=f"Найдено: {len(self.pdf_files)}")

        # чистим правый список (материалы добавим позже)
        self.material_rows = []
        for iid in self.mat_tree.get_children():
            self.mat_tree.delete(iid)
        self.mat_count_lbl.config(text="Позиций: 0")
        
        self._parse_materials()

        self._set_status(f"Папка: {folder} | PDF за {yy}-{mm_str}: {len(self.pdf_files)}")
        self._calc_m2()

    def _parse_materials(self):
    
        # очистка списка перед циклом (важно не чистить внутри)
        for iid in self.mat_tree.get_children():
            self.mat_tree.delete(iid)
        self.material_rows = []

        if not self.pdf_files:
            self._set_status("Сначала нажмите «Сканировать PDF».")
            self.mat_count_lbl.config(text="Позиций: 0")
            return

        # --- паттерны ---
        # размеры: 22x22x1000, 40*67*1000, 20x20 (mm необяз.)
        size_rx = re.compile(r"\b\d+\s*([xX*])\s*\d+(?:\s*\1\s*\d+)?(?:\s*mm\b)?")
        # qty на той же/след. строке: qty=2000, QTY 2000, Quantity: 70
        qty_any_rx = re.compile(r"\b(?:qty|quantity)\s*[:=]?\s*(\d{1,7})\b", re.IGNORECASE)
        # табличный вариант в строке выше: "1 70056 2000 ..." (индекс, partno, qty)
        qty_row_above_rx = re.compile(r"^\s*\d+\s+\S+\s+(\d{1,7})\b")
        # PO и дата (разные варианты написания)
        po_rx = re.compile(r"\bPO(?:\s*(?:Number|No\.?)|)\s*[:#]?\s*([A-Za-z0-9_-]+)")
        date_rx = re.compile(r"\b(?:Order\s*Date|Date)\s*[:#]?\s*([0-9]{1,2}[.\-/][0-9]{1,2}[.\-/][0-9]{2,4})")

        total = 0

        for p in self.pdf_files:
            text = read_pdf_text(str(p))
            if not text:
                continue

            po = (po_rx.search(text).group(1) if po_rx.search(text) else "?")
            od = (date_rx.search(text).group(1) if date_rx.search(text) else "?")

            lines = [ln.strip() for ln in text.splitlines()]
            n = len(lines)

            for i, line in enumerate(lines):
                # ищем строку-описание по наличию размеров
                if not size_rx.search(line):
                    continue

                desc = line
                qty = None

                # 1) qty на этой же строке
                m = qty_any_rx.search(line)
                if m:
                    qty = int(m.group(1))

                # 2) если не нашли — ищем на 1..3 строках ниже
                if qty is None:
                    for j in range(1, 4):
                        if i + j >= n:
                            break
                        m = qty_any_rx.search(lines[i + j])
                        if m:
                            qty = int(m.group(1))
                            break

                # 3) если не нашли — пробуем табличный вариант в 1..2 строках выше
                if qty is None:
                    for j in (1, 2):
                        if i - j >= 0:
                            m = qty_row_above_rx.match(lines[i - j])
                            if m:
                                qty = int(m.group(1))
                                break

                if qty is None:
                    # не смогли уверенно найти количество — пропускаем позицию
                    continue

                # сохраняем как словарь (на будущее для отчётов)
                # определяем материал
                mat_name = ""
                if "valge" in desc.lower():
                    mat_name = "POM Valge"
                elif "must" in desc.lower() or "õhuke" in desc.lower():
                    mat_name = "POM Must"

                # сохраняем как словарь (для последующего расчёта)
                self.material_rows.append({
                    "desc": desc,
                    "qty": qty,
                    "po": po,
                    "date": od,
                    "material": mat_name
                })
                # добавляем строку в таблицу
                self.mat_tree.insert("", "end", values=(desc, qty, po, od))
                total += 1

        self.mat_count_lbl.config(text=f"Позиций: {total}")
        self._set_status(f"Разобрано PDF: {len(self.pdf_files)} | Найдено позиций: {total}")

    def _on_date_change(self, *_):
        self._sync_month_var()
        self._scan_pdfs()

    def _report_stub(self):
        # Заглушка: здесь позже будет агрегатор + генерация отчёта
        messagebox.showinfo("Отчёт", "Формирование отчёта (заглушка).")

    def _sync_month_var(self):
        yy = self.year_var.get()
        mm = self.month_num_var.get()
        if len(yy) == 4 and mm in {f"{i:02d}" for i in range(1,13)}:
            self.month_var.set(f"{yy}-{mm}")
            
    def _calc_m2(self):
        """Пересчитывает площади (м²) по материалам на основе таблицы заказов."""
        import re
        from decimal import Decimal, ROUND_HALF_UP

        totals = {"POM Valge": Decimal("0.0"), "POM Must": Decimal("0.0")}

        for row in self.material_rows:
            desc = str(row.get("desc", ""))
            qty = int(row.get("qty", 0) or 0)
            material = row.get("material", "")

            # считаем только нужные материалы
            if not any(x in material for x in ["POM Valge", "POM Must"]):
                continue

            # извлечь все числа
            nums = [int(x) for x in re.findall(r"\d+", desc)]
            if len(nums) < 3:
                continue

            # считаем количество "52"
            cnt_52 = nums.count(52)

            if cnt_52 == 1:
                sides = [x for x in nums if x != 52]
                if len(sides) >= 2:
                    A, B = sides[0], sides[1]
                else:
                    continue
            elif cnt_52 == 2:
                non52 = [x for x in nums if x != 52]
                if len(non52) == 1:
                    A, B = 52, non52[0]
                else:
                    continue
            elif cnt_52 == 3:
                A, B = 52, 52
            else:  # ни одного 52 → берём 2 самые большие стороны
                sides = sorted(nums, reverse=True)
                A, B = sides[0], sides[1]

            # площадь одной детали
            S1 = Decimal(A) * Decimal(B) / Decimal(1_000_000)
            # площадь всех qty
            Spos = S1 * qty

            # округление только для вывода (сумму храним точно)
            if "Valge" in material:
                totals["POM Valge"] += Spos
            elif "Must" in material:
                totals["POM Must"] += Spos

        # обновляем GUI (с двумя знаками)
        for name, value in totals.items():
            rounded = value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            self.conv_totals[name].set(str(rounded))

    def _apply_stub(self):
        # Заглушка: здесь позже будет применение к kumex_stock.json
        messagebox.showinfo("Применение", "Применение отчёта к остаткам (заглушка).")

    def _on_exit(self):
        # При выходе всегда сохраняем last_month и текущий pdf_dir
        try:
            self._save_config()
        finally:
            self.master.destroy()

    def _set_status(self, text: str):
        self.status_var.set(text)
