

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
        # Толщина пилы (мм) — влияет на конвертацию в м²
        self.kerf_mm_var = tk.StringVar(value="0.00")
        # При изменении значения пересчитываем конвертацию
        self.kerf_mm_var.trace_add("write", lambda *a: self._calc_m2())
        # Окно "Настройка склада" (пересоздаём по мере закрытия)
        self._stock_win = None


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
        
        # при старте подтянуть остатки из JSON
        _data = self._load_stock_data()
        self._recompute_balances_from_ledger(_data)
        self._save_stock_data(_data)
        
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
        ttk.Label(cfg_group, text="Остаток, м²").grid(row=0, column=3, padx=8, pady=(6, 2), sticky="w")

        # Ряды по материалам
        # Ряды по материалам ИЗ JSON (с сохранением флага enabled)
        _row = 1
        for name, cfg in self.materials_cfg.items():
            ttk.Label(cfg_group, text=name).grid(row=_row, column=0, padx=8, pady=2, sticky="w")

            e_rem = ttk.Entry(cfg_group, textvariable=cfg["remain_m3"], width=10, state="readonly")
            e_rem.grid(row=_row, column=3, padx=8, pady=2, sticky="w")

            _row += 1
            
        # Кнопка "Настройка склада…" по центру снизу левого блока
        btn_row = ttk.Frame(cfg_group)
        btn_row.grid(row=_row, column=0, columnspan=3, sticky="ew", padx=8, pady=(8, 6))
        btn_row.columnconfigure(0, weight=1)
        ttk.Button(btn_row, text="Настройка склада…", command=self._open_stock_dialog).pack(anchor="center")


        conv_group = ttk.LabelFrame(bottom_frame, text="Конвертация (по выбранному месяцу)")
        conv_group.grid(row=0, column=1, sticky="nsew")
        # 0: имя материала, 1: значение, 2: "м²", 3: спейсер для растяжения
        conv_group.columnconfigure(0, weight=0)
        conv_group.columnconfigure(1, weight=0)
        conv_group.columnconfigure(2, weight=0)


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
            
                # Толщина пилы (мм)# Разделитель и поле "Толщина пилы"
        ttk.Separator(conv_group, orient="horizontal").grid(
            row=_row_conv, column=0, columnspan=3, sticky="ew", padx=8, pady=(6, 4)
        )

        ttk.Label(conv_group, text="Толщина пилы:").grid(
            row=_row_conv + 1, column=0, padx=8, pady=2, sticky="w"
        )
        e_kerf = ttk.Entry(conv_group, textvariable=self.kerf_mm_var, width=6)
        e_kerf.grid(
            row=_row_conv + 1, column=1, padx=(8, 2), pady=2, sticky="w"
        )
        # автосейв при выходе из поля и при Enter
        e_kerf.bind("<FocusOut>", lambda _e: self._save_config())
        e_kerf.bind("<Return>",  lambda _e: (self._save_config(), "break"))

        ttk.Label(conv_group, text="мм").grid(
            row=_row_conv + 1, column=2, padx=(0, 8), pady=2, sticky="w"
        )

                # Кнопка "Рассчитать" внутри рамки конвертации, снизу справа
        calc_btn = ttk.Button(conv_group, text="Рассчитать", command=self._apply_stub)
        calc_btn.grid(row=_row_conv + 2, column=0, columnspan=3, sticky="e", padx=8, pady=(10, 6))
        self._calc_btn = calc_btn  # пригодится позже, когда будем блокировать закрытые месяцы

        # первичная проверка (на случай, если месяц уже закрыт)
        self._update_calc_button_state()


        # Статус-бар
        
        status = ttk.Label(self, textvariable=self.status_var, anchor="w")
        status.pack(fill="x", padx=16, pady=(4, 8))


        # колонки
        container.columnconfigure(1, weight=1)
        
        self._scan_pdfs()
        # первичная проверка (на случай, если месяц уже закрыт)
        

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
        self.kerf_mm_var.set(str(cfg.get("kerf_mm", self.kerf_mm_var.get() or "1")))

        # Папка PDF: по умолчанию data/input_pdf
        default_pdf_dir = (self.project_root / "data" / "input_pdf").as_posix()
        self.pdf_dir_var.set(cfg.get("pdf_dir", default_pdf_dir))

    def _save_config(self):
        self.state_dir.mkdir(parents=True, exist_ok=True)
        data = {
            "pdf_dir": self.pdf_dir_var.get().strip()
            , "kerf_mm": float(self.kerf_mm_var.get().strip() or "1")
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
        self._update_calc_button_state()

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
                d = desc.lower()
                if "pom" in d and "valge" in d:
                    mat_name = "POM Valge"
                elif "pom" in d and ("must" in d or "õhuke" in d):
                    mat_name = "POM Must"
                else:
                    mat_name = ""  # всё остальное (PET, Messing, ESD-only и т.д.) не считать
                # + отдельно: если "esd" in d: mat_name = ""


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
        self._update_calc_button_state()

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
            # Исключаем ESD-материалы из калькуляции (например, "ESD_POM_valge 52x52x1000")
            

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

            # Толщина пилы (мм), безопасный парсинг
            try:
                kerf = Decimal(str(self.kerf_mm_var.get()).replace(",", "."))
                if kerf < 0:
                    kerf = Decimal("0")
            except Exception:
                kerf = Decimal("0")

            A_eff = Decimal(A) + kerf
            B_eff = Decimal(B) + kerf

            # Площадь одной детали с учётом пилы
            S1 = (A_eff * B_eff) / Decimal(1_000_000)
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

    def _open_stock_dialog(self):
        """Открывает окно 'Настройка склада' с таблицей журнала операций."""
        if self._stock_win and tk.Toplevel.winfo_exists(self._stock_win):
            self._stock_win.focus_set()
            return

        self._stock_win = tk.Toplevel(self.master)
        self._stock_win.title("Настройка склада")
        self._stock_win.transient(self.master)
        self._stock_win.grab_set()
        self._stock_win.geometry("720x420")

        def _on_close():
            try:
                self._stock_win.grab_release()
            except Exception:
                pass
            self._stock_win.destroy()
            self._stock_win = None

        self._stock_win.protocol("WM_DELETE_WINDOW", _on_close)

                # Корневой фрейм
        frm = ttk.Frame(self._stock_win)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        # ===== ФОРМА ОПЕРАЦИИ (+/- м²) =====
        op_box = ttk.LabelFrame(frm, text="Операция со складом")
        op_box.pack(fill="x", padx=0, pady=(0, 8))

        # переменные формы
        self._op_material = tk.StringVar(value="POM Valge")
        self._op_amount = tk.StringVar(value="0.00")
        self._op_note = tk.StringVar(value="")
        self._op_type = tk.StringVar(value="add")  # add | sub

        ttk.Label(op_box, text="Материал:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Combobox(op_box, textvariable=self._op_material, state="readonly",
                    values=("POM Valge", "POM Must"), width=12).grid(row=0, column=1, padx=(0, 8), pady=6, sticky="w")

        ttk.Label(op_box, text="Количество:").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        ttk.Entry(op_box, textvariable=self._op_amount, width=10).grid(row=0, column=3, padx=(0, 4), pady=6, sticky="w")
        ttk.Label(op_box, text="м²").grid(row=0, column=4, padx=(0, 8), pady=6, sticky="w")

        ttk.Label(op_box, text="Тип действия:").grid(row=0, column=5, padx=10, pady=6, sticky="w")
        ttk.Radiobutton(op_box, text="+ Пополнение", variable=self._op_type, value="add").grid(row=0, column=6, padx=(0, 8), pady=6, sticky="w")
        ttk.Radiobutton(op_box, text="– Списание",   variable=self._op_type, value="sub").grid(row=0, column=7, padx=(0, 8), pady=6, sticky="w")


        ttk.Label(op_box, text="Комментарий:").grid(row=1, column=0, padx=8, pady=(0, 8), sticky="w")
        ttk.Entry(op_box, textvariable=self._op_note, width=60).grid(row=1, column=1, columnspan=5, padx=(0, 8), pady=(0, 8), sticky="w")

        ttk.Button(op_box, text="Добавить операцию", command=self._add_stock_operation)\
            .grid(row=1, column=6, padx=8, pady=(0, 8), sticky="e")

        for i in range(8):
            op_box.columnconfigure(i, weight=0)
        op_box.columnconfigure(5, weight=1)  # чтобы поле комментария имело место

        # ===== УДАЛЕНИЕ РАСЧЁТА МЕСЯЦА =====
        rm_box = ttk.LabelFrame(frm, text="Удалить расчёт месяца (month_calc)")
        rm_box.pack(fill="x", padx=0, pady=(0, 8))

        self._rm_year = tk.StringVar(value=self.year_var.get())
        self._rm_month = tk.StringVar(value=self.month_num_var.get())

        ttk.Label(rm_box, text="Год:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Combobox(rm_box, textvariable=self._rm_year, values=[str(y) for y in range(2017, 2031)],
                    state="readonly", width=6).grid(row=0, column=1, padx=(0, 8), pady=6, sticky="w")
        ttk.Label(rm_box, text="Месяц:").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        ttk.Combobox(rm_box, textvariable=self._rm_month, values=[f"{m:02d}" for m in range(1, 13)],
                    state="readonly", width=4).grid(row=0, column=3, padx=(0, 8), pady=6, sticky="w")

        ttk.Button(rm_box, text="Удалить расчёт месяца", command=self._delete_month_calc)\
            .grid(row=0, column=4, padx=8, pady=6, sticky="e")

        rm_box.columnconfigure(5, weight=1)

        # ===== ЖУРНАЛ ОПЕРАЦИЙ =====
        topbar = ttk.Frame(frm)
        topbar.pack(fill="x", pady=(6, 0))
        ttk.Label(topbar, text="Журнал операций склада").pack(side="left")

        # Кнопка Undo (отмена выбранной записи)
        ttk.Button(topbar, text="Отменить действие", command=self._undo_selected).pack(side="right", padx=(4, 0))
        # Кнопка Обновить — правее Undo
        ttk.Button(topbar, text="Обновить", command=self._reload_ledger).pack(side="right", padx=(4, 0))

        cols = ("ts", "month", "material", "action", "amount", "note")
        self._ledger_tree = ttk.Treeview(frm, columns=cols, show="headings", height=12)
        self._ledger_tree.heading("ts", text="Дата/время")
        self._ledger_tree.heading("month", text="Месяц")
        self._ledger_tree.heading("material", text="Материал")
        self._ledger_tree.heading("action", text="Действие")
        self._ledger_tree.heading("amount", text="м²")
        self._ledger_tree.heading("note", text="Комментарий")

        self._ledger_tree.column("ts", width=140, anchor="center")
        self._ledger_tree.column("month", width=80, anchor="center")
        self._ledger_tree.column("material", width=120, anchor="w")
        self._ledger_tree.column("action", width=120, anchor="w")
        self._ledger_tree.column("amount", width=80, anchor="e")
        self._ledger_tree.column("note", width=220, anchor="w")
        
        # Цветовые теги для строк журнала
        self._ledger_tree.tag_configure("t_add",   foreground="#0A7D00")  # зелёный  (пополнение)
        self._ledger_tree.tag_configure("t_sub",   foreground="#A40000")  # красный  (списание)
        self._ledger_tree.tag_configure("t_month", foreground="#004A9F")  # синий    (расчёт месяца)

        scr = ttk.Scrollbar(frm, orient="vertical", command=self._ledger_tree.yview)
        self._ledger_tree.configure(yscrollcommand=scr.set)
        self._ledger_tree.pack(side="left", fill="both", expand=True, pady=(6, 0))
        scr.pack(side="left", fill="y", pady=(6, 0))

        # статус внизу
        self._stock_status = tk.StringVar(value="")
        ttk.Label(self._stock_win, textvariable=self._stock_status).pack(fill="x", padx=10, pady=(6, 10))

        # загрузка данных в таблицу
        self._reload_ledger()

    def _load_stock_data(self):
        """Гарантированно читаем JSON структуры склада."""
        data = load_json(self.stock_path, default={})
        data.setdefault("materials", {})
        for k in ("POM Valge", "POM Must"):
            data["materials"].setdefault(k, {})
            data["materials"][k].setdefault("stock_m3", 0.0)   # тут «м3» — просто имя ключа; фактически м²
            data["materials"][k].setdefault("remain_m3", 0.0)  # фактически м²
        data.setdefault("ledger", [])
        data.setdefault("closed_months", [])
        return data

    def _save_stock_data(self, data):
        save_json(self.stock_path, data)

    def _recompute_balances_from_ledger(self, data):
    
        from decimal import Decimal, ROUND_HALF_UP
        sums = {"POM Valge": Decimal("0.0"), "POM Must": Decimal("0.0")}
        for rec in data.get("ledger", []):
            mat = rec.get("material")
            if mat not in sums:
                continue
            amt = Decimal(str(rec.get("amount_m2", 0)))
            typ = (rec.get("type") or "").lower()
            if typ == "manual_add":
                sums[mat] += amt
            elif typ == "manual_sub":
                sums[mat] -= amt
            elif typ == "month_calc":
                sums[mat] -= amt
        # обновим JSON + GUI
        for mat in ("POM Valge", "POM Must"):
            v = sums[mat].quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            data["materials"][mat]["remain_m3"] = float(v)
            # если есть поля на форме — обновим
            if mat in self.materials_cfg:
                self.materials_cfg[mat]["remain_m3"].set(str(v))

    def _add_stock_operation(self):
        """Добавить запись в журнал: manual_add/manual_sub, пересчитать остатки."""
        from decimal import Decimal
        import datetime as _dt

        mat = self._op_material.get()
        raw = str(self._op_amount.get()).replace(",", ".").strip()
        note = self._op_note.get().strip()
        typ = self._op_type.get()  # add | sub

        # валидация
        try:
            amount = Decimal(raw)
        except Exception:
            messagebox.showerror("Ошибка", "Введите корректное число в поле 'Количество, м²'.")
            return
        if amount <= 0:
            messagebox.showerror("Ошибка", "Количество должно быть больше 0.")
            return
        if mat not in ("POM Valge", "POM Must"):
            messagebox.showerror("Ошибка", "Выберите материал.")
            return

        # читаем/обновляем JSON
        data = self._load_stock_data()
        rec = {
            "ts": _dt.datetime.now().isoformat(timespec="seconds"),
            "month": f"{self.year_var.get()}-{self.month_num_var.get()}",
            "material": mat,
            "type": "manual_add" if typ == "add" else "manual_sub",
            "amount_m2": float(amount),
            "note": note or "Ручная операция"
        }
        data["ledger"].append(rec)

        # пересчёт и сохранение
        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)

        # обновим GUI
        self._reload_ledger()
        self._stock_status.set("Операция добавлена.")

    def _delete_month_calc(self):
        """Удалить все записи type=month_calc за выбранный месяц и пересчитать остатки."""
        import datetime as _dt

        yy = str(self._rm_year.get())
        mm = str(self._rm_month.get())
        month = f"{yy}-{mm}"

        if not messagebox.askyesno("Подтверждение", f"Удалить расчёт за {month}?"):
            return

        data = self._load_stock_data()
        # удалить month_calc за месяц
        before = len(data.get("ledger", []))
        data["ledger"] = [r for r in data.get("ledger", []) if not (r.get("type") == "month_calc" and r.get("month") == month)]
        after = len(data["ledger"])

        # убрать месяц из закрытых
        if "closed_months" in data and month in data["closed_months"]:
            data["closed_months"].remove(month)

        # пересчёт и сохранение
        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)

        self._reload_ledger()
        self._stock_status.set(f"Удалено записей: {before - after}.")
        self._update_calc_button_state()

    def _undo_selected(self):
        """Инвертировать выбранную запись журнала (manual_add <-> manual_sub)."""
        sel = getattr(self, "_ledger_tree", None).selection()
        if not sel:
            messagebox.showwarning("Отмена", "Выберите запись в журнале.")
            return

        iid = sel[0]
        meta = getattr(self, "_ledger_index", {}).get(iid)
        if not meta:
            messagebox.showerror("Ошибка", "Не удалось найти данные по выбранной записи.")
            return

        raw_type = (meta.get("type") or "").lower()
        if raw_type not in ("manual_add", "manual_sub"):
            messagebox.showinfo("Отмена записи",
                                "Эту запись нельзя отменить таким образом.\n"
                                "Для расчётов месяца используйте 'Удалить расчёт месяца'.")
            return

        material = meta.get("material")
        month = meta.get("month")
        amount = meta.get("amount")
        if amount <= 0:
            messagebox.showerror("Ошибка", "Количество в записи равно 0.")
            return

        # Подтверждение
        pretty_action = "Пополнение" if raw_type == "manual_add" else "Списание"
        if not messagebox.askyesno("Подтверждение",
                                f"Отменить запись:\n"
                                f"Материал: {material}\n"
                                f"Действие: {pretty_action}\n"
                                f"Объём: {amount} м²\n\n"
                                f"Будет создана обратная корректировка."):
            return

        # Готовим инверсию
        inverse_type = "manual_sub" if raw_type == "manual_add" else "manual_add"

        # Читаем JSON, добавляем корректировку, пересчитываем
        data = self._load_stock_data()

        import datetime as _dt
        data["ledger"].append({
            "ts": _dt.datetime.now().isoformat(timespec="seconds"),
            "month": month,
            "material": material,
            "type": inverse_type,
            "amount_m2": float(amount),
            "note": f"Undo выбранной записи ({pretty_action})",
        })

        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)

        # Обновление UI
        self._reload_ledger()
        self._stock_status.set("Добавлена обратная корректировка.")

    def _reload_ledger(self):
        """Читает ledger из JSON и перерисовывает таблицу журнала."""
        try:
            data = load_json(self.stock_path, default={"ledger": []})
            ledger = data.get("ledger", []) or []
        except Exception as e:
            ledger = []
            if hasattr(self, "_stock_status"):
                self._stock_status.set(f"Ошибка чтения JSON: {e}")
        # Маппинг iid -> «сырые» значения для undo
        self._ledger_index = {}
        
        # дерево журнала может не существовать (диалог закрыт)
        tree = getattr(self, "_ledger_tree", None)
        tree_exists = bool(tree) and tree.winfo_exists()

        status_var = getattr(self, "_stock_status", None)
        stock_win_alive = bool(getattr(self, "_stock_win", None)) and \
                        tk.Toplevel.winfo_exists(getattr(self, "_stock_win"))


        if tree_exists:
            for iid in tree.get_children():
                tree.delete(iid)


        from decimal import Decimal, ROUND_HALF_UP
        import datetime as _dt

        def _fmt_amount(v):
            try:
                return str(Decimal(v).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
            except Exception:
                return "0.00"

        cnt = 0
        for rec in ledger:
            ts = rec.get("ts") or ""
            # форматируем ISO в читабельный
            try:
                if "T" in ts:
                    dt = _dt.datetime.fromisoformat(ts)
                    ts_disp = dt.strftime("%Y-%m-%d %H:%M:%S")
                else:
                    ts_disp = ts
            except Exception:
                ts_disp = ts

            month = rec.get("month", "")
            material = rec.get("material", "")
            typ = (rec.get("type", "") or "").lower()
            if   typ == "manual_add":
                action = "Пополнение"
            elif typ == "manual_sub":
                action = "Списание"
            elif typ == "month_calc":
                action = "Расчёт месяца"
            else:
                action = rec.get("type", "")

            amt = _fmt_amount(rec.get("amount_m2", 0))
            note = rec.get("note", "")

            if hasattr(self, "_ledger_tree"):
                # сырой тип для undo
                raw_type = (rec.get("type", "") or "").lower()

                # тег по типу записи
                if   raw_type == "manual_add":
                    row_tag = "t_add"
                elif raw_type == "manual_sub":
                    row_tag = "t_sub"
                elif raw_type == "month_calc":
                    row_tag = "t_month"
                else:
                    row_tag = ""

                # стабильный iid (чтобы строку можно было однозначно выбрать)
                iid = f"{ts}-{cnt}"

                # положим метаданные в индекс для Undo
                try:
                    amt_dec = Decimal(str(rec.get("amount_m2", 0) or 0))
                except Exception:
                    amt_dec = Decimal("0")
                self._ledger_index[iid] = {
                    "ts": ts,
                    "month": month,
                    "material": material,
                    "type": raw_type,
                    "amount": amt_dec,
                    "note": note,
                }

                # ... подготовка iid, row_tag, values как у тебя уже есть ...

                # Вставляем строку безопасно: окно могли закрыть прямо сейчас
                if tree_exists and tree.winfo_exists():
                    try:
                        tree.insert(
                            "", "end", iid=iid, tags=(row_tag,) if row_tag else (),
                            values=(ts_disp, month, material, action, amt, note)
                        )
                        cnt += 1
                    except tk.TclError:
                        # диалог закрылся в момент вставки — тихо выходим из цикла
                        break



        if status_var is not None and stock_win_alive:
            status_var.set(f"Записей в журнале: {cnt}")

    def _month_key(self) -> str:
        """Ключ месяца вида YYYY-MM из текущих комбобоксов."""
        return f"{self.year_var.get()}-{self.month_num_var.get()}"

    def _update_calc_button_state(self):
        """Включить/выключить кнопку 'Рассчитать' в зависимости от закрытого месяца."""
        if not hasattr(self, "_calc_btn"):
            return
        data = self._load_stock_data()
        closed = set(data.get("closed_months", []))
        mkey = self._month_key()
        if mkey in closed:
            self._calc_btn.configure(state="disabled")
            self.status_var.set(f"Месяц {mkey} закрыт: расчёт уже выполнен.")
        else:
            self._calc_btn.configure(state="normal")

    def _apply_stub(self):
        """Фиксируем расчёт месяца: пишем month_calc в журнал, закрываем месяц."""
        from decimal import Decimal, ROUND_HALF_UP
        import datetime as _dt

        mkey = self._month_key()
        data = self._load_stock_data()

        # Если месяц уже закрыт — не даём повторно
        if mkey in set(data.get("closed_months", [])):
            messagebox.showinfo("Расчёт уже выполнен", f"Месяц {mkey} закрыт. Расчёт был произведён ранее.")
            self._update_calc_button_state()
            return

        # Берём рассчитанные значения из правого блока (конвертация, м²)
        def _get(name: str) -> Decimal:
            var = self.conv_totals.get(name)
            if not var:
                return Decimal("0")
            raw = (var.get() or "0").replace(",", ".").strip()
            try:
                return Decimal(raw)
            except Exception:
                return Decimal("0")

        valge = _get("POM Valge")
        must  = _get("POM Must")

        # Если оба нули — предупредим и выйдем
        if valge == 0 and must == 0:
            messagebox.showwarning("Нет данных", "За выбранный месяц нет материалов для списания.")
            return

        # Запишем операции month_calc (только ненулевые)
        ts_now = _dt.datetime.now().isoformat(timespec="seconds")
        note = f"Авто: расчёт {mkey}"

        if valge > 0:
            data["ledger"].append({
                "ts": ts_now,
                "month": mkey,
                "material": "POM Valge",
                "type": "month_calc",
                "amount_m2": float(valge.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
                "note": note
            })
        if must > 0:
            data["ledger"].append({
                "ts": ts_now,
                "month": mkey,
                "material": "POM Must",
                "type": "month_calc",
                "amount_m2": float(must.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
                "note": note
            })

        # Закрываем месяц, если что-то записали
        if "closed_months" not in data:
            data["closed_months"] = []
        if any(r for r in data["ledger"] if r.get("ts") == ts_now and r.get("type") == "month_calc"):
            if mkey not in data["closed_months"]:
                data["closed_months"].append(mkey)

        # Пересчёт остатков и сохранение
        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)
        self._save_config()

        # Обновить журнал (если окно открыто) и заблокировать кнопку
        self._reload_ledger()
        self._update_calc_button_state()
        messagebox.showinfo("Готово", f"Расчёт за {mkey} зафиксирован.")

    def _on_exit(self):
        # При выходе всегда сохраняем last_month и текущий pdf_dir
        try:
            self._save_config()
        finally:
            self.master.destroy()

    def _set_status(self, text: str):
        self.status_var.set(text)
