
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
import time
import re
from decimal import Decimal, ROUND_HALF_UP
from kumex.io.pdf_reader import read_pdf_text
from kumex.io.file_ops import load_json, save_json

# --- Kumex расчет: базовые константы по спецификации ---
PLATE_W_MM = 1000
PLATE_L_MM = 2000
BASE_THICKNESS_MM = 52
STOCK_PLATES = 10
# Длина реза фиксирована на 1000 мм
DEFAULT_CUT_LEN_MM = 1000


class MainWindow(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master: tk.Tk = master
        self.year_var = tk.StringVar()
        self.month_num_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Valmis")
        # Толщина пилы (мм) — влияет на конвертацию в м²
        self.kerf_mm_var = tk.StringVar(value="0.00")
        self.cut_len_var = tk.StringVar(value=str(DEFAULT_CUT_LEN_MM))
        self.period_mode_var = tk.BooleanVar(value=False)
        self.period_from_var = tk.StringVar()
        self.period_to_var = tk.StringVar()
        # При изменении значения пересчитываем конвертацию
        self.kerf_mm_var.trace_add("write", lambda *a: self._calc_m2())
        # Окно "Настройка склада" (пересоздаём по мере закрытия)
        self._stock_win = None
        self._stock_win_binds = {}


        # Настройки склада Kumex (минимум: два материала) — пока без логики

        self.pdf_files = []          # список путей найденных PDF
        self.material_rows = []      # сюда позже положим строки из PDF-парсера


            # --- пути и состояние ---
        # База состояния: %APPDATA%\Kumex
        appdata = Path(os.getenv("APPDATA") or Path.home() / "AppData" / "Roaming")
        self.state_dir = appdata / "Kumex"
        self.state_dir.mkdir(parents=True, exist_ok=True)

        # Папка по умолчанию для PDF внутри APPDATA (можно менять в GUI)
        (self.state_dir / "input_pdf").mkdir(parents=True, exist_ok=True)

        self.config_path = self.state_dir / "kumex_config.json"
        self.stock_path = self.state_dir / "kumex_stock.json"
        stock_data = load_json(self.stock_path, default={"materials": {}})

        # Оставляем в интерфейсе только два материала
        allowed = {"POM Valge", "POM Must"}
        mats = stock_data.get("materials", {})
        filtered = {k: v for k, v in mats.items() if k in allowed}
        if filtered != mats:
            stock_data["materials"] = filtered
            save_json(self.stock_path, stock_data)
        if not stock_data["materials"]:
            stock_data["materials"] = {
                "POM Valge": {"enabled": True, "stock_m3": 0.0, "remain_m3": 0.0},
                "POM Must":  {"enabled": True, "stock_m3": 0.0, "remain_m3": 0.0}
            }
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
        # Храним фактическое списание (m2_used по плитам), чтобы Ledger брал реальные значения, а в UI можно показывать идеальную потребность
        self._m2_used_totals = {name: Decimal("0.00") for name in self.materials_cfg.keys()}
        # Отходы по ширине (м²)


        # --- переменные GUI ---
        self.month_var = tk.StringVar()
        self.pdf_dir_var = tk.StringVar()

        # --- загрузка конфига / дефолтов ---
        self._load_defaults()
        
        # при старте подтянуть остатки из JSON
        _data = self._load_stock_data()
        self._recompute_balances_from_ledger(_data)
        self._save_stock_data(_data)
        self._update_negative_highlight()
        
        # --- построение интерфейса ---
        self._build_ui()

        # --- обработчик закрытия: сохранить конфиг ---
        self.master.protocol("WM_DELETE_WINDOW", self._on_exit)

    # ---------------- UI ----------------

    def _build_ui(self):
        self.master.minsize(720, 360)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=16, pady=16)
        # стиль для отрицательных остатков
        style = ttk.Style(self)
        style.configure("Neg.TEntry", foreground="#C62828")   # тёмно-красный


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
        ttk.Label(container, text="PDF kaust:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        pdf_entry = ttk.Entry(container, textvariable=self.pdf_dir_var, width=60)
        pdf_entry.grid(row=2, column=1, sticky="we", pady=(8, 0))
        choose_btn = ttk.Button(container, text="Vali…", command=self._choose_pdf_dir)
        choose_btn.grid(row=2, column=2, sticky="w", padx=(8, 0), pady=(8, 0))

        # Ряд 5: Два списка (PDF и материалы)
        lists_frame = ttk.Frame(container)
        lists_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(16, 0))
        container.rowconfigure(5, weight=1)
        container.columnconfigure(1, weight=1)
        lists_frame.columnconfigure(0, weight=1)
        lists_frame.columnconfigure(1, weight=1)
        lists_frame.rowconfigure(1, weight=1)

        # Левая панель — PDF
        pdf_group = ttk.LabelFrame(lists_frame, text="PDF-failid (filtreeritud kuu järgi)")
        pdf_group.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        pdf_group.columnconfigure(0, weight=1)
        pdf_group.rowconfigure(1, weight=1)

        self.pdf_count_lbl = ttk.Label(pdf_group, text="Leitud: 0")
        self.pdf_count_lbl.grid(row=0, column=0, sticky="w", padx=8, pady=(4, 0))

        self.pdf_list = tk.Listbox(pdf_group, height=12)
        pdf_scroll = ttk.Scrollbar(pdf_group, orient="vertical", command=self.pdf_list.yview)
        self.pdf_list.configure(yscrollcommand=pdf_scroll.set)
        self.pdf_list.grid(row=1, column=0, sticky="nsew", padx=(8, 0), pady=8)
        pdf_scroll.grid(row=1, column=1, sticky="ns", pady=8)

        # Правая панель — Материалы из PDF (позже заполним реальным парсером)
        mat_group = ttk.LabelFrame(lists_frame, text="Tellitud materjalide loetelu (valitud kuu järgi)")
        mat_group.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        mat_group.columnconfigure(0, weight=1)
        mat_group.rowconfigure(1, weight=1)

        self.mat_count_lbl = ttk.Label(mat_group, text="Positsioone: 0")
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

        # --- Таблица "Расчёты" с фильтром по effective_date ---
        calc_frame = ttk.LabelFrame(container, text="Расчёты (по effective_date)")
        calc_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(12, 0))
        container.rowconfigure(6, weight=1)
        calc_frame.columnconfigure(0, weight=0)
        calc_frame.columnconfigure(1, weight=0)
        calc_frame.columnconfigure(2, weight=0)
        calc_frame.columnconfigure(3, weight=1)

        self.calc_month_var = tk.StringVar(value=datetime.now().strftime("%m"))
        self.calc_year_var = tk.StringVar(value=str(datetime.now().year))

        ttk.Label(calc_frame, text="Aasta:").grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ttk.Combobox(calc_frame, textvariable=self.calc_year_var,
                     values=[str(y) for y in range(datetime.now().year - 5, datetime.now().year + 2)],
                     width=6, state="readonly").grid(row=0, column=1, padx=(0, 12), pady=4, sticky="w")

        ttk.Label(calc_frame, text="Kuu:").grid(row=0, column=2, padx=6, pady=4, sticky="w")
        ttk.Combobox(calc_frame, textvariable=self.calc_month_var,
                     values=[f"{m:02d}" for m in range(1, 13)],
                     width=4, state="readonly").grid(row=0, column=3, padx=(0, 12), pady=4, sticky="w")

        ttk.Button(calc_frame, text="Filtreeri", command=self._reload_calculations)\
            .grid(row=0, column=4, padx=(0, 8), pady=4, sticky="w")

        cols_calc = ("eff_date", "material", "op", "qty", "period", "source")
        self.calc_tree = ttk.Treeview(calc_frame, columns=cols_calc, show="headings", height=8)
        self.calc_tree.heading("eff_date", text="Effective date")
        self.calc_tree.heading("material", text="Material")
        self.calc_tree.heading("op", text="Operation")
        self.calc_tree.heading("qty", text="m²")
        self.calc_tree.heading("period", text="Period")
        self.calc_tree.heading("source", text="Source/PO")
        self.calc_tree.column("eff_date", width=110, anchor="center")
        self.calc_tree.column("material", width=120, anchor="w")
        self.calc_tree.column("op", width=90, anchor="center")
        self.calc_tree.column("qty", width=80, anchor="e")
        self.calc_tree.column("period", width=140, anchor="center")
        self.calc_tree.column("source", width=200, anchor="w")

        scr_calc = ttk.Scrollbar(calc_frame, orient="vertical", command=self.calc_tree.yview)
        self.calc_tree.configure(yscrollcommand=scr_calc.set)
        self.calc_tree.grid(row=1, column=0, columnspan=5, sticky="nsew", padx=(6, 0), pady=(4, 6))
        scr_calc.grid(row=1, column=5, sticky="ns", pady=(4, 6))
        calc_frame.rowconfigure(1, weight=1)

        # контекстное меню для расчётов
        self._calc_menu = tk.Menu(self, tearoff=0)
        self._calc_menu.add_command(label="Edit", command=self._calc_edit_selected)
        self._calc_menu.add_command(label="Delete", command=self._calc_delete_selected)
        self.calc_tree.bind("<Button-3>", self._on_calc_context)

        # --- Нижняя панель: слева — склад Kumex (ввод), справа — конвертация в м² (вывод) ---
        bottom_frame = ttk.Frame(container)
        bottom_frame.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=(8, 0))

        bottom_frame.columnconfigure(0, weight=0)  # левая узкая колонка
        bottom_frame.columnconfigure(1, weight=1)  # правая широкая колонка

        # Левая группа (под левым списком): настройки склада Kumex
        cfg_group = ttk.LabelFrame(bottom_frame, text="Kumex ladu — materjalide arvestus (m²)")
        cfg_group.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        for c in range(4):
            cfg_group.columnconfigure(c, weight=0)

        # Заголовки
        ttk.Label(cfg_group, text="Materjal").grid(row=0, column=0, padx=8, pady=(6, 2), sticky="w")
        ttk.Label(cfg_group, text="Jääk, m²").grid(row=0, column=3, padx=8, pady=(6, 2), sticky="w")

        # Ряды по материалам
        # Ряды по материалам ИЗ JSON (с сохранением флага enabled)
        _row = 1
        for name, cfg in self.materials_cfg.items():
            ttk.Label(cfg_group, text=name).grid(row=_row, column=0, padx=8, pady=2, sticky="w")

            e_rem = ttk.Entry(cfg_group, textvariable=cfg["remain_m3"], width=10, state="readonly")
            e_rem.grid(row=_row, column=3, padx=8, pady=2, sticky="w")
            cfg["remain_entry"] = e_rem

            _row += 1
            
        # Кнопка "Настройка склада…" по центру снизу левого блока
        btn_row = ttk.Frame(cfg_group)
        btn_row.grid(row=_row, column=0, columnspan=3, sticky="ew", padx=8, pady=(8, 6))
        btn_row.columnconfigure(0, weight=1)
        ttk.Button(btn_row, text="Lao seaded……", command=self._open_stock_dialog).pack(anchor="center")


        conv_group = ttk.LabelFrame(bottom_frame, text="Konverteerimine (valitud kuu järgi)")
        conv_group.grid(row=0, column=1, sticky="nsew")
        # 0: имя материала, 1: значение, 2: "м²", 3: спейсер для растяжения
        conv_group.columnconfigure(0, weight=0)
        conv_group.columnconfigure(1, weight=0)
        conv_group.columnconfigure(2, weight=0)
        # Conversion totals (read-only)
        _row_conv = 0
        for name in self.conv_totals.keys():
            pad_top = (6 if _row_conv == 0 else 2, 2)
            ttk.Label(conv_group, text=f"{name}:").grid(row=_row_conv, column=0, padx=8, pady=pad_top, sticky="w")
            ttk.Entry(conv_group, textvariable=self.conv_totals[name], width=10, state="readonly").grid(row=_row_conv, column=1, padx=(8, 2), pady=pad_top, sticky="w")
            ttk.Label(conv_group, text="m²").grid(row=_row_conv, column=2, padx=(0, 8), pady=pad_top, sticky="w")
            _row_conv += 1

        kerf_row = _row_conv + 1
        ttk.Label(conv_group, text="Saetera paksus:").grid(row=kerf_row, column=0, padx=8, pady=(10, 2), sticky="w")
        e_kerf = ttk.Entry(conv_group, textvariable=self.kerf_mm_var, width=6)
        e_kerf.grid(row=kerf_row, column=1, padx=(8, 2), pady=(10, 2), sticky="w")
        e_kerf.bind("<FocusOut>", lambda _e: self._save_config())
        e_kerf.bind("<Return>",  lambda _e: (self._save_config(), "break"))
        ttk.Label(conv_group, text="mm").grid(row=kerf_row, column=2, padx=(0, 8), pady=(10, 2), sticky="w")

        # Period controls near Arvuta
        pm_row = kerf_row + 1
        period_cb = ttk.Checkbutton(conv_group, text="Period mode", variable=self.period_mode_var, command=self._toggle_period_mode)
        period_cb.grid(row=pm_row, column=0, sticky="w", padx=8, pady=(10, 2))

        ttk.Label(conv_group, text="From:").grid(row=pm_row, column=1, sticky="w", padx=(4,0))
        self.period_from_year = tk.StringVar(value=str(datetime.now().year))
        self.period_from_month = tk.StringVar(value=datetime.now().strftime("%m"))
        self._cmb_from_year = ttk.Combobox(conv_group, textvariable=self.period_from_year,
                     values=[str(y) for y in range(datetime.now().year-5, datetime.now().year+2)],
                     width=6, state="disabled")
        self._cmb_from_year.grid(row=pm_row, column=2, sticky="w")
        self._cmb_from_month = ttk.Combobox(conv_group, textvariable=self.period_from_month,
                     values=[f"{m:02d}" for m in range(1,13)],
                     width=4, state="disabled")
        self._cmb_from_month.grid(row=pm_row, column=3, sticky="w", padx=(4,0))

        ttk.Label(conv_group, text="To:").grid(row=pm_row+1, column=1, sticky="w", padx=(4,0))
        self.period_to_year = tk.StringVar(value=str(datetime.now().year))
        self.period_to_month = tk.StringVar(value=datetime.now().strftime("%m"))
        self._cmb_to_year = ttk.Combobox(conv_group, textvariable=self.period_to_year,
                     values=[str(y) for y in range(datetime.now().year-5, datetime.now().year+2)],
                     width=6, state="disabled")
        self._cmb_to_year.grid(row=pm_row+1, column=2, sticky="w")
        self._cmb_to_month = ttk.Combobox(conv_group, textvariable=self.period_to_month,
                     values=[f"{m:02d}" for m in range(1,13)],
                     width=4, state="disabled")
        self._cmb_to_month.grid(row=pm_row+1, column=3, sticky="w", padx=(4,0))

        calc_btn = ttk.Button(conv_group, text="Arvuta", command=self._apply_stub)
        calc_btn.grid(row=pm_row+2, column=0, columnspan=4, sticky="e", padx=8, pady=(10, 6))
        self._calc_btn = calc_btn

        self._period_widgets = [self._cmb_from_year, self._cmb_from_month, self._cmb_to_year, self._cmb_to_month]
        # ---------------- Служебные обработчики ----------------

    def _toggle_period_mode(self):
        pm = self.period_mode_var.get()
        state = "readonly" if pm else "disabled"
        for cmb in getattr(self, "_period_widgets", []):
            try:
                cmb.configure(state=state)
            except tk.TclError:
                pass

    # ------------- Архивация журнала -------------

    def _ensure_archive_dir(self):
        arch = self.state_dir / "archive"
        arch.mkdir(parents=True, exist_ok=True)
        return arch

    def _archive_dialog(self):
        tree = getattr(self, "_ledger_tree", None)
        if not tree or not tree.winfo_exists():
            return
        win = tk.Toplevel(self)
        win.title("Archive ledger")
        win.grab_set()

        tk.Label(win, text="Archive until month (YYYY-MM):").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        until_var = tk.StringVar()
        tk.Entry(win, textvariable=until_var, width=10).grid(row=0, column=1, padx=8, pady=6, sticky="w")

        sel_var = tk.BooleanVar(value=False)
        tk.Checkbutton(win, text="Archive selected rows", variable=sel_var).grid(row=1, column=0, columnspan=2, padx=8, pady=4, sticky="w")

        status = tk.StringVar(value="")
        tk.Label(win, textvariable=status, fg="#C62828").grid(row=2, column=0, columnspan=2, padx=8, pady=4, sticky="w")

        def _do_archive():
            to_archive = []
            data = self._load_stock_data()
            ledger = data.get("ledger", [])

            # selected rows
            if sel_var.get():
                sel = tree.selection()
                idx_map = getattr(self, "_ledger_index", {})
                for iid in sel:
                    meta = idx_map.get(iid)
                    if not meta:
                        continue
                    for rec in ledger:
                        if rec.get("ts") == meta.get("ts") and rec.get("material") == meta.get("material") and rec.get("type") == meta.get("type"):
                            to_archive.append(rec)
                            break

            # until month
            mtext = until_var.get().strip()
            if mtext:
                try:
                    _dt.datetime.strptime(mtext + "-01", "%Y-%m-%d")
                except Exception:
                    status.set("Kuupäev peab olema kujul YYYY-MM")
                    return
                for rec in ledger:
                    month = rec.get("month") or ""
                    if month and month <= mtext and rec not in to_archive:
                        to_archive.append(rec)

            if not to_archive:
                status.set("Valige read või sisestage kuu.")
                return

            arch_dir = self._ensure_archive_dir()
            fname = arch_dir / f"ledger_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jsonl"
            import json
            with fname.open("w", encoding="utf-8") as f:
                for rec in to_archive:
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")

            # удалить их из ledger
            data["ledger"] = [r for r in ledger if r not in to_archive]
            # пересчитать closed_months (month_calc)
            data["closed_months"] = sorted({r.get("month") for r in data["ledger"] if r.get("type") == "month_calc" and r.get("month")})
            self._recompute_balances_from_ledger(data)
            self._save_stock_data(data)
            self._update_negative_highlight()
            self._reload_ledger()
            self._reload_calculations()
            status.set(f"Arhiveeritud: {len(to_archive)}, fail: {fname.name}")

        tk.Button(win, text="Archive", command=_do_archive).grid(row=3, column=0, padx=8, pady=8, sticky="e")
        tk.Button(win, text="Close", command=win.destroy).grid(row=3, column=1, padx=8, pady=8, sticky="w")

    def _open_archive_viewer(self):
        arch_dir = self._ensure_archive_dir()
        files = sorted(arch_dir.glob("ledger_*.jsonl"))
        win = tk.Toplevel(self)
        win.title("Archive viewer")
        win.geometry("640x400")
        win.grab_set()

        listbox = tk.Listbox(win, height=10)
        listbox.pack(side="left", fill="y", padx=8, pady=8)
        for f in files:
            listbox.insert(tk.END, f.name)

        txt = tk.Text(win, wrap="none")
        txt.pack(side="right", fill="both", expand=True, padx=8, pady=8)

        def _on_select(evt=None):
            sel = listbox.curselection()
            if not sel:
                return
            fname = files[sel[0]]
            try:
                content = fname.read_text(encoding="utf-8")
            except Exception as e:
                content = f"Viga: {e}"
            txt.delete("1.0", tk.END)
            txt.insert(tk.END, content)

        listbox.bind("<<ListboxSelect>>", _on_select)

    def _choose_pdf_dir(self):
        initial = self.pdf_dir_var.get() or (self.state_dir / "input_pdf").as_posix()
        chosen = filedialog.askdirectory(initialdir=initial, title="Выбрать папку с PDF")
        if chosen:
            self.pdf_dir_var.set(Path(chosen).as_posix())
            # всегда сохраняем последний путь
            self._save_config()

    def _scan_pdfs(self):
    
        folder = Path(self.pdf_dir_var.get()).expanduser()
        if not folder.exists():
            messagebox.showerror("Kaust pole kättesaadav", f"Kausta ei eksisteeri:\n{folder}")
            return

        # читаем год/месяц из комбобоксов
        yy_str = self.year_var.get().strip()
        mm_str = self.month_num_var.get().strip()
        if not (yy_str.isdigit() and mm_str.isdigit()):
            messagebox.showerror("Vale kuupäev", f"Oodatud formaat YYYY-MM, saadi: {yy_str}-{mm_str}")
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
        self.pdf_count_lbl.config(text=f"Leitud: {len(self.pdf_files)}")

        # чистим правый список (материалы добавим позже)
        self.material_rows = []
        for iid in self.mat_tree.get_children():
            self.mat_tree.delete(iid)
        self.mat_count_lbl.config(text="Positsioone: 0")
        
        self._parse_materials()

        self._set_status(f"Kaust: {folder} | PDF kuu {yy}-{mm_str}: {len(self.pdf_files)}")
        self._calc_m2()
        self._update_calc_button_state()

    def _parse_materials(self):
    
        # очистка списка перед циклом (важно не чистить внутри)
        for iid in self.mat_tree.get_children():
            self.mat_tree.delete(iid)
        self.material_rows = []

        if not self.pdf_files:
            self._set_status("Kõigepealt vajutage „Skaneeri PDF“.")
            self.mat_count_lbl.config(text="Positsioone: 0")
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
                qty_text = ""

                # 1) qty ?? ???? ?? ??????
                m = qty_any_rx.search(line)
                if m:
                    qty = int(m.group(1))
                    qty_text = line

                # 2) ???? ?? ????? ? ???? ?? 1..3 ??????? ????
                if qty is None:
                    for j in range(1, 4):
                        if i + j >= n:
                            break
                        m = qty_any_rx.search(lines[i + j])
                        if m:
                            qty = int(m.group(1))
                            qty_text = lines[i + j]
                            break

                # 3) ???? ?? ????? ? ??????? ????????? ??????? ? 1..2 ??????? ????
                if qty is None:
                    for j in (1, 2):
                        if i - j >= 0:
                            m = qty_row_above_rx.match(lines[i - j])
                            if m:
                                qty = int(m.group(1))
                                qty_text = lines[i - j]
                                break

                if qty is None:
                    # ?? ?????? ???????? ????? ?????????? ? ?????????? ???????
                    continue

                # ?????????? uom: ???? ????? ? ??????????? ???? ??????? mm ? ???????? ??? ????? ?????, ?? ?????
                # ?????????? uom: ??????? ??? ? ???? ?????????? >= 1000, ??????? ??? ????? ?????? (??), ????? ?????
                uom = "mm" if qty >= 1000 else "tk"
                # ???????? ??????: ???? ? ?????? ???-???? ???? ??????? mm, ???? ??????? ????????
                if uom == "tk":
                    qty_ctx = f"{desc}\n{qty_text}"
                    if re.search(rf"{qty}\\s*mm\\b", qty_ctx, re.IGNORECASE):
                        uom = "mm"

                # ????????? ??? ??????? (?? ??????? ??? ???????)
                # ?????????? ????????

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
                    "uom": uom,
                    "po": po,
                    "date": od,
                    "material": mat_name
                })
                
                # добавляем строку в таблицу
                self.mat_tree.insert("", "end", values=(desc, qty, po, od))
                total += 1

        self.mat_count_lbl.config(text=f"Positsioone: {total}")
        self._set_status(f"Töödeldud PDF: {len(self.pdf_files)} | Leitud positsioone: {total}")
        self._calc_m2()

    def _on_date_change(self, *_):
        self._sync_month_var()
        self._scan_pdfs()
        self._update_calc_button_state()

    def _report_stub(self):
        # Заглушка: здесь позже будет агрегатор + генерация отчёта
        messagebox.showinfo("Aruanne", "Aruande koostamine (kohatäide)")

    def _sync_month_var(self):
        yy = self.year_var.get()
        mm = self.month_num_var.get()
        if len(yy) == 4 and mm in {f"{i:02d}" for i in range(1,13)}:
            self.month_var.set(f"{yy}-{mm}")

    # ------------- Загрузка/сохранение конфигурации -------------

    def _load_defaults(self):
        cfg = load_json(self.config_path, default={})
        current_month = datetime.now().strftime("%Y-%m")
        target = cfg.get("last_month", current_month)
        if "-" not in target:
            target = current_month
        try:
            yy, mm = target.split("-")
        except ValueError:
            yy, mm = current_month.split("-")
        self.month_var.set(f"{yy}-{mm}")
        self.year_var.set(yy)
        self.month_num_var.set(mm)
        self.kerf_mm_var.set(str(cfg.get("kerf_mm", self.kerf_mm_var.get() or "0")))
        default_pdf_dir = (self.state_dir / "input_pdf").as_posix()
        self.pdf_dir_var.set(cfg.get("pdf_dir", default_pdf_dir))

    def _save_config(self):
        """Persist current UI configuration (month, kerf, pdf dir)."""
        try:
            kerf_val = float(str(self.kerf_mm_var.get()).replace(",", ".") or 0)
        except Exception:
            kerf_val = 0.0
        data = {
            "last_month": self._month_key(),
            "pdf_dir": self.pdf_dir_var.get().strip(),
            "kerf_mm": kerf_val,
        }
        save_json(self.config_path, data)

    def _calc_m2(self):
        """Compute cutting plan and summary according to Codex rules."""
        import re
        from decimal import Decimal, ROUND_HALF_UP

        # длина реза фиксирована
        cut_len = DEFAULT_CUT_LEN_MM
        self.cut_len_var.set(str(cut_len))

        try:
            kerf = Decimal(str(self.kerf_mm_var.get()).replace(',', '.'))
            if kerf < 0:
                kerf = Decimal('0')
        except Exception:
            kerf = Decimal('0')

        S_plate = Decimal(PLATE_W_MM * PLATE_L_MM) / Decimal(1_000_000)
        layers_per_plate = PLATE_L_MM // cut_len

        issues = []
        plan_layers_rows = []
        results = {}

        def add_issue(material, dims, uom, qty, code, detail):
            issues.append({
                'material': material,
                'dims': dims,
                'uom': uom,
                'qty': qty,
                'issue_code': code,
                'issue_detail': detail,
            })

        def parse_dims(desc: str):
            nums = [int(x) for x in re.findall(r'\d+', desc)]
            if len(nums) < 3:
                return []
            return sorted(nums, reverse=True)[:3]

        items_by_mat = {m: [] for m in ('POM Valge', 'POM Must')}
        for row in self.material_rows:
            mat = row.get('material', '')
            if mat not in items_by_mat:
                continue
            uom = str(row.get('uom', 'tk') or 'tk')
            try:
                qty = int(row.get('qty', 0) or 0)
            except Exception:
                qty = 0
            dims = parse_dims(str(row.get('desc', '')))
            if qty <= 0:
                add_issue(mat, dims, uom, qty, 'ISSUE_QTY_NON_POSITIVE', 'Qty must be positive')
                continue
            if len(dims) < 3:
                add_issue(mat, dims, uom, qty, 'ISSUE_NO_DIM_LE_BASE_THICKNESS', 'Need three dimensions')
                continue
            items_by_mat[mat].append({'dims': dims, 'qty': qty, 'uom': uom})

        for mat, items in items_by_mat.items():
            strips_widths = []
            s_ideal_sum = Decimal('0')
            waste_width_sum = Decimal('0')
            waste_thickness_sum = Decimal('0')
            for item in items:
                dims = item['dims']
                qty = item['qty']
                uom = item['uom']

                candidates = [d for d in dims if d <= BASE_THICKNESS_MM]
                if not candidates:
                    add_issue(mat, dims, uom, qty, 'ISSUE_NO_DIM_LE_BASE_THICKNESS', 'All dimensions exceed base thickness')
                    continue
                t = max(candidates)
                rem = dims.copy()
                rem.remove(t)
                rem_sorted = sorted(rem, reverse=True)
                if len(rem_sorted) < 2:
                    add_issue(mat, dims, uom, qty, 'ISSUE_NO_FEASIBLE_ORIENTATION', 'Not enough sides for footprint')
                    continue
                p, q = rem_sorted[:2]
                # kerf ?????? ?? ?????????? ????????; ???? ?????? ????? ????? ???? (cut_len), ?? ????????? kerf
                def _add_kerf(dim: Decimal) -> Decimal:
                    return dim if dim == Decimal(cut_len) else dim + kerf

                p_eff = _add_kerf(Decimal(p))
                q_eff = _add_kerf(Decimal(q))
                orientations = []
                width_issue = False

                for w, L in ((p_eff, q_eff), (q_eff, p_eff)):
                    if w > PLATE_W_MM:
                        width_issue = True
                        continue
                    if L > cut_len:
                        continue
                    if uom == 'tk':
                        k = int(Decimal(cut_len) // Decimal(L))
                        if k <= 0:
                            continue
                        strips = (qty + k - 1) // k
                        W_demand = strips * w
                        L_waste = Decimal(strips * cut_len) - Decimal(qty) * Decimal(L)
                        S_ideal = (Decimal(strips) * Decimal(w) * Decimal(cut_len)) / Decimal(1_000_000)
                    else:
                        # uom=mm: qty ??? ????? ?????; ?????? ??????????? = min ?????????? ??????
                        w_use = min(w, q_eff, p_eff)
                        strips = (qty + cut_len - 1) // cut_len
                        W_demand = strips * w_use
                        L_waste = Decimal(strips * cut_len - qty)
                        S_ideal = (Decimal(strips) * Decimal(w_use) * Decimal(cut_len)) / Decimal(1_000_000)
                        w = w_use
                    orientations.append((W_demand, L_waste, Decimal(BASE_THICKNESS_MM - t), w, strips, S_ideal))
                if not orientations:
                    if width_issue:
                        add_issue(mat, dims, uom, qty, 'ISSUE_WIDTH_GT_PLATE', f'w > {PLATE_W_MM}')
                    add_issue(mat, dims, uom, qty, 'ISSUE_NO_FEASIBLE_ORIENTATION', 'No orientation fits plate/cut length')
                    continue

                orientations.sort(key=lambda x: (x[0], x[1], x[2]))
                _, _, _, w_sel, strips, S_ideal = orientations[0]
                s_ideal_sum += S_ideal
                strips_widths.extend([float(w_sel)] * int(strips))

                if t < BASE_THICKNESS_MM:
                    waste_thickness_sum += (Decimal(BASE_THICKNESS_MM - t) * Decimal(p) * Decimal(q) * qty)

            layers = []
            for w in sorted(strips_widths, reverse=True):
                placed = False
                for layer in layers:
                    if sum(layer) + w <= PLATE_W_MM:
                        layer.append(w)
                        placed = True
                        break
                if not placed:
                    layers.append([w])

            for idx, layer in enumerate(layers, start=1):
                remaining = PLATE_W_MM - sum(layer)
                plan_layers_rows.append({
                    'material': mat,
                    'layer_id': idx,
                    'remaining_width_mm': remaining,
                    'strip_widths': layer,
                })
                waste_width_sum += Decimal(remaining)

            layers_used = len(layers)
            plates_used = (layers_used + layers_per_plate - 1) // layers_per_plate
            plates_left = STOCK_PLATES - plates_used
            s_used = Decimal(plates_used) * S_plate
            s_left = Decimal(max(plates_left, 0)) * S_plate
            s_waste_width = (waste_width_sum * Decimal(cut_len)) / Decimal(1_000_000)

            results[mat] = {
                'layers_used': layers_used,
                'plates_used': plates_used,
                'plates_left': plates_left,
                'm2_used': s_used,
                'm2_left': s_left,
                'layers_per_plate': layers_per_plate,

                's_ideal_sum': s_ideal_sum,
                'waste_width_m2': s_waste_width,
                'waste_thickness_mm3': waste_thickness_sum,
            }

        for mat in ('POM Valge', 'POM Must'):
            res = results.get(mat)
            ideal_val = Decimal('0') if not res else res['s_ideal_sum'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            used_val = Decimal('0') if not res else res['m2_used'].quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
            self._m2_used_totals[mat] = used_val
            if mat in self.conv_totals:
                self.conv_totals[mat].set(str(ideal_val))



        self._write_reports(results, plan_layers_rows, issues)

    def _write_reports(self, results, plan_layers_rows, issues):
        """Save summary, layer plan, and issues to files."""
        import json
        import csv

        report_dir = self.state_dir / 'reports'
        report_dir.mkdir(parents=True, exist_ok=True)

        summary_rows = []
        for mat, res in results.items():
            summary_rows.append({
                'material': mat,
                'plates_used': res.get('plates_used', 0),
                'm2_used': float(res.get('m2_used', 0)),
                'plates_left': res.get('plates_left', 0),
                'm2_left': float(res.get('m2_left', 0)),
                'layers_used': res.get('layers_used', 0),
                'layers_per_plate': res.get('layers_per_plate', 0),

                's_ideal_sum': float(res.get('s_ideal_sum', 0)),
                'waste_width_m2': float(res.get('waste_width_m2', 0)),
            })

        ts_label = datetime.now().strftime("%Y%m%d_%H%M%S")
        (report_dir / f'summary_{ts_label}.json').write_text(json.dumps(summary_rows, ensure_ascii=False, indent=2), encoding='utf-8')
        if summary_rows:
            fieldnames = sorted({k for row in summary_rows for k in row.keys()})
            with (report_dir / f'summary_{ts_label}.csv').open('w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(summary_rows)

        if plan_layers_rows:
            with (report_dir / f'plan_layers_{ts_label}.csv').open('w', newline='', encoding='utf-8') as f:
                fieldnames = sorted({k for row in plan_layers_rows for k in row.keys()})
                writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
                writer.writeheader()
                for row in plan_layers_rows:
                    writer.writerow(row)

        with (report_dir / 'issues.csv').open('w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=['material', 'dims', 'uom', 'qty', 'issue_code', 'issue_detail'])
            writer.writeheader()
            for row in issues:
                writer.writerow(row)


    def _open_stock_dialog(self):
        """Открывает окно 'Настройка склада' с таблицей журнала операций."""
        if self._stock_win and tk.Toplevel.winfo_exists(self._stock_win):
            self._stock_win.focus_set()
            return

        self._stock_win = tk.Toplevel(self.master)
        self._stock_win.withdraw()  # спрятать до настройки
        self._stock_win.title("Lao seaded… ")
        self._stock_win.transient(self.master)
        self._stock_win.geometry("720x420")

        def _release_modal_grab(_=None):
            try:
                self._stock_win.grab_release()
            except Exception:
                pass

        def _restore_modal_grab(_=None):
            if self._stock_win and tk.Toplevel.winfo_exists(self._stock_win):
                try:
                    self._stock_win.grab_set()
                except Exception:
                    pass

        # Следим за сворачиванием/разворачиванием, чтобы не оставлять grab на скрытом окне
        self._stock_win_binds = {
            "master_unmap": self.master.bind("<Unmap>", _release_modal_grab, add="+"),
            "master_map": self.master.bind("<Map>", _restore_modal_grab, add="+"),
            "stock_unmap": self._stock_win.bind("<Unmap>", _release_modal_grab, add="+"),
            "stock_map": self._stock_win.bind("<Map>", _restore_modal_grab, add="+"),
        }

        def _on_close():
            try:
                self._stock_win.grab_release()
            except Exception:
                pass
            try:
                binds = self._stock_win_binds or {}
                if binds.get("master_unmap"):
                    self.master.unbind("<Unmap>", binds["master_unmap"])
                if binds.get("master_map"):
                    self.master.unbind("<Map>", binds["master_map"])
                if binds.get("stock_unmap"):
                    self._stock_win.unbind("<Unmap>", binds["stock_unmap"])
                if binds.get("stock_map"):
                    self._stock_win.unbind("<Map>", binds["stock_map"])
            except Exception:
                pass
            self._stock_win_binds = {}
            self._stock_win.destroy()
            self._stock_win = None

        self._stock_win.protocol("WM_DELETE_WINDOW", _on_close)

                # Корневой фрейм
        frm = ttk.Frame(self._stock_win)
        frm.pack(fill="both", expand=True, padx=10, pady=10)

        # ===== ФОРМА ОПЕРАЦИИ (+/- м²) =====
        op_box = ttk.LabelFrame(frm, text="Lao toiming")
        op_box.pack(fill="x", padx=0, pady=(0, 8))

        # переменные формы
        self._op_material = tk.StringVar(value="POM Valge")
        self._op_amount = tk.StringVar(value="0.00")
        self._op_effective_date = tk.StringVar(value=datetime.now().strftime("%d.%m.%Y"))

        ttk.Label(op_box, text="Materjal:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Combobox(op_box, textvariable=self._op_material, state="readonly",
                    values=("POM Valge", "POM Must"), width=12).grid(row=0, column=1, padx=(0, 8), pady=6, sticky="w")

        ttk.Label(op_box, text="Kogus:").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        self._op_amount_entry = tk.Entry(op_box, textvariable=self._op_amount, width=12)
        self._op_amount_entry.grid(row=0, column=3, padx=(0, 4), pady=6, sticky="w")
        ttk.Label(op_box, text="m²").grid(row=0, column=4, padx=(0, 8), pady=6, sticky="w")

        ttk.Label(op_box, text="Order/Arrival date:").grid(row=1, column=0, padx=8, pady=(0, 6), sticky="w")
        ttk.Entry(op_box, textvariable=self._op_effective_date, width=14).grid(row=1, column=1, padx=(0, 8), pady=(0, 6), sticky="w")

        # даём пространство перед кнопкой добавления
        op_box.columnconfigure(3, weight=1)

        ttk.Button(op_box, text="Lisa toiming", command=self._add_stock_operation)\
        .grid(row=1, column=4, padx=8, pady=(0, 6), sticky="e")


        # Колонки: растягиваем пространство между радиокнопками и кнопкой
        for i in range(9):
            op_box.columnconfigure(i, weight=0)
        op_box.columnconfigure(7, weight=1)   # даёт место заголовку «Тип действия» и кнопке справа
        # чтобы поле комментария имело место

        # ===== УДАЛЕНИЕ РАСЧЁТА МЕСЯЦА =====
        rm_box = ttk.LabelFrame(frm, text="Kustuta kuu arvestus")
        rm_box.pack(fill="x", padx=0, pady=(0, 8))

        self._rm_year = tk.StringVar(value=self.year_var.get())
        self._rm_month = tk.StringVar(value=self.month_num_var.get())

        ttk.Label(rm_box, text="Aasta:").grid(row=0, column=0, padx=8, pady=6, sticky="w")
        ttk.Combobox(rm_box, textvariable=self._rm_year, values=[str(y) for y in range(2017, 2031)],
                    state="readonly", width=6).grid(row=0, column=1, padx=(0, 8), pady=6, sticky="w")
        ttk.Label(rm_box, text="Kuu:").grid(row=0, column=2, padx=8, pady=6, sticky="w")
        ttk.Combobox(rm_box, textvariable=self._rm_month, values=[f"{m:02d}" for m in range(1, 13)],
                    state="readonly", width=4).grid(row=0, column=3, padx=(0, 8), pady=6, sticky="w")

        ttk.Button(rm_box, text="Kustuta kuu arvestus", command=self._delete_month_calc)\
            .grid(row=0, column=4, padx=8, pady=6, sticky="e")

        rm_box.columnconfigure(5, weight=1)

        # ===== ЖУРНАЛ ОПЕРАЦИЙ =====
        topbar = ttk.Frame(frm)
        topbar.pack(fill="x", pady=(6, 0))
        ttk.Label(topbar, text="Lao toimingute logi").pack(side="left")
        ttk.Button(topbar, text="Archive…", command=self._archive_dialog).pack(side="right", padx=(4, 0))
        ttk.Button(topbar, text="View archive", command=self._open_archive_viewer).pack(side="right", padx=(4, 0))

        cols = ("ts", "month", "material", "action", "amount", "note")
        self._ledger_tree = ttk.Treeview(frm, columns=cols, show="headings", height=12)
        self._ledger_tree.heading("ts", text="Kuupäev/aeg")
        self._ledger_tree.heading("month", text="Kuu")
        self._ledger_tree.heading("material", text="Materjal")
        self._ledger_tree.heading("action", text="Toiming")
        self._ledger_tree.heading("amount", text="m²")
        self._ledger_tree.heading("note", text="Kommentaar")

        self._ledger_tree.column("ts", width=140, anchor="center")
        self._ledger_tree.column("month", width=80, anchor="center")
        self._ledger_tree.column("material", width=120, anchor="w")
        self._ledger_tree.column("action", width=120, anchor="w")
        self._ledger_tree.column("amount", width=80, anchor="center")
        self._ledger_tree.column("note", width=220, anchor="w")
        
        # Цветовые теги для строк журнала
        self._ledger_tree.tag_configure("t_add",   foreground="#0A7D00")  # зелёный  (пополнение)
        self._ledger_tree.tag_configure("t_sub",   foreground="#A40000")  # красный  (списание)
        self._ledger_tree.tag_configure("t_month", foreground="#004A9F")  # синий    (расчёт месяца/периода)
        self._ledger_tree.tag_configure("t_set",   foreground="#6A1B9A")  # фиолетовый (установка)
        self._ledger_tree.tag_configure("t_calc",  foreground="#004A9F")

        scr = ttk.Scrollbar(frm, orient="vertical", command=self._ledger_tree.yview)
        self._ledger_tree.configure(yscrollcommand=scr.set)
        self._ledger_tree.pack(side="left", fill="both", expand=True, pady=(6, 0))
        scr.pack(side="left", fill="y", pady=(6, 0))

        # статус внизу
        self._stock_status = tk.StringVar(value="")
        ttk.Label(self._stock_win, textvariable=self._stock_status).pack(fill="x", padx=10, pady=(6, 10))

        # загрузка данных в таблицу
        self._reload_ledger()
        
        # --- центрирование окна "Настройка склада" над главным окном ---
        # --- центрирование без мигания ---
        # посчитать нужные размеры, окно пока скрыто
        self._stock_win.update_idletasks()

        # геометрия главного окна
        px = self.winfo_rootx()
        py = self.winfo_rooty()
        pw = self.winfo_width()
        ph = self.winfo_height()

        # запрошенные размеры диалога (точнее, чем winfo_width/height для скрытого окна)
        ww = self._stock_win.winfo_reqwidth()
        wh = self._stock_win.winfo_reqheight()

        # позиция по центру главного окна
        x = px + (pw - ww) // 2
        y = py + (ph - wh) // 2

        # сперва выставляем геометрию, а уже потом показываем окно
        self._stock_win.geometry(f"+{x}+{y}")
        self._stock_win.transient(self.master)
        self._stock_win.deiconify()        # ← показать окно (теперь уже по центру)
        self._stock_win.grab_set()         # если у тебя был grab_set — оставь
        self._stock_win.focus_set()
        # --- конец центрирования ---

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
            elif typ == "manual_set":
                sums[mat] = amt
            elif typ == "month_calc":
                sums[mat] -= amt
            elif typ == "calc_deduct":
                sums[mat] -= amt
        # обновим JSON + GUI
        for mat in ("POM Valge", "POM Must"):
            v = sums[mat].quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            data["materials"][mat]["remain_m3"] = float(v)
            # если есть поля на форме — обновим
            if mat in self.materials_cfg:
                self.materials_cfg[mat]["remain_m3"].set(str(v))
            self._update_negative_highlight()

    def _add_stock_operation(self):
        """Добавить запись в журнал: manual_add/manual_sub, пересчитать остатки."""
        from decimal import Decimal
        import datetime as _dt

        mat = self._op_material.get()
        raw = str(self._op_amount.get()).replace(",", ".").strip()
        note = "Käsitsi toiming"


        # --- валидация количества с поддержкой + / - / set ---
        def _mark_amount_error(msg: str = None):
            try:
                self._op_amount_entry.configure(fg="#C62828")
            except tk.TclError:
                pass
            if msg:
                messagebox.showerror("Viga", msg)

        def _clear_amount_error():
            try:
                self._op_amount_entry.configure(fg="#000000")
            except tk.TclError:
                pass

        _clear_amount_error()

        # онлайн-валидация по вводу
        def _live_check(_evt=None):
            raw_live = self._op_amount.get().replace(" ", "")
            if not raw_live:
                _mark_amount_error()
                return
            if raw_live[0] in "+-":
                candidate = raw_live[1:]
            else:
                candidate = raw_live
            import re as _re
            if _re.fullmatch(r"\d+(?:[.,]\d+)?", candidate):
                _clear_amount_error()
            else:
                _mark_amount_error()

        self._op_amount_entry.bind("<KeyRelease>", _live_check)
        raw_clean = raw.replace(" ", "")
        if not raw_clean:
            _mark_amount_error("Väli 'Kogus, m²' on tühi.")
            return

        op = "set"
        if raw_clean.startswith(("+", "-")):
            op = "add" if raw_clean[0] == "+" else "sub"
            raw_clean = raw_clean[1:]

        import re as _re
        if not _re.fullmatch(r"\d+(?:[.,]\d+)?", raw_clean):
            _mark_amount_error("Sisestage number kujul +12.5, -3, 10 või 2,5.")
            return

        try:
            amount_val = Decimal(raw_clean.replace(",", "."))
        except Exception:
            _mark_amount_error("Sisestage korrektne number väljale 'Kogus, m²'.")
            return
        if amount_val < 0:
            _mark_amount_error("Kogus ei tohi olla negatiivne (kasutage prefiksit + või -).")
            return
        if mat not in ("POM Valge", "POM Must"):
            messagebox.showerror("Viga", "Valige materjal.")
            return

        # дата заказа/поступления (effective_date)
        eff_raw = (self._op_effective_date.get() or "").strip()
        try:
            eff_dt = _dt.datetime.strptime(eff_raw, "%d.%m.%Y").date()
            eff_iso = eff_dt.isoformat()
        except Exception:
            messagebox.showerror("Viga", "Sisestage kuupäev kujul dd.mm.yyyy väljale 'Order/Arrival date'.")
            return

        # читаем/обновляем JSON
        data = self._load_stock_data()
        rec = {
            "ts": _dt.datetime.now().isoformat(timespec="seconds"),
            "month": eff_dt.strftime("%Y-%m"),
            "material": mat,
            "type": "manual_add" if op == "add" else "manual_sub" if op == "sub" else "manual_set",
            "amount_m2": float(amount_val),
            "effective_date": eff_iso,
            "note": note or "Käsitsi toiming"
        }
        data["ledger"].append(rec)

        # пересчёт и сохранение
        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)
        self._update_negative_highlight()
        
        # обновим GUI
        self._reload_ledger()
        self._stock_status.set("Toiming lisatud.")

    def _delete_month_calc(self):
        """Удалить все записи type=month_calc за выбранный месяц и пересчитать остатки."""
        import datetime as _dt

        yy = str(self._rm_year.get())
        mm = str(self._rm_month.get())
        month = f"{yy}-{mm}"

        if not messagebox.askyesno("Kinnitus", f"Kas kustutada kuu {month} arvestus?"):
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
        self._update_negative_highlight()

        self._reload_ledger()
        self._stock_status.set(f"Kustutatud kirjeid: {before - after}.")
        self._update_calc_button_state()

    def _undo_selected(self):
        """Инвертировать выбранную запись журнала (manual_add <-> manual_sub)."""
        sel = getattr(self, "_ledger_tree", None).selection()
        if not sel:
            messagebox.showwarning("Tühistamine", "Valige kirje logis.")
            return

        iid = sel[0]
        meta = getattr(self, "_ledger_index", {}).get(iid)
        if not meta:
            messagebox.showerror("Viga", "Valitud kirje andmeid ei leitud.")
            return

        raw_type = (meta.get("type") or "").lower()
        if raw_type not in ("manual_add", "manual_sub"):
            messagebox.showinfo("Kirje tühistamine",
                                "Seda kirjet ei saa sel viisil tühistada.\n"
                                "Kuu arvestuse jaoks kasutage 'Kustuta kuu arvestus'.")
            return

        material = meta.get("material")
        month = meta.get("month")
        amount = meta.get("amount")
        if amount <= 0:
            messagebox.showerror("Viga", "Kirje kogus on 0")
            return

        # Подтверждение
        pretty_action = "Täiendus" if raw_type == "manual_add" else "Mahakandmine"
        if not messagebox.askyesno("Kinnitus",
                                f"Kirje tühistamine:\n"
                                f"Materjal: {material}\n"
                                f"Toiming: {pretty_action}\n"
                                f"Объём: {amount} m²\n\n"
                                f"Luua vastupidine korrigeerimine."):
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
            "note": f"Valitud kirje tühistamine ({pretty_action})",
        })

        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)
        self._update_negative_highlight()
        
        # Обновление UI
        self._reload_ledger()
        self._stock_status.set("Lisatud vastupidine korrigeerimine.")

    def _reload_ledger(self):
        """Читает ledger из JSON и перерисовывает таблицу журнала."""
        try:
            data = load_json(self.stock_path, default={"ledger": []})
            ledger = data.get("ledger", []) or []
        except Exception as e:
            ledger = []
            if hasattr(self, "_stock_status"):
                self._stock_status.set(f"Viga чтения JSON: {e}")
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

            month = rec.get("month") or (rec.get("effective_date") or "")[:7]
            material = rec.get("material", "")
            typ = (rec.get("type", "") or "").lower()
            if   typ == "manual_add":
                action = "Täiendus"
            elif typ == "manual_sub":
                action = "Mahakandmine"
            elif typ == "manual_set":
                action = "Määramine"
            elif typ == "month_calc":
                action = "Kuu arvestus"
            elif typ == "calc_deduct":
                action = "Perioodi arvestus"
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
                elif raw_type == "manual_set":
                    row_tag = "t_set"
                elif raw_type == "calc_deduct":
                    row_tag = "t_calc"
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
            status_var.set(f"Kirjeid logis: {cnt}")
            
    def _update_negative_highlight(self):
        """Покрасить отрицательные остатки и показать предупреждение в статусе."""
        from decimal import Decimal, InvalidOperation

        any_negative = False
        for name, cfg in self.materials_cfg.items():
            var = cfg.get("remain_m3")
            ent = cfg.get("remain_entry")
            if not var or not ent:
                continue
            raw = (var.get() or "0").replace(",", ".").strip()
            try:
                val = Decimal(raw)
            except (InvalidOperation, TypeError):
                val = Decimal("0")

            if val < 0:
                any_negative = True
                try:
                    ent.configure(style="Neg.TEntry")
                except tk.TclError:
                    pass
            else:
                try:
                    ent.configure(style="TEntry")
                except tk.TclError:
                    pass

        # короткое предупреждение в статусе
        if any_negative:
            base = self.status_var.get() or ""
            warn = " | Tähelepanu: negatiivsed jäägid"
            if warn not in base:
                self.status_var.set(base + warn)
        else:
            # уберём предупреждение, если всё уже ок
            self.status_var.set((self.status_var.get() or "").replace(" | Tähelepanu: negatiivsed jäägid", ""))

    def _month_key(self) -> str:
        """Ключ месяца вида YYYY-MM из текущих комбобоксов."""
        return f"{self.year_var.get()}-{self.month_num_var.get()}"

    def _update_calc_button_state(self):
        """Включить/выключить кнопку 'Рассчитать' в зависимости от закрытого месяца."""
        if not hasattr(self, "_calc_btn"):
            return
        data = self._load_stock_data()
        mkey = self._month_key()
        has_calc = any(r for r in data.get("ledger", []) if r.get("type") == "month_calc" and r.get("month") == mkey)
        # Кнопку не блокируем, но пишем статус, если месяц уже посчитан
        self._calc_btn.configure(state="normal")
        if has_calc:
            self.status_var.set(f"Kuu {mkey} on juba arvestatud.")

    def _compute_locked_months(self):
        data = self._load_stock_data()
        locked = set()
        for rec in data.get("ledger", []):
            typ = rec.get("type")
            if typ == "month_calc" and rec.get("month"):
                locked.add(rec.get("month"))
            if typ == "calc_deduct":
                for m in rec.get("months_covered", []):
                    locked.add(m)
        return locked

    # ---------------- Calculations table ----------------
    def _reload_calculations(self):
        """Заполнить таблицу 'Расчёты' по фильтру effective_date (месяц/год)."""
        tree = getattr(self, "calc_tree", None)
        if not tree or not tree.winfo_exists():
            return

        for iid in tree.get_children():
            tree.delete(iid)

        year = self.calc_year_var.get().strip()
        month = self.calc_month_var.get().strip()

        data = self._load_stock_data()
        ledger = data.get("ledger", []) or []

        from datetime import datetime as _dt

        def _parse_eff(rec):
            eff = rec.get("effective_date")
            if eff:
                try:
                    return _dt.fromisoformat(eff).date()
                except Exception:
                    pass
            ts = rec.get("ts")
            if ts:
                try:
                    return _dt.fromisoformat(ts).date()
                except Exception:
                    pass
            return None

        filtered = []
        for rec in ledger:
            d = _parse_eff(rec)
            if not d:
                continue
            if d.year == int(year) and f"{d.month:02d}" == month:
                filtered.append((d, rec))

        # сортировка по effective_date
        filtered.sort(key=lambda x: x[0])

        for idx, (d, rec) in enumerate(filtered):
            typ_raw = (rec.get("type") or "").lower()
            if typ_raw == "manual_add":
                op = "ADD"
            elif typ_raw == "manual_sub":
                op = "SUB"
            elif typ_raw == "manual_set":
                op = "SET"
            elif typ_raw == "month_calc":
                op = "CALC_DEDUCT"
            elif typ_raw == "calc_deduct":
                op = "CALC_DEDUCT"
            else:
                op = typ_raw.upper() if typ_raw else ""

            qty = rec.get("amount_m2", "")
            period = ""
            if typ_raw == "month_calc":
                period = rec.get("month") or f"{d.year}-{d.month:02d}"
                pf = rec.get("period_from")
                pt = rec.get("period_to")
                if pf or pt:
                    period = f"{pf or ''} — {pt or ''}"
            if typ_raw == "calc_deduct":
                pf = rec.get("period_from")
                pt = rec.get("period_to")
                if pf or pt:
                    period = f"{pf or ''} — {pt or ''}"
                else:
                    months = rec.get("months_covered") or []
                    if months:
                        period = f"{months[0]}…{months[-1]}"
            source = rec.get("sources") or rec.get("note") or ""
            if isinstance(source, list):
                source = ", ".join(map(str, source))

            tree.insert(
                "", "end", iid=f"{d.isoformat()}-{idx}",
                values=(
                    d.isoformat(),
                    rec.get("material", ""),
                    op,
                    qty,
                    period,
                    source,
                ),
                tags=(typ_raw,)
            )

        # раскраска по типам (если есть стили)
        tree.tag_configure("manual_add", foreground="#0A7D00")
        tree.tag_configure("manual_sub", foreground="#A40000")
        tree.tag_configure("manual_set", foreground="#6A1B9A")
        tree.tag_configure("month_calc", foreground="#004A9F")
        tree.tag_configure("calc_deduct", foreground="#004A9F")

    def _on_calc_context(self, event):
        tree = getattr(self, "calc_tree", None)
        if not tree or not tree.winfo_exists():
            return
        iid = tree.identify_row(event.y)
        if iid:
            tree.selection_set(iid)
            try:
                self._calc_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self._calc_menu.grab_release()

    def _calc_get_selected(self):
        tree = getattr(self, "calc_tree", None)
        if not tree or not tree.winfo_exists():
            return None
        sel = tree.selection()
        if not sel:
            return None
        iid = sel[0]
        vals = tree.item(iid, "values")
        return iid, vals

    def _calc_delete_selected(self):
        sel = self._calc_get_selected()
        if not sel:
            messagebox.showwarning("Valik", "Valige kirje tabelist 'Расчёты'.")
            return
        iid, vals = sel
        eff_date, material, op, qty, period, source = vals

        data = self._load_stock_data()
        ledger = data.get("ledger", [])

        # находим запись по eff_date/material/operation/qty/period (первое совпадение)
        target = None
        for rec in ledger:
            eff = rec.get("effective_date") or ""
            mat = rec.get("material") or ""
            typ = (rec.get("type") or "").lower()
            if typ == "manual_add":
                op_raw = "ADD"
            elif typ == "manual_sub":
                op_raw = "SUB"
            elif typ == "manual_set":
                op_raw = "SET"
            elif typ == "month_calc":
                op_raw = "CALC_DEDUCT"
            elif typ == "calc_deduct":
                op_raw = "CALC_DEDUCT"
            else:
                op_raw = typ.upper()
            if eff == eff_date and mat == material and op_raw == op and str(rec.get("amount_m2", "")) == str(qty):
                target = rec
                break

        if not target:
            messagebox.showerror("Kustutamine", "Kirjet ei leitud ledgeris.")
            return

        if not messagebox.askyesno("Kinnitus", f"Kustutada valitud kirje?\n{eff_date} {material} {op} {qty}"):
            return

        ledger.remove(target)
        data["ledger"] = ledger
        self._recompute_balances_from_ledger(data)
        self._save_stock_data(data)
        self._update_negative_highlight()
        self._reload_calculations()
        self._reload_ledger()

    def _calc_edit_selected(self):
        sel = self._calc_get_selected()
        if not sel:
            messagebox.showwarning("Valik", "Valige kirje tabelist 'Расчёты'.")
            return
        iid, vals = sel
        eff_date, material, op, qty, period, source = vals
        data = self._load_stock_data()
        ledger = data.get("ledger", [])

        # найти первую подходящую запись
        target = None
        target_idx = None
        for i, rec in enumerate(ledger):
            typ = (rec.get("type") or "").lower()
            op_raw = "CALC_DEDUCT" if typ == "month_calc" else "ADD" if typ == "manual_add" else "SUB" if typ == "manual_sub" else "SET" if typ == "manual_set" else typ.upper()
            if rec.get("effective_date") == eff_date and rec.get("material") == material and op_raw == op and str(rec.get("amount_m2", "")) == str(qty):
                target = rec
                target_idx = i
                break

        if target is None:
            messagebox.showerror("Muutmine", "Kirjet ei leitud ledgeris.")
            return

        if op == "CALC_DEDUCT":
            messagebox.showinfo("Muutmine", "CALC_DEDUCT redigeerimine lisatakse hiljem.")
            return

        # Создаём простую форму редактирования
        win = tk.Toplevel(self)
        win.title("Muuda kirje")
        win.grab_set()

        tk.Label(win, text="Effective date (dd.mm.yyyy)").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        eff_var = tk.StringVar(value=datetime.strptime(eff_date, "%Y-%m-%d").strftime("%d.%m.%Y") if eff_date else "")
        tk.Entry(win, textvariable=eff_var, width=16).grid(row=0, column=1, sticky="w", padx=8, pady=6)

        tk.Label(win, text="Materjal").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        mat_var = tk.StringVar(value=material or "POM Valge")
        tk.OptionMenu(win, mat_var, "POM Valge", "POM Must").grid(row=1, column=1, sticky="w", padx=8, pady=6)

        tk.Label(win, text="Operatsioon").grid(row=2, column=0, sticky="w", padx=8, pady=6)
        op_var = tk.StringVar(value=op)
        tk.OptionMenu(win, op_var, "ADD", "SUB", "SET").grid(row=2, column=1, sticky="w", padx=8, pady=6)

        tk.Label(win, text="Kogus m²").grid(row=3, column=0, sticky="w", padx=8, pady=6)
        qty_var = tk.StringVar(value=str(qty))
        tk.Entry(win, textvariable=qty_var, width=12).grid(row=3, column=1, sticky="w", padx=8, pady=6)

        status = tk.StringVar(value="")
        tk.Label(win, textvariable=status, fg="#C62828").grid(row=4, column=0, columnspan=2, sticky="w", padx=8, pady=4)

        def _save_edit():
            # validate date
            try:
                eff_dt = datetime.strptime(eff_var.get().strip(), "%d.%m.%Y").date()
                eff_iso = eff_dt.isoformat()
            except Exception:
                status.set("Kuupäev peab olema dd.mm.yyyy")
                return
            # validate qty
            try:
                amt = Decimal(str(qty_var.get()).replace(",", "."))
            except Exception:
                status.set("Kogus peab olema number")
                return
            if amt < 0:
                status.set("Kogus ei tohi olla negatiivne")
                return
            op_choice = op_var.get()
            typ_new = "manual_add" if op_choice == "ADD" else "manual_sub" if op_choice == "SUB" else "manual_set"

            # apply
            rec = dict(target)
            rec["effective_date"] = eff_iso
            rec["material"] = mat_var.get()
            rec["type"] = typ_new
            rec["amount_m2"] = float(amt)
            ledger[target_idx] = rec
            data["ledger"] = ledger
            self._recompute_balances_from_ledger(data)
            self._save_stock_data(data)
            self._update_negative_highlight()
            self._reload_calculations()
            self._reload_ledger()
            win.destroy()

        tk.Button(win, text="Salvesta", command=_save_edit).grid(row=5, column=0, padx=8, pady=8, sticky="e")
        tk.Button(win, text="Loobu", command=win.destroy).grid(row=5, column=1, padx=8, pady=8, sticky="w")

    def _apply_stub(self):
        """Фиксируем расчёт месяца: пишем month_calc в журнал, закрываем месяц."""
        from decimal import Decimal, ROUND_HALF_UP
        import datetime as _dt

        mkey = self._month_key()
        data = self._load_stock_data()

        # Если уже есть расчёт за этот месяц — предупреждаем и даём пользователю сменить месяц
        locked = self._compute_locked_months()
        if mkey in locked:
            messagebox.showinfo("Arvestus on juba tehtud", f"Kuu {mkey} on juba arvestatud. Valige teine kuu või kustutage arvestus.")
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

        if self.period_mode_var.get():
            try:
                dt_from = _dt.date(int(self.period_from_year.get()), int(self.period_from_month.get()), 1)
                dt_to = _dt.date(int(self.period_to_year.get()), int(self.period_to_month.get()), 1)
            except Exception:
                messagebox.showerror("Periood", "Valige perioodi algus ja lõpp (YYYY-MM).")
                return
            if dt_from > dt_to:
                messagebox.showerror("Periood", "Alguskuupäev peab olema enne lõppkuupäeva.")
                return
            # блокировка пересечений
            months_covered = []
            cur = dt_from
            while cur <= dt_to:
                months_covered.append(f"{cur.year}-{cur.month:02d}")
                # следующий месяц
                if cur.month == 12:
                    cur = _dt.date(cur.year + 1, 1, 1)
                else:
                    cur = _dt.date(cur.year, cur.month + 1, 1)

            locked = self._compute_locked_months()
            overlap = [m for m in months_covered if m in locked]
            if overlap:
                messagebox.showerror("Lukus", f"Perioод kattub lukustatud kuudega: {', '.join(overlap)}")
                return

            # суммы берём из conv_totals (пока агрегат)
            data = self._load_stock_data()
            ts_now = _dt.datetime.now().isoformat(timespec="seconds")
            note = f"Perioodi arvestus {months_covered[0]}…{months_covered[-1]}"

            valge = _get("POM Valge")
            must  = _get("POM Must")

            # если конвертация дала 0, попробуем взять фактическое списание из ledger по effective_date
            if valge == 0 and must == 0:
                cons = {"POM Valge": Decimal("0"), "POM Must": Decimal("0")}
                def _eff(rec):
                    e = rec.get("effective_date")
                    if not e:
                        return None
                    try:
                        return _dt.date.fromisoformat(e)
                    except Exception:
                        return None
                for rec in data.get("ledger", []):
                    d = _eff(rec)
                    if not d or d < dt_from or d > dt_to:
                        continue
                    typ = (rec.get("type") or "").lower()
                    if typ in ("manual_sub", "month_calc", "calc_deduct"):
                        mat = rec.get("material")
                        if mat in cons:
                            cons[mat] += Decimal(str(rec.get("amount_m2", 0)))
                valge = cons["POM Valge"]
                must = cons["POM Must"]

            if valge == 0 and must == 0:
                messagebox.showwarning("Andmed puuduvad", "Valitud perioodi kohta ei ole materjale mahakandmiseks.")
                return

            def _append(mat_name, amt_dec):
                if amt_dec <= 0:
                    return
                data["ledger"].append({
                    "ts": ts_now,
                    "effective_date": dt_to.isoformat(),
                    "period_from": months_covered[0],
                    "period_to": months_covered[-1],
                    "months_covered": months_covered,
                    "material": mat_name,
                    "type": "calc_deduct",
                    "amount_m2": float(amt_dec.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
                    "note": note
                })

            _append("POM Valge", valge)
            _append("POM Must", must)

            self._recompute_balances_from_ledger(data)
            self._save_stock_data(data)
            self._update_negative_highlight()
            self._reload_ledger()
            # показать расчёт в таблице за месяц окончания периода
            self.calc_year_var.set(str(dt_to.year))
            self.calc_month_var.set(f"{dt_to.month:02d}")
            self._reload_calculations()
            self._update_calc_button_state()
            self.status_var.set(f"Periood salvestatud: {months_covered[0]}…{months_covered[-1]}")
            messagebox.showinfo("Valmis", f"Periood salvestatud: {months_covered[0]}…{months_covered[-1]}")
            return

        valge = _get("POM Valge")
        must  = _get("POM Must")

        # Если оба нули — предупредим и выйдем
        if valge == 0 and must == 0:
            messagebox.showwarning("Andmed puuduvad", "Valitud kuu kohta ei ole materjale mahakandmiseks.")
            return

        # Запишем операции month_calc (только ненулевые)
        ts_now = _dt.datetime.now().isoformat(timespec="seconds")
        note = f"Auto: kuu {mkey} arvestus"

        if valge > 0:
            data["ledger"].append({
                "ts": ts_now,
                "month": mkey,
                "material": "POM Valge",
                "type": "month_calc",
                "effective_date": f"{mkey}-01",
                "amount_m2": float(valge.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)),
                "note": note
            })
        if must > 0:
            data["ledger"].append({
                "ts": ts_now,
                "month": mkey,
                "material": "POM Must",
                "type": "month_calc",
                "effective_date": f"{mkey}-01",
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
        self._update_negative_highlight()
        self._save_config()

        # Обновить журнал (если окно открыто) и заблокировать кнопку
        self._reload_ledger()
        self.calc_year_var.set(self.year_var.get())
        self.calc_month_var.set(self.month_num_var.get())
        self._reload_calculations()
        self._update_calc_button_state()
        messagebox.showinfo("Valmis", f"Kuu {mkey} arvestus on salvestatud.")

    def _on_exit(self):
        # При выходе всегда сохраняем last_month и текущий pdf_dir
        try:
            self._save_config()
        finally:
            self.master.destroy()

    def _set_status(self, text: str):
        self.status_var.set(text)
