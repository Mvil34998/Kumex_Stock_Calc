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

from kumex.io.file_ops import load_json, save_json


class MainWindow(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master: tk.Tk = master
        self.year_var = tk.StringVar()
        self.month_num_var = tk.StringVar()
        
        
        self.pdf_files = []          # список путей найденных PDF
        self.material_rows = []      # сюда позже положим строки из PDF-парсера


        # --- пути и состояние ---
        # <корень проекта> = ../../../.. от этого файла
        self.project_root = Path(__file__).resolve().parents[4]
        self.state_dir = self.project_root / "state"
        self.config_path = self.state_dir / "kumex_config.json"

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

        # Ряд 0: Заголовок
        title = ttk.Label(container, text="Kumex — минимальный каркас", font=("Arial", 14))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))
        
        # Разделитель «-»
        ttk.Label(container, text="-").grid(row=1, column=1, padx=(64, 0), sticky="w")

        # Ряд 1: Месяц (YYYY-MM)
        ttk.Label(container, text="Месяц (YYYY-MM):").grid(row=1, column=0, sticky="w")
        month_entry = ttk.Entry(container, textvariable=self.month_var, width=12)
        month_entry.grid(row=1, column=1, sticky="w")
        
        # Месяцы 01..12
        months = [f"{m:02d}" for m in range(1, 13)]
        cb_month = ttk.Combobox(container, textvariable=self.month_num_var, values=months, width=4, state="readonly")
        cb_month.grid(row=1, column=1, padx=(80, 0), sticky="w")
        cb_month.bind("<<ComboboxSelected>>", lambda e: self._sync_month_var())
        
        # Годы: текущий-10 … текущий+1
        current_year = datetime.now().year
        years = [str(y) for y in range(current_year - 10, current_year + 2)]
        cb_year = ttk.Combobox(container, textvariable=self.year_var, values=years, width=6, state="readonly")
        cb_year.grid(row=1, column=1, sticky="w")
        cb_year.bind("<<ComboboxSelected>>", lambda e: self._sync_month_var())

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

        # Ряд 4: Кнопки действий
        scan_btn = ttk.Button(container, text="Сканировать PDF", command=self._scan_pdfs_stub)
        report_btn = ttk.Button(container, text="Сформировать отчёт", command=self._report_stub)
        apply_btn = ttk.Button(container, text="Применить", command=self._apply_stub)

        scan_btn.grid(row=4, column=0, sticky="w", pady=(8, 0))
        report_btn.grid(row=4, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        apply_btn.grid(row=4, column=2, sticky="w", pady=(8, 0))
        
        
        

        # колонки
        container.columnconfigure(1, weight=1)

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
            

    def _scan_pdfs_stub(self):
        # Заглушка: здесь позже будет вызов парсера
        messagebox.showinfo("Сканирование", f"Сканируем PDF в:\n{self.pdf_dir_var.get()}\nза {self.month_var.get()}")

    def _report_stub(self):
        # Заглушка: здесь позже будет агрегатор + генерация отчёта
        messagebox.showinfo("Отчёт", "Формирование отчёта (заглушка).")
        
    def _sync_month_var(self):
        yy = self.year_var.get()
        mm = self.month_num_var.get()
        if len(yy) == 4 and mm in {f"{i:02d}" for i in range(1,13)}:
            self.month_var.set(f"{yy}-{mm}")

    def _apply_stub(self):
        # Заглушка: здесь позже будет применение к kumex_stock.json
        messagebox.showinfo("Применение", "Применение отчёта к остаткам (заглушка).")

    def _on_exit(self):
        # При выходе всегда сохраняем last_month и текущий pdf_dir
        try:
            self._save_config()
        finally:
            self.master.destroy()
