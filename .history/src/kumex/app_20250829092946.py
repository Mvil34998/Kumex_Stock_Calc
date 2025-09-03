"""
Kumex — точка входа утилиты.
"""

import tkinter as tk
from kumex.ui.main_window import MainWindow


def main():
    root = tk.Tk()
    root.title("Kumex")
    app = MainWindow(root)
    app.pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    main()
