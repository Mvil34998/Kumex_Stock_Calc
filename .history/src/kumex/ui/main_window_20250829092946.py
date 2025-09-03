"""
Главное окно GUI.
"""
import tkinter as tk

class MainWindow(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.create_widgets()

    def create_widgets(self):
        label = tk.Label(self, text="Kumex — GUI каркас", font=("Arial", 14))
        label.pack(pady=20)
