import tkinter as tk
from tkinter import ttk
import open

if __name__ == "__main__":
    root = tk.Tk()

    # Forzar theme compatible en macOS
    style = ttk.Style()
    style.theme_use("clam")

    app = open.caja_valentia_app(root)
    root.mainloop()