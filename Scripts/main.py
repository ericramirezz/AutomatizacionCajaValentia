import customtkinter as ctk
import tkinter as tk
from tkinter import ttk
import open

if __name__ == "__main__":
    root = ctk.CTk()
    # Forzar theme compatible en macOS
    style = ttk.Style()
    style.theme_use("clam")

    app = open.caja_valentia_app(root)
    root.mainloop()