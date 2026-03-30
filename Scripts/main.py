import customtkinter as ctk
from tkinter import ttk
from app import caja_valentia_app

if __name__ == "__main__":
    root = ctk.CTk()
    style = ttk.Style()
    style.theme_use("clam")

    app = caja_valentia_app(root)
    root.mainloop()
