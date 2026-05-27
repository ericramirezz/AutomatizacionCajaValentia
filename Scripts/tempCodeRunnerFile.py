import customtkinter as ctk
from tkinter import ttk
from app import caja_valentia_app

#punto de entrada de la aplicacion, arranca la ventana y entra al loop principal
if __name__ == "__main__":
    root = ctk.CTk()
    style = ttk.Style()
    #se usa el tema clam para que los widgets de tkinter clasico se vean mejor
    style.theme_use("clam")

    app = caja_valentia_app(root)
    root.mainloop()
