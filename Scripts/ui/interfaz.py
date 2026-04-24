import tkinter as tk
import customtkinter as ctk


def construir_interfaz(app):
    """
    Construye todos los widgets de la ventana principal y los asocia a `app`.
    Usa app.root como ventana raíz y deja referencias en el objeto app.

    Widgets creados en app:
        app.lbl_imagen      — label del logo en el header
        app.lista_archivos  — tk.Listbox de archivos
        app.log_area        — CTkTextbox de consola
    """
    root = app.root

    main_frame = ctk.CTkFrame(
        root, fg_color="#1a1a1a", corner_radius=20,
        border_width=1, border_color="#333333"
    )
    main_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.9, relheight=0.95)

    content = ctk.CTkFrame(main_frame, fg_color="transparent")
    content.pack(fill="both", expand=True, padx=30, pady=30)

    # ── HEADER ──────────────────────────────────────────────────────────────
    header_frame = ctk.CTkFrame(content, fg_color="transparent")
    header_frame.pack(fill="x", pady=(0, 15))

    ctk.CTkLabel(
        header_frame, text="Análisis de ensayos con cruces",
        font=("Arial", 24, "bold")
    ).pack(side="left")

    app.lbl_imagen = ctk.CTkLabel(header_frame, text="")
    app.lbl_imagen.pack(side="right")

    # ── SECCIÓN 1: CARGA DE ARCHIVOS ─────────────────────────────────────
    carga_frame = ctk.CTkFrame(content, fg_color="#242424", corner_radius=15)
    carga_frame.pack(fill="x", pady=(0, 10), ipady=10)

    ctk.CTkLabel(
        carga_frame, text="1. Carga de Archivos",
        font=("Arial", 14, "bold")
    ).pack(anchor="w", padx=20, pady=(10, 5))

    btn_container_1 = ctk.CTkFrame(carga_frame, fg_color="transparent")
    btn_container_1.pack(fill="x", padx=20, pady=5)
    btn_container_1.columnconfigure(0, weight=1)
    btn_container_1.columnconfigure(1, weight=1)

    _tarjeta(
        btn_container_1, "Seleccionar Archivos",
        "Sube los archivos de MATLAB originales para procesar latencias.",
        app.seleccionar_archivos
    ).grid(row=0, column=0, padx=(0, 10), sticky="ew")

    _tarjeta(
        btn_container_1, "Seleccionar para Gráfica",
        "Sube los archivos Excel generados para comparar días.",
        app.seleccionar_archivos
    ).grid(row=0, column=1, padx=(10, 0), sticky="ew")

    # ── LISTA DE ARCHIVOS ────────────────────────────────────────────────
    lista_frame = ctk.CTkFrame(content, fg_color="transparent")
    lista_frame.pack(fill="x", pady=10)

    app.lista_archivos = tk.Listbox(
        lista_frame, height=5, bg="#121212", fg="#ffffff",
        selectbackground="#800607", selectforeground="white",
        borderwidth=1, relief="flat", highlightcolor="#333333",
        font=("Arial", 10), selectmode=tk.EXTENDED
    )
    app.lista_archivos.pack(side="left", fill="x", expand=True, padx=(0, 10))

    botones_orden = ctk.CTkFrame(lista_frame, fg_color="transparent")
    botones_orden.pack(side="right")

    ctk.CTkButton(
        botones_orden, text="▲ Subir", width=80,
        command=app.mover_arriba, fg_color="#333333", hover_color="#555555"
    ).pack(pady=(0, 5))

    ctk.CTkButton(
        botones_orden, text="▼ Bajar", width=80,
        command=app.mover_abajo, fg_color="#333333", hover_color="#555555"
    ).pack(pady=(0, 10))

    ctk.CTkButton(
        botones_orden, text="✖ Eliminar", width=80,
        command=app.eliminar_archivo, fg_color="#800607", hover_color="#a10b0c"
    ).pack()

    # ── SECCIÓN 2: ACCIONES ──────────────────────────────────────────────
    procesamiento_frame = ctk.CTkFrame(content, fg_color="#242424", corner_radius=15)
    procesamiento_frame.pack(fill="x", pady=(10, 15), ipady=10)

    ctk.CTkLabel(
        procesamiento_frame, text="2. Acciones",
        font=("Arial", 14, "bold")
    ).pack(anchor="w", padx=20, pady=(10, 5))

    btn_container_2 = ctk.CTkFrame(procesamiento_frame, fg_color="transparent")
    btn_container_2.pack(fill="x", padx=20, pady=5)
    btn_container_2.columnconfigure(0, weight=1)
    btn_container_2.columnconfigure(1, weight=1)

    _tarjeta(
        btn_container_2, "Iniciar",
        "Limpia datos, elimina lados consecutivos y genera Excel consolidado.",
        app.iniciar_proceso_latencias
    ).grid(row=0, column=0, padx=(0, 10), sticky="ew")

    _tarjeta(
        btn_container_2, "Graficar",
        "Genera gráfica de dispersión comparando promedios en 5 días.",
        app.procesar_xlsx
    ).grid(row=0, column=1, padx=(10, 0), sticky="ew")

    # ── CONSOLA DE LOGS ──────────────────────────────────────────────────
    ctk.CTkLabel(
        content, text="Presiona ESC para salir del programa.",
        font=("Arial", 11, "bold"), text_color="#a3a3a3"
    ).pack(side="bottom", anchor="w", pady=(5, 0))

    app.log_area = ctk.CTkTextbox(
        content, height=90, corner_radius=10,
        fg_color="#0d0d0d", text_color="#00ff00"
    )
    app.log_area.pack(side="bottom", fill="both", expand=True)
    app.log_area.configure(state="disabled")


# ── Helpers privados ────────────────────────────────────────────────────────

def _tarjeta(parent, texto_boton, texto_subtitulo, comando):
    """Crea un frame con botón + subtítulo (widget reutilizable)."""
    frame = ctk.CTkFrame(parent, fg_color="transparent")

    ctk.CTkButton(
        frame, text=texto_boton, command=comando,
        fg_color="#800607", hover_color="#a10b0c",
        font=("Arial", 14, "bold"), corner_radius=15, height=40
    ).pack(fill="x", pady=(0, 5))

    ctk.CTkLabel(
        frame, text=texto_subtitulo, font=("Arial", 11),
        text_color="#a3a3a3", wraplength=280, justify="center"
    ).pack()

    return frame