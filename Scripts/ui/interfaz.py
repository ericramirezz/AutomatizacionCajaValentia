import tkinter as tk
import customtkinter as ctk

#funciones que construyen todos los widgets visuales de la ventana principal

def construir_interfaz(app):
    """
    construye y empaqueta la totalidad de los elementos visuales que componen la ventana principal de la aplicacion
    organizando los contenedores botones y areas de texto mediante los gestores de geometria de customtkinter
    Args: app -> instancia principal de la clase que contiene la raiz de la ventana y los metodos logicos a ejecutar
    Return: no regresa valores pero inicializa y guarda las referencias de los widgets clave dentro del mismo objeto app
    """
    root = app.root

    main_frame = ctk.CTkFrame(
        root, fg_color="#1a1a1a", corner_radius=20,
        border_width=1, border_color="#333333"
    )
    main_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.9, relheight=0.95)

    content = ctk.CTkFrame(main_frame, fg_color="transparent")
    content.pack(fill="both", expand=True, padx=30, pady=30)

    #construimos el contenedor superior transparente que alojara el titulo principal de la herramienta y el espacio reservado para el logo institucional
    header_frame = ctk.CTkFrame(content, fg_color="transparent")
    header_frame.pack(fill="x", pady=(0, 15))

    ctk.CTkLabel(
        header_frame, text="Análisis de ensayos con cruces",
        font=("Arial", 24, "bold")
    ).pack(side="left")

    app.lbl_imagen = ctk.CTkLabel(header_frame, text="")
    app.lbl_imagen.pack(side="right")

    #creamos un marco visualmente diferenciado con bordes redondeados para agrupar los controles relacionados con la importacion de datos al sistema
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

    #implementamos un cuadro de lista interactivo de tkinter nativo acompañado de controles laterales para gestionar y ordenar la cola de procesamiento
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

    #diseñamos un segundo bloque de controles para alojar los botones de ejecucion principal que desencadenaran los algoritmos de analisis y graficado
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

    #configuramos un area de texto de solo lectura en la parte inferior para mostrar el registro de eventos y un texto de ayuda para atajos de teclado
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


def _tarjeta(parent, texto_boton, texto_subtitulo, comando):
    """
    genera un modulo de interfaz de usuario empaquetado y reutilizable que combina un boton interactivo de accion principal
    con una etiqueta de texto descriptiva debajo para guiar al usuario sobre el proposito exacto de la funcion a ejecutar
    Args: parent -> widget o frame contenedor padre donde se anidara y dibujara este nuevo bloque visual
        texto_boton -> cadena de texto principal que se mostrara centrada y resaltada dentro del boton clickeable
        texto_subtitulo -> texto de ayuda secundario que aparecera con una fuente mas pequeña y de color tenue en la parte inferior
        comando -> referencia directa a la funcion o metodo interno que se disparara automaticamente al hacer clic en el boton
    Return: regresa la instancia del frame transparente ya construido con todos sus elementos internos empaquetados y listos para inyectarse
    """
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