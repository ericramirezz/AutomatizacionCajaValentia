import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import scrolledtext
import pandas as pd
import scipy.io
import os


class MatToPandasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de .mat a Excel")
        self.root.geometry("800x600")

        self.archivos_seleccionados = []

        # ---------- ESTILOS ----------
        style = ttk.Style()
        
        style.configure(
            "Titulo.TLabel",
            font=("Arial", 16, "bold")
        )

        style.configure(
            "Red.TButton",
            background="#800607",
            foreground="white",
            font=("Arial", 12, "bold"),
            padding=6
        )

        style.map(
            "Red.TButton",
            background=[("active", "#a10b0c")]
        )

        # ---------- INTERFAZ ----------
        
        # Frame principal con padding
        main_frame = ttk.Frame(self.root, padding=20)
        main_frame.pack(fill="both", expand=True)

        # Título
        lbl_titulo = ttk.Label(
            main_frame,
            text="Análisis de ensayos con cruces",
            style="Titulo.TLabel"
        )
        lbl_titulo.pack(pady=10)

        # Botón seleccionar
        btn_seleccionar = ttk.Button(
            main_frame,
            text="Seleccionar Archivos .mat",
            command=self.seleccionar_archivos,
            style="Red.TButton"
        )
        btn_seleccionar.pack(pady=5)

        # Lista de archivos
        self.lista_archivos = tk.Listbox(
            main_frame,
            selectmode=tk.MULTIPLE,
            width=70,
            height=10
        )
        self.lista_archivos.pack(pady=10)

        # Botón procesar
        btn_procesar = ttk.Button(
            main_frame,
            text="Iniciar Procesamiento de Datos",
            command=self.iniciar_proceso,
            style="Red.TButton"
        )
        btn_procesar.pack(pady=10)

        # Área de logs
        self.log_area = scrolledtext.ScrolledText(
            main_frame,
            width=100,
            height=10,
            state='disabled'
        )
        self.log_area.pack(pady=10)

    # ---------- FUNCIONES ----------

    def log(self, mensaje):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, mensaje + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def seleccionar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos .mat",
            filetypes=[("Archivos MATLAB", "*.mat")]
        )

        if archivos:
            self.archivos_seleccionados = archivos
            self.lista_archivos.delete(0, tk.END)

            for archivo in archivos:
                nombre_base = os.path.basename(archivo)
                self.lista_archivos.insert(tk.END, nombre_base)

            self.log(f"-> Se han cargado {len(archivos)} archivos listos para procesar.")

    def leer_mat_a_df(self, ruta_archivo):
        try:
            columnas = [
                'Ensayo', 'Lado', 'Estim Electrico', 'Latencia',
                'Tiempo Absoluto', 'Palancas Izq',
                'Palancas Der', 'Desplazamiento'
            ]

            mat_data = scipy.io.loadmat(ruta_archivo)

            matriz_datos = None
            for key, value in mat_data.items():
                if not key.startswith('__'):
                    matriz_datos = value
                    break

            if matriz_datos is None:
                self.log(f"Error: No se encontraron datos válidos en {os.path.basename(ruta_archivo)}")
                return None

            if matriz_datos.shape[1] == len(columnas):
                df = pd.DataFrame(matriz_datos, columns=columnas)
                df['archivo_origen'] = os.path.basename(ruta_archivo)
                return df
            else:
                self.log(
                    f"Error en dimensiones: {os.path.basename(ruta_archivo)} "
                    f"tiene {matriz_datos.shape[1]} columnas, "
                    f"pero esperamos {len(columnas)}."
                )
                return None

        except Exception as e:
            self.log(f"Error al procesar {os.path.basename(ruta_archivo)}: {str(e)}")
            return None

    def iniciar_proceso(self):
        if not self.archivos_seleccionados:
            messagebox.showwarning("Atención", "Por favor selecciona archivos primero.")
            return

        self.log("--- Iniciando Procesamiento ---")

        self.dfs_mat = {}

        for ruta in self.archivos_seleccionados:
            df = self.leer_mat_a_df(ruta)
            if df is not None:
                nombre_archivo = os.path.basename(ruta)
                self.dfs_mat[nombre_archivo] = df
                self.log(f"OK: {nombre_archivo} guardado (Filas: {len(df)})")

        if self.dfs_mat:
            self.log(f"Se cargaron {len(self.dfs_mat)} DataFrames distintos ---")

            primer_archivo = list(self.dfs_mat.keys())[0]
            resumen = self.dfs_mat[primer_archivo].head().to_string()
            self.log(f"\nVista previa de: {primer_archivo}\n" + resumen)

            messagebox.showinfo(
                "Proceso Terminado",
                f"Se crearon {len(self.dfs_mat)} DataFrames individuales correctamente."
            )
        else:
            self.log("No se pudieron procesar los archivos.")