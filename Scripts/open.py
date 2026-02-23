import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import scipy.io
import os

class MatToPandasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Automatización .mat a Pandas")
        self.root.geometry("600x500")

        # Variable para almacenar las rutas de los archivos seleccionados
        self.archivos_seleccionados = []

        # --- INTERFAZ GRÁFICA ---
        
        # Título
        lbl_titulo = tk.Label(root, text="Procesador de Archivos MATLAB (.mat)", font=("Arial", 14, "bold"))
        lbl_titulo.pack(pady=10)

        # Botón Seleccionar
        btn_seleccionar = tk.Button(root, text="Seleccionar Archivos .mat", command=self.seleccionar_archivos, bg="#4CAF50", fg="white", font=("Arial", 11))
        btn_seleccionar.pack(pady=5)

        # Lista para mostrar archivos seleccionados
        self.lista_archivos = tk.Listbox(root, selectmode=tk.MULTIPLE, width=70, height=10)
        self.lista_archivos.pack(pady=10)

        # Botón Iniciar
        btn_procesar = tk.Button(root, text="2. Iniciar Procesamiento", command=self.iniciar_proceso, bg="#2196F3", fg="white", font=("Arial", 12, "bold"))
        btn_procesar.pack(pady=10)

        # Área de texto para logs/resultados
        self.log_area = scrolledtext.ScrolledText(root, width=70, height=8, state='disabled')
        self.log_area.pack(pady=10)

    def log(self, mensaje):
        """Función auxiliar para escribir en la consola de la GUI"""
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, mensaje + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def seleccionar_archivos(self):
        """Abre el explorador de archivos para selección múltiple"""
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos .mat",
            filetypes=[("Archivos MATLAB", "*.mat")]
        )
        
        if archivos:
            self.archivos_seleccionados = archivos
            self.lista_archivos.delete(0, tk.END) # Limpiar lista anterior
            for archivo in archivos:
                nombre_base = os.path.basename(archivo)
                self.lista_archivos.insert(tk.END, nombre_base)
            
            self.log(f"-> Se han cargado {len(archivos)} archivos listos para procesar.")

    def leer_mat_a_df(self, ruta_archivo):
        """Lógica interna para limpiar el .mat y pasarlo a DataFrame"""
        try:
            mat_data = scipy.io.loadmat(ruta_archivo)
            data_limpia = {}
            
            # Filtramos metadata de MATLAB (empiezan con __)
            for key, value in mat_data.items():
                if not key.startswith('__'):
                    # Aplanamos arrays para que entren bien en columnas
                    if hasattr(value, 'flatten'):
                        data_limpia[key] = value.flatten()
                    else:
                        data_limpia[key] = value

            df = pd.DataFrame(data_limpia)
            df['archivo_origen'] = os.path.basename(ruta_archivo)
            return df
        except Exception as e:
            self.log(f"Error en {os.path.basename(ruta_archivo)}: {str(e)}")
            return None

    def iniciar_proceso(self):
        """Ejecuta la automatización"""
        if not self.archivos_seleccionados:
            messagebox.showwarning("Atención", "Por favor selecciona archivos primero.")
            return

        self.log("--- Iniciando Procesamiento ---")
        lista_dfs = []

        for ruta in self.archivos_seleccionados:
            df = self.leer_mat_a_df(ruta)
            if df is not None:
                lista_dfs.append(df)
                self.log(f"OK: {os.path.basename(ruta)} (Filas: {len(df)})")

        if lista_dfs:

            df_final = pd.concat(lista_dfs, ignore_index=True)
            
            self.log(f"--- ÉXITO: DataFrame Final creado ---")
            self.log(f"Total Filas: {df_final.shape[0]}, Total Columnas: {df_final.shape[1]}")
            
            # Ejemplo de operación posterior: Guardar a Excel para verificar
            # O mostrar las primeras filas en un mensaje
            resumen = df_final.head().to_string()
            self.log("Vista previa de los primeros datos:\n" + resumen)
            
            messagebox.showinfo("Proceso Terminado", "Los archivos han sido procesados y unidos correctamente.")
            
            # Opcional: Aquí podrías retornar df_final para usarlo en otra parte del código
            return df_final
            
        else:
            self.log("No se pudieron procesar los DataFrames.")
