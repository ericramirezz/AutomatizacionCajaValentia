import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import scrolledtext
import pandas as pd
import scipy.io
import openpyxl
import os
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

class caja_valentia_app:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de .mat a Excel")
        self.root.geometry("800x600")
        self.root.bind("<Escape>", lambda e: self.root.destroy())


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
            background=[("active", "#a10b0c")] #no megusto el color
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

        # Frame para los botones de selección
        buttons_selection_frame = ttk.Frame(main_frame)
        buttons_selection_frame.pack(pady=5, fill="x")

        # Botón seleccionar .mats (izquierda)
        btn_seleccionar_mat = ttk.Button(
            buttons_selection_frame,
            text="Seleccionar Archivos .mat",
            command=self.seleccionar_archivos,
            style="Red.TButton"
        )
        btn_seleccionar_mat.pack(side="left", padx=80)
        
        # Botón seleccionar excel (derecha)
        btn_seleccionar_xlsx = ttk.Button(
            buttons_selection_frame,
            text="Seleccionar Archivos .xlsx",
            command=self.seleccionar_archivos,
            style="Red.TButton"
        )
        btn_seleccionar_xlsx.pack(side="right", padx=80)

        # Lista de archivos
        self.lista_archivos = tk.Listbox(
            main_frame,
            selectmode=tk.MULTIPLE,
            width=70,
            height=10
        )
        self.lista_archivos.pack(pady=10)

        # Frame para los botones de procesamiento
        buttons_process_frame = ttk.Frame(main_frame)
        buttons_process_frame.pack(pady=5, fill="x")

        # Botón procesar .mat (izquierda)
        btn_procesar = ttk.Button(
            buttons_process_frame,
            text="Iniciar Procesamiento de .mat (misma día, diferentes ratas)",
            command=self.iniciar_proceso,
            style="Red.TButton"
        )
        btn_procesar.pack(side="left", padx=20)

        # Botón procesar .xlsx (derecha)
        btn_procesar_xlsx = ttk.Button(
            buttons_process_frame,
            text="Graficar .xlsx (diferentes días)",
            command=self.procesar_xlsx,
            style="Red.TButton"
        )
        btn_procesar_xlsx.pack(side="right", padx=0)

        # Área de logs
        self.log_area = scrolledtext.ScrolledText(
            main_frame,
            width=100,
            height=10,
            state='disabled'
        )
        self.log_area.pack(pady=10)
        
        lbl_esc = ttk.Label(
            self.root,
            text="Presiona ESC para cerrar el programa",
            font=("Arial", 8),
            foreground="black"
            )
        lbl_esc.place(x=10, rely=1.0, anchor="sw", y=-5)

    # ---------- FUNCIONES ----------

    def log(self, mensaje):
        """Muestra mensajes en la consola de la interfaz gráfica."""
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, mensaje + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def procesar_xlsx(self):
        pass  

    def seleccionar_archivos(self):
        """Abre el explorador para seleccionar múltiples archivos .mat."""
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
        """Convierte un archivo .mat específico a un DataFrame de Pandas."""
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
                # Opcional: dejamos el archivo origen por si se requiere rastrear
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

    def eliminar_lados_consecutivos(self, df):
        """
        Elimina filas consecutivas duplicadas en la columna 'Lado'.
        Mantiene solo la primera fila de cada grupo de valores consecutivos iguales.
        """
        if 'Lado' not in df.columns:
            return df
        
        # Detectar cambios en la columna 'Lado'
        cambios = df['Lado'].ne(df['Lado'].shift()).cumsum()
        
        # Mantener solo la primera fila de cada grupo
        df_filtrado = df.groupby(cambios).first().reset_index(drop=True)
        
        return df_filtrado

    def modificar_dataframe(self, df):
        df = df.drop(['Tiempo Absoluto' ,'Palancas Izq', 'Palancas Der', 'archivo_origen'], axis=1)  # Elimina columnas que no se necesitan para el análisis de cruces
        df = df[df['Desplazamiento'] > 1]  # Cuando el desplazamiento es menor a 1, se interpreta que la rata no cruzó 
        # Eliminar filas consecutivas duplicadas en 'Lado' para cuando el estimulo es del mismo lado.
        df = self.eliminar_lados_consecutivos(df)
        df['Ensayo'] = range(1, len(df) + 1)  # Agregar columna de número de ensayo (cruce) al inicio
        
        return df

    def calcular_promedios_latencia(self, dfs_dict):
            """
            Calcula el promedio de Latencia para Estim Electrico == 1 y == 0
            por cada archivo procesado. Retorna un DataFrame resumen.
            """
            filas = []
            for nombre_archivo, df in dfs_dict.items():
                if 'Latencia' not in df.columns or 'Estim Electrico' not in df.columns:
                    self.log(f"Advertencia: {nombre_archivo} no tiene columnas Latencia/Estim Electrico.")
                    continue

                promedio_con = df.loc[df['Estim Electrico'] == 1, 'Latencia'].mean()
                promedio_sin = df.loc[df['Estim Electrico'] == 0, 'Latencia'].mean()

                filas.append({
                    'Nombre de archivo': nombre_archivo,
                    'Promedio Seguro': round(promedio_sin, 1) if pd.notna(promedio_sin) else 0,
                    'Promedio Riesgo': round(promedio_con, 1) if pd.notna(promedio_con) else 0,
                })

            return pd.DataFrame(filas)
    
    def iniciar_proceso(self):
        """Ejecuta la lectura, modificación y exportación a un solo Excel."""
        if not self.archivos_seleccionados:
            messagebox.showwarning("Atención", "Por favor selecciona archivos primero.")
            return

        self.dfs_mat = {}

        # 1. Leer y modificar cada archivo
        for ruta in self.archivos_seleccionados:
            df = self.leer_mat_a_df(ruta)
            
            if df is not None:
                # AQUÍ LLAMAMOS A LA FUNCIÓN MODIFICADORA
                df_modificado = self.modificar_dataframe(df)
                
                nombre_archivo = os.path.basename(ruta)
                self.dfs_mat[nombre_archivo] = df_modificado
                self.log(f"OK: {nombre_archivo} procesado correctamente con {len(df_modificado)} filas finales.")

        if not self.dfs_mat:
            self.log("No se pudieron procesar los archivos.")
            return

        # 2. Guardar en un solo Excel con múltiples hojas
        self.log(f"Se generaron {len(self.dfs_mat)} DataFrames. Preparando exportación...")
        
        # Pedir al usuario dónde guardar el archivo Excel
        ruta_guardado = filedialog.asksaveasfilename(
            title="Guardar datos procesados en Excel",
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")],
            initialfile="Resultados_Ensayos.xlsx"
        )

        if not ruta_guardado:
            self.log("Exportación cancelada por el usuario.")
            return

        try:
            # Usar pd.ExcelWriter para poder crear múltiples hojas
            with pd.ExcelWriter(ruta_guardado, engine='openpyxl') as writer:
                
                # Hoja extra con promedios de latencia por archivo
                df_promedios = self.calcular_promedios_latencia(self.dfs_mat)
                promedio_riesgo = round(df_promedios['Promedio Riesgo'].mean(), 1)
                promedio_seguro = round(df_promedios['Promedio Seguro'].mean(), 1)
                indice_nueva_fila = len(df_promedios) + 2  # Dejar una fila vacía después de los datos
                df_promedios.loc[indice_nueva_fila, 'Promedio Riesgo'] = promedio_riesgo
                df_promedios.loc[indice_nueva_fila, 'Promedio Seguro'] = promedio_seguro
                
                self.dfs_mat['Promedios Latencia'] = df_promedios  # Agregar el DataFrame de promedios al diccionario para que se guarde como hoja también
                for nombre_archivo, df_final in self.dfs_mat.items():
                    
                    # Excel tiene un límite de 31 caracteres para el nombre de las hojas
                    # Limpiamos la extensión .mat y recortamos si es necesario
                    nombre_hoja = nombre_archivo.replace('.mat', '')[:31]
                    
                    # Guardamos el DataFrame en su propia hoja, sin incluir el índice (0,1,2...)
                    df_final.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    
                    worksheet = writer.sheets[nombre_hoja]
                    
                    for cell in worksheet[1]: # worksheet[1] representa la primera fila
                        cell.font = Font(bold=True)
                    
                    for col_idx, column in enumerate(df_final.columns, start=1):
                        # Obtenemos la letra de la columna en Excel (A, B, C...)
                        col_letter = get_column_letter(col_idx)
                        
                        # Iniciamos asumiendo que el ancho máximo es el largo del nombre de la columna
                        max_length = len(str(column))
                        
                        # Revisamos celda por celda en esa columna para ver si hay un texto más largo
                        for cell in worksheet[col_letter]:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        
                        # Asignamos el nuevo ancho (le sumamos 2 para que no quede muy apretado)
                        worksheet.column_dimensions[col_letter].width = max_length + 2    
            
            self.log(f"Archivo guardado en {ruta_guardado}")
            messagebox.showinfo(
                "Proceso Terminado",
                f"Se procesaron {len(self.dfs_mat)} archivos y se guardaron correctamente en el Excel."
            )

        except Exception as e:
            self.log(f"Error crítico al intentar guardar el Excel: {str(e)}")
            messagebox.showerror("Error", "No se pudo guardar el archivo Excel. Revisa los logs.")

