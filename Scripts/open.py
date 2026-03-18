import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import scipy.io
import openpyxl
import os
import math
import random
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt

import customtkinter as ctk
from PIL import Image

# Configuración global del tema
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class caja_valentia_app:
    def __init__(self, root):
        self.root = root
        self.root.title("Procesador de .mat a Excel")
        self.root.minsize(900, 750) 
        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.FILENAME_LOGO = "ifc_logo.png"
        self.archivos_seleccionados = []
        self.imagen_cargada = None

        # Variables para el efecto neuronal
        self.particles = []
        self.num_particles = 60
        self.mouse_x = -1000
        self.mouse_y = -1000

        self._construir_fondo_neuronal()
        self._construir_interfaz()
        
        self.intentar_cargar_logo_automatico()
        self._animar_neuronas()

    # ---------- EFECTO DE NEURONAS (PARTÍCULAS) ----------
    def _construir_fondo_neuronal(self):
        self.canvas = tk.Canvas(self.root, bg="#0d0d0d", highlightthickness=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)
        
        self.root.bind("<Motion>", self._actualizar_mouse)
        
        for _ in range(self.num_particles):
            self.particles.append({
                'x': random.randint(0, 1200),
                'y': random.randint(0, 900),
                'vx': random.uniform(-0.5, 0.5),
                'vy': random.uniform(-0.5, 0.5)
            })

    def _actualizar_mouse(self, event):
        self.mouse_x = event.x
        self.mouse_y = event.y

    def _animar_neuronas(self):
        self.canvas.delete("all")
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        for p in self.particles:
            p['x'] += p['vx']
            p['y'] += p['vy']
            
            if p['x'] < 0 or p['x'] > width: p['vx'] *= -1
            if p['y'] < 0 or p['y'] > height: p['vy'] *= -1

            self.canvas.create_oval(
                p['x']-2, p['y']-2, p['x']+2, p['y']+2, 
                fill="#4a0404", outline=""
            )

        for i in range(len(self.particles)):
            p1 = self.particles[i]
            dist_mouse = math.hypot(p1['x'] - self.mouse_x, p1['y'] - self.mouse_y)
            
            if dist_mouse < 150:
                self.canvas.create_line(
                    p1['x'], p1['y'], self.mouse_x, self.mouse_y, 
                    fill="#ff1a1a", width=1.5
                )

            for j in range(i + 1, len(self.particles)):
                p2 = self.particles[j]
                dist = math.hypot(p1['x'] - p2['x'], p1['y'] - p2['y'])
                
                if dist < 100:
                    self.canvas.create_line(
                        p1['x'], p1['y'], p2['x'], p2['y'], 
                        fill="#330b0b", width=1
                    )

        self.root.after(30, self._animar_neuronas)

    # ---------- INTERFAZ MODERNA CON CUSTOMTKINTER ----------
    def _crear_tarjeta_accion(self, parent, texto_boton, texto_subtitulo, comando):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        
        btn = ctk.CTkButton(
            frame, text=texto_boton, command=comando, 
            fg_color="#800607", hover_color="#a10b0c",
            font=("Arial", 14, "bold"), corner_radius=15, height=40
        )
        btn.pack(fill="x", pady=(0, 5))
        
        lbl = ctk.CTkLabel(
            frame, text=texto_subtitulo, font=("Arial", 11), 
            text_color="#a3a3a3", wraplength=280, justify="center"
        )
        lbl.pack()
        return frame

    def _construir_interfaz(self):
        main_frame = ctk.CTkFrame(self.root, fg_color="#1a1a1a", corner_radius=20, border_width=1, border_color="#333333")
        main_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.9, relheight=0.95)

        content = ctk.CTkFrame(main_frame, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=30, pady=30)

        # ---------- HEADER ----------
        header_frame = ctk.CTkFrame(content, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 15))
        
        lbl_titulo = ctk.CTkLabel(header_frame, text="Análisis de ensayos con cruces", font=("Arial", 24, "bold"))
        lbl_titulo.pack(side="left")

        self.lbl_imagen = ctk.CTkLabel(header_frame, text="")
        self.lbl_imagen.pack(side="right")

        # ---------- SECCIÓN DE CARGA ----------
        carga_frame = ctk.CTkFrame(content, fg_color="#242424", corner_radius=15)
        carga_frame.pack(fill="x", pady=(0, 10), ipady=10)
        
        ctk.CTkLabel(carga_frame, text="1. Carga de Archivos", font=("Arial", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 5))

        btn_container_1 = ctk.CTkFrame(carga_frame, fg_color="transparent")
        btn_container_1.pack(fill="x", padx=20, pady=5)
        btn_container_1.columnconfigure(0, weight=1)
        btn_container_1.columnconfigure(1, weight=1)

        tarjeta_mat = self._crear_tarjeta_accion(
            btn_container_1, "Seleccionar Archivos .mat", 
            "Sube los archivos de MATLAB originales para procesar latencias.", self.seleccionar_archivos
        )
        tarjeta_mat.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        tarjeta_xlsx = self._crear_tarjeta_accion(
            btn_container_1, "Seleccionar Archivos .xlsx", 
            "Sube los archivos Excel generados para comparar días.", self.seleccionar_archivos
        )
        tarjeta_xlsx.grid(row=0, column=1, padx=(10, 0), sticky="ew")

        # ---------- LISTA DE ARCHIVOS Y ORDENAMIENTO ----------
        lista_frame = ctk.CTkFrame(content, fg_color="transparent")
        lista_frame.pack(fill="x", pady=10)
        
        self.lista_archivos = tk.Listbox(
            lista_frame, height=5, bg="#121212", fg="#ffffff", 
            selectbackground="#800607", selectforeground="white",
            borderwidth=1, relief="flat", highlightcolor="#333333", font=("Arial", 10)
        )
        self.lista_archivos.pack(side="left", fill="x", expand=True, padx=(0, 10))

        # Botones para ordenar y eliminar
        botones_orden_frame = ctk.CTkFrame(lista_frame, fg_color="transparent")
        botones_orden_frame.pack(side="right")

        btn_subir = ctk.CTkButton(botones_orden_frame, text="▲ Subir", width=80, command=self.mover_arriba, fg_color="#333333", hover_color="#555555")
        btn_subir.pack(pady=(0, 5))

        btn_bajar = ctk.CTkButton(botones_orden_frame, text="▼ Bajar", width=80, command=self.mover_abajo, fg_color="#333333", hover_color="#555555")
        btn_bajar.pack(pady=(0, 10))

        btn_eliminar = ctk.CTkButton(botones_orden_frame, text="✖ Eliminar", width=80, command=self.eliminar_archivo, fg_color="#800607", hover_color="#a10b0c")
        btn_eliminar.pack()

        # ---------- SECCIÓN DE PROCESAMIENTO ----------
        procesamiento_frame = ctk.CTkFrame(content, fg_color="#242424", corner_radius=15)
        procesamiento_frame.pack(fill="x", pady=(10, 15), ipady=10)
        
        ctk.CTkLabel(procesamiento_frame, text="2. Acciones", font=("Arial", 14, "bold")).pack(anchor="w", padx=20, pady=(10, 5))

        btn_container_2 = ctk.CTkFrame(procesamiento_frame, fg_color="transparent")
        btn_container_2.pack(fill="x", padx=20, pady=5)
        btn_container_2.columnconfigure(0, weight=1)
        btn_container_2.columnconfigure(1, weight=1)

        tarjeta_procesar_mat = self._crear_tarjeta_accion(
            btn_container_2, "Iniciar Procesamiento de .mat",
            "Limpia datos, elimina lados consecutivos y genera Excel consolidado.", self.iniciar_proceso_latencias
        )
        tarjeta_procesar_mat.grid(row=0, column=0, padx=(0, 10), sticky="ew")

        tarjeta_procesar_xlsx = self._crear_tarjeta_accion(
            btn_container_2, "Graficar .xlsx",
            "Genera gráfica de dispersión comparando promedios en 5 días.", self.procesar_xlsx
        )
        tarjeta_procesar_xlsx.grid(row=0, column=1, padx=(10, 0), sticky="ew")

        # ---------- AVISO ESC Y LOGS (ORDEN INVERTIDO PARA QUE NO SE CORTE) ----------
        
        # 1. Empaquetamos el texto ESC al fondo del contenedor
        lbl_esc = ctk.CTkLabel(
            content, 
            text="Presiona ESC para salir del programa.", 
            font=("Arial", 11, "bold"), 
            text_color="#a3a3a3"
        )
        lbl_esc.pack(side="bottom", anchor="w", pady=(5, 0))

        # 2. Empaquetamos la consola de logs justo arriba del texto ESC
        self.log_area = ctk.CTkTextbox(content, height=90, corner_radius=10, fg_color="#0d0d0d", text_color="#00ff00")
        self.log_area.pack(side="bottom", fill="both", expand=True)
        self.log_area.configure(state="disabled")

    def intentar_cargar_logo_automatico(self):
        ruta_script = os.path.dirname(os.path.abspath(__file__))
        ruta_imagen = os.path.join(ruta_script, self.FILENAME_LOGO)

        if os.path.exists(ruta_imagen):
            try:
                img = Image.open(ruta_imagen)
                self.imagen_cargada = ctk.CTkImage(light_image=img, dark_image=img, size=(60, 60))
                self.lbl_imagen.configure(image=self.imagen_cargada, text="")
            except Exception as e:
                self.log(f"Aviso: No se pudo cargar el logo: {str(e)}")
        else:
            self.lbl_imagen.configure(text="[Valentía Lab]", text_color="gray")


    # =========================================================================
    # LÓGICA DE GESTIÓN DE LA LISTA
    # =========================================================================

    def _actualizar_lista_visual(self):
        self.lista_archivos.delete(0, tk.END)
        for archivo in self.archivos_seleccionados:
            self.lista_archivos.insert(tk.END, f"  {os.path.basename(archivo)}")

    def mover_arriba(self):
        seleccion = self.lista_archivos.curselection()
        if not seleccion:
            return
        idx = seleccion[0]
        if idx > 0:
            self.archivos_seleccionados[idx-1], self.archivos_seleccionados[idx] = self.archivos_seleccionados[idx], self.archivos_seleccionados[idx-1]
            self._actualizar_lista_visual()
            self.lista_archivos.selection_set(idx-1)

    def mover_abajo(self):
        seleccion = self.lista_archivos.curselection()
        if not seleccion:
            return
        idx = seleccion[0]
        if idx < len(self.archivos_seleccionados) - 1:
            self.archivos_seleccionados[idx+1], self.archivos_seleccionados[idx] = self.archivos_seleccionados[idx], self.archivos_seleccionados[idx+1]
            self._actualizar_lista_visual()
            self.lista_archivos.selection_set(idx+1)

    def eliminar_archivo(self):
        seleccion = self.lista_archivos.curselection()
        if not seleccion:
            messagebox.showinfo("Información", "Por favor, selecciona un archivo de la lista para eliminar.")
            return
        
        idx = seleccion[0]
        archivo_eliminado = os.path.basename(self.archivos_seleccionados[idx])
        
        del self.archivos_seleccionados[idx]
        self._actualizar_lista_visual()
        
        if self.archivos_seleccionados:
            nuevo_idx = min(idx, len(self.archivos_seleccionados) - 1)
            self.lista_archivos.selection_set(nuevo_idx)
            
        self.log(f"-> Se eliminó de la lista: {archivo_eliminado}")


    # =========================================================================
    # LÓGICA DE PROCESAMIENTO
    # =========================================================================

    def log(self, mensaje):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, mensaje + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled') 
        
    def seleccionar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=[("Archivos MATLAB", "*.mat"), ("Archivos Excel", "*.xlsx")]
        )

        if archivos:
            self.archivos_seleccionados.extend(list(archivos))
            self._actualizar_lista_visual()
            self.log(f"-> Se han añadido {len(archivos)} archivos. Total en lista: {len(self.archivos_seleccionados)}.")

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

    def eliminar_lados_consecutivos(self, df):
        if 'Lado' not in df.columns: return df
        cambios = df['Lado'].ne(df['Lado'].shift()).cumsum()
        return df.groupby(cambios).first().reset_index(drop=True)

    def modificar_dataframe(self, df):
        df = df.drop(['Tiempo Absoluto' ,'Palancas Izq', 'Palancas Der', 'archivo_origen'], axis=1)
        df = df[df['Desplazamiento'] > 1]
        df = self.eliminar_lados_consecutivos(df)
        df['Ensayo'] = range(1, len(df) + 1)
        return df

    def calcular_promedios_latencia(self, dfs_dict):
        filas = []
        for nombre_archivo, df in dfs_dict.items():
            if 'Latencia' not in df.columns or 'Estim Electrico' not in df.columns:
                self.log(f"Advertencia: {nombre_archivo} no tiene columnas.")
                continue
            promedio_con = df.loc[df['Estim Electrico'] == 1, 'Latencia'].mean()
            promedio_sin = df.loc[df['Estim Electrico'] == 0, 'Latencia'].mean()
            filas.append({
                'Nombre de archivo': nombre_archivo,
                'Promedio Seguro': round(promedio_sin, 1) if pd.notna(promedio_sin) else 0,
                'Promedio Riesgo': round(promedio_con, 1) if pd.notna(promedio_con) else 0,
            })
        return pd.DataFrame(filas)
    
    def procesar_xlsx(self):
        if not self.archivos_seleccionados:
            self.log("Operación cancelada. No se seleccionaron archivos.")
            return
        
        archivos = self.archivos_seleccionados
    
        dias_seguros, valores_seguros = [], []
        dias_riesgo, valores_riesgo = [], []
        
        for dia_idx, ruta in enumerate(archivos, start=1):
            try:
                archivo_excel = pd.ExcelFile(ruta)
                ultima_hoja = archivo_excel.sheet_names[-1]
                df = archivo_excel.parse(ultima_hoja)
                columna_seguros = 'Promedio Seguro'
                columna_riesgo = 'Promedio Riesgo'
                
                if columna_seguros in df.columns and columna_riesgo in df.columns:
                    val_seguro = float(df[columna_seguros].iloc[-1])
                    val_riesgo = float(df[columna_riesgo].iloc[-1])
                    
                    if val_seguro != 0 and val_riesgo == 0:
                        dias_seguros.append(dia_idx)
                        valores_seguros.append(val_seguro)
                    elif val_riesgo != 0 and val_seguro == 0:
                        dias_riesgo.append(dia_idx)
                        valores_riesgo.append(val_riesgo)
                    elif val_seguro != 0 and val_riesgo != 0:
                        dias_seguros.append(dia_idx)
                        valores_seguros.append(val_seguro)
                        dias_riesgo.append(dia_idx)
                        valores_riesgo.append(val_riesgo)
                    else:
                        print(f"Día {dia_idx}: Ambos valores son 0.")
                else:
                    print(f"Error en {os.path.basename(ruta)}: Faltan columnas.")
            except Exception as e:
                print(f"Error procesando {os.path.basename(ruta)}: {str(e)}")

        plt.figure(figsize=(9, 6))
        if dias_seguros: plt.scatter(dias_seguros, valores_seguros, color='blue', label='Promedio Seguros', s=120, alpha=0.8)
        if dias_riesgo: plt.scatter(dias_riesgo, valores_riesgo, color='red', label='Promedio Riesgo', s=120, alpha=0.8)
            
        plt.title('Evolución de Promedios de Latencia', fontsize=15, fontweight='bold')
        plt.xlabel('Día de Prueba', fontsize=12)
        plt.ylabel('Promedio de Latencia', fontsize=12)
        plt.xticks(range(1, len(archivos) + 1))
        plt.grid(True, linestyle='--', alpha=0.6)
        plt.legend()
        
        ruta_png = 'grafica.png'
        plt.savefig(ruta_png, dpi=300, bbox_inches='tight')
        plt.close()
        
        self.log(f"¡Éxito! Gráfica guardada como: {ruta_png}")
        messagebox.showinfo("Gráfica Generada", f"La gráfica se guardó como:\n{ruta_png}")
    
    def iniciar_proceso_latencias(self):
        if not self.archivos_seleccionados:
            messagebox.showwarning("Atención", "Por favor selecciona archivos primero.")
            return

        self.dfs_mat = {}
        for ruta in self.archivos_seleccionados:
            df = self.leer_mat_a_df(ruta)
            if df is not None:
                df_modificado = self.modificar_dataframe(df)
                nombre_archivo = os.path.basename(ruta)
                self.dfs_mat[nombre_archivo] = df_modificado
                self.log(f"OK: {nombre_archivo} procesado.")

        if not self.dfs_mat:
            self.log("No se pudieron procesar los archivos.")
            return

        ruta_guardado = filedialog.asksaveasfilename(
            title="Guardar datos procesados",
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")],
            initialfile="Resultados_Ensayos.xlsx"
        )

        if not ruta_guardado: return

        try:
            with pd.ExcelWriter(ruta_guardado, engine='openpyxl') as writer:
                df_promedios = self.calcular_promedios_latencia(self.dfs_mat)
                promedio_riesgo = round(df_promedios['Promedio Riesgo'].mean(), 1)
                promedio_seguro = round(df_promedios['Promedio Seguro'].mean(), 1)
                indice_nueva_fila = len(df_promedios) + 2
                df_promedios.loc[indice_nueva_fila, 'Promedio Riesgo'] = promedio_riesgo
                df_promedios.loc[indice_nueva_fila, 'Promedio Seguro'] = promedio_seguro
                
                self.dfs_mat['Promedios Latencia'] = df_promedios
                
                for nombre_archivo, df_final in self.dfs_mat.items():
                    nombre_hoja = nombre_archivo.replace('.mat', '')[:31]
                    df_final.to_excel(writer, sheet_name=nombre_hoja, index=False)
                    
                    worksheet = writer.sheets[nombre_hoja]
                    for cell in worksheet[1]: cell.font = Font(bold=True)
                    
                    for col_idx, column in enumerate(df_final.columns, start=1):
                        col_letter = get_column_letter(col_idx)
                        max_length = len(str(column))
                        for cell in worksheet[col_letter]:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except: pass
                        worksheet.column_dimensions[col_letter].width = max_length + 2    
            
            self.log(f"Archivo guardado en {ruta_guardado}")
            messagebox.showinfo("Terminado", f"Se procesaron {len(self.dfs_mat)} archivos.")

        except Exception as e:
            self.log(f"Error crítico: {str(e)}")
            messagebox.showerror("Error", "No se pudo guardar el archivo Excel.")