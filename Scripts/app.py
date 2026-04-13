import os
import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk
import matplotlib.pyplot as plt
from PIL import Image

from ui.fondo_neuronal import FondoNeuronal
from ui.interfaz import construir_interfaz
from ui.lista_archivos import GestorLista
from logica.procesador_mat import leer_mat_a_df, modificar_dataframe, guardar_excel
from logica.graficador import generar_grafica, guardar_excel_resumen

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

        # 1. Fondo animado (debe ir primero para quedar detrás de todo)
        self.fondo = FondoNeuronal(root)

        # 2. Interfaz (crea los widgets y los asigna a self.*)
        construir_interfaz(self)

        # 3. Gestor de lista (necesita self.lista_archivos ya creado)
        self.gestor = GestorLista(
            lista_widget=self.lista_archivos,
            archivos_seleccionados=self.archivos_seleccionados,
            log_fn=self.log
        )

        # 4. Logo y animación
        self._cargar_logo()
        self.fondo.animar()

    # ── LOGO ────────────────────────────────────────────────────────────────
    def _cargar_logo(self):
        ruta_script = os.path.dirname(os.path.abspath(__file__))
        ruta_imagen = os.path.join(ruta_script, self.FILENAME_LOGO)

        if os.path.exists(ruta_imagen):
            try:
                img = Image.open(ruta_imagen)
                self.imagen_cargada = ctk.CTkImage(
                    light_image=img, dark_image=img, size=(60, 60)
                )
                self.lbl_imagen.configure(image=self.imagen_cargada, text="")
            except Exception as e:
                self.log(f"Aviso: No se pudo cargar el logo: {str(e)}")
        else:
            self.lbl_imagen.configure(text="[Valentía Lab]", text_color="gray")

    # ── CONSOLA ──────────────────────────────────────────────────────────────
    def _limpiar_log(self):
        self.log_area.configure(state='normal')
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state='disabled')

    def log(self, mensaje):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, mensaje + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    # ── GESTIÓN DE ARCHIVOS (delegada al GestorLista) ─────────────────────
    def seleccionar_archivos(self):
        archivos = filedialog.askopenfilenames(
            title="Seleccionar archivos",
            filetypes=[
                ("Archivos compatibles", "*.mat *.xlsx"),
                ("Archivos MATLAB", "*.mat"),
                ("Archivos Excel", "*.xlsx"),
            ]
        )
        if archivos:
            self.gestor.anadir(archivos)

    def mover_arriba(self):
        self.gestor.mover_arriba()

    def mover_abajo(self):
        self.gestor.mover_abajo()

    def eliminar_archivo(self):
        self.gestor.eliminar()

    # ── PROCESAMIENTO .MAT ───────────────────────────────────────────────
    def iniciar_proceso_latencias(self):
        self._limpiar_log()
        if not self.archivos_seleccionados:
            messagebox.showwarning("Atención", "Por favor selecciona archivos primero.")
            return

        dfs_mat = {}
        for ruta in self.archivos_seleccionados:
            df = leer_mat_a_df(ruta, log_fn=self.log)
            if df is not None:
                df_modificado = modificar_dataframe(df)
                nombre_archivo = os.path.basename(ruta)
                dfs_mat[nombre_archivo] = df_modificado
                self.log(f"OK: {nombre_archivo} procesado ({len(df_modificado)} ensayos).")

        if not dfs_mat:
            self.log("No se pudieron procesar los archivos.")
            return

        ruta_guardado = filedialog.asksaveasfilename(
            title="Guardar datos procesados",
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx")],
            initialfile="Resultados_Ensayos.xlsx"
        )
        if not ruta_guardado:
            return

        ok = guardar_excel(ruta_guardado, dfs_mat, log_fn=self.log)
        if ok:
            messagebox.showinfo("Terminado", f"Se procesaron {len(dfs_mat)} archivos.")
            self.gestor.vaciar()
        else:
            messagebox.showerror("Error", "No se pudo guardar el archivo Excel.")

    # ── GRAFICACIÓN .XLSX ────────────────────────────────────────────────
    def procesar_xlsx(self):
        self._limpiar_log()
        if not self.archivos_seleccionados:
            self.log("Operación cancelada. No se seleccionaron archivos.")
            return

        num_dias = len(self.archivos_seleccionados)
        fig, df_resumen, dias_ratas_raw = generar_grafica(
            self.archivos_seleccionados, log_fn=self.log
        )
        if fig is None:
            return

        # 1. Mostrar ventana interactiva (se queda abierta hasta que el usuario la cierre)
        plt.show(block=False)

        # 2. Guardar imagen PNG
        ruta_png = filedialog.asksaveasfilename(
            title="Guardar gráfica",
            defaultextension=".png",
            filetypes=[("Imagen PNG", "*.png"), ("Imagen JPEG", "*.jpg")],
            initialfile="grafica_latencias.png"
        )
        if ruta_png:
            fig.savefig(ruta_png, dpi=300, bbox_inches='tight')
            self.log(f"Gráfica guardada en: {ruta_png}")
        else:
            self.log("Guardado de imagen cancelado.")

        # 3. Guardar Excel de resumen (Resumen Tercios + Datos por Rata)
        ruta_xlsx = filedialog.asksaveasfilename(
            title="Guardar Excel de resumen",
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
            initialfile="resumen_tercios.xlsx"
        )
        if ruta_xlsx:
            guardar_excel_resumen(
                df_resumen, ruta_xlsx,
                dias_ratas_raw=dias_ratas_raw,
                num_dias=num_dias,
                log_fn=self.log
            )
        else:
            self.log("Guardado de Excel cancelado.")

        self.gestor.vaciar()