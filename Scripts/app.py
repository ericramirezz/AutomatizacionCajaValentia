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

#configuracion global del tema visual de la aplicacion
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

#clase principal que coordina toda la aplicacion: interfaz, logica y archivos

class caja_valentia_app:

    def __init__(self, root):
        """
        inicializa la ventana principal de la aplicacion y orquesta el montaje de todos los componentes visuales
        y logicos estableciendo el tamaño minimo de la interfaz y preparando el arreglo de la sesion de archivos
        Args: root -> ventana raiz principal generada por customtkinter que contendra absolutamente toda la aplicacion
        Return: no regresa nada pero deja la aplicacion completamente construida interactiva y con el fondo animado corriendo
        """
        self.root = root
        self.root.title("Procesador de .mat a Excel")
        self.root.minsize(900, 750)
        self.root.bind("<Escape>", lambda e: self.root.destroy())

        self.FILENAME_LOGO = "C:\AutomatizacionCajaValentia\Scripts\ifc_logo.png"
        self.archivos_seleccionados = []

        #el fondo se crea primero para que quede detras de todos los demas widgets
        self.fondo = FondoNeuronal(root)

        #se construye la interfaz, que crea y asigna los widgets al objeto app
        construir_interfaz(self)

        #el gestor de lista necesita que lista_archivos ya exista, por eso va despues
        self.gestor = GestorLista(
            lista_widget=self.lista_archivos,
            archivos_seleccionados=self.archivos_seleccionados,
            log_fn=self.log
        )

        #se carga el logo y se arranca la animacion del fondo
        self._cargar_logo()
        self.fondo.animar()

    def _cargar_logo(self):
        """
        intenta localizar y cargar la imagen del logotipo institucional desde una ruta absoluta predefinida en el disco duro
        redimensionando el archivo para inyectarlo al encabezado o mostrando un texto alternativo discreto si no se encuentra
        Args: no recibe parametros externos ya que utiliza la constante de ruta definida en la misma clase
        Return: no regresa nada unicamente modifica el widget grafico de la etiqueta superior
        """
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
            #si no se encuentra el archivo se muestra un texto alternativo
            self.lbl_imagen.configure(text="[Valentía Lab]", text_color="gray")

    def _limpiar_log(self):
        """
        habilita de manera temporal los permisos de escritura del area de texto de la consola para purgar todo
        su contenido actual de forma inmediata y vuelve a bloquear el componente para mantener la seguridad visual
        Args: no recibe parametros
        Return: no regresa nada
        """
        self.log_area.configure(state='normal')
        self.log_area.delete("1.0", tk.END)
        self.log_area.configure(state='disabled')

    def log(self, mensaje):
        """
        inserta una nueva linea de texto al final de la consola grafica desplazando automaticamente el scroll
        para garantizar que el usuario siempre pueda leer las ultimas notificaciones del sistema sin interactuar
        Args: mensaje -> cadena de texto que contiene la informacion de diagnostico o aviso que se desea proyectar
        Return: no regresa nada
        """
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, mensaje + "\n")
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')

    def seleccionar_archivos(self):
        """
        despliega una ventana de exploracion nativa del sistema operativo forzando filtros de busqueda para asegurar
        que el usuario unicamente pueda seleccionar e importar archivos en los formatos procesables del experimento
        Args: no recibe parametros
        Return: no regresa nada pero manda las rutas capturadas directamente hacia el modulo gestor de la lista
        """
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
        """
        sirve como puente de interconexion que transmite la intencion del usuario de recorrer un archivo hacia
        la parte superior de la cola directamente al gestor logico encargado de actualizar la tabla visual
        Args: no recibe parametros
        Return: no regresa nada
        """
        self.gestor.mover_arriba()

    def mover_abajo(self):
        """
        actua como un canal delegador que comunica el clic del boton inferior con la clase especializada de la lista
        para desplazar el elemento seleccionado hacia el final de la secuencia de procesamiento
        Args: no recibe parametros
        Return: no regresa nada
        """
        self.gestor.mover_abajo()

    def eliminar_archivo(self):
        """
        retransmite el evento de borrado al controlador de la lista visual para expulsar el archivo de la memoria
        evitando que sea tomado en cuenta durante las iteraciones de los analisis estadisticos o generacion de graficas
        Args: no recibe parametros
        Return: no regresa nada
        """
        self.gestor.eliminar()

    def iniciar_proceso_latencias(self):
        """
        orquesta el flujo de trabajo principal de conversion iterando sobre el lote completo de archivos matlab crudos
        para limpiarlos individualmente estandarizarlos en diccionarios y finalmente condensar toda la data en un solo excel
        Args: no recibe parametros externos ya que opera basandose enteramente en la memoria de self.archivos_seleccionados
        Return: no regresa nada pero puede desplegar alertas modales si el usuario olvido cargar documentos o fallo el guardado
        """
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

        #se pide al usuario donde quiere guardar el archivo de resultados
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

    def procesar_xlsx(self):
        """
        coordina el proceso de analisis visual extrayendo los datos diarios compilados para mapearlos en una grafica
        de dispersion estadistica permitiendo guardar el canvas generado como imagen de alta resolucion y exportar tablas
        Args: no recibe parametros
        Return: no regresa nada
        """
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

        #se muestra la ventana interactiva de la grafica sin bloquear la aplicacion
        plt.show(block=False)

        #se pide donde guardar la imagen png
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

        #se pide donde guardar el excel con el resumen por bloques y por rata
        ruta_xlsx = filedialog.asksaveasfilename(
            title="Guardar Excel de resumen",
            defaultextension=".xlsx",
            filetypes=[("Archivo Excel", "*.xlsx")],
            initialfile="resumen_bloques.xlsx"
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