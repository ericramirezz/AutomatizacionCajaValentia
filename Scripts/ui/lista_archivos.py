import tkinter as tk
import os
from tkinter import messagebox


class GestorLista:
    """
    Maneja la lógica de la lista de archivos:
    añadir, eliminar, reordenar y sincronizar con el widget Listbox.
    """

    def __init__(self, lista_widget, archivos_seleccionados, log_fn=print):
        """
        Args:
            lista_widget:         el tk.Listbox de la interfaz
            archivos_seleccionados: lista compartida con la app principal
            log_fn:               función de log de la app principal
        """
        self.lista_widget = lista_widget
        self.archivos = archivos_seleccionados
        self.log = log_fn

    def actualizar_visual(self):
        """Sincroniza el Listbox con el contenido de self.archivos."""
        self.lista_widget.delete(0, tk.END)
        for archivo in self.archivos:
            self.lista_widget.insert(tk.END, f"  {os.path.basename(archivo)}")

    def anadir(self, nuevos_archivos):
        """Añade una lista de rutas y refresca el widget."""
        self.archivos.extend(list(nuevos_archivos))
        self.actualizar_visual()
        self.log(f"-> Se han añadido {len(nuevos_archivos)} archivos. Total: {len(self.archivos)}.")

    def mover_arriba(self):
        seleccion = self.lista_widget.curselection()
        if not seleccion:
            return
        idx = seleccion[0]
        if idx > 0:
            self.archivos[idx - 1], self.archivos[idx] = self.archivos[idx], self.archivos[idx - 1]
            self.actualizar_visual()
            self.lista_widget.selection_set(idx - 1)

    def mover_abajo(self):
        seleccion = self.lista_widget.curselection()
        if not seleccion:
            return
        idx = seleccion[0]
        if idx < len(self.archivos) - 1:
            self.archivos[idx + 1], self.archivos[idx] = self.archivos[idx], self.archivos[idx + 1]
            self.actualizar_visual()
            self.lista_widget.selection_set(idx + 1)

    def eliminar(self):
        seleccion = self.lista_widget.curselection()
        if not seleccion:
            messagebox.showinfo("Información", "Selecciona un archivo de la lista para eliminar.")
            return
        idx = seleccion[0]
        nombre = os.path.basename(self.archivos[idx])
        del self.archivos[idx]
        self.actualizar_visual()
        if self.archivos:
            nuevo_idx = min(idx, len(self.archivos) - 1)
            self.lista_widget.selection_set(nuevo_idx)
        self.log(f"-> Se eliminó de la lista: {nombre}")
