import tkinter as tk
import os
from tkinter import messagebox

#clase que controla la lista de archivos interactuando con la interfaz grafica para permitir añadir eliminar reordenar y visualizar las rutas seleccionadas

class GestorLista:

    def __init__(self, lista_widget, archivos_seleccionados, log_fn=print):
        """
        inicializa el gestor de la lista vinculando el componente visual de la interfaz con la estructura de datos interna
        y estableciendo el canal de comunicacion para emitir mensajes hacia la consola de la aplicacion
        Args: lista_widget -> referencia al componente listbox nativo de tkinter donde se mostraran los nombres
              archivos_seleccionados -> lista mutable compartida con la aplicacion principal que almacena las rutas absolutas
              log_fn -> funcion inyectada para registrar y mostrar mensajes de estado en la interfaz grafica
        Return: no regresa valores pero deja configurados los atributos de la clase para su posterior manipulacion
        """
        self.lista_widget = lista_widget
        self.archivos = archivos_seleccionados
        self.log = log_fn

    def actualizar_visual(self):
        """
        sincroniza el contenido del componente visual listbox borrando su estado anterior y repoblandolo
        iterativamente con los nombres de archivo extraidos de la lista interna actualizada
        Args: no recibe parametros externos ya que opera directamente sobre el estado actual de self.archivos
        Return: no regresa nada unicamente modifica el estado visual del widget en la pantalla
        """
        self.lista_widget.delete(0, tk.END) #limpiamos absolutamente todo el contenido visual actual del componente listbox para prepararlo para una insercion limpia desde cero
        for archivo in self.archivos: #iteramos sobre la lista de rutas absolutas extrayendo unicamente el nombre final del archivo mediante os.path para insertarlo en la interfaz de forma legible
            self.lista_widget.insert(tk.END, f"  {os.path.basename(archivo)}")

    def anadir(self, nuevos_archivos):
        """
        recibe una coleccion de rutas de archivos seleccionadas por el usuario y las concatena al final de la lista
        interna actualizando inmediatamente la vista e imprimiendo un registro en la consola
        Args: nuevos_archivos -> tupla o lista iterables que contiene las rutas absolutas seleccionadas desde el explorador
        Return: no regresa nada
        """
        self.archivos.extend(list(nuevos_archivos)) #convertimos la tupla de seleccion nativa del sistema a una lista de python y la concatenamos al final de nuestra estructura de datos interna
        self.actualizar_visual() #invocamos el metodo de actualizacion para que los nuevos elementos se rendericen inmediatamente en la pantalla
        self.log(f"-> Se han añadido {len(nuevos_archivos)} archivos. Total: {len(self.archivos)}.") #enviamos un mensaje informativo a la consola indicando el exito de la operacion matematica y el total acumulado

    def mover_arriba(self):
        """
        desplaza el elemento actualmente seleccionado una posicion hacia arriba dentro de la lista interna
        intercambiando su indice con el elemento predecesor y actualizando la seleccion visual
        Args: no recibe parametros, lee directamente el estado del widget
        Return: no regresa nada
        """
        seleccion = self.lista_widget.curselection()
        if not seleccion: #verificamos si el usuario realmente ha seleccionado algun elemento del listbox haciendo click antes de intentar realizar cualquier calculo de movimiento para evitar errores fatales de indice
            return
        idx = seleccion[0]
        if idx > 0: #comprobamos logicamente que el elemento seleccionado no sea el primero de la lista ya que si su indice es cero es matematicamente imposible desplazarlo mas hacia arriba
            self.archivos[idx - 1], self.archivos[idx] = self.archivos[idx], self.archivos[idx - 1] #realizamos un intercambio directo de variables en python para invertir las posiciones de los archivos en la memoria
            self.actualizar_visual() #refrescamos el componente grafico para que el cambio de orden logico se refleje ante los ojos del usuario
            self.lista_widget.selection_set(idx - 1) #restauramos la seleccion visual sobre el mismo elemento que acabamos de mover hacia arriba para que el usuario no pierda el foco de trabajo

    def mover_abajo(self):
        """
        desplaza el elemento actualmente resaltado una posicion hacia abajo dentro de la estructura de datos
        permutando su lugar con el siguiente elemento y manteniendo el enfoque visual para facilitar multiples movimientos
        Args: no recibe parametros
        Return: no regresa nada
        """
        seleccion = self.lista_widget.curselection()
        if not seleccion: #comprobamos que exista una seleccion activa en el widget grafico para evitar un fallo critico al intentar extraer el indice cero de una tupla vacia
            return
        idx = seleccion[0]
        if idx < len(self.archivos) - 1: #validamos mediante comparacion que el indice actual sea estrictamente menor al ultimo indice disponible para garantizar que el archivo tenga espacio fisico hacia donde bajar
            self.archivos[idx + 1], self.archivos[idx] = self.archivos[idx], self.archivos[idx + 1] #permutamos el archivo seleccionado con el que se encuentra inmediatamente despues ocupando su lugar en la lista interna
            self.actualizar_visual() #solicitamos una reconstruccion completa del widget visual para hacer efectivo el cambio posicional en la interfaz
            self.lista_widget.selection_set(idx + 1) #movemos el cursor de seleccion hacia abajo acompañando al archivo desplazado para permitir la continuidad operativa

    def eliminar(self):
        """
        procesa la eliminacion de uno o multiples archivos seleccionados simultaneamente desde el listbox
        limpiandolos de la lista interna en orden inverso para preservar la integridad de los indices y ajustando la seleccion remanente
        Args: no recibe parametros
        Return: no regresa nada
        """
        seleccion = self.lista_widget.curselection()
        if not seleccion: #evaluamos si hay elementos marcados y de no ser asi lanzamos inmediatamente una ventana modal interactiva de advertencia indicandole al usuario la accion correctiva que debe tomar
            messagebox.showinfo("Información", "Selecciona un archivo de la lista para eliminar.")
            return
        #iteramos de forma inversa para garantizar que la alteracion de los indices al borrar un elemento no afecte las posiciones de los siguientes elementos a eliminar
        nombres = []
        for idx in reversed(seleccion): #recorremos la tupla de indices seleccionados comenzando obligatoriamente desde el ultimo hacia el primero para garantizar que las eliminaciones no desplacen los indices restantes
            nombres.append(os.path.basename(self.archivos[idx])) #extraemos el nombre corto del archivo antes de destruirlo y lo guardamos temporalmente para poder referenciarlo en el registro de la consola al finalizar
            del self.archivos[idx] #borramos definitivamente la ruta del archivo de nuestra lista maestra almacenada en memoria usando el comando de eliminacion por indice
        self.actualizar_visual() #ordenamos repintar el listbox que ahora contendra menos elementos reflejando la limpieza
        if self.archivos: #verificamos de forma segura si todavia quedan elementos sobrevivientes en la lista despues de la purga masiva para intentar reasignar el cursor de seleccion a una posicion valida
            nuevo_idx = min(seleccion[0], len(self.archivos) - 1)
            self.lista_widget.selection_set(nuevo_idx)
        if len(nombres) == 1: #evaluamos la longitud exacta de la lista de nombres borrados para decidir con precision gramatical si enviamos un mensaje en singular o un resumen en plural a la consola
            self.log(f"-> Se eliminó de la lista: {nombres[0]}")
        else:
            self.log(f"-> Se eliminaron {len(nombres)} archivos de la lista.")

    def vaciar(self):
        """
        ejecuta un borrado completo y absoluto de la memoria interna que almacena los archivos seleccionados
        y sincroniza este estado de vacio con la interfaz visual
        Args: no recibe parametros
        Return: no regresa nada
        """
        self.archivos.clear() #vaciamos completamente la lista de referencias en un solo paso invocando el metodo nativo que libera la memoria y reinicia el estado del gestor
        self.actualizar_visual() #notificamos inmediatamente al componente grafico que la lista logica ha sido borrada para que elimine todos los items visuales de la pantalla