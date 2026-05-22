import tkinter as tk
import math
import random

class FondoNeuronal:

    def __init__(self, root, num_particles=60):
        """
        inicializa el gestor del fondo animado estableciendo la conexion directa con la ventana principal de la interfaz grafica
        y preparando las estructuras de datos fundamentales para almacenar y manipular el enjambre de particulas en tiempo real
        Args: root -> referencia a la ventana o frame principal de tkinter donde se renderizara el canvas oscuro
                num_particles -> cantidad entera de particulas simultaneas a generar para el fondo (valor por defecto de 60)
        Return: no regresa nada, pero deja inicializadas las variables de estado y enlaza los eventos del puntero
        """
        self.root = root
        self.num_particles = num_particles
        self.particles = []
        #inicializamos las coordenadas del cursor fuera del area visible para evitar que las particulas reaccionen antes de tiempo
        self.mouse_x = -1000
        self.mouse_y = -1000

        self._construir_canvas()
        self._poblar_particulas()
        #registramos un escuchador de eventos a nivel raiz que intercepta cualquier desplazamiento del puntero sobre la aplicacion
        self.root.bind("<Motion>", self._actualizar_mouse)

    def _construir_canvas(self):
        """
        crea y empaqueta un widget de tipo canvas con un color base profundamente oscuro que se expande dinamicamente
        para cubrir la totalidad del area de la ventana padre sirviendo como el lienzo de renderizado continuo
        Args: no recibe parametros externos, opera sobre la instancia root
        Return: no regresa nada, asigna la instancia del widget creado directamente a la variable de clase self.canvas
        """
        #instanciamos el lienzo eliminando cualquier borde resaltado nativo para lograr un acabado inmersivo sin cortes visuales
        self.canvas = tk.Canvas(self.root, bg="#0d0d0d", highlightthickness=0)
        #utilizamos el gestor de geometria place con coordenadas relativas al cien por ciento para que el fondo sea completamente responsivo
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)

    def _poblar_particulas(self):
        """
        genera el estado inicial del sistema de particulas calculando coordenadas de aparicion aleatorias distribuidas
        por el espacio de la ventana y asignando vectores de velocidad independientes para simular un movimiento organico
        Args: no recibe parametros, consume la variable de configuracion self.num_particles
        Return: no regresa nada, inyecta los diccionarios con coordenadas y velocidades dentro de la lista self.particles
        """
        for _ in range(self.num_particles): #iteramos segun la cantidad configurada para inyectar diccionarios independientes con la fisica de cada nodo
            self.particles.append({
                'x': random.randint(0, 1200),
                'y': random.randint(0, 900),
                'vx': random.uniform(-0.5, 0.5),
                'vy': random.uniform(-0.5, 0.5)
            })

    def _actualizar_mouse(self, event):
        """
        actua como un manejador de eventos asincrono que intercepta constantemente las coordenadas del cursor del usuario
        dentro de la ventana para alimentar el algoritmo de atraccion y trazado de conexiones en el ciclo de animacion principal
        Args: event -> objeto de evento nativo de tkinter que empaqueta las coordenadas x e y actuales del puntero fisico
        Return: no regresa nada, unicamente actualiza las variables de estado interno del mouse para su posterior lectura
        """
        #extraemos la posicion bidimensional exacta del evento del sistema operativo y la guardamos en la memoria de la clase
        self.mouse_x = event.x
        self.mouse_y = event.y

    def animar(self):
        """
        representa el motor grafico principal del fondo neuronal que se ejecuta recursivamente limpiando el canvas en cada fotograma
        calculando la nueva fisica de rebote de cada particula en los bordes y trazando geometricamente las lineas de interaccion
        Args: no recibe parametros externos, procesa la geometria global actual
        Return: no regresa nada, pero programa de forma autonoma su propio siguiente cuadro de ejecucion utilizando after
        """
        #borramos absolutamente todos los trazos del fotograma anterior para evitar que el canvas se sature de lineas residuales infinitas
        self.canvas.delete("all")
        #consultamos dinamicamente las dimensiones reales de la ventana en este instante preciso para garantizar que los rebotes sean exactos aunque se redimensione
        width  = self.root.winfo_width()
        height = self.root.winfo_height()

        for p in self.particles: #iteramos sobre el enjambre completo para actualizar la cinetica individual sumando el vector de velocidad a su posicion actual
            p['x'] += p['vx']
            p['y'] += p['vy']
            #evaluamos mediante limites condicionales si el nodo sobrepaso los margenes visibles del lienzo e invertimos su aceleracion para simular un rebote elastico
            if p['x'] < 0 or p['x'] > width:  p['vx'] *= -1
            if p['y'] < 0 or p['y'] > height: p['vy'] *= -1

            #renderizamos el cuerpo de la particula como un pequeño circulo solido de color rojizo oscuro sin bordes aparentes para mantener la estetica neuronal
            self.canvas.create_oval(
                p['x'] - 2, p['y'] - 2, p['x'] + 2, p['y'] + 2,
                fill="#4a0404", outline=""
            )

        for i in range(len(self.particles)): #implementamos un algoritmo de anidacion para comparar distancias vectoriales iterando desde la particula actual hacia adelante
            p1 = self.particles[i]
            dist_mouse = math.hypot(p1['x'] - self.mouse_x, p1['y'] - self.mouse_y)

            #comprobamos si el calculo de la hipotenusa indica que el cursor humano esta dentro del radio de influencia para trazar el enlace de atraccion rojo brillante
            if dist_mouse < 150:
                self.canvas.create_line(
                    p1['x'], p1['y'], self.mouse_x, self.mouse_y,
                    fill="#ff1a1a", width=1.5
                )

            #ciclo secundario optimizado que compara la particula base unicamente con las particulas restantes en la lista evitando recalcular conexiones inversas ya evaluadas
            for j in range(i + 1, len(self.particles)):
                p2 = self.particles[j]
                dist = math.hypot(p1['x'] - p2['x'], p1['y'] - p2['y'])
                if dist < 100: #si la proximidad entre ambos nodos es suficiente dibujamos una linea sutil y oscura para simular una sinapsis estable en la red de fondo
                    self.canvas.create_line(
                        p1['x'], p1['y'], p2['x'], p2['y'],
                        fill="#330b0b", width=1
                    )

        #le indicamos al gestor de eventos de tkinter que vuelva a invocar esta misma funcion en treinta milisegundos logrando una animacion fluida de aproximadamente treinta cuadros por segundo
        self.root.after(30, self.animar)