import tkinter as tk
import math
import random


class FondoNeuronal:
    """
    Maneja el canvas animado con partículas (efecto neuronal).
    Se instancia pasándole el root de tkinter.
    """

    def __init__(self, root, num_particles=60):
        self.root = root
        self.num_particles = num_particles
        self.particles = []
        self.mouse_x = -1000
        self.mouse_y = -1000

        self._construir_canvas()
        self._poblar_particulas()
        self.root.bind("<Motion>", self._actualizar_mouse)

    def _construir_canvas(self):
        self.canvas = tk.Canvas(self.root, bg="#0d0d0d", highlightthickness=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)

    def _poblar_particulas(self):
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

    def animar(self):
        """Dibuja un frame de la animación y se auto-programa para el siguiente."""
        self.canvas.delete("all")
        width  = self.root.winfo_width()
        height = self.root.winfo_height()

        for p in self.particles:
            p['x'] += p['vx']
            p['y'] += p['vy']
            if p['x'] < 0 or p['x'] > width:  p['vx'] *= -1
            if p['y'] < 0 or p['y'] > height: p['vy'] *= -1

            self.canvas.create_oval(
                p['x'] - 2, p['y'] - 2, p['x'] + 2, p['y'] + 2,
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

        self.root.after(30, self.animar)
