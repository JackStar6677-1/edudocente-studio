import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches

base_dir = os.path.dirname(os.path.abspath(__file__))
out_dir = os.path.join(base_dir, "assets", "circuit_components")
os.makedirs(out_dir, exist_ok=True)

# 1. SÍMBOLO FUENTE O GENERADOR (PILA) - Placa larga (+), placa corta gruesa (-)
fig, ax = plt.subplots(figsize=(3.0, 1.5), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(2, 8)
ax.axis('off')

# Cable izquierdo
ax.plot([0.5, 4.3], [5.0, 5.0], color="#173F73", lw=3)
# Placa positiva (+) - Larga y delgada
ax.plot([4.3, 4.3], [2.2, 7.8], color="#173F73", lw=3.5)
# Placa negativa (-) - Corta y gruesa
ax.plot([5.7, 5.7], [3.3, 6.7], color="#173F73", lw=7.5)
# Cable derecho
ax.plot([5.7, 9.5], [5.0, 5.0], color="#173F73", lw=3)

# Signos + y -
ax.text(3.6, 7.3, "+", fontsize=16, fontweight='bold', color='#173F73', ha='center', va='center')
ax.text(6.4, 7.3, "–", fontsize=18, fontweight='bold', color='#173F73', ha='center', va='center')

plt.savefig(os.path.join(out_dir, "simbolo_pila.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

# 2. SÍMBOLO CABLES CONDUCTORES - Línea horizontal continua
fig, ax = plt.subplots(figsize=(3.0, 1.2), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(3, 7)
ax.axis('off')
ax.plot([0.5, 9.5], [5.0, 5.0], color="#173F73", lw=4.5)
plt.savefig(os.path.join(out_dir, "simbolo_cable.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

# 3. SÍMBOLO INTERRUPTOR (Muestra Abierto y Cerrado)
fig, ax = plt.subplots(figsize=(3.2, 2.0), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(0.5, 8.5)
ax.axis('off')

# --- Estado Abierto (arriba) ---
ax.plot([0.5, 3.2], [6.5, 6.5], color="#173F73", lw=3)
c1 = patches.Circle((3.5, 6.5), 0.35, ec="#173F73", fc="white", lw=2.5)
c2 = patches.Circle((6.5, 6.5), 0.35, ec="#173F73", fc="white", lw=2.5)
ax.add_patch(c1)
ax.add_patch(c2)
ax.plot([3.5, 6.5], [6.8, 8.2], color="#173F73", lw=3.2) # Palanca abierta
ax.plot([6.8, 9.5], [6.5, 6.5], color="#173F73", lw=3)
ax.text(5.0, 5.1, "Interruptor Abierto", fontsize=8.5, fontweight='bold', color='#173F73', ha='center')

# --- Estado Cerrado (abajo) ---
ax.plot([0.5, 3.2], [2.5, 2.5], color="#173F73", lw=3)
c3 = patches.Circle((3.5, 2.5), 0.35, ec="#173F73", fc="white", lw=2.5)
c4 = patches.Circle((6.5, 2.5), 0.35, ec="#173F73", fc="white", lw=2.5)
ax.add_patch(c3)
ax.add_patch(c4)
ax.plot([3.5, 6.5], [2.5, 2.5], color="#173F73", lw=3.2) # Palanca cerrada
ax.plot([6.8, 9.5], [2.5, 2.5], color="#173F73", lw=3)
ax.text(5.0, 1.1, "Interruptor Cerrado", fontsize=8.5, fontweight='bold', color='#173F73', ha='center')

plt.savefig(os.path.join(out_dir, "simbolo_interruptor.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

# 4. SÍMBOLO RECEPTOR / AMPOLLETA (Círculo con cruz X y dos líneas laterales)
fig, ax = plt.subplots(figsize=(3.0, 1.6), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(2, 8)
ax.axis('off')

# Cable izquierdo
ax.plot([0.5, 3.2], [5.0, 5.0], color="#173F73", lw=3.5)
# Círculo central perfectamente redondo
bulb_circle = patches.Circle((5.0, 5.0), 1.8, ec="#173F73", fc="white", lw=3.2)
ax.add_patch(bulb_circle)
# Cruz "X" en su interior
# r = 1.8 -> delta = 1.8 * 0.7071 = 1.27
dx = 1.27
ax.plot([5.0 - dx, 5.0 + dx], [5.0 - dx, 5.0 + dx], color="#173F73", lw=3.2)
ax.plot([5.0 - dx, 5.0 + dx], [5.0 + dx, 5.0 - dx], color="#173F73", lw=3.2)
# Cable derecho
ax.plot([6.8, 9.5], [5.0, 5.0], color="#173F73", lw=3.5)

plt.savefig(os.path.join(out_dir, "simbolo_ampolleta.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

# 5. SÍMBOLO RESISTENCIA (Opción complementaria: Zigzag / Rectángulo)
fig, ax = plt.subplots(figsize=(3.0, 1.5), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(2, 8)
ax.axis('off')
# Cable izquierdo
ax.plot([0.5, 3.0], [5.0, 5.0], color="#173F73", lw=3.5)
# Rectángulo de resistencia estándar
res_rect = patches.Rectangle((3.0, 3.8), 4.0, 2.4, ec="#173F73", fc="white", lw=3.2)
ax.add_patch(res_rect)
# Cable derecho
ax.plot([7.0, 9.5], [5.0, 5.0], color="#173F73", lw=3.5)
ax.text(5.0, 7.0, "Resistencia (R)", fontsize=9, fontweight='bold', color='#173F73', ha='center')

plt.savefig(os.path.join(out_dir, "simbolo_resistencia.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

print("Perfect circular symbols generated successfully!")
