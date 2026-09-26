import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

base_dir = os.path.dirname(os.path.abspath(__file__))
out_dir = os.path.join(base_dir, "assets", "matter_states")
os.makedirs(out_dir, exist_ok=True)

# 1. SÓLIDO (Partículas ordenadas y compactas)
fig, ax = plt.subplots(figsize=(2.2, 2.0), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis('off')

# Vaso / Recipiente sutil
ax.plot([1.5, 1.5, 8.5, 8.5], [8.5, 1.5, 1.5, 8.5], color="#B0C4DE", lw=2)

# Grid de partículas sólidas
for x in np.linspace(2.5, 7.5, 5):
    for y in np.linspace(2.2, 6.2, 4):
        p = patches.Circle((x, y), 0.5, ec="#173F73", fc="#6BA4D9", lw=1.5)
        ax.add_patch(p)

ax.text(5.0, 0.4, "Estado Sólido", fontsize=9, fontweight='bold', color='#173F73', ha='center')
plt.savefig(os.path.join(out_dir, "estado_solido.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

# 2. LÍQUIDO (Partículas cercanas, desordenadas en el fondo)
fig, ax = plt.subplots(figsize=(2.2, 2.0), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis('off')

ax.plot([1.5, 1.5, 8.5, 8.5], [8.5, 1.5, 1.5, 8.5], color="#B0C4DE", lw=2)

# Nivel de líquido
ax.plot([1.7, 8.3], [5.5, 5.5], color="#2980B9", lw=1.5, ls="--")

coords_liq = [
    (2.4, 2.2), (3.6, 2.0), (5.0, 2.3), (6.4, 2.1), (7.6, 2.4),
    (2.8, 3.2), (4.2, 3.4), (5.8, 3.1), (7.1, 3.3),
    (3.2, 4.3), (4.8, 4.5), (6.2, 4.2), (7.5, 4.4)
]
for x, y in coords_liq:
    p = patches.Circle((x, y), 0.48, ec="#173F73", fc="#5DADE2", lw=1.5)
    ax.add_patch(p)

ax.text(5.0, 0.4, "Estado Líquido", fontsize=9, fontweight='bold', color='#173F73', ha='center')
plt.savefig(os.path.join(out_dir, "estado_liquido.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

# 3. GASEOSO (Partículas muy separadas, ocupando todo el volumen)
fig, ax = plt.subplots(figsize=(2.2, 2.0), dpi=300)
ax.set_aspect('equal')
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis('off')

# Recipiente cerrado
ax.plot([1.5, 1.5, 8.5, 8.5, 1.5], [8.5, 1.5, 1.5, 8.5, 8.5], color="#B0C4DE", lw=2)

coords_gas = [
    (2.5, 7.2, 0.4, 0.3), (7.0, 7.5, -0.4, -0.2), (4.8, 5.5, 0.5, -0.4),
    (2.8, 3.5, -0.3, 0.5), (7.2, 3.2, 0.3, 0.4), (4.5, 2.5, -0.4, -0.3)
]
for x, y, dx, dy in coords_gas:
    p = patches.Circle((x, y), 0.48, ec="#173F73", fc="#AED6F1", lw=1.5)
    ax.add_patch(p)
    # Flechitas de movimiento
    ax.arrow(x, y, dx, dy, head_width=0.25, head_length=0.25, fc="#E74C3C", ec="#E74C3C", lw=1.2)

ax.text(5.0, 0.4, "Estado Gaseoso", fontsize=9, fontweight='bold', color='#173F73', ha='center')
plt.savefig(os.path.join(out_dir, "estado_gaseoso.png"), bbox_inches='tight', transparent=True, pad_inches=0.05)
plt.close()

print("Particle diagrams generated successfully!")
