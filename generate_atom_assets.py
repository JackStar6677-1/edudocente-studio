"""
Generador de Recursos Gráficos del Átomo (Esqueleto Alumno y Modelo Resuelto Docente)
Colegio Luis Pasteur Anexo - Ciencias Naturales 8° Básico
"""

import os
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

base_dir = os.path.dirname(os.path.abspath(__file__))
out_dir = os.path.join(base_dir, "assets", "atom_model")
os.makedirs(out_dir, exist_ok=True)

# ------------------------------------------------------------------------------
# 1. ESQUELETO PARA EL ESTUDIANTE (Líneas guía tenues para completar y colorear)
# ------------------------------------------------------------------------------
def generate_student_atom_skeleton():
    fig, ax = plt.subplots(figsize=(6.5, 4.8), dpi=300)
    ax.set_aspect('equal')
    ax.set_xlim(-6.5, 6.5)
    ax.set_ylim(-4.8, 4.8)
    ax.axis('off')

    # Órbitas elípticas (guías punteadas)
    e1 = patches.Ellipse((0, 0), width=9.5, height=3.4, angle=25,
                         fill=False, edgecolor='#B0C4DE', linestyle='--', linewidth=1.8)
    e2 = patches.Ellipse((0, 0), width=9.5, height=3.4, angle=-25,
                         fill=False, edgecolor='#B0C4DE', linestyle='--', linewidth=1.8)
    e3 = patches.Ellipse((0, 0), width=9.5, height=3.4, angle=90,
                         fill=False, edgecolor='#B0C4DE', linestyle='--', linewidth=1.8)
    ax.add_patch(e1)
    ax.add_patch(e2)
    ax.add_patch(e3)

    # Círculo del Núcleo central tenue
    nuc_circle = patches.Circle((0, 0), radius=1.35, fill=True,
                                facecolor='#F6FAFE', edgecolor='#173F73',
                                linestyle='-.', linewidth=2.0)
    ax.add_patch(nuc_circle)

    # Pequeños círculos guía punteados en el núcleo para sugerir dibujar protones y neutrones
    pos_nuc = [(-0.5, 0.4), (0.5, 0.4), (0.0, -0.5), (-0.5, -0.4), (0.5, -0.4)]
    for x, y in pos_nuc:
        p_c = patches.Circle((x, y), radius=0.32, fill=False,
                             edgecolor='#CBD5E1', linestyle=':', linewidth=1.2)
        ax.add_patch(p_c)

    # Círculos guía tenue en las órbitas para sugerir electrones
    pos_elec = [(-3.8, -1.8), (3.8, 1.8), (-3.8, 1.8), (3.8, -1.8), (0, 3.8), (0, -3.8)]
    for x, y in pos_elec:
        el_c = patches.Circle((x, y), radius=0.28, fill=False,
                              edgecolor='#CBD5E1', linestyle=':', linewidth=1.2)
        ax.add_patch(el_c)

    # Líneas y cuadros guía para rotular partes (Callout boxes en blanco para que el estudiante responda)
    # 1. Zona exterior / corteza
    ax.annotate("1. ___________________",
                xy=(-2.5, 2.5), xytext=(-5.8, 3.5),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.5),
                fontsize=9.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FFFFFF", ec="#B0C4DE", lw=1.3))

    # 2. Partícula en órbita
    ax.annotate("2. ___________________",
                xy=(3.8, 1.8), xytext=(4.2, 3.2),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.5),
                fontsize=9.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FFFFFF", ec="#B0C4DE", lw=1.3))

    # 3. Región central
    ax.annotate("3. ___________________",
                xy=(-0.8, -0.4), xytext=(-5.8, -2.5),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.5),
                fontsize=9.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FFFFFF", ec="#B0C4DE", lw=1.3))

    # 4. Partícula nuclear
    ax.annotate("4. ___________________",
                xy=(0.5, 0.4), xytext=(4.0, -1.8),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.5),
                fontsize=9.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FFFFFF", ec="#B0C4DE", lw=1.3))

    # 5. Otra partícula nuclear
    ax.annotate("5. ___________________",
                xy=(0.0, -0.5), xytext=(4.0, -3.6),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.5),
                fontsize=9.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FFFFFF", ec="#B0C4DE", lw=1.3))

    out_file = os.path.join(out_dir, "atom_skeleton_student.png")
    plt.savefig(out_file, bbox_inches='tight', transparent=False, facecolor='white', pad_inches=0.08)
    plt.close()
    print(f"Esqueleto del átomo para alumno generado: {out_file}")

# ------------------------------------------------------------------------------
# 2. MODELO RESUELTO PARA LA PAUTA DOCENTE (A todo color y rotulado perfecto)
# ------------------------------------------------------------------------------
def generate_teacher_atom_solved():
    fig, ax = plt.subplots(figsize=(6.5, 4.8), dpi=300)
    ax.set_aspect('equal')
    ax.set_xlim(-6.5, 6.5)
    ax.set_ylim(-4.8, 4.8)
    ax.axis('off')

    # Órbitas elípticas
    e1 = patches.Ellipse((0, 0), width=9.5, height=3.4, angle=25,
                         fill=False, edgecolor='#6BA4D9', linestyle='-', linewidth=2.0)
    e2 = patches.Ellipse((0, 0), width=9.5, height=3.4, angle=-25,
                         fill=False, edgecolor='#6BA4D9', linestyle='-', linewidth=2.0)
    e3 = patches.Ellipse((0, 0), width=9.5, height=3.4, angle=90,
                         fill=False, edgecolor='#6BA4D9', linestyle='-', linewidth=2.0)
    ax.add_patch(e1)
    ax.add_patch(e2)
    ax.add_patch(e3)

    # Núcleo delimitado
    nuc_circle = patches.Circle((0, 0), radius=1.35, fill=True,
                                facecolor='#EAF3FB', edgecolor='#173F73',
                                linestyle='--', linewidth=2.2)
    ax.add_patch(nuc_circle)

    # Partículas en el núcleo: Protones (Rojos / Coral con +) y Neutrones (Amarillos / Gris con 0)
    protones = [(-0.45, 0.45), (0.45, -0.35), (0.0, -0.55)]
    neutrones = [(0.45, 0.45), (-0.45, -0.35), (0.1, 0.1)]

    # Protones
    for x, y in protones:
        p = patches.Circle((x, y), radius=0.34, fill=True, facecolor='#E74C3C', edgecolor='#922B21', lw=1.5)
        ax.add_patch(p)
        ax.text(x, y, "+", fontsize=11, fontweight='bold', color='white', ha='center', va='center')

    # Neutrones
    for x, y in neutrones:
        n = patches.Circle((x, y), radius=0.34, fill=True, facecolor='#F39C12', edgecolor='#B9770E', lw=1.5)
        ax.add_patch(n)
        ax.text(x, y, "n", fontsize=9, fontweight='bold', color='white', ha='center', va='center')

    # Electrones en órbitas (Azules con signo -)
    pos_elec = [(-3.8, -1.8), (3.8, 1.8), (-3.8, 1.8), (3.8, -1.8), (0, 3.8), (0, -3.8)]
    for x, y in pos_elec:
        el = patches.Circle((x, y), radius=0.28, fill=True, facecolor='#2980B9', edgecolor='#173F73', lw=1.5)
        ax.add_patch(el)
        ax.text(x, y, "–", fontsize=11, fontweight='bold', color='white', ha='center', va='center')

    # Rotulaciones completas
    # 1. Corteza
    ax.annotate("Corteza o Electrosfera\n(Zona exterior donde giran los electrones)",
                xy=(-2.5, 2.5), xytext=(-6.2, 3.6),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.8),
                fontsize=8.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#F6FAFE", ec="#173F73", lw=1.2))

    # 2. Electrón
    ax.annotate("Electrón (e–)\n(Carga negativa, masa despreciable)",
                xy=(3.8, 1.8), xytext=(3.6, 3.3),
                arrowprops=dict(arrowstyle="->", color="#2980B9", lw=1.8),
                fontsize=8.5, fontweight='bold', color="#1A5276",
                bbox=dict(boxstyle="round,pad=0.35", fc="#EAF2F8", ec="#2980B9", lw=1.2))

    # 3. Núcleo
    ax.annotate("Núcleo Atómico\n(Zona central densa y positiva)",
                xy=(-0.8, -0.4), xytext=(-6.2, -2.5),
                arrowprops=dict(arrowstyle="->", color="#173F73", lw=1.8),
                fontsize=8.5, fontweight='bold', color="#173F73",
                bbox=dict(boxstyle="round,pad=0.35", fc="#F6FAFE", ec="#173F73", lw=1.2))

    # 4. Protón
    ax.annotate("Protón (p+)\n(Carga eléctrica positiva, en el núcleo)",
                xy=(-0.45, 0.45), xytext=(3.5, -1.8),
                arrowprops=dict(arrowstyle="->", color="#E74C3C", lw=1.8),
                fontsize=8.5, fontweight='bold', color="#922B21",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FDEDEC", ec="#E74C3C", lw=1.2))

    # 5. Neutrón
    ax.annotate("Neutrón (n0)\n(Sin carga / neutro, en el núcleo)",
                xy=(0.45, 0.45), xytext=(3.5, -3.6),
                arrowprops=dict(arrowstyle="->", color="#F39C12", lw=1.8),
                fontsize=8.5, fontweight='bold', color="#7E5109",
                bbox=dict(boxstyle="round,pad=0.35", fc="#FEF9E7", ec="#F39C12", lw=1.2))

    out_file = os.path.join(out_dir, "atom_solved_teacher.png")
    plt.savefig(out_file, bbox_inches='tight', transparent=False, facecolor='white', pad_inches=0.08)
    plt.close()
    print(f"Modelo resuelto del átomo generado: {out_file}")

if __name__ == "__main__":
    generate_student_atom_skeleton()
    generate_teacher_atom_solved()
