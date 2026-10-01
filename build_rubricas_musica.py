"""
Generador de Rúbricas de Evaluación Final de Música (1° Básico A y 2° Básico A)
Colegio Castelgandolfo - Docencia Institucional
Año Escolar 2026
"""

import os
import sys
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

DOWNLOADS_ROOT = r"C:\Users\Jack\Downloads\Evaluaciones Finales Profesora Margarita"

from generator_core import (
    create_base_doc, add_header, add_title_banner, add_section_header,
    set_cell_shading, set_cell_borders, set_cell_margins, COLOR_NAVY_HEX,
    COLOR_ICE_HEX, COLOR_BOX_BG_HEX, COLOR_BORDER_HEX, COLOR_NAVY_RGB,
    COLOR_DARK_RGB, COLOR_WHITE_RGB
)

def build_rubrica_doc(curso_str, curso_folder_name, oa_desc):
    doc = create_base_doc()

    # Encabezado institucional con logo
    add_header(doc, school="COLEGIO CASTELGANDOLFO", subject=f"MÚSICA — {curso_str}")
    add_title_banner(doc, "RÚBRICA DE EVALUACIÓN FINAL: PERCUSIÓN METALÓFONO O SONAJAS")

    # Cuadro de datos del estudiante
    t_info = doc.add_table(rows=2, cols=2)
    t_info.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_info.autofit = False
    w_left, w_right = Pt(255), Pt(256.2)

    filas_info = [
        ("Nombre del/la Estudiante: ____________________________________", f"Curso: {curso_str}"),
        ("Fecha de Aplicación: ______ de _________________ de 2026", "Puntaje Total: 25 puntos   |   Puntaje Obtenido: _____ pts")
    ]

    for r_idx, (d1, d2) in enumerate(filas_info):
        c1, c2 = t_info.cell(r_idx, 0), t_info.cell(r_idx, 1)
        c1.width, c2.width = w_left, w_right
        for c, text in [(c1, d1), (c2, d2)]:
            set_cell_margins(c, top=55, bottom=55, left=80, right=80)
            set_cell_borders(c, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
            if r_idx == 0:
                set_cell_shading(c, COLOR_ICE_HEX)
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            r.font.name = "Arial"
            r.font.size = Pt(8.5)
            if "Puntaje" in text or "Curso" in text:
                r.bold = True

    # Cuadro de descripción de la actividad práctica
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    t_inst = doc.add_table(rows=1, cols=1)
    t_inst.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_inst.autofit = False
    c_inst = t_inst.cell(0, 0)
    c_inst.width = Pt(511.2)
    set_cell_shading(c_inst, COLOR_BOX_BG_HEX)
    set_cell_margins(c_inst, top=50, bottom=50, left=80, right=80)
    set_cell_borders(c_inst, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                             left={'val': 'single', 'sz': '12', 'color': COLOR_NAVY_HEX},
                             right={'color': COLOR_BORDER_HEX})
    p_ins = c_inst.paragraphs[0]
    p_ins.paragraph_format.space_before = Pt(0)
    p_ins.paragraph_format.space_after = Pt(0)

    r_oa = p_ins.add_run(f"Objetivo de Aprendizaje: {oa_desc}\n")
    r_oa.bold = True
    r_oa.font.name = "Arial"
    r_oa.font.size = Pt(8.5)
    r_oa.font.color.rgb = COLOR_NAVY_RGB

    r_ins = p_ins.add_run(
        "Modalidad: Evaluación práctica de interpretación musical. Cada estudiante demuestra sus habilidades "
        "en canto coral al unísono y ejecución rítmica/melódica de la canción 'Estrellita', utilizando el instrumento "
        "a su elección (sonaja o metalófono) según lo trabajado durante las clases."
    )
    r_ins.font.name = "Arial"
    r_ins.font.size = Pt(8)

    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # SECCIÓN RÚBRICA
    add_section_header(doc, "TABLA DE RÚBRICA Y CRITERIOS DE LOGRO (5 puntos c/u — Total: 25 puntos)")

    # 5 criterios solicitados:
    criterios = [
        ("1. Cantar canción Estrellita",
         "No participa del canto o no recuerda la melodía ni la letra de la canción.",
         "Canta frases aisladas con dificultad para sostener el ritmo y la melodía.",
         "Canta la mayor parte de la canción con adecuada afinación y dicción.",
         "Canta al unísono con afinación precisa, dicción clara y seguridad constante."),

        ("2. Percuten en forma individual la canción Estrellita",
         "Presenta severa descoordinación y no logra mantener el pulso individualmente.",
         "Percute en forma individual pero pierde el pulso o tempo con frecuencia.",
         "Percute individualmente con pulso estable en la mayor parte de la obra.",
         "Percute en forma individual con pulso rítmico exacto, seguro y constante."),

        ("3. Percuten en forma grupal la canción Estrellita",
         "No se acopla a la ejecución del grupo curso y toca fuera de tiempo.",
         "Intenta coordinarse con el grupo pero se desfasa constantemente.",
         "Percute de manera grupal manteniendo la coordinación con sus compañeros.",
         "Percute en forma grupal en perfecta sincronía, respetando entradas y cortes."),

        ("4. Percuten con metalófono o sonajas la canción",
         "Dificultad evidente en el agarre y percusión del instrumento elegido.",
         "Utiliza el instrumento con errores técnicos en baquetas o movimientos de sonaja.",
         "Ejecuta el instrumento elegido (sonaja o metalófono) con técnica adecuada.",
         "Ejecuta con técnica óptima, sonoridad limpia y dominio total del instrumento."),

        ("5. Mantiene el orden y conducta adecuada durante la evaluación",
         "Muestra desinterés, interrumpe a compañeros o descuida el instrumento.",
         "Requiere constantes llamados de atención para guardar silencio y concentrarse.",
         "Mantiene una conducta adecuada y respetuosa durante la evaluación.",
         "Demuestra excelente concentración, cuidado del instrumento y respeto ejemplar.")
    ]

    t_rub = doc.add_table(rows=len(criterios) + 2, cols=6)
    t_rub.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_rub.autofit = False

    col_widths = [Pt(135), Pt(75), Pt(75), Pt(75), Pt(85), Pt(66.2)]
    headers_rub = [
        "Criterio / Indicador",
        "1 pto\nInsuficiente",
        "2 pts\nElemental",
        "3 pts\nBueno",
        "5 pts\nExcelente",
        "Puntaje\nObtenido"
    ]

    # Encabezado
    for i, h in enumerate(headers_rub):
        c = t_rub.cell(0, i)
        c.width = col_widths[i]
        set_cell_shading(c, COLOR_NAVY_HEX)
        set_cell_margins(c, top=50, bottom=50, left=45, right=45)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        r.font.name = "Arial"
        r.font.size = Pt(8)
        r.bold = True
        r.font.color.rgb = COLOR_WHITE_RGB

    # Filas de criterios
    for row_idx, (crit_tit, d1, d2, d3, d5) in enumerate(criterios, 1):
        bg_col = "FFFFFF" if row_idx % 2 != 0 else COLOR_BOX_BG_HEX
        fila_datos = [crit_tit, d1, d2, d3, d5, ""]

        for col_idx, text_c in enumerate(fila_datos):
            cell = t_rub.cell(row_idx, col_idx)
            cell.width = col_widths[col_idx]
            set_cell_shading(cell, bg_col)
            set_cell_margins(cell, top=45, bottom=45, left=50, right=50)
            set_cell_borders(cell, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                   left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)

            if col_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(text_c)
                r.bold = True
                r.font.name = "Arial"
                r.font.size = Pt(8)
                r.font.color.rgb = COLOR_NAVY_RGB
            elif col_idx in [1, 2, 3, 4]:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r = p.add_run(text_c)
                r.font.name = "Arial"
                r.font.size = Pt(7.5)
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                # Casilla para puntaje obtenido
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(8)

    # Fila de Total
    row_tot = len(criterios) + 1
    c_tot_label = t_rub.cell(row_tot, 0)
    c_tot_label.width = Pt(135)
    set_cell_shading(c_tot_label, COLOR_NAVY_HEX)
    set_cell_margins(c_tot_label, top=50, bottom=50, left=50, right=50)
    p_tl = c_tot_label.paragraphs[0]
    p_tl.paragraph_format.space_before = Pt(0)
    p_tl.paragraph_format.space_after = Pt(0)
    r_tl = p_tl.add_run("PUNTAJE FINAL TOTAL")
    r_tl.bold = True
    r_tl.font.name = "Arial"
    r_tl.font.size = Pt(8.5)
    r_tl.font.color.rgb = COLOR_WHITE_RGB

    # Celdas intermedias vacías / unificadas
    for mid_i in range(1, 5):
        c_mid = t_rub.cell(row_tot, mid_i)
        c_mid.width = col_widths[mid_i]
        set_cell_shading(c_mid, COLOR_ICE_HEX)
        set_cell_margins(c_mid, top=50, bottom=50, left=40, right=40)
        set_cell_borders(c_mid, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
        p_m = c_mid.paragraphs[0]
        p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if mid_i == 4:
            r_sc = p_m.add_run("Máx: 25 pts")
            r_sc.bold = True
            r_sc.font.name = "Arial"
            r_sc.font.size = Pt(8)
            r_sc.font.color.rgb = COLOR_NAVY_RGB

    c_tot_val = t_rub.cell(row_tot, 5)
    c_tot_val.width = Pt(66.2)
    set_cell_shading(c_tot_val, COLOR_ICE_HEX)
    set_cell_margins(c_tot_val, top=50, bottom=50, left=40, right=40)
    set_cell_borders(c_tot_val, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
    p_tv = c_tot_val.paragraphs[0]
    p_tv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_tv = p_tv.add_run("/ 25")
    r_tv.bold = True
    r_tv.font.name = "Arial"
    r_tv.font.size = Pt(9)
    r_tv.font.color.rgb = COLOR_NAVY_RGB

    # Observaciones y retroalimentación docente
    p_sp_obs = doc.add_paragraph()
    p_sp_obs.paragraph_format.space_before = Pt(3)
    p_sp_obs.paragraph_format.space_after = Pt(2)

    t_obs = doc.add_table(rows=1, cols=1)
    t_obs.alignment = WD_TABLE_ALIGNMENT.CENTER
    t_obs.autofit = False
    c_obs = t_obs.cell(0, 0)
    c_obs.width = Pt(511.2)
    set_cell_shading(c_obs, COLOR_BOX_BG_HEX)
    set_cell_margins(c_obs, top=50, bottom=50, left=70, right=70)
    set_cell_borders(c_obs, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                            left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
    p_ob = c_obs.paragraphs[0]
    p_ob.paragraph_format.space_before = Pt(0)
    p_ob.paragraph_format.space_after = Pt(0)
    r_ob = p_ob.add_run("Observaciones / Retroalimentación Pedagógica del Desempeño:\n")
    r_ob.bold = True
    r_ob.font.name = "Arial"
    r_ob.font.size = Pt(8.5)
    r_ob.font.color.rgb = COLOR_NAVY_RGB

    r_lin = p_ob.add_run(
        "_____________________________________________________________________________________________________\n"
        "_____________________________________________________________________________________________________\n"
        "_____________________________________________________________________________________________________"
    )
    r_lin.font.name = "Arial"
    r_lin.font.size = Pt(8)
    r_lin.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)

    # Firma docente
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.paragraph_format.space_before = Pt(6)
    p_sig.paragraph_format.space_after = Pt(0)

    sig_runs = [
        ("Docencia Institucional\n", True, 9),
        ("Docente de Música | Colegio Castelgandolfo\n", False, 8),
        ("Firma del Docente Evaluador: ____________________________", False, 8)
    ]
    for text, bold, sz in sig_runs:
        r = p_sig.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(sz)
        r.bold = bold
        r.font.color.rgb = COLOR_NAVY_RGB if bold else COLOR_DARK_RGB

    return doc

def generar_rubricas_musica():
    print("\n=======================================================")
    print(" GENERANDO RÚBRICAS DE MÚSICA 1° BÁSICO A Y 2° BÁSICO A")
    print("=======================================================")

    # 1° Básico A
    doc1 = build_rubrica_doc(
        curso_str="1° BÁSICO A",
        curso_folder_name=r"1° Básico (1°A)",
        oa_desc="Cantar al unísono y tocar instrumentos de percusión convencionales o no convencionales."
    )
    f1_name = "Rubrica_Evaluacion_Final_Musica_1Basico.docx"
    p1_local = os.path.join(OUTPUT_DIR, f1_name)
    doc1.save(p1_local)
    print(f"  [OK] Rúbrica 1° Básico guardada en local: {p1_local}")

    dst1_dir = os.path.join(DOWNLOADS_ROOT, "02 - Música", "1° Básico (1°A)")
    os.makedirs(dst1_dir, exist_ok=True)
    p1_dst = os.path.join(dst1_dir, f1_name)
    shutil.copy2(p1_local, p1_dst)
    print(f"  [OK] Copiado a Descargas: {p1_dst}")

    # 2° Básico A
    doc2 = build_rubrica_doc(
        curso_str="2° BÁSICO A",
        curso_folder_name=r"2° Básico (2°A)",
        oa_desc="Cantar al unísono y tocar instrumentos de percusión convencionales o no convencionales."
    )
    f2_name = "Rubrica_Evaluacion_Final_Musica_2Basico.docx"
    p2_local = os.path.join(OUTPUT_DIR, f2_name)
    doc2.save(p2_local)
    print(f"  [OK] Rúbrica 2° Básico guardada en local: {p2_local}")

    dst2_dir = os.path.join(DOWNLOADS_ROOT, "02 - Música", "2° Básico (2°A)")
    os.makedirs(dst2_dir, exist_ok=True)
    p2_dst = os.path.join(dst2_dir, f2_name)
    shutil.copy2(p2_local, p2_dst)
    print(f"  [OK] Copiado a Descargas: {p2_dst}")

    print("\n[COMPLETADO] Rúbricas de música de 1°A y 2°A creadas y desplegadas exitosamente.")

if __name__ == "__main__":
    generar_rubricas_musica()
