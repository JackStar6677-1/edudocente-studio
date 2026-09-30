"""
EduDocente-Engine: Motor Declarativo de Evaluaciones y Temarios Institucionales
Colegio Luis Pasteur Anexo - Framework Estandarizado para Evaluaciones Escolares

Permite generar a partir de una estructura JSON/Dict:
  1. Evaluación del Estudiante (.docx)
  2. Pauta Oficial de Corrección con Solucionario y Justificaciones (.docx)
  3. Temario e Informativo Oficial para Apoderados (.docx)
"""

import os
import sys
import json
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

from generator_core import (
    create_base_doc, add_header, add_title_banner, add_section_header,
    set_cell_shading, set_cell_borders, set_cell_margins,
    COLOR_NAVY_HEX, COLOR_ICE_HEX, COLOR_BOX_BG_HEX, COLOR_BORDER_HEX,
    COLOR_CORRECT_HEX, COLOR_NAVY_RGB, COLOR_CORRECT_RGB, COLOR_DARK_RGB,
    COLOR_WHITE_RGB, LOGO_PATH, safe_save
)

class DocenteEngine:
    def __init__(self, config):
        """
        Inicializa el motor con una configuración de evaluación (Dict o JSON path)
        """
        if isinstance(config, str):
            with open(config, 'r', encoding='utf-8') as f:
                self.cfg = json.load(f)
        else:
            self.cfg = config

        self.school = self.cfg.get("colegio", "COLEGIO CASTELGANDOLFO")
        self.subject = self.cfg.get("asignatura", "CIENCIAS NATURALES").upper()
        self.course = self.cfg.get("curso", "").strip()
        self.title = self.cfg.get("titulo", "EVALUACIÓN FINAL").upper()
        self.oa = self.cfg.get("oa", "")
        self.contents = self.cfg.get("contenidos", "")
        self.total_pts = self.cfg.get("puntaje_total", 25)
        self.teacher = self.cfg.get("docente", "Docente de Asignatura")
        self.teacher_email = self.cfg.get("email_docente", "docente@colegiocastelgandolfo.cl")

    # --------------------------------------------------------------------------
    # COMPONENTE: Cuadro de Datos y OA
    # --------------------------------------------------------------------------
    def _add_info_box(self, doc, is_pauta=False):
        t_info = doc.add_table(rows=2, cols=2)
        t_info.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_info.autofit = False
        w_l, w_r = Pt(255), Pt(256.2)

        filas = [
            ("Nombre del/la Estudiante: " + ("PAUTA OFICIAL DOCENTE" if is_pauta else "____________________________________"),
             f"Curso: {self.course}"),
            ("Fecha: ______ de _________________ de 2026",
             f"Puntaje Total: {self.total_pts} puntos   |   Puntaje Obtenido: " + (f"{self.total_pts} pts" if is_pauta else "_____ pts"))
        ]

        for r_idx, (d1, d2) in enumerate(filas):
            c1, c2 = t_info.cell(r_idx, 0), t_info.cell(r_idx, 1)
            c1.width, c2.width = w_l, w_r
            for c, text in [(c1, d1), (c2, d2)]:
                set_cell_margins(c, top=55, bottom=55, left=80, right=80)
                set_cell_borders(c, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                    left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
                if r_idx == 0:
                    set_cell_shading(c, "EAF9E6" if is_pauta else COLOR_ICE_HEX)
                p = c.paragraphs[0]
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(text)
                r.font.name = "Arial"
                r.font.size = Pt(8.5)
                if is_pauta and "PAUTA" in text:
                    r.bold = True
                    r.font.color.rgb = COLOR_CORRECT_RGB
                elif "Puntaje" in text or "Curso" in text:
                    r.bold = True

        doc.add_paragraph().paragraph_format.space_after = Pt(2)

        # Instrucciones y OA
        t_inst = doc.add_table(rows=1, cols=1)
        t_inst.alignment = WD_TABLE_ALIGNMENT.CENTER
        t_inst.autofit = False
        c_ins = t_inst.cell(0, 0)
        c_ins.width = Pt(511.2)
        set_cell_shading(c_ins, COLOR_BOX_BG_HEX)
        set_cell_margins(c_ins, top=50, bottom=50, left=80, right=80)
        set_cell_borders(c_ins, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                left={'val': 'single', 'sz': '12', 'color': COLOR_NAVY_HEX},
                                right={'color': COLOR_BORDER_HEX})
        p_ins = c_ins.paragraphs[0]
        p_ins.paragraph_format.space_before = Pt(0)
        p_ins.paragraph_format.space_after = Pt(0)

        if self.oa:
            r_oa = p_ins.add_run(f"Objetivo de Aprendizaje: {self.oa}\n")
            r_oa.bold = True
            r_oa.font.name = "Arial"
            r_oa.font.size = Pt(8.5)
            r_oa.font.color.rgb = COLOR_NAVY_RGB

        if self.contents:
            r_co = p_ins.add_run(f"Contenidos Evaluados: {self.contents}\n")
            r_co.bold = True
            r_co.font.name = "Arial"
            r_co.font.size = Pt(8)

        inst_txt = self.cfg.get("instrucciones",
            "Instrucciones: Lee atentamente cada ítem antes de contestar. Responde con claridad y letra legible. "
            "En caso de dudas, levanta tu mano en silencio para que la profesora se acerque a orientarte."
        )
        r_txt = p_ins.add_run(inst_txt)
        r_txt.font.name = "Arial"
        r_txt.font.size = Pt(8)

        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # --------------------------------------------------------------------------
    # ÍTEM I: SELECCIÓN MÚLTIPLE
    # --------------------------------------------------------------------------
    def _render_item1(self, doc, is_pauta=False):
        items = self.cfg.get("item1_seleccion_multiple", [])
        if not items:
            return

        pts_c_u = self.cfg.get("item1_pts_cada_una", 2)
        total_pts = len(items) * pts_c_u
        header_text = f"ÍTEM I: SELECCIÓN MÚLTIPLE ({pts_c_u} puntos c/u — Total: {total_pts} puntos)"
        if is_pauta:
            header_text = f"SOLUCIONARIO {header_text}"
        add_section_header(doc, header_text)

        if not is_pauta:
            p_inst = doc.add_paragraph()
            p_inst.paragraph_format.space_before = Pt(1)
            p_inst.paragraph_format.space_after = Pt(4)
            p_inst.add_run("Lee comprensivamente cada pregunta y marca con una ").font.name = "Arial"
            rx = p_inst.add_run("X")
            rx.bold = True
            rx.font.name = "Arial"
            p_inst.add_run(" la alternativa que consideres correcta:").font.name = "Arial"

            for item in items:
                q_text = item["pregunta"]
                opts = item["alternativas"] # List of (letra, texto) or dicts

                t_q = doc.add_table(rows=1, cols=1)
                t_q.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_q.autofit = False
                c_q = t_q.cell(0, 0)
                c_q.width = Pt(511.2)
                set_cell_shading(c_q, COLOR_BOX_BG_HEX)
                set_cell_margins(c_q, top=45, bottom=45, left=60, right=60)
                set_cell_borders(c_q, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                      left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
                pq = c_q.paragraphs[0]
                pq.paragraph_format.space_before = Pt(0)
                pq.paragraph_format.space_after = Pt(0)
                rq = pq.add_run(q_text)
                rq.bold = True
                rq.font.name = "Arial"
                rq.font.size = Pt(8.5)
                rq.font.color.rgb = COLOR_NAVY_RGB

                t_opts = doc.add_table(rows=len(opts), cols=1)
                t_opts.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_opts.autofit = False

                for opt_idx, opt in enumerate(opts):
                    letra = opt[0] if isinstance(opt, (list, tuple)) else opt.get("letra", "")
                    texto_opt = opt[1] if isinstance(opt, (list, tuple)) else opt.get("texto", "")

                    c_opt = t_opts.cell(opt_idx, 0)
                    c_opt.width = Pt(511.2)
                    set_cell_margins(c_opt, top=30, bottom=30, left=50, right=50)
                    set_cell_borders(c_opt, top={'val': 'none'}, bottom={'val': 'none'}, left={'val': 'none'}, right={'val': 'none'})
                    p_opt = c_opt.paragraphs[0]
                    p_opt.paragraph_format.space_before = Pt(0)
                    p_opt.paragraph_format.space_after = Pt(0)

                    r_let = p_opt.add_run(f"{letra})  ")
                    r_let.bold = True
                    r_let.font.name = "Arial"
                    r_let.font.size = Pt(8.5)
                    r_let.font.color.rgb = COLOR_NAVY_RGB

                    r_txt = p_opt.add_run(texto_opt)
                    r_txt.font.name = "Arial"
                    r_txt.font.size = Pt(8.5)

                doc.add_paragraph().paragraph_format.space_after = Pt(2)
        else:
            # Solucionario docente
            for item in items:
                q_text = item["pregunta"]
                correcta = item.get("correcta", "")
                justif = item.get("justificacion", "Respuesta oficial según contenidos abordados en clases y texto escolar.")

                t_s = doc.add_table(rows=1, cols=1)
                t_s.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_s.autofit = False
                cs = t_s.cell(0, 0)
                cs.width = Pt(511.2)
                set_cell_shading(cs, "F6FAFE")
                set_cell_margins(cs, top=45, bottom=45, left=60, right=60)
                set_cell_borders(cs, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                       left={'val': 'single', 'sz': '12', 'color': COLOR_CORRECT_HEX},
                                       right={'color': COLOR_BORDER_HEX})
                ps = cs.paragraphs[0]
                ps.paragraph_format.space_before = Pt(0)
                ps.paragraph_format.space_after = Pt(0)

                rp = ps.add_run(f"{q_text}\n")
                rp.bold = True
                rp.font.name = "Arial"
                rp.font.size = Pt(8.5)
                rp.font.color.rgb = COLOR_NAVY_RGB

                rr = ps.add_run(f"Respuesta Correcta: {correcta}\n")
                rr.bold = True
                rr.font.name = "Arial"
                rr.font.size = Pt(8.5)
                rr.font.color.rgb = COLOR_CORRECT_RGB

                rj = ps.add_run(f"Justificación Pedagógica: {justif}")
                rj.font.name = "Arial"
                rj.font.size = Pt(8)

                doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # --------------------------------------------------------------------------
    # ÍTEM II: VERDADERO O FALSO
    # --------------------------------------------------------------------------
    def _render_item2(self, doc, is_pauta=False):
        items = self.cfg.get("item2_verdadero_falso", [])
        if not items:
            return

        pts_c_u = self.cfg.get("item2_pts_cada_una", 1)
        total_pts = len(items) * pts_c_u
        header_text = f"ÍTEM II: VERDADERO O FALSO ({pts_c_u} punto{'s' if pts_c_u > 1 else ''} c/u — Total: {total_pts} puntos)"
        if is_pauta:
            header_text = f"SOLUCIONARIO {header_text}"
        add_section_header(doc, header_text)

        if not is_pauta:
            p_inst = doc.add_paragraph()
            p_inst.paragraph_format.space_before = Pt(1)
            p_inst.paragraph_format.space_after = Pt(4)
            p_inst.add_run("De las siguientes oraciones, escribe en el paréntesis una ").font.name = "Arial"
            rv = p_inst.add_run("V")
            rv.bold = True
            rv.font.name = "Arial"
            p_inst.add_run(" si es verdadera o una ").font.name = "Arial"
            rf = p_inst.add_run("F")
            rf.bold = True
            rf.font.name = "Arial"
            p_inst.add_run(" si es falsa:").font.name = "Arial"

            t_vf = doc.add_table(rows=len(items), cols=3)
            t_vf.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_vf.autofit = False

            for idx, item in enumerate(items, 1):
                texto = item if isinstance(item, str) else item.get("oracion", "")
                c_num, c_box, c_txt = t_vf.cell(idx - 1, 0), t_vf.cell(idx - 1, 1), t_vf.cell(idx - 1, 2)
                c_num.width, c_box.width, c_txt.width = Pt(22), Pt(40), Pt(449.2)

                bg_col = "FFFFFF" if idx % 2 != 0 else COLOR_BOX_BG_HEX
                for c in (c_num, c_box, c_txt):
                    set_cell_shading(c, bg_col)
                    set_cell_margins(c, top=35, bottom=35, left=35, right=35)
                    set_cell_borders(c, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                        left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})

                pn = c_num.paragraphs[0]
                pn.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pn.paragraph_format.space_before = Pt(0)
                pn.paragraph_format.space_after = Pt(0)
                rn = pn.add_run(f"{idx}.")
                rn.font.name = "Arial"
                rn.font.size = Pt(8.5)
                rn.bold = True
                rn.font.color.rgb = COLOR_NAVY_RGB

                pb = c_box.paragraphs[0]
                pb.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pb.paragraph_format.space_before = Pt(0)
                pb.paragraph_format.space_after = Pt(0)
                rb = pb.add_run("(     )")
                rb.font.name = "Arial"
                rb.font.size = Pt(8.5)
                rb.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

                pt = c_txt.paragraphs[0]
                pt.paragraph_format.space_before = Pt(0)
                pt.paragraph_format.space_after = Pt(0)
                rt = pt.add_run(texto)
                rt.font.name = "Arial"
                rt.font.size = Pt(8.5)

            doc.add_paragraph().paragraph_format.space_after = Pt(3)
        else:
            # Solucionario docente V/F
            t_s = doc.add_table(rows=len(items) + 1, cols=4)
            t_s.alignment = WD_TABLE_ALIGNMENT.CENTER
            t_s.autofit = False
            w_cols = [Pt(25), Pt(45), Pt(230), Pt(211.2)]
            headers = ["N°", "Resp.", "Afirmación Evaluada", "Justificación Pedagógica"]

            for i, h in enumerate(headers):
                c = t_s.cell(0, i)
                c.width = w_cols[i]
                set_cell_shading(c, COLOR_NAVY_HEX)
                set_cell_margins(c, top=45, bottom=45, left=45, right=45)
                p = c.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                r = p.add_run(h)
                r.bold = True
                r.font.name = "Arial"
                r.font.size = Pt(8)
                r.font.color.rgb = COLOR_WHITE_RGB

            for row_idx, item in enumerate(items, 1):
                letra = item.get("resp", "V") if isinstance(item, dict) else "V"
                oracion = item.get("oracion", "") if isinstance(item, dict) else str(item)
                justif = item.get("justificacion", "Conforme a los contenidos curriculares oficiales.") if isinstance(item, dict) else ""

                bg = "FFFFFF" if row_idx % 2 != 0 else COLOR_BOX_BG_HEX
                for col_idx, text_val in enumerate([f"{row_idx}", f"[ {letra} ]", oracion, justif]):
                    cell = t_s.cell(row_idx, col_idx)
                    cell.width = w_cols[col_idx]
                    set_cell_shading(cell, bg)
                    set_cell_margins(cell, top=35, bottom=35, left=40, right=40)
                    set_cell_borders(cell, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                           left={'color': COLOR_BORDER_HEX}, right={'color': COLOR_BORDER_HEX})
                    p = cell.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    r = p.add_run(text_val)
                    r.font.name = "Arial"
                    r.font.size = Pt(8)
                    if col_idx == 0:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r.bold = True
                    elif col_idx == 1:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        r.bold = True
                        r.font.color.rgb = COLOR_CORRECT_RGB
                    elif col_idx == 3:
                        r.font.size = Pt(7.5)
                        r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

            doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # --------------------------------------------------------------------------
    # ÍTEM III: DIBUJO / ESQUEMA / APLICACIÓN PRÁCTICA
    # --------------------------------------------------------------------------
    def _render_item3(self, doc, is_pauta=False):
        cfg3 = self.cfg.get("item3_aplicacion", {})
        if not cfg3:
            return

        titulo3 = cfg3.get("titulo", "ÍTEM III: APLICACIÓN PRÁCTICA Y DIBUJO")
        pts = cfg3.get("puntaje", 5)
        instruccion = cfg3.get("instruccion", "")
        tipo = cfg3.get("tipo", "drawing_boxes") # drawing_boxes, diagram_image, table_symbols

        header_text = f"{titulo3} ({pts} puntos)"
        if is_pauta:
            header_text = f"SOLUCIONARIO {header_text}"
        add_section_header(doc, header_text)

        p_i3 = doc.add_paragraph()
        p_i3.paragraph_format.space_before = Pt(1)
        p_i3.paragraph_format.space_after = Pt(4)
        r_i3 = p_i3.add_run(instruccion)
        r_i3.font.name = "Arial"
        r_i3.font.size = Pt(8.5)

        if tipo == "drawing_boxes":
            cajas = cfg3.get("cajas", [])
            for c_info in cajas:
                t_box = doc.add_table(rows=2, cols=1)
                t_box.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_box.autofit = False

                c_draw = t_box.cell(0, 0)
                c_draw.width = Pt(511.2)
                set_cell_shading(c_draw, "FAFCFF" if not is_pauta else "F6FAFE")
                set_cell_margins(c_draw, top=40, bottom=40, left=70, right=70)
                set_cell_borders(c_draw, top={'val': 'dashed', 'sz': '6', 'color': '888888'},
                                         bottom={'val': 'dashed', 'sz': '6', 'color': '888888'},
                                         left={'val': 'dashed', 'sz': '6', 'color': '888888'},
                                         right={'val': 'dashed', 'sz': '6', 'color': '888888'})
                p_d = c_draw.paragraphs[0]
                p_d.paragraph_format.space_before = Pt(0)
                p_d.paragraph_format.space_after = Pt(10 if is_pauta else 45)

                nom = c_info.get("titulo", "")
                r_n = p_d.add_run(f"• {nom}\n")
                r_n.bold = True
                r_n.font.name = "Arial"
                r_n.font.size = Pt(8.5)
                r_n.font.color.rgb = COLOR_CORRECT_RGB if is_pauta else COLOR_NAVY_RGB

                if is_pauta:
                    pauta_exp = c_info.get("ejemplo_pauta", "")
                    r_exp = p_d.add_run(f"Respuesta esperada: {pauta_exp}\n")
                    r_exp.font.name = "Arial"
                    r_exp.font.size = Pt(8)
                else:
                    r_sub = p_d.add_run("(Dibuja aquí tu representación)")
                    r_sub.font.name = "Arial"
                    r_sub.font.size = Pt(7.5)
                    r_sub.italic = True
                    r_sub.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)

                # Celda 1: Línea de descripción
                c_desc = t_box.cell(1, 0)
                c_desc.width = Pt(511.2)
                set_cell_margins(c_desc, top=25, bottom=25, left=70, right=70)
                set_cell_borders(c_desc, top={'val': 'none'}, bottom={'val': 'none'}, left={'val': 'none'}, right={'val': 'none'})
                p_dc = c_desc.paragraphs[0]
                p_dc.paragraph_format.space_before = Pt(1)
                p_dc.paragraph_format.space_after = Pt(2)
                p_dc.add_run("Descripción / Explicación: ____________________________________________________________________________________").font.size = Pt(8)

                doc.add_paragraph().paragraph_format.space_after = Pt(2)

        elif tipo == "diagram_image":
            img_key = "imagen_pauta" if is_pauta else "imagen_alumno"
            img_path = cfg3.get(img_key, "")
            if img_path and os.path.exists(img_path):
                p_img = doc.add_paragraph()
                p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_img.paragraph_format.space_before = Pt(2)
                p_img.paragraph_format.space_after = Pt(4)
                p_img.add_run().add_picture(img_path, width=Inches(cfg3.get("ancho_pulgadas", 5.5)))

            # Criterios / Desglose
            criterios = cfg3.get("criterios", [])
            if criterios:
                t_cr = doc.add_table(rows=len(criterios), cols=1)
                t_cr.alignment = WD_TABLE_ALIGNMENT.CENTER
                t_cr.autofit = False
                for idx_cr, crit in enumerate(criterios):
                    c = t_cr.cell(idx_cr, 0)
                    c.width = Pt(511.2)
                    set_cell_shading(c, "F6FAFE" if is_pauta else COLOR_BOX_BG_HEX)
                    set_cell_margins(c, top=35, bottom=35, left=60, right=60)
                    set_cell_borders(c, top={'color': COLOR_BORDER_HEX}, bottom={'color': COLOR_BORDER_HEX},
                                        left={'val': 'single', 'sz': '12', 'color': COLOR_CORRECT_HEX if is_pauta else COLOR_NAVY_HEX},
                                        right={'color': COLOR_BORDER_HEX})
                    p = c.paragraphs[0]
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    r = p.add_run(f"• {crit}")
                    r.font.name = "Arial"
                    r.font.size = Pt(8)
                    if is_pauta:
                        r.font.color.rgb = COLOR_CORRECT_RGB

        doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # --------------------------------------------------------------------------
    # MÉTODOS PÚBLICOS DE COMPILACIÓN
    # --------------------------------------------------------------------------
    def build_evaluacion_estudiante(self, out_path=None):
        doc = create_base_doc()
        add_header(doc, school=self.school, subject=f"{self.subject} — {self.course}")
        add_title_banner(doc, self.title)
        self._add_info_box(doc, is_pauta=False)

        self._render_item1(doc, is_pauta=False)
        self._render_item2(doc, is_pauta=False)
        self._render_item3(doc, is_pauta=False)

        if not out_path:
            clean_sub = self.subject.replace(" ", "_").capitalize()
            clean_crs = self.course.replace(" ", "_").replace("°", "").replace("—", "")
            out_path = os.path.join(OUTPUT_DIR, f"Evaluacion_{clean_sub}_{clean_crs}.docx")

        safe_save(doc, out_path)
        print(f"[OK] Prueba Estudiante generada: {out_path}")
        return out_path

    def build_pauta_correccion(self, out_path=None):
        doc = create_base_doc()
        add_header(doc, school=self.school, subject=f"{self.subject} — {self.course}", subtitle="DOCUMENTO DOCENTE")
        add_title_banner(doc, f"PAUTA DE CORRECCIÓN: {self.title}", is_pauta=True)
        self._add_info_box(doc, is_pauta=True)

        self._render_item1(doc, is_pauta=True)
        self._render_item2(doc, is_pauta=True)
        self._render_item3(doc, is_pauta=True)

        if not out_path:
            clean_sub = self.subject.replace(" ", "_").capitalize()
            clean_crs = self.course.replace(" ", "_").replace("°", "").replace("—", "")
            out_path = os.path.join(OUTPUT_DIR, f"Pauta_Correccion_{clean_sub}_{clean_crs}.docx")

        safe_save(doc, out_path)
        print(f"[OK] Pauta Docente generada: {out_path}")
        return out_path

    def build_all(self, target_folder=None):
        """Genera tanto la prueba como la pauta y las despliega en la carpeta deseada"""
        p_est = self.build_evaluacion_estudiante()
        p_pau = self.build_pauta_correccion()

        if target_folder and os.path.exists(target_folder):
            shutil.copy2(p_est, os.path.join(target_folder, os.path.basename(p_est)))
            shutil.copy2(p_pau, os.path.join(target_folder, os.path.basename(p_pau)))
            print(f"[OK] Desplegado con éxito en: {target_folder}")

        return p_est, p_pau

if __name__ == "__main__":
    print("DocenteEngine: Módulo de arquitectura pedagógica cargado exitosamente.")
