# -*- coding: utf-8 -*-
"""Formato editable de la evaluacion: ponderacion por item, exigencia %, etiquetas y bloques visibles.

    python -m unittest tests.test_engine_formato -v
"""
import copy
import os
import shutil
import tempfile
import unittest

from docx import Document
from docx.oxml.ns import qn

from engine import DocenteEngine

BASE = {
    "colegio": "COLEGIO DE PRUEBA", "asignatura": "CIENCIAS NATURALES", "curso": "5° Básico", "titulo": "EVALUACIÓN FINAL",
    "oa": "OA 1: Objetivo de prueba", "contenidos": "Contenido de prueba", "puntaje_total": 20, "anio": 2031,
    "item1_seleccion_multiple": [
        {"pregunta": f"{i}. ¿Pregunta {i}?", "alternativas": [["A", "uno"], ["B", "dos"], ["C", "tres"], ["D", "cuatro"]],
         "correcta": "A) uno", "justificacion": "porque sí"} for i in range(1, 5)],
    "item1_pts_cada_una": 2,
    "item2_verdadero_falso": [{"oracion": f"Afirmación {i}", "resp": "V", "justificacion": "ok"} for i in range(1, 9)],
    "item2_pts_cada_una": 1,
    "item3_aplicacion": {"titulo": "ÍTEM III: DIBUJO", "puntaje": 4, "instruccion": "Dibuja.", "tipo": "drawing_boxes"},
}


def _texto(ruta):
    d = Document(ruta)
    return ' | '.join(''.join(t.text or '' for t in p.iter(qn('w:t'))) for p in d.element.body.iter(qn('w:p')))


class PruebasFormatoEvaluacion(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _generar(self, formato=None):
        cfg = copy.deepcopy(BASE)
        if formato is not None:
            cfg["formato"] = formato
        ruta = os.path.join(self.tmp, "e.docx")
        DocenteEngine(cfg).build_evaluacion_estudiante(ruta)
        return _texto(ruta)

    def test_sin_formato_no_cambia_nada(self):
        t = self._generar()
        self.assertIn("2 puntos c/u — Total: 8 puntos", t)
        self.assertIn("Puntaje Total: 20 puntos", t)
        self.assertNotIn("Exigencia", t)
        self.assertIn("2031", t)

    def test_ponderacion_recalcula_puntajes(self):
        t = self._generar({"ponderacion": {"item1": 40, "item2": 40, "item3": 20}})
        self.assertIn("2 puntos c/u — Total: 8 puntos", t)             # 40% de 20 = 8 pts en 4 preguntas
        self.assertIn("1 punto c/u — Total: 8 puntos", t)              # 40% de 20 = 8 pts en 8 afirmaciones
        self.assertIn("(4 puntos)", t)                                  # 20% de 20
        self.assertIn("Puntaje Total: 20 puntos", t)

    def test_ponderacion_que_no_cuadra_muestra_el_total_real(self):
        t = self._generar({"ponderacion": {"item1": 50, "item2": 30, "item3": 20}})
        self.assertIn("2,5 puntos c/u — Total: 10 puntos", t)          # 50% de 20 en 4 preguntas
        self.assertIn("Puntaje Total: 22 puntos", t)                    # 30% en 8 afirmaciones se redondea a 1 pto c/u (multiplos de 0,5)

    def test_exigencia_y_etiquetas_y_ocultar_oa(self):
        t = self._generar({"exigencia_pct": 60, "etiquetas": {"nombre": "Alumno/a:", "puntaje_total": "Total:"},
                           "mostrar_oa": False, "mostrar_contenidos": False})
        self.assertIn("Exigencia: 60% (puntaje mínimo de aprobación: 12 pts)", t)
        self.assertIn("Alumno/a:", t)
        self.assertIn("Total: 20 puntos", t)
        self.assertNotIn("Objetivo de Aprendizaje", t)
        self.assertNotIn("Contenidos Evaluados", t)


if __name__ == "__main__":
    unittest.main()
