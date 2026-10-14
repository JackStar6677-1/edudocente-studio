# -*- coding: utf-8 -*-
"""Pruebas rapidas del modulo:  python -m unittest planificacion.test_planificacion -v

Usan el banco curricular local (planificacion/datos/banco/2026). Si no existe, se omiten.
"""
import glob
import json
import os
import shutil
import tempfile
import unittest

from docx import Document
from docx.oxml.ns import qn

from planificacion import anual, banco, ia, leer, mensual

RAIZ = os.path.dirname(os.path.abspath(__file__))
MENSUALES = sorted(glob.glob(os.path.join(RAIZ, 'datos', 'mensual', '2026-11_12', '*.json')))
BANCOS = sorted(glob.glob(os.path.join(RAIZ, 'datos', 'banco', '2026', '*.json')))


def _cargar(ruta):
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


@unittest.skipUnless(MENSUALES, 'sin datos locales de ejemplo')
class PruebasMensual(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_todas_se_generan_con_su_numeracion(self):
        for ruta in MENSUALES:
            plan = _cargar(ruta)
            destino = mensual.construir(plan, os.path.join(self.tmp, 'x.docx'))
            leido = leer.leer_mensual(destino, plan['curso_codigo'], plan['anio'])
            esperado = [len(p['clases']) for p in plan['periodos']]
            self.assertEqual([len(p['clases']) for p in leido['periodos']], esperado, ruta)
            self.assertEqual([str(x) for x in leido['oa']], [str(x) for x in plan['oa']], ruta)

    def test_escribir_semana_y_volver_a_leer(self):
        plan = _cargar(MENSUALES[0])
        plan['periodos'][0]['clases'][2].update(objetivo='Analizar X', actividad='Observan Y', evaluacion='Formativa', recursos='Texto págs. 1-2')
        destino = mensual.construir(plan, os.path.join(self.tmp, 'a.docx'))
        leido = leer.leer_mensual(destino, plan['curso_codigo'], plan['anio'])
        self.assertEqual(leido['periodos'][0]['clases'][2]['objetivo'], 'Analizar X')
        self.assertEqual(leido['periodos'][0]['clases'][0]['objetivo'], '')

    def test_sugerencias_offline_no_pisan_lo_escrito(self):
        plan = _cargar(MENSUALES[0])
        plan['periodos'][0]['clases'][0]['objetivo'] = 'Mi objetivo'
        n = ia.sugerir(plan, 'offline')
        self.assertGreater(n, 0)
        self.assertNotIn('sugerencia', plan['periodos'][0]['clases'][0])
        self.assertIn('sugerencia', plan['periodos'][0]['clases'][1])
        mensual.construir(plan, os.path.join(self.tmp, 's.docx'), incluir_sugerencias=True)


@unittest.skipUnless(BANCOS, 'sin banco curricular local')
class PruebasAnual(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_el_banco_reproduce_las_marcas_x(self):
        for ruta in BANCOS:
            plan = _cargar(ruta)
            destino = anual.construir(plan, os.path.join(self.tmp, 'a.docx'))
            d = Document(destino)
            xs = sum(1 for tc in d.element.body.iter(qn('w:tc'))
                     if ''.join(t.text or '' for t in tc.iter(qn('w:t'))).strip() == 'X')
            esperado = sum(len(v) for o in plan['oa'] for v in o['marcas'].values())
            self.assertEqual(xs, esperado, ruta)

    def test_plan_anual_en_blanco(self):
        plan = anual.plan_en_blanco('1°A', '1°A', 'Música', 2027, 'Docente')
        destino = anual.construir(plan, os.path.join(self.tmp, 'b.docx'))
        self.assertTrue(os.path.getsize(destino) > 5000)

    def test_buscar_oa_cae_a_otra_unidad(self):
        b = banco.cargar('5°AB', 'Orientación', 2026)
        self.assertEqual(banco.buscar_oa(b, 4, unidad=4)['codigo'], '4')


if __name__ == '__main__':
    unittest.main()
