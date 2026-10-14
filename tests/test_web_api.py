# -*- coding: utf-8 -*-
"""Saneamiento de lo que envia el navegador y formatos editables por la API web.

    python -m unittest tests.test_web_api -v
"""
import os
import shutil
import tempfile
import unittest

from planificacion import formato, web_api


class PruebasWebApi(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self._carpeta = formato.CARPETA
        formato.CARPETA = os.path.join(self.tmp, 'formatos')

    def tearDown(self):
        formato.CARPETA = self._carpeta
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_plan_se_sanea_y_se_acota(self):
        sucio = {'anio': 99999, 'asignatura': 'X' * 1000, 'curso': 'A', 'unidad': {'numero': '1' * 50},
                 'periodos': [{'nombre': 'Mes', 'clases': [{'obj<>etivo!': 'y' * 9000}] * 100}] * 20,
                 'oa': [{'codigo': '1', 'texto': 't', 'indicadores': ['i'] * 100}] * 30,
                 'colegio': {'linea1': 'NO DEBE ACEPTARSE'}}
        p = web_api.limpiar_plan(sucio)
        self.assertEqual(p['anio'], 2100)
        self.assertEqual(len(p['asignatura']), 80)
        self.assertEqual(len(p['periodos']), 6)
        self.assertEqual(len(p['periodos'][0]['clases']), 40)
        self.assertEqual(len(p['oa']), 12)
        self.assertEqual(len(p['oa'][0]['indicadores']), 30)
        self.assertIn('objetivo', p['periodos'][0]['clases'][0])          # la clave se limpia (solo a-z0-9_)
        self.assertEqual(len(p['periodos'][0]['clases'][0]['objetivo']), 3000)
        self.assertNotEqual(p['colegio'].get('linea1'), 'NO DEBE ACEPTARSE')   # el colegio lo define el servidor

    def test_plan_sin_datos_minimos_se_rechaza(self):
        with self.assertRaises(ValueError):
            web_api.limpiar_plan({'asignatura': '', 'periodos': []})

    def test_formato_con_nombre_peligroso_queda_en_la_carpeta_de_formatos(self):
        nombre = web_api.guardar_formato({'tipo': 'mensual', 'nombre': '../../etc/passwd', 'formato': {'titulo': 'X'}})
        self.assertNotIn('/', nombre)
        archivos = os.listdir(formato.CARPETA)
        self.assertEqual(len(archivos), 1)
        self.assertTrue(archivos[0].startswith('mensual.'))

    def test_ver_formato_con_ajustes_en_edicion(self):
        f = web_api.ver_formato({'tipo': 'mensual', 'formato': {'titulo': 'EDITADO', 'colores': {'primario': 'no'}}})
        self.assertEqual(f['titulo'], 'EDITADO')
        self.assertEqual(f['colores']['primario'], '1F3864')

    def test_tipo_invalido(self):
        with self.assertRaises(ValueError):
            web_api.ver_formato({'tipo': 'otro'})
        with self.assertRaises(ValueError):
            web_api.guardar_formato({'tipo': 'otro', 'nombre': 'a'})

    def test_formato_de_evaluacion_validado(self):
        f = web_api.limpiar_formato_evaluacion({'exigencia_pct': 500, 'ponderacion': {'item1': 'x', 'item2': 30}, 'etiquetas': {'nombre': 'A', 'hack': 'b'}})
        self.assertEqual(f['exigencia_pct'], 100)
        self.assertEqual(f['ponderacion'], {'item1': 0, 'item2': 30})
        self.assertEqual(f['etiquetas'], {'nombre': 'A'})


if __name__ == '__main__':
    unittest.main()
