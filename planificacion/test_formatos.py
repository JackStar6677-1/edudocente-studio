# -*- coding: utf-8 -*-
"""Pruebas de formatos editables:  python -m unittest planificacion.test_formatos -v

No dependen de datos locales: los OA se escriben a mano en el plan.
"""
import os
import shutil
import tempfile
import unittest

from docx import Document
from docx.oxml.ns import qn

from planificacion import anual, formato, leer, mensual, temario


def _texto(ruta):
    d = Document(ruta)
    return ' | '.join(''.join(t.text or '' for t in p.iter(qn('w:t'))) for p in d.element.body.iter(qn('w:p')))


def _plan(**extra):
    p = {'anio': 2030, 'profesor': 'Docente Ejemplo', 'curso': '3°A', 'curso_codigo': '3°A', 'asignatura': 'Lenguaje',
         'unidad': {'numero': '2', 'nombre': 'Cuentos'}, 'horas': '4',
         'oa': [{'codigo': '7', 'texto': 'Texto del objetivo', 'puntos': ['punto uno'], 'indicadores': ['Leen cuentos']}],
         'periodos': [{'nombre': 'Mayo', 'clases': mensual.clases_vacias(2)}, {'nombre': 'Junio', 'clases': mensual.clases_vacias(1)}]}
    p.update(extra)
    return p


class PruebasFormatos(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self._carpeta = formato.CARPETA
        formato.CARPETA = os.path.join(self.tmp, 'formatos')          # los perfiles de prueba no tocan los datos reales

    def tearDown(self):
        formato.CARPETA = self._carpeta
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_formato_por_defecto_y_validacion(self):
        f = formato.cargar('mensual', None, {'colores': {'primario': 'zzz'}, 'fuente': 'Comic'})
        self.assertEqual(f['colores']['primario'], '1F3864')          # color invalido -> se conserva el valor por defecto
        self.assertEqual(f['fuente'], 'Arial')                        # fuente no permitida -> Arial

    def test_columnas_personalizadas_en_mensual(self):
        fmt = {'titulo': 'PLAN DE CLASES', 'colores': {'primario': '7A1F1F'},
               'clases': {'columnas': [{'clave': 'n', 'titulo': 'Sesión', 'ancho': 900},
                                       {'clave': 'tema', 'titulo': 'TEMA', 'ancho': 3000},
                                       {'clave': 'tiempo', 'titulo': 'MINUTOS', 'ancho': 1200}]}}
        plan = _plan(formato=fmt)
        plan['periodos'][0]['clases'] = [{'tema': 'Fábulas', 'tiempo': '90'}, {'tema': '', 'tiempo': ''}]
        plan['periodos'][1]['clases'] = [{'tema': 'Poemas', 'tiempo': '45'}]
        ruta = mensual.construir(plan, os.path.join(self.tmp, 'm.docx'))
        t = _texto(ruta)
        for esperado in ('PLAN DE CLASES', 'Sesión', 'MINUTOS', 'Fábulas', 'Poemas', 'JUNIO'):
            self.assertIn(esperado, t)
        self.assertNotIn('ESTRATEGIA DE ACTIVIDAD', t)
        self.assertIn('7A1F1F', Document(ruta).element.xml)            # el color elegido llega al documento
        leido = leer.leer_mensual(ruta, '3°A', 2030, formato.cargar('mensual', None, fmt))
        self.assertEqual(leido['periodos'][1]['clases'][0]['tema'], 'Poemas')

    def test_bloques_y_cabecera_editables(self):
        fmt = {'bloques': ['clases'], 'datos': [[['DOCENTE', '{profesor}']]]}
        ruta = mensual.construir(_plan(formato=fmt), os.path.join(self.tmp, 'm2.docx'))
        t = _texto(ruta)
        self.assertNotIn('OBJETIVOS DE APRENDIZAJE', t)               # bloque oculto
        self.assertNotIn('Indicadores de Evaluación', t)

    def test_perfil_guardado_se_aplica_y_se_restablece(self):
        nombre = formato.guardar('mensual', 'Mi colegio', {'titulo': 'MI FORMATO'})
        self.assertIn(nombre, formato.listar('mensual'))
        ruta = mensual.construir(_plan(formato_perfil=nombre), os.path.join(self.tmp, 'm3.docx'))
        self.assertIn('MI FORMATO', _texto(ruta))
        self.assertTrue(formato.restablecer('mensual', nombre))
        self.assertNotIn(nombre, formato.listar('mensual'))

    def test_anual_con_otros_semestres_y_sin_cobertura(self):
        fmt = {'semestres': [{'nombre': 'TRIMESTRE 1', 'etiqueta': 'T1', 'meses': ['MARZO', 'ABRIL'], 'semanas': [3, 3], 'unidades': [1]}],
               'cobertura': [], 'tabla': {'marca': '✔'}}
        plan = anual.plan_en_blanco('3°A', '3°A', 'Lenguaje', 2030, 'Docente Ejemplo')
        plan['formato'] = fmt
        plan['oa'] = [{'semestre': 1, 'unidad': 1, 'codigo': '5', 'texto': 'Texto', 'puntos': [], 'indicadores': ['Ind'], 'marcas': {'MARZO': [1, 2]}}]
        t = _texto(anual.construir(plan, os.path.join(self.tmp, 'a.docx')))
        self.assertIn('TRIMESTRE 1', t)
        self.assertEqual(t.count('✔'), 2)
        self.assertNotIn('OA Enseñados', t)

    def test_temario_con_secciones_propias(self):
        datos = {'curso': '3°A', 'asignatura': 'Lenguaje', 'profesor': 'Docente Ejemplo', 'fecha': '10 de junio', 'puntaje': '30',
                 'contenidos': ['Fábulas', 'Poemas'],
                 'formato': {'secciones': [{'titulo': 'Qué estudiar', 'tipo': 'lista', 'contenido': '{contenidos}'},
                                           {'titulo': 'Pauta', 'tipo': 'texto', 'contenido': 'Lectura en voz alta.'}]}}
        t = _texto(temario.construir(datos, os.path.join(self.tmp, 't.docx')))
        for esperado in ('Qué estudiar', 'Fábulas', 'Poemas', 'Pauta', 'Lectura en voz alta.', '10 de junio'):
            self.assertIn(esperado, t)


if __name__ == '__main__':
    unittest.main()
