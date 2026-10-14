# -*- coding: utf-8 -*-
"""Lee una planificacion mensual ya generada (y escrita en Word por el docente) de vuelta a JSON.

Sirve para que una IA vea que semanas ya estan escritas y complete o sugiera las que faltan.
"""
import re

from docx import Document
from docx.oxml.ns import qn

from .mensual import CAMPOS


def _tx(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t'))).strip()


def _celdas(fila):
    return [' '.join(''.join(t.text or '' for t in p.iter(qn('w:t'))).strip() for p in tc.iter(qn('w:p'))).strip()
            for tc in fila.iter(qn('w:tc'))]


def _celda_dividida(tc):
    """(texto escrito, texto sugerido). La sugerencia se reconoce por su color gris (7F7F7F)."""
    normal, sug = [], []
    for r in tc.iter(qn('w:r')):
        t = ''.join(x.text or '' for x in r.iter(qn('w:t')))
        if not t:
            continue
        col = r.find(qn('w:rPr'))
        col = col.find(qn('w:color')) if col is not None else None
        (sug if col is not None and col.get(qn('w:val')) == '7F7F7F' else normal).append(t)
    return ''.join(normal).strip(), ''.join(sug).strip()


def leer_mensual(ruta_docx, curso_codigo, anio):
    d = Document(ruta_docx)
    plan = {'anio': anio, 'curso_codigo': curso_codigo, 'periodos': [], 'oa': []}
    for tbl in d.element.body.iter(qn('w:tbl')):
        filas = list(tbl.iter(qn('w:tr')))
        c0 = _celdas(filas[0])
        txt = _tx(tbl)
        if c0 and c0[0].startswith('PROFESOR'):
            plan['profesor'], plan['curso'], plan['asignatura'] = c0[1], c0[3], c0[5]
            plan['mes'] = c0[7]
        elif c0 and c0[0].startswith('UNIDAD PROGRAM'):
            plan['unidad'] = {'numero': c0[1], 'nombre': c0[3]}
            plan['horas'] = c0[5]
        elif c0 and c0[0].startswith('OBJETIVOS DE APRENDIZAJE QUE'):
            plan['oa'] = re.findall(r'(?:^|\s)OA\s+(\d+):', ' '.join(_celdas(filas[1])))
        elif c0 and c0[0].startswith('N° de clases'):
            nombres = [x.strip() for x in plan.get('mes', '').split(' - ')] or ['Mes']
            periodo = {'nombre': nombres[0] if nombres else 'Mes', 'clases': []}
            plan['periodos'].append(periodo)
            for fila in filas[1:]:
                c = _celdas(fila)
                if len(c) < 5 or (len(c) == 1) or (c[0].isalpha() and c[0].upper() == c[0]):
                    nombre = c[0].capitalize() if c else 'Mes'
                    periodo = {'nombre': nombre, 'clases': []}
                    plan['periodos'].append(periodo)
                    continue
                tcs = list(fila.iter(qn('w:tc')))[1:5]
                partes = [_celda_dividida(tc) for tc in tcs]
                clase = {k: p[0] for k, p in zip(CAMPOS, partes)}
                sug = {k: p[1] for k, p in zip(CAMPOS, partes) if p[1]}
                if sug:
                    clase['sugerencia'] = sug
                periodo['clases'].append(clase)
    return plan
