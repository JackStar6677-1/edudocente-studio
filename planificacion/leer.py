# -*- coding: utf-8 -*-
"""Lee una planificacion mensual ya generada (y escrita en Word por el docente) de vuelta a JSON.

Sirve para que una IA vea que semanas ya estan escritas y complete o sugiera las que faltan.
Funciona con cualquier formato (ver formato.py): reconoce la cabecera por las etiquetas del formato.
"""
import re

from docx import Document
from docx.oxml.ns import qn

from . import formato as F
from .mensual import campos_de


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


def leer_mensual(ruta_docx, curso_codigo, anio, fmt=None):
    fmt = fmt or F.defecto('mensual')
    claves = campos_de(fmt)
    cols = fmt['clases']['columnas']
    primera = cols[0]['titulo']
    etiquetas = {}                         # etiqueta de la cabecera -> campo del plan ({profesor}, {curso}, ...)
    for fila in fmt['datos']:
        for etiqueta, valor in fila:
            m = re.fullmatch(r'\{(\w+)\}', valor.strip())
            if m:
                etiquetas[etiqueta.strip()] = m.group(1)
    d = Document(ruta_docx)
    plan = {'anio': anio, 'curso_codigo': curso_codigo, 'periodos': [], 'oa': [], 'unidad': {}}
    for tbl in d.element.body.iter(qn('w:tbl')):
        filas = list(tbl.iter(qn('w:tr')))
        c0 = _celdas(filas[0])
        if not c0:
            continue
        if any(c0[i] in etiquetas for i in range(0, len(c0) - 1, 2)):
            for i in range(0, len(c0) - 1, 2):
                campo = etiquetas.get(c0[i])
                if campo in ('profesor', 'curso', 'asignatura', 'horas'):
                    plan[campo] = c0[i + 1]
                elif campo == 'meses':
                    plan['mes'] = c0[i + 1]
                elif campo == 'unidad_numero':
                    plan['unidad']['numero'] = c0[i + 1]
                elif campo == 'unidad_nombre':
                    plan['unidad']['nombre'] = c0[i + 1]
        elif c0[0] == fmt['oa']['titulo'] and len(filas) > 1:
            pref = re.escape(fmt['oa'].get('prefijo', 'OA'))
            plan['oa'] = re.findall(rf'(?:^|\s){pref}\s+(\w+):', ' '.join(_celdas(filas[1])))
        elif c0[0] == primera:
            nombres = [x.strip() for x in plan.get('mes', '').split(' - ')] or ['Mes']
            periodo = {'nombre': nombres[0] if nombres else 'Mes', 'clases': []}
            plan['periodos'].append(periodo)
            for fila in filas[1:]:
                c = _celdas(fila)
                if len(c) < len(cols):                          # franja de mes (una sola celda combinada)
                    periodo = {'nombre': (c[0] if c else 'Mes').capitalize(), 'clases': []}
                    plan['periodos'].append(periodo)
                    continue
                tcs = [tc for tc, col in zip(fila.iter(qn('w:tc')), cols) if col['clave'] != 'n']
                partes = [_celda_dividida(tc) for tc in tcs]
                clase = {k: p[0] for k, p in zip(claves, partes)}
                sug = {k: p[1] for k, p in zip(claves, partes) if p[1]}
                if sug:
                    clase['sugerencia'] = sug
                periodo['clases'].append(clase)
    return plan
