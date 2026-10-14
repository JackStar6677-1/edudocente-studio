# -*- coding: utf-8 -*-
"""Temario de evaluacion (informativo para estudiantes y apoderados): JSON -> .docx segun un FORMATO editable.

Las secciones del temario (titulo, tipo lista/texto, contenido) y los datos de la cabecera son totalmente editables
en el formato; el contenido puede usar campos {contenidos}, {como_sera}, {fecha}, {puntaje}, {curso}, {asignatura}, {profesor}.
"""
import os

from . import docx_base as X
from . import formato as F


def valores(datos):
    cont = datos.get('contenidos', '')
    if isinstance(cont, list):
        cont = '\n'.join(str(x) for x in cont)
    return {'profesor': datos.get('profesor', ''), 'curso': datos.get('curso', ''), 'asignatura': datos.get('asignatura', ''),
            'anio': datos.get('anio', ''), 'fecha': datos.get('fecha', ''), 'puntaje': datos.get('puntaje', ''),
            'contenidos': cont, 'como_sera': datos.get('como_sera', '')}


def construir(datos, destino):
    fmt = F.cargar('temario', datos.get('formato_perfil'), datos.get('formato'))
    X.usar_tema(fmt['colores'], fmt['fuente'])
    pg = fmt['pagina']
    doc = X.nuevo_documento(pg['ancho_cm'], pg['alto_cm'], pg['margen_lateral_cm'], pg['margen_superior_cm'], pg['margen_inferior_cm'])
    ancho = int((pg['ancho_cm'] - 2 * pg['margen_lateral_cm']) * 567)
    v = valores(datos)
    colegio = datos.get('colegio', {})
    logo = (colegio.get('logo') or X.LOGO_DEFECTO) if fmt['mostrar_logo'] else None
    lineas = [F.rellenar(t, v) for t in fmt['institucion']] or ['']
    for i, k in enumerate(('linea1', 'linea2')):
        if i < len(lineas) and not lineas[i] and colegio.get(k):
            lineas[i] = colegio[k]
    estilos = [(10, X.NAVY, True), (10, X.NAVY, True), (8.5, '595959', False), (8.5, '595959', False)]
    X.franja_titulo(doc, logo, [(t, estilos[i][0], estilos[i][1], estilos[i][2]) for i, t in enumerate(lineas) if t],
                    F.rellenar(fmt['titulo'], v), F.rellenar(fmt.get('subtitulo', ''), v), ancho)
    if fmt['datos']:
        t = X.tabla(doc, len(fmt['datos']), [int(ancho * 0.35), int(ancho * 0.65)])
        for (k, val), row in zip(fmt['datos'], t.rows):
            X.sombrear(row.cells[0], X.ICE); X.parrafo(row.cells[0], F.rellenar(k, v), True, 10, X.NAVY)
            X.parrafo(row.cells[1], F.rellenar(val, v), False, 10)
        X.espaciador(doc, 8)
    for sec in fmt['secciones']:
        t = X.tabla(doc, 1, [ancho])
        X.sombrear(t.rows[0].cells[0], X.MID)
        X.parrafo(t.rows[0].cells[0], F.rellenar(sec['titulo'], v), True, 10.5, 'FFFFFF').paragraph_format.keep_with_next = True
        texto = F.rellenar(sec['contenido'], v)
        lineas_txt = [x.strip() for x in texto.split('\n') if x.strip()]
        for x in lineas_txt:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = X.Pt(3)
            if sec['tipo'] == 'lista':
                p.paragraph_format.left_indent = X.Cm(0.9); p.paragraph_format.first_line_indent = X.Cm(-0.45)
                X.run(p, '•  ' + x.lstrip('-•· '), False, 11)
            else:
                X.run(p, x, False, 11)
        X.espaciador(doc, 8)
    if fmt.get('cierre'):
        for x in F.rellenar(fmt['cierre'], v).split('\n'):
            p = doc.add_paragraph(); p.paragraph_format.space_after = X.Pt(0)
            X.run(p, x, False, 10.5)
    X.pie_de_pagina(doc, F.rellenar(fmt['pie'], v))
    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    X.finalizar(doc, 'Temario de evaluación')
    doc.save(destino)
    return destino
