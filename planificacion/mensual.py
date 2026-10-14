# -*- coding: utf-8 -*-
"""Planificacion mensual (o de varios meses): JSON -> .docx en el formato del colegio.

El docente solo escribe, por cada clase: objetivo, actividad, evaluacion y recursos.
Los OA y sus indicadores se toman del banco curricular segun los numeros indicados.
"""
import json
import os

from . import banco as B
from . import docx_base as X

RAIZ = os.path.dirname(os.path.abspath(__file__))
ANCHO = 16400
CAMPOS = ('objetivo', 'actividad', 'evaluacion', 'recursos')


def cargar_json(ruta):
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def clases_vacias(n):
    return [{c: '' for c in CAMPOS} for _ in range(n)]


def resolver_oa(plan):
    """Lista de OA ya con texto e indicadores (del banco, o escritos a mano en el JSON)."""
    anio = plan.get('anio_banco', plan['anio'])
    banco = None
    out = []
    unidad = plan['unidad'].get('numero')
    for item in plan['oa']:
        if isinstance(item, dict):                      # OA escrito a mano en el JSON
            out.append({'codigo': str(item['codigo']), 'texto': item['texto'],
                        'puntos': item.get('puntos', []), 'indicadores': item.get('indicadores', [])})
            continue
        if banco is None:
            banco = B.cargar(plan.get('banco_curso', plan['curso_codigo']), plan['asignatura'], anio)
        out.append(B.buscar_oa(banco, item, unidad))
    return out


def construir(plan, destino, incluir_sugerencias=False):
    oas = resolver_oa(plan)
    doc = X.nuevo_documento(margen_sup=1.0, margen_inf=1.1)
    meses = ' - '.join(p['nombre'] for p in plan['periodos'])

    # ---- franja de titulo
    colegio = plan.get('colegio', {})
    logo = colegio.get('logo') or X.LOGO_DEFECTO
    X.franja_titulo(doc, logo,
                    [(colegio.get('linea1', 'Complejo Educacional Particular'), 10, X.NAVY, True),
                     (colegio.get('linea2', ''), 10, X.NAVY, True),
                     (colegio.get('linea3', f'UTP {plan["anio"]}'), 8.5, '595959', False)],
                    'PLANIFICACIÓN MENSUAL', f'{meses.upper()} {plan["anio"]}', ANCHO)

    # ---- datos generales
    def pares(anchos, textos):
        t = X.tabla(doc, 1, anchos)
        for i, x in enumerate(textos):
            c = t.rows[0].cells[i]
            if i % 2 == 0:
                X.sombrear(c, X.ICE); X.parrafo(c, x, True, 9.5, X.NAVY)
            else:
                X.parrafo(c, x, False, 10)
            X.valign(c)
    pares([1350, 3300, 1000, 2200, 1800, 2900, 800, 3050],
          ['PROFESOR:', plan['profesor'], 'CURSO:', plan['curso'], 'ASIGNATURA:', plan['asignatura'], 'MES:', meses])
    pares([2650, 700, 2700, 5000, 1900, 3450],
          ['UNIDAD PROGRAMÁTICA N°:', str(plan['unidad']['numero']), 'NOMBRE DE LA UNIDAD:', plan['unidad']['nombre'],
           'N° HORAS MENSUAL:', plan.get('horas', '')])
    X.espaciador(doc, 5)

    # ---- objetivos de aprendizaje
    t = X.tabla(doc, 2, [ANCHO])
    X.sombrear(t.rows[0].cells[0], X.MID)
    X.parrafo(t.rows[0].cells[0], 'OBJETIVOS DE APRENDIZAJE QUE SE ABORDARÁN EN EL MES:', True, 10, 'FFFFFF')
    cuerpo = t.rows[1].cells[0]; X.sombrear(cuerpo, X.PALE)
    for i, oa in enumerate(oas):
        p = X.parrafo(cuerpo, '', first=(i == 0), after=3 if not oa['puntos'] else 1)
        X.run(p, f'OA {oa["codigo"]}: ', True, 10, X.NAVY)
        X.run(p, oa['texto'], False, 10)
        for k, pt in enumerate(oa['puntos']):
            X.parrafo(cuerpo, '–  ' + pt, False, 10, first=False, left=510, hanging=200,
                      after=3 if k == len(oa['puntos']) - 1 else 0)
    X.fila_fija(t.rows[1])
    X.espaciador(doc, 6)

    # ---- tabla de clases
    anchos = [1150, 2700, 5400, 3950, 3200]
    n_filas = 1 + sum(len(p['clases']) for p in plan['periodos']) + max(0, len(plan['periodos']) - 1)
    t = X.tabla(doc, n_filas, anchos)
    for i, h in enumerate(['N° de clases', 'OBJETIVO DE LA CLASE', 'ESTRATEGIA DE ACTIVIDAD',
                           'ESTRATEGIA DE EVALUACIÓN', 'RECURSOS O MATERIALES']):
        c = t.rows[0].cells[i]; X.sombrear(c, X.NAVY)
        X.parrafo(c, h, True, 9.5, 'FFFFFF', X.CENTER if i == 0 else None); X.valign(c)
    X.fila_fija(t.rows[0], encabezado=True)
    ri = 1
    for k, periodo in enumerate(plan['periodos']):
        if k > 0:                                         # franja del mes (DICIEMBRE, etc.)
            fila = t.rows[ri]; ri += 1
            m = fila.cells[0].merge(fila.cells[4]); X.sombrear(m, X.MID)
            X.parrafo(m, periodo['nombre'].upper(), True, 10, 'FFFFFF', X.CENTER).paragraph_format.keep_with_next = True
            X.fila_fija(fila)
        for n, cl in enumerate(periodo['clases']):
            fila = t.rows[ri]; ri += 1
            relleno = X.PALE if n % 2 else 'FFFFFF'
            sug = cl.get('sugerencia', {}) if incluir_sugerencias else {}
            textos = [str(n + 1)] + [cl.get(c, '') for c in CAMPOS]
            for i, (cel, txt) in enumerate(zip(fila.cells, textos)):
                X.sombrear(cel, relleno)
                if i > 0 and not txt and sug.get(CAMPOS[i - 1]):
                    X.parrafo(cel, sug[CAMPOS[i - 1]], False, 9, '7F7F7F', italic=True)   # sugerencia en gris cursiva
                else:
                    X.parrafo(cel, txt, i == 0, 9.5, X.NAVY if i == 0 else None, X.CENTER if i == 0 else None)
                X.valign(cel, 'center' if i == 0 else 'top')
            X.fila_fija(fila, alto_min=760)
    X.espaciador(doc, 6)

    # ---- indicadores
    t = X.tabla(doc, 1 + len(oas), [1500, ANCHO - 1500])
    m = t.rows[0].cells[0].merge(t.rows[0].cells[1]); X.sombrear(m, X.MID)
    X.parrafo(m, 'Indicadores de Evaluación:', True, 10, 'FFFFFF').paragraph_format.keep_with_next = True
    X.fila_fija(t.rows[0], encabezado=True)
    for i, oa in enumerate(oas):
        a, b = t.rows[1 + i].cells
        X.sombrear(a, X.ICE); X.parrafo(a, f'OA {oa["codigo"]}', True, 10, X.NAVY, X.CENTER); X.valign(a)
        for j, ind in enumerate(oa['indicadores']):
            X.parrafo(b, '•  ' + ind, False, 9, first=(j == 0), after=2, left=255, hanging=255)
        if len(oa['indicadores']) <= 6:
            X.fila_fija(t.rows[1 + i])
    fin = X.espaciador(doc, 1); fin.runs[0].font.size = X.Pt(1)
    X.pie_de_pagina(doc, f'Planificación {meses} {plan["anio"]}  ·  {plan["asignatura"]}  ·  {plan["curso"]}')
    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    doc.save(destino)
    return destino


def nombre_archivo(plan):
    meses = '-'.join(p['nombre'] for p in plan['periodos'])
    return f'{plan["curso_codigo"]} - {plan["asignatura"]} - Planificación {meses}.docx'
