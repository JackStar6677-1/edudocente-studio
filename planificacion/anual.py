# -*- coding: utf-8 -*-
"""Plan Anual: JSON -> .docx en el formato del colegio (tabla OA/indicadores con grilla de semanas y marcas X).

Una "X" en una semana significa que ese objetivo se aborda en esa semana.
Si el JSON no trae OA, se genera la plantilla en blanco (para escribir directamente en Word).
"""
import json
import os

from . import banco as B
from . import docx_base as X

RAIZ = os.path.dirname(os.path.abspath(__file__))
ANCHO = 17400
W_OA, W_IND, W_SEM = 3510, 3855, 425
SEMESTRES = {1: ('1er SEMESTRE', ['MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO']),
             2: ('2do SEMESTRE', ['AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE'])}
UNIDADES_BASE = {1: [1, 2], 2: [3, 4]}
FILAS_EN_BLANCO = 4


def semanas_por_mes(plan, meses):
    dado = plan.get('semanas', {})
    return [int(dado.get(m, 4 if i == len(meses) - 1 else 5)) for i, m in enumerate(meses)]


def agrupar(plan, sem):
    """[(numero_unidad | None, [oa, ...]), ...] en orden."""
    grupos = {}
    for o in plan.get('oa', []):
        if o['semestre'] == sem:
            grupos.setdefault(o.get('unidad'), []).append(o)
    if not grupos:
        return [(u, []) for u in UNIDADES_BASE[sem]]
    return sorted(grupos.items(), key=lambda kv: (kv[0] is None, kv[0] or 0))


def tabla_oa(doc, meses, semanas, etiqueta_semestre, etiqueta_unidad, oas, filas_blanco=FILAS_EN_BLANCO):
    n_filas = len(oas) if oas else filas_blanco
    n_cols = 2 + sum(semanas)
    anchos = [W_OA, W_IND] + [W_SEM] * sum(semanas)
    t = X.tabla(doc, 2 + n_filas, anchos, borde=None, margenes=(40, 70))
    h0, h1 = t.rows[0], t.rows[1]
    a = h0.cells[0].merge(h1.cells[0]); b = h0.cells[1].merge(h1.cells[1])
    for c, txt in ((a, 'OBJETIVOS DE APRENDIZAJE'), (b, 'INDICADORES DE EVALUACIÓN')):
        X.sombrear(c, X.NAVY); X.parrafo(c, txt, True, 9, 'FFFFFF', X.CENTER); X.borde_celda(c); X.valign(c)
    X.parrafo(a, etiqueta_semestre, True, 9, 'FFFFFF', X.CENTER, first=False)
    if etiqueta_unidad:
        X.parrafo(a, etiqueta_unidad, False, 8.5, X.ICE, X.CENTER, first=False)
    col = 2
    spans = []
    for m, (mes, n) in enumerate(zip(meses, semanas)):
        c0 = h0.cells[col].merge(h0.cells[col + n - 1])
        X.sombrear(c0, X.NAVY); X.parrafo(c0, mes, True, 9, 'FFFFFF', X.CENTER); X.borde_celda(c0)
        for k in range(n):
            c = h1.cells[col + k]
            X.sombrear(c, X.ICE if m % 2 else 'FFFFFF')
            X.parrafo(c, str(k + 1), True, 8, X.NAVY, X.CENTER); X.borde_celda(c)
        spans.append((mes, col, n)); col += n
    for p in a.paragraphs + b.paragraphs:
        p.paragraph_format.keep_with_next = True
    for r in range(n_filas):
        fila = t.rows[2 + r]
        for c in fila.cells:
            X.borde_celda(c)
        oa = oas[r] if oas else None
        c0, c1 = fila.cells[0], fila.cells[1]
        if oa:
            p = X.parrafo(c0, '', after=1)
            X.run(p, (f'OA {oa["codigo"]}: ' if oa.get('codigo') else ''), True, 9, X.NAVY)
            X.run(p, oa['texto'], False, 9)
            for pt in oa.get('puntos', []):
                X.parrafo(c0, '–  ' + pt, False, 9, first=False, left=300, hanging=200)
            for j, ind in enumerate(oa['indicadores']):
                X.parrafo(c1, '•  ' + ind, False, 9, first=(j == 0), after=2, left=240, hanging=240)
        marcas = oa.get('marcas', {}) if oa else {}
        for m, (mes, c_ini, n) in enumerate(spans):
            for k in range(n):
                cel = fila.cells[c_ini + k]
                if m % 2:
                    X.sombrear(cel, X.ICE)
                X.parrafo(cel, 'X' if (k + 1) in marcas.get(mes, []) else '', True, 9, X.NAVY, X.CENTER)
                X.valign(cel)
        X.fila_fija(fila, alto_min=1100 if not oa else None)
    for tr in list(t.rows)[:2]:
        X.fila_fija(tr, encabezado=True)
    return t


def tabla_datos(doc, plan):
    filas = [('Sector o Subsector', plan['asignatura']), ('Profesor/a', plan['profesor']),
             ('Curso', plan['curso']), ('Año', str(plan['anio'])),
             ('Nº de Horas', plan.get('horas', '')), ('Nº de Clases', plan.get('clases', ''))]
    t = X.tabla(doc, len(filas), [3510, 8000])
    for (k, v), row in zip(filas, t.rows):
        X.sombrear(row.cells[0], X.ICE); X.parrafo(row.cells[0], k, True, 9.5, X.NAVY)
        X.parrafo(row.cells[1], v, False, 9.5)
    return t


def tabla_cobertura(doc):
    t = X.tabla(doc, 1, [2860] * 6)
    for i, txt in enumerate(['OA Enseñados', '', 'OA No Enseñados', '', '% Cobertura Curricular', '']):
        c = t.rows[0].cells[i]
        if i % 2 == 0:
            X.sombrear(c, X.ICE); X.parrafo(c, txt, True, 9.5, X.NAVY, X.CENTER)
        X.valign(c)
    X.fila_fija(t.rows[0], alto_min=620)
    return t


def construir(plan, destino):
    doc = X.nuevo_documento(margen_sup=1.0, margen_inf=1.1)
    colegio = plan.get('colegio', {})
    logo = colegio.get('logo') or X.LOGO_DEFECTO
    for sem in (1, 2):
        nombre, meses = SEMESTRES[sem]
        sem_meses = plan.get('meses', {}).get(str(sem)) or meses
        semanas = semanas_por_mes(plan, sem_meses)
        X.franja_titulo(doc, logo,
                        [(colegio.get('linea1', 'Complejo Educacional Particular'), 10, X.NAVY, True),
                         (colegio.get('linea2', ''), 10, X.NAVY, True),
                         (colegio.get('linea3', f'UTP {plan["anio"]}'), 8.5, '595959', False)],
                        f'PLAN ANUAL {plan["anio"]} · PRIORIZACIÓN CURRICULAR', nombre, ANCHO, salto_pagina=(sem == 2))
        tabla_datos(doc, plan); X.espaciador(doc, 5)
        nombres_unidad = plan.get('unidades', {})
        for unidad, oas in agrupar(plan, sem):
            etiqueta = None
            if unidad is not None:
                nom = nombres_unidad.get(str(unidad), '')
                etiqueta = f'Unidad N° {unidad}' + (f': {nom}' if nom else '')
            tabla_oa(doc, sem_meses, semanas, 'PRIMER SEMESTRE' if sem == 1 else 'SEGUNDO SEMESTRE', etiqueta, oas)
            X.espaciador(doc, 5)
        p = doc.add_paragraph(); p.paragraph_format.space_after = X.Pt(0)
        X.run(p, 'X = semana en que se aborda el objetivo. Ajustar el número de semanas de cada mes según el calendario escolar.',
              False, 8, '595959', italic=True)
        p.paragraph_format.keep_with_next = True
        X.espaciador(doc, 4).paragraph_format.keep_with_next = True
        tabla_cobertura(doc); X.espaciador(doc, 2)
    X.pie_de_pagina(doc, f'Plan Anual {plan["anio"]}  ·  {plan["asignatura"]}  ·  {plan["curso"]}')
    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    doc.save(destino)
    return destino


def plan_en_blanco(curso, curso_codigo, asignatura, anio, profesor, horas='', clases=''):
    return {'anio': anio, 'curso': curso, 'curso_codigo': curso_codigo, 'asignatura': asignatura,
            'profesor': profesor, 'horas': horas, 'clases': clases, 'oa': []}
