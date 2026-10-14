# -*- coding: utf-8 -*-
"""Plan Anual: JSON -> .docx segun un FORMATO editable (ver formato.py).

Por defecto: tabla de OA/indicadores con grilla de semanas y marcas "X" (= el objetivo se aborda esa semana).
El formato permite cambiar semestres, meses, semanas por mes, unidades, textos, colores, cabecera y cuadro de cobertura.
Si el JSON no trae OA, se genera la plantilla en blanco (para escribir directamente en Word).
"""
import os

from . import docx_base as X
from . import formato as F


def formato_de(plan):
    return F.cargar('anual', plan.get('formato_perfil'), plan.get('formato'))


def valores_plan(plan):
    return {'profesor': plan.get('profesor', ''), 'curso': plan.get('curso', ''), 'asignatura': plan.get('asignatura', ''),
            'anio': plan.get('anio', ''), 'horas': plan.get('horas', ''), 'clases': plan.get('clases', '')}


def agrupar(plan, sem_idx, unidades_base):
    """[(numero_unidad | None, [oa, ...]), ...] del semestre `sem_idx` (1, 2, ...)."""
    grupos = {}
    for o in plan.get('oa', []):
        if o['semestre'] == sem_idx:
            grupos.setdefault(o.get('unidad'), []).append(o)
    if not grupos:
        return [(u, []) for u in unidades_base] or [(None, [])]
    return sorted(grupos.items(), key=lambda kv: (kv[0] is None, kv[0] or 0))


def tabla_oa(doc, fmt, sem, etiqueta_unidad, oas):
    t_ = fmt['tabla']
    meses, semanas = sem['meses'], sem['semanas']
    n_filas = len(oas) if oas else t_['filas_en_blanco']
    anchos = [t_['ancho_oa'], t_['ancho_indicadores']] + [t_['ancho_semana']] * sum(semanas)
    t = X.tabla(doc, 2 + n_filas, anchos, borde=None, margenes=(40, 70))
    h0, h1 = t.rows[0], t.rows[1]
    a = h0.cells[0].merge(h1.cells[0]); b = h0.cells[1].merge(h1.cells[1])
    for c, txt in ((a, t_['titulo_oa']), (b, t_['titulo_indicadores'])):
        X.sombrear(c, X.NAVY); X.parrafo(c, txt, True, 9, 'FFFFFF', X.CENTER); X.borde_celda(c); X.valign(c)
    X.parrafo(a, sem['etiqueta'], True, 9, 'FFFFFF', X.CENTER, first=False)
    if etiqueta_unidad:
        X.parrafo(a, etiqueta_unidad, False, 8.5, X.ICE, X.CENTER, first=False)
    col, spans = 2, []
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
                X.parrafo(cel, t_['marca'] if (k + 1) in marcas.get(mes, []) else '', True, 9, X.NAVY, X.CENTER)
                X.valign(cel)
        X.fila_fija(fila, alto_min=1100 if not oa else None)
    for tr in list(t.rows)[:2]:
        X.fila_fija(tr, encabezado=True)
    return t


def tabla_datos(doc, fmt, v):
    filas = fmt['datos']
    if not filas:
        return
    t = X.tabla(doc, len(filas), [3510, 8000])
    for (k, val), row in zip(filas, t.rows):
        X.sombrear(row.cells[0], X.ICE); X.parrafo(row.cells[0], F.rellenar(k, v), True, 9.5, X.NAVY)
        X.parrafo(row.cells[1], F.rellenar(val, v), False, 9.5)


def tabla_cobertura(doc, etiquetas):
    cols = []
    for e in etiquetas:
        cols += [e, '']
    t = X.tabla(doc, 1, [int(17400 / len(cols))] * len(cols))
    for i, txt in enumerate(cols):
        c = t.rows[0].cells[i]
        if i % 2 == 0:
            X.sombrear(c, X.ICE); X.parrafo(c, txt, True, 9.5, X.NAVY, X.CENTER)
        X.valign(c)
    X.fila_fija(t.rows[0], alto_min=620)


def construir(plan, destino):
    fmt = formato_de(plan)
    X.usar_tema(fmt['colores'], fmt['fuente'])
    pg = fmt['pagina']
    doc = X.nuevo_documento(pg['ancho_cm'], pg['alto_cm'], pg['margen_lateral_cm'], pg['margen_superior_cm'], pg['margen_inferior_cm'])
    ancho = int((pg['ancho_cm'] - 2 * pg['margen_lateral_cm']) * 567)
    v = valores_plan(plan)
    colegio = plan.get('colegio', {})
    logo = (colegio.get('logo') or X.LOGO_DEFECTO) if fmt['mostrar_logo'] else None
    lineas = [F.rellenar(t, v) for t in fmt['institucion']] or ['']
    for i, k in enumerate(('linea1', 'linea2')):
        if i < len(lineas) and not lineas[i] and colegio.get(k):
            lineas[i] = colegio[k]
    estilos = [(10, X.NAVY, True), (10, X.NAVY, True), (8.5, '595959', False), (8.5, '595959', False)]
    inst = [(t, estilos[i][0], estilos[i][1], estilos[i][2]) for i, t in enumerate(lineas) if t]
    nombres_unidad = plan.get('unidades', {})
    for idx, sem in enumerate(fmt['semestres'], start=1):
        sem_def = dict(sem)
        meses_plan = plan.get('meses', {}).get(str(idx))        # el plan puede traer sus propios meses/semanas
        if meses_plan:
            sem_def['meses'] = [m.upper() for m in meses_plan]
            dado = plan.get('semanas', {})
            sem_def['semanas'] = [int(dado.get(m, sem['semanas'][i] if i < len(sem['semanas']) else 5)) for i, m in enumerate(sem_def['meses'])]
        X.franja_titulo(doc, logo, inst, F.rellenar(fmt['titulo'], v), sem['nombre'], ancho, salto_pagina=(idx > 1))
        tabla_datos(doc, fmt, v); X.espaciador(doc, 5)
        for unidad, oas in agrupar(plan, idx, sem['unidades']):
            etiqueta = None
            if unidad is not None:
                nom = nombres_unidad.get(str(unidad), '')
                etiqueta = f'{fmt["tabla"]["prefijo_unidad"]} {unidad}'.strip() + (f': {nom}' if nom else '')
            tabla_oa(doc, fmt, sem_def, etiqueta, oas)
            X.espaciador(doc, 5)
        if fmt.get('leyenda'):
            p = doc.add_paragraph(); p.paragraph_format.space_after = X.Pt(0)
            X.run(p, fmt['leyenda'], False, 8, '595959', italic=True)
            p.paragraph_format.keep_with_next = True
            X.espaciador(doc, 4).paragraph_format.keep_with_next = True
        if fmt['cobertura']:
            tabla_cobertura(doc, fmt['cobertura'])
        X.espaciador(doc, 2)
    X.pie_de_pagina(doc, F.rellenar(fmt['pie'], v))
    os.makedirs(os.path.dirname(os.path.abspath(destino)), exist_ok=True)
    X.finalizar(doc, 'Plan anual')
    doc.save(destino)
    return destino


def plan_en_blanco(curso, curso_codigo, asignatura, anio, profesor, horas='', clases=''):
    return {'anio': anio, 'curso': curso, 'curso_codigo': curso_codigo, 'asignatura': asignatura,
            'profesor': profesor, 'horas': horas, 'clases': clases, 'oa': []}
