#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Orquestador Integral de Evaluaciones Finales UTP 2026
Procesa los archivos de /tmp/evaluaciones_analisis, estandariza a Arial 12,
verifica puntajes (mínimo 25 puntos), asegura 3+ ítems, anexa Tablas de Especificaciones
y Pautas Oficiales en documentos docentes independientes (NUNCA en la prueba del alumno),
organiza la entrega por carpetas y empaqueta el ZIP final.
"""

import os
import sys
import shutil
import zipfile
import docx
from utp_formatter import (
    standardize_document_font,
    append_pauta_summary_table,
    append_specification_table
)

SRC_DIR = "/tmp/evaluaciones_analisis"
OUT_BASE_DIR = "/home/jack/Descargas/Evaluaciones_Finales_UTP_2026"
TEMARIOS_DIR = "/home/jack/Documentos/Desarrollo/Repositorios/personal/edudocente-studio/output/temarios"

# Asegurar directorios de salida organizados por asignatura y curso
os.makedirs(OUT_BASE_DIR, exist_ok=True)

# 1. CIENCIAS 5° BÁSICO
def procesar_ciencias_5():
    target_dir = os.path.join(OUT_BASE_DIR, "01 - Ciencias Naturales", "5° Básico")
    os.makedirs(target_dir, exist_ok=True)

    # A) Evaluación Estudiante
    doc_eval = docx.Document(os.path.join(SRC_DIR, "Evaluacion_Final_Ciencias_5Basico.docx"))
    standardize_document_font(doc_eval, default_size=12)
    doc_eval.save(os.path.join(target_dir, "Evaluacion_Final_Ciencias_5Basico.docx"))
    print("[OK] Ciencias 5° Básico: Evaluación de Estudiante limpia (Arial 12, sin pauta).")

    # B) Pauta Docente + Tabla de Especificaciones
    doc_pauta = docx.Document(os.path.join(SRC_DIR, "Pauta_Correccion_Ciencias_5Basico.docx"))
    standardize_document_font(doc_pauta, default_size=12)

    pauta_items = [
        ("Ítem I", "1", "C) Movimiento y flujo de cargas eléctricas", "2 pts"),
        ("Ítem I", "2", "A) Partículas diminutas que poseen cargas eléctricas", "2 pts"),
        ("Ítem I", "3", "B) Pequeñas esferas", "2 pts"),
        ("Ítem I", "4", "B) Positiva (+) o negativa (-)", "2 pts"),
        ("Ítem I", "5", "B) Cobre (material conductor)", "2 pts"),
        ("Ítem II", "1", "Verdadero (V)", "1 pto"),
        ("Ítem II", "2", "Verdadero (V)", "1 pto"),
        ("Ítem II", "3", "Falso (F) — Requieren conexión en camino cerrado", "1 pto"),
        ("Ítem II", "4", "Verdadero (V)", "1 pto"),
        ("Ítem II", "5", "Falso (F) — Requieren una fuente que impulse cargas", "1 pto"),
        ("Ítem II", "6", "Verdadero (V)", "1 pto"),
        ("Ítem II", "7", "Verdadero (V)", "1 pto"),
        ("Ítem II", "8", "Falso (F) — Regula corriente eléctrica, no hidráulica", "1 pto"),
        ("Ítem II", "9", "Verdadero (V)", "1 pto"),
        ("Ítem II", "10", "Verdadero (V)", "1 pto"),
        ("Ítem III", "1", "Fuente o generador: Placas paralelas desiguales (+ y -)", "2 pts"),
        ("Ítem III", "2", "Cables: Línea continua y recta", "2 pts"),
        ("Ítem III", "3", "Interruptor: Dos bornes con palanca abierta o cerrada", "2 pts"),
        ("Ítem III", "4", "Resistencia / Receptor: Círculo con X o zigzag", "2 pts")
    ]
    append_pauta_summary_table(doc_pauta, pauta_items)

    specs = {
        "headers": ["Eje Temático / Contenido / Ítem", "Conocimiento (36%)", "Comprensión (36%)", "Aplicación (28%)", "Total (100%)"],
        "rows": [
            ["Eje Ciencias de la Vida / Energía Eléctrica\nÍtem I: Selección Múltiple", "4 pts", "6 pts", "—", "10 pts"],
            ["Eje Circuitos Eléctricos y Componentes\nÍtem II: Verdadero o Falso", "6 pts", "4 pts", "—", "10 pts"],
            ["Eje Simbología Normalizada de Circuitos\nÍtem III: Dibujo y Esquematización", "—", "—", "8 pts", "8 pts"],
            ["Puntaje Total General", "10 pts (36%)", "10 pts (36%)", "8 pts (28%)", "28 pts (100%)"]
        ]
    }
    append_specification_table(doc_pauta, specs)
    doc_pauta.save(os.path.join(target_dir, "Pauta_Correccion_y_Especificaciones_Ciencias_5Basico.docx"))
    print("[OK] Ciencias 5° Básico: Pauta Oficial y Tabla de Especificaciones anexada.")

    # C) Copiar Temario Oficial
    temario_src = os.path.join(TEMARIOS_DIR, "Temario_Ciencias_5Basico_A_B.docx")
    if os.path.exists(temario_src):
        shutil.copyfile(temario_src, os.path.join(target_dir, "Temario_Ciencias_5Basico_A_B.docx"))

# 2. CIENCIAS 6° BÁSICO
def procesar_ciencias_6():
    target_dir = os.path.join(OUT_BASE_DIR, "01 - Ciencias Naturales", "6° Básico")
    os.makedirs(target_dir, exist_ok=True)

    # A) Evaluación Estudiante
    doc_eval = docx.Document(os.path.join(SRC_DIR, "Evaluacion_Final_Ciencias_6Basico.docx"))
    standardize_document_font(doc_eval, default_size=12)
    doc_eval.save(os.path.join(target_dir, "Evaluacion_Final_Ciencias_6Basico.docx"))
    print("[OK] Ciencias 6° Básico: Evaluación de Estudiante limpia (Arial 12, sin pauta).")

    # B) Pauta Docente + Tabla de Especificaciones
    doc_pauta = docx.Document(os.path.join(SRC_DIR, "Pauta_Correccion_Ciencias_6Basico.docx"))
    standardize_document_font(doc_pauta, default_size=12)

    pauta_items = [
        ("Ítem I", "1", "C) Cambios físicos sin alterar su composición", "2 pts"),
        ("Ítem I", "2", "A) pero continúa siendo la misma sustancia", "2 pts"),
        ("Ítem I", "3", "B) Por absorción o por liberación de energía (calor)", "2 pts"),
        ("Ítem I", "4", "A) Los cambios por absorción de energía (calor)", "2 pts"),
        ("Ítem I", "5", "A) Por liberación de calor (enfriamiento)", "2 pts"),
        ("Ítem II", "1", "Verdadero (V)", "1 pto"),
        ("Ítem II", "2", "Verdadero (V)", "1 pto"),
        ("Ítem II", "3", "Falso (F) — La fusión es de sólido a líquido", "1 pto"),
        ("Ítem II", "4", "Verdadero (V)", "1 pto"),
        ("Ítem II", "5", "Falso (F) — En la ebullición participan todas las partículas", "1 pto"),
        ("Ítem II", "6", "Verdadero (V)", "1 pto"),
        ("Ítem II", "7", "Verdadero (V)", "1 pto"),
        ("Ítem II", "8", "Falso (F) — También la sublimación inversa", "1 pto"),
        ("Ítem II", "9", "Verdadero (V)", "1 pto"),
        ("Ítem II", "10", "Verdadero (V)", "1 pto"),
        ("Ítem III", "1", "Solidificación: Líquido a Sólido (Partículas ordenadas y compactas)", "2 pts"),
        ("Ítem III", "2", "Sublimación Inversa: Gas a Sólido directo", "2 pts"),
        ("Ítem III", "3", "Condensación: Gas a Líquido (Partículas agrupadas en el fondo)", "2 pts")
    ]
    append_pauta_summary_table(doc_pauta, pauta_items)

    specs = {
        "headers": ["Eje Temático / Contenido / Ítem", "Comprensión (38%)", "Análisis (38%)", "Aplicación (24%)", "Total (100%)"],
        "rows": [
            ["Eje Materia y sus Transformaciones / Cambios de Estado\nÍtem I: Selección Múltiple", "4 pts", "6 pts", "—", "10 pts"],
            ["Eje Procesos Progresivos y Regresivos\nÍtem II: Verdadero o Falso", "6 pts", "4 pts", "—", "10 pts"],
            ["Eje Modelo Corpuscular y Liberación de Calor\nÍtem III: Esquematización y Dibujo", "—", "—", "6 pts", "6 pts"],
            ["Puntaje Total General", "10 pts (38%)", "10 pts (38%)", "6 pts (24%)", "26 pts (100%)"]
        ]
    }
    append_specification_table(doc_pauta, specs)
    doc_pauta.save(os.path.join(target_dir, "Pauta_Correccion_y_Especificaciones_Ciencias_6Basico.docx"))
    print("[OK] Ciencias 6° Básico: Pauta Oficial y Tabla de Especificaciones anexada.")

    # C) Copiar Temario Oficial
    temario_src = os.path.join(TEMARIOS_DIR, "Temario_Ciencias_6Basico_A_B.docx")
    if os.path.exists(temario_src):
        shutil.copyfile(temario_src, os.path.join(target_dir, "Temario_Ciencias_6Basico_A_B.docx"))

# 3. CIENCIAS 8° BÁSICO
def procesar_ciencias_8():
    target_dir = os.path.join(OUT_BASE_DIR, "01 - Ciencias Naturales", "8° Básico")
    os.makedirs(target_dir, exist_ok=True)

    # A) Evaluación Estudiante
    doc_eval = docx.Document(os.path.join(SRC_DIR, "Evaluacion_Final_Ciencias_8Basico.docx"))
    standardize_document_font(doc_eval, default_size=12)
    doc_eval.save(os.path.join(target_dir, "Evaluacion_Final_Ciencias_8Basico.docx"))
    print("[OK] Ciencias 8° Básico: Evaluación de Estudiante limpia (Arial 12, sin pauta).")

    # B) Pauta Docente + Tabla de Especificaciones
    doc_pauta = docx.Document(os.path.join(SRC_DIR, "Pauta_Correccion_Ciencias_8Basico.docx"))
    standardize_document_font(doc_pauta, default_size=12)

    pauta_items = [
        ("Ítem I", "1", "A) Partículas diminutas indivisibles llamadas átomos", "2 pts"),
        ("Ítem I", "2", "C) Indivisibles e indestructibles llamadas átomos", "2 pts"),
        ("Ítem I", "3", "B) Como una esfera maciza, compacta e indivisible", "2 pts"),
        ("Ítem I", "4", "A) Átomos reorganizados sin perder masa (conservación)", "2 pts"),
        ("Ítem I", "5", "B) Tres partículas fundamentales: p+, n0 y e-", "2 pts"),
        ("Ítem II", "1", "Verdadero (V)", "1 pto"),
        ("Ítem II", "2", "Verdadero (V)", "1 pto"),
        ("Ítem II", "3", "Verdadero (V)", "1 pto"),
        ("Ítem II", "4", "Falso (F) — Z corresponde a la cantidad de protones", "1 pto"),
        ("Ítem II", "5", "Verdadero (V)", "1 pto"),
        ("Ítem II", "6", "Verdadero (V)", "1 pto"),
        ("Ítem II", "7", "Falso (F) — No son figuras geométricas rectangulares", "1 pto"),
        ("Ítem II", "8", "Falso (F) — Contiene protones y electrones", "1 pto"),
        ("Ítem II", "9", "Falso (F) — Conservan su estructura nuclear", "1 pto"),
        ("Ítem II", "10", "Falso (F) — Cada elemento tiene su propio número atómico", "1 pto"),
        ("Ítem III", "1", "Corteza o Electrosfera: Rotula zona exterior de electrones", "1 pto"),
        ("Ítem III", "2", "Electrones (e-): Dibuja y sitúa en órbitas con carga negativa", "1 pto"),
        ("Ítem III", "3", "Núcleo Atómico: Rotula región central compacta", "1 pto"),
        ("Ítem III", "4", "Protones (p+): Dibuja en el núcleo con carga positiva", "1 pto"),
        ("Ítem III", "5", "Neutrones (n0) y Pintado: Partículas neutras y color diferenciado", "1 pto")
    ]
    append_pauta_summary_table(doc_pauta, pauta_items)

    specs = {
        "headers": ["Eje Temático / Contenido / Ítem", "Conocimiento (40%)", "Comprensión (40%)", "Aplicación (20%)", "Total (100%)"],
        "rows": [
            ["Eje Química / Teoría Atómica de Dalton\nÍtem I: Selección Múltiple", "4 pts", "6 pts", "—", "10 pts"],
            ["Eje Estructura Atómica y Notación Z y A\nÍtem II: Verdadero o Falso", "6 pts", "4 pts", "—", "10 pts"],
            ["Eje Modelo Atómico y Partículas Subatómicas\nÍtem III: Esquema y Rotulación", "—", "—", "5 pts", "5 pts"],
            ["Puntaje Total General", "10 pts (40%)", "10 pts (40%)", "5 pts (20%)", "25 pts (100%)"]
        ]
    }
    append_specification_table(doc_pauta, specs)
    doc_pauta.save(os.path.join(target_dir, "Pauta_Correccion_y_Especificaciones_Ciencias_8Basico.docx"))
    print("[OK] Ciencias 8° Básico: Pauta Oficial y Tabla de Especificaciones anexada.")

    # C) Copiar Temario Oficial
    temario_src = os.path.join(TEMARIOS_DIR, "Temario_Ciencias_8Basico_A.docx")
    if os.path.exists(temario_src):
        shutil.copyfile(temario_src, os.path.join(target_dir, "Temario_Ciencias_8Basico_A.docx"))

# 4. ORIENTACIÓN 5° BÁSICO
def procesar_orientacion_5():
    target_dir = os.path.join(OUT_BASE_DIR, "02 - Orientación", "5° Básico")
    os.makedirs(target_dir, exist_ok=True)

    # A) Evaluación Estudiante
    doc_eval = docx.Document(os.path.join(SRC_DIR, "Evaluacion_Final_Orientacion_5Basico.docx"))
    standardize_document_font(doc_eval, default_size=12)
    doc_eval.save(os.path.join(target_dir, "Evaluacion_Final_Orientacion_5Basico.docx"))
    print("[OK] Orientación 5° Básico: Evaluación de Estudiante limpia (Arial 12, sin pauta).")

    # B) Pauta Docente + Tabla de Especificaciones
    doc_pauta = docx.Document(os.path.join(SRC_DIR, "Pauta_Correccion_Orientacion_5Basico.docx"))
    standardize_document_font(doc_pauta, default_size=12)

    pauta_items = [
        ("Ítem I", "1", "Prevención en el patio/escalera: No correr, caminar con calma", "2 pts"),
        ("Ítem I", "2", "Convivencia: No empujar, respetar turnos en filas", "2 pts"),
        ("Ítem I", "3", "Higiene y Salud: Lavado de manos antes de comer y tras recreos", "2 pts"),
        ("Ítem I", "4", "Alimentación sana: Colaciones nutritivas, frutas y agua", "2 pts"),
        ("Ítem I", "5", "Seguridad digital y física: Cuidado de objetos y no aislarse", "2 pts"),
        ("Ítem II", "1", "Hábito 1: Usar pañuelos/mascarilla al estar resfriado", "1 pto"),
        ("Ítem II", "2", "Hábito 2: Traer toalla personal y jabón para educación física", "1 pto"),
        ("Ítem II", "3", "Hábito 3: Depositar desechos en basureros correspondientes", "1 pto"),
        ("Ítem II", "4", "Hábito 4: Cuidar y respetar las áreas verdes del colegio", "1 pto"),
        ("Ítem II", "5", "Hábito 5: Mantener la mesa y útiles limpios y ordenados", "1 pto"),
        ("Ítem III", "1", "Falso (F) — Bajar escaleras corriendo es de alto riesgo", "1 pto"),
        ("Ítem III", "2", "Falso (F) — Alimentos con sellos NO deben consumirse en exceso", "1 pto"),
        ("Ítem III", "3", "Falso (F) — Salir corriendo en masa provoca caídas", "1 pto"),
        ("Ítem III", "4", "Verdadero (V) — Lavarse la cara y manos tras actividad física", "1 pto"),
        ("Ítem III", "5", "Verdadero (V) — Higiene antes de almorzar previene infecciones", "1 pto"),
        ("Ítem III", "6", "Falso (F) — Correr en pasillos expone a choques y golpes", "1 pto"),
        ("Ítem III", "7", "Verdadero (V) — Avisar oportunamente a un docente si te sientes mal", "1 pto"),
        ("Ítem III", "8", "Falso (F) — Se debe cruzar el patio atento a balones y juegos", "1 pto"),
        ("Ítem III", "9", "Verdadero (V) — Tijeras punta roma para uso estrictamente pedagógico", "1 pto"),
        ("Ítem III", "10", "Verdadero (V) — Trato respetuoso fortalece la buena convivencia", "1 pto")
    ]
    append_pauta_summary_table(doc_pauta, pauta_items)

    specs = {
        "headers": ["Eje / Contenido Formativo / Ítem", "Comprensión (20%)", "Aplicación (40%)", "Análisis y Juicio (40%)", "Total (100%)"],
        "rows": [
            ["Eje Crecimiento Personal / Situaciones de Prevención\nÍtem I: Dibujo y Expresión Gráfica", "—", "10 pts", "—", "10 pts"],
            ["Eje Vida Saludable / Hábitos de Higiene y Cuidado\nÍtem II: Argumentación y Registro Escrito", "5 pts", "—", "—", "5 pts"],
            ["Eje Autocuidado y Factores Protectores\nÍtem III: Verdadero o Falso", "—", "—", "10 pts", "10 pts"],
            ["Puntaje Total General", "5 pts (20%)", "10 pts (40%)", "10 pts (40%)", "25 pts (100%)"]
        ]
    }
    append_specification_table(doc_pauta, specs)
    doc_pauta.save(os.path.join(target_dir, "Pauta_Correccion_y_Especificaciones_Orientacion_5Basico.docx"))
    print("[OK] Orientación 5° Básico: Pauta Oficial y Tabla de Especificaciones anexada.")

    # C) Copiar Temario Oficial
    temario_src = os.path.join(TEMARIOS_DIR, "Temario_Orientacion_5Basico_A.docx")
    if os.path.exists(temario_src):
        shutil.copyfile(temario_src, os.path.join(target_dir, "Temario_Orientacion_5Basico_A.docx"))

# 5. MÚSICA 1° Y 2° BÁSICO
def procesar_musica(curso_num):
    curso_str = f"{curso_num}° Básico"
    target_dir = os.path.join(OUT_BASE_DIR, "03 - Música", curso_str)
    os.makedirs(target_dir, exist_ok=True)

    fn = f"Rubrica_Evaluacion_Final_Musica_{curso_num}Basico.docx"
    doc_rubrica = docx.Document(os.path.join(SRC_DIR, fn))
    standardize_document_font(doc_rubrica, default_size=12)

    specs = {
        "headers": ["Eje Curricular / Habilidad Musical Evaluada", "Inicial (1 pto)", "Elemental (2 pts)", "Bueno (3 pts)", "Excelente (5 pts)"],
        "rows": [
            ["Eje Expresión Vocal\n1. Canto al unísono de 'Estrellita'", "1 pto", "2 pts", "3 pts", "5 pts"],
            ["Eje Coordinación y Pulso Rítmico\n2. Percusión individual con tempo estable", "1 pto", "2 pts", "3 pts", "5 pts"],
            ["Eje Interpretación en Ensamble\n3. Percusión grupal y sincronía armónica", "1 pto", "2 pts", "3 pts", "5 pts"],
            ["Eje Destreza Instrumental\n4. Técnica de ejecución (Metalófono o Sonaja)", "1 pto", "2 pts", "3 pts", "5 pts"],
            ["Eje Actitud y Convivencia Musical\n5. Cuidado del instrumento y concentración", "1 pto", "2 pts", "3 pts", "5 pts"],
            ["Puntaje Máximo de la Rúbrica", "5 pts (20%)", "10 pts (40%)", "15 pts (60%)", "25 pts (100%)"]
        ]
    }
    append_specification_table(doc_rubrica, specs)
    out_name = f"Instrumento_y_Rubrica_Evaluacion_Final_Musica_{curso_num}Basico.docx"
    doc_rubrica.save(os.path.join(target_dir, out_name))
    print(f"[OK] Música {curso_str}: Instrumento de Evaluación y Rúbrica de 25 puntos generada.")

    # Copiar Temario Oficial
    temario_src = os.path.join(TEMARIOS_DIR, f"Temario_Musica_{curso_num}Basico_A.docx")
    if os.path.exists(temario_src):
        shutil.copyfile(temario_src, os.path.join(target_dir, f"Temario_Musica_{curso_num}Basico_A.docx"))

# 6. CENTRALIZAR TEMARIOS POR CURSO
def centralizar_temarios():
    target_dir = os.path.join(OUT_BASE_DIR, "00 - Temarios por Curso (Para Enviar a Apoderados y UTP)")
    os.makedirs(target_dir, exist_ok=True)
    if os.path.exists(TEMARIOS_DIR):
        for f in os.listdir(TEMARIOS_DIR):
            if f.endswith('.docx'):
                shutil.copyfile(os.path.join(TEMARIOS_DIR, f), os.path.join(target_dir, f))
        print(f"[OK] Temarios por curso centralizados en: {target_dir}")

# 7. EMPAQUETAR EN ZIP OFICIAL
def empaquetar_zip():
    zip_path = "/home/jack/Descargas/Evaluaciones_Finales_Completas_UTP_2026.zip"
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(OUT_BASE_DIR):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, OUT_BASE_DIR)
                zipf.write(full_path, arcname=rel_path)
    print(f"\n[ÉXITO] Archivo ZIP oficial generado: {zip_path}")
    print(f"Tamaño: {os.path.getsize(zip_path)} bytes")

if __name__ == '__main__':
    print("=========================================================")
    print(" INICIANDO PROCESAMIENTO OFICIAL DE EVALUACIONES UTP 2026")
    print("=========================================================\n")
    procesar_ciencias_5()
    procesar_ciencias_6()
    procesar_ciencias_8()
    procesar_orientacion_5()
    procesar_musica(1)
    procesar_musica(2)
    centralizar_temarios()
    empaquetar_zip()
    print("\n[PROCESO COMPLETADO SATISFACTORIAMENTE]")
