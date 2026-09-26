"""
Sistema Automatizado de Generación Docente
Colegio Luis Pasteur Anexo - Profesora Margarita Miranda B.
Repositorio Local de Evaluaciones, Pautas, Temarios y Planillas Excel.
"""

import os
import sys
import shutil
import argparse

# Asegurar codificación utf-8 en consola de Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
DOWNLOADS_ROOT = r"C:\Users\Jack\Downloads\Evaluaciones Finales Profesora Margarita"

# Importar generadores
from build_student_test import build_evaluacion_estudiante
from build_teacher_answer_key import build_pauta_correccion
from build_comunicado_temario import build_comunicado

from build_student_test_6basico import build_evaluacion_estudiante_6b
from build_teacher_answer_key_6basico import build_pauta_correccion_6b
from build_comunicado_temario_6basico import build_comunicado_6b

from crear_excel_entrevistas import generar_cronograma_excel
from perfect_symbols import *  # Circuit symbols generator
from generate_matter_states import *  # Matter states generator

def deploy_file(filename, target_subpath):
    src = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(src):
        return
    dst_dir = os.path.join(DOWNLOADS_ROOT, target_subpath)
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, filename)
    try:
        shutil.copy2(src, dst)
        print(f"  [OK] Copiado a Descargas: {dst}")
    except PermissionError:
        print(f"  [AVISO] No se pudo sobrescribir '{filename}' en Descargas porque est\u00e1 abierto en Word/Excel.")
    except Exception as e:
        print(f"  [ERROR] Al desplegar '{filename}': {e}")

def run_ciencias_5basico():
    print("\n--- Generando Ciencias Naturales 5° Básico (OA 11 - 28 pts) ---")
    build_evaluacion_estudiante()
    build_pauta_correccion()
    build_comunicado()
    dest = os.path.join("01 - Ciencias Naturales", "5° Básico (5°A y 5°B)")
    deploy_file("Evaluacion_Final_Ciencias_5Basico.docx", dest)
    deploy_file("Pauta_Correccion_Ciencias_5Basico.docx", dest)
    deploy_file("Temario_Evaluacion_Final_Ciencias_5Basico.docx", dest)
    print("[COMPLETADO] Evaluación, Pauta y Temario de 5° Básico generados.")

def run_ciencias_6basico():
    print("\n--- Generando Ciencias Naturales 6° Básico (OA 13 - 26 pts) ---")
    build_evaluacion_estudiante_6b()
    build_pauta_correccion_6b()
    build_comunicado_6b()
    dest = os.path.join("01 - Ciencias Naturales", "6° Básico (6°A y 6°B)")
    deploy_file("Evaluacion_Final_Ciencias_6Basico.docx", dest)
    deploy_file("Pauta_Correccion_Ciencias_6Basico.docx", dest)
    deploy_file("Temario_Evaluacion_Final_Ciencias_6Basico.docx", dest)
    print("[COMPLETADO] Evaluación, Pauta y Temario de 6° Básico generados.")

def run_excel_entrevistas():
    print("\n--- Generando Cronograma de Entrevistas de Apoderados 2026 ---")
    generar_cronograma_excel()
    dest = os.path.join("03 - Registro y Gestión Docente", "Nóminas y Listas")
    deploy_file("Cronograma_Entrevistas_Apoderados_2026.xlsx", dest)
    print("[COMPLETADO] Planilla Excel de entrevistas generada.")

def run_assets():
    print("\n--- Regenerando Gr\u00e1ficos y Recursos Visuales ---")
    import perfect_symbols
    import generate_matter_states
    print("[COMPLETADO] S\u00edmbolos el\u00e9ctricos y modelos de part\u00edculas regenerados en assets/.")

def run_all():
    print("\n=======================================================")
    print(" EJECUTANDO GENERACI\u00d3N COMPLETA DE RECURSOS DOCENTES ")
    print("=======================================================")
    run_assets()
    run_ciencias_5basico()
    run_ciencias_6basico()
    run_excel_entrevistas()
    print("\n=======================================================")
    print(" \u00a1TODO EL MATERIAL FUE GENERADO Y ACTUALIZADO CON \u00c9XITO! ")
    print("=======================================================\n")

def menu_interactivo():
    while True:
        print("\n" + "=" * 55)
        print("  SISTEMA DOCENTE LUIS PASTEUR - PROFESORA MARGARITA  ")
        print("=" * 55)
        print("1. Generar 5\u00b0 B\u00e1sico (Prueba, Pauta, Temario de Ciencias)")
        print("2. Generar 6\u00b0 B\u00e1sico (Prueba, Pauta, Temario de Ciencias)")
        print("3. Generar Planilla Excel de Entrevistas a Apoderados")
        print("4. Regenerar Recursos Gr\u00e1ficos (Circuitos y Part\u00edculas)")
        print("5. Generar y Desplegar TODO a Descargas")
        print("0. Salir")
        print("-" * 55)
        opcion = input("Selecciona una opci\u00f3n (0-5): ").strip()

        if opcion == "1":
            run_ciencias_5basico()
        elif opcion == "2":
            run_ciencias_6basico()
        elif opcion == "3":
            run_excel_entrevistas()
        elif opcion == "4":
            run_assets()
        elif opcion == "5":
            run_all()
        elif opcion == "0":
            print("Cerrando el sistema docente. \u00a1Hasta pronto!")
            break
        else:
            print("Opci\u00f3n inv\u00e1lida. Por favor, ingresa un n\u00famero del 0 al 5.")

def main():
    parser = argparse.ArgumentParser(description="Generador de Material Docente Colegio Luis Pasteur")
    parser.add_argument("--all", action="store_true", help="Genera todas las evaluaciones, pautas, temarios y Excel")
    parser.add_argument("--ciencias5", action="store_true", help="Genera material de 5to b\u00e1sico")
    parser.add_argument("--ciencias6", action="store_true", help="Genera material de 6to b\u00e1sico")
    parser.add_argument("--excel", action="store_true", help="Genera el Excel de entrevistas de apoderados")
    parser.add_argument("--assets", action="store_true", help="Regenera las im\u00e1genes y s\u00edmbolos gr\u00e1ficos")

    args = parser.parse_args()

    if args.all:
        run_all()
    elif args.ciencias5:
        run_ciencias_5basico()
    elif args.ciencias6:
        run_ciencias_6basico()
    elif args.excel:
        run_excel_entrevistas()
    elif args.assets:
        run_assets()
    else:
        menu_interactivo()

if __name__ == "__main__":
    main()
