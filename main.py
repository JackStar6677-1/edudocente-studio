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

from build_student_test_6basico import build_evaluacion_estudiante_6b
from build_teacher_answer_key_6basico import build_pauta_correccion_6b

from build_temarios import generar_todos_los_temarios
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
        print(f"  [AVISO] No se pudo sobrescribir '{filename}' en Descargas porque está abierto en Word/Excel.")
    except Exception as e:
        print(f"  [ERROR] Al desplegar '{filename}': {e}")

def run_ciencias_5basico():
    print("\n--- Generando Evaluación y Pauta de Ciencias Naturales 5° Básico (OA 11 - 28 pts) ---")
    build_evaluacion_estudiante()
    build_pauta_correccion()
    dest = os.path.join("01 - Ciencias Naturales", "5° Básico (5°A y 5°B)")
    deploy_file("Evaluacion_Final_Ciencias_5Basico.docx", dest)
    deploy_file("Pauta_Correccion_Ciencias_5Basico.docx", dest)
    print("[COMPLETADO] Evaluación y Pauta de 5° Básico generadas.")

def run_ciencias_6basico():
    print("\n--- Generando Evaluación y Pauta de Ciencias Naturales 6° Básico (OA 13 - 26 pts) ---")
    build_evaluacion_estudiante_6b()
    build_pauta_correccion_6b()
    dest = os.path.join("01 - Ciencias Naturales", "6° Básico (6°A y 6°B)")
    deploy_file("Evaluacion_Final_Ciencias_6Basico.docx", dest)
    deploy_file("Pauta_Correccion_Ciencias_6Basico.docx", dest)
    print("[COMPLETADO] Evaluación y Pauta de 6° Básico generadas.")

def run_temarios():
    generar_todos_los_temarios()

def run_excel_entrevistas():
    print("\n--- Generando Cronograma de Entrevistas de Apoderados 2026 ---")
    generar_cronograma_excel()
    dest = os.path.join("03 - Registro y Gestión Docente", "Nóminas y Listas")
    deploy_file("Cronograma_Entrevistas_Apoderados_2026.xlsx", dest)
    print("[COMPLETADO] Planilla Excel de entrevistas generada.")

def run_assets():
    print("\n--- Regenerando Gráficos y Recursos Visuales ---")
    import perfect_symbols
    import generate_matter_states
    print("[COMPLETADO] Símbolos eléctricos y modelos de partículas regenerados en assets/.")

def run_all():
    print("\n=======================================================")
    print(" EJECUTANDO GENERACIÓN COMPLETA DE RECURSOS DOCENTES ")
    print("=======================================================")
    run_assets()
    run_ciencias_5basico()
    run_ciencias_6basico()
    run_temarios()
    run_excel_entrevistas()
    print("\n=======================================================")
    print(" ¡TODO EL MATERIAL FUE GENERADO Y ACTUALIZADO CON ÉXITO! ")
    print("=======================================================\n")

def menu_interactivo():
    while True:
        print("\n" + "=" * 58)
        print("   SISTEMA DOCENTE LUIS PASTEUR - PROFESORA MARGARITA   ")
        print("=" * 58)
        print("1. Generar 5° Básico (Prueba y Pauta de Ciencias)")
        print("2. Generar 6° Básico (Prueba y Pauta de Ciencias)")
        print("3. Generar TODOS los Temarios por Curso (00 - Temarios)")
        print("4. Generar Planilla Excel de Entrevistas a Apoderados")
        print("5. Regenerar Recursos Gráficos (Circuitos y Partículas)")
        print("6. Generar y Desplegar TODO a Descargas")
        print("0. Salir")
        print("-" * 58)
        opcion = input("Selecciona una opción (0-6): ").strip()

        if opcion == "1":
            run_ciencias_5basico()
        elif opcion == "2":
            run_ciencias_6basico()
        elif opcion == "3":
            run_temarios()
        elif opcion == "4":
            run_excel_entrevistas()
        elif opcion == "5":
            run_assets()
        elif opcion == "6":
            run_all()
        elif opcion == "0":
            print("Cerrando el sistema docente. ¡Hasta pronto!")
            break
        else:
            print("Opción inválida. Por favor, ingresa un número del 0 al 6.")

def main():
    parser = argparse.ArgumentParser(description="Generador de Material Docente Colegio Luis Pasteur")
    parser.add_argument("--all", action="store_true", help="Genera todas las evaluaciones, pautas, temarios y Excel")
    parser.add_argument("--ciencias5", action="store_true", help="Genera material de 5to básico")
    parser.add_argument("--ciencias6", action="store_true", help="Genera material de 6to básico")
    parser.add_argument("--temarios", action="store_true", help="Genera todos los temarios por curso (00 - Temarios)")
    parser.add_argument("--excel", action="store_true", help="Genera el Excel de entrevistas de apoderados")
    parser.add_argument("--assets", action="store_true", help="Regenera las imágenes y símbolos gráficos")

    args = parser.parse_args()

    if args.all:
        run_all()
    elif args.ciencias5:
        run_ciencias_5basico()
    elif args.ciencias6:
        run_ciencias_6basico()
    elif args.temarios:
        run_temarios()
    elif args.excel:
        run_excel_entrevistas()
    elif args.assets:
        run_assets()
    else:
        menu_interactivo()

if __name__ == "__main__":
    main()
