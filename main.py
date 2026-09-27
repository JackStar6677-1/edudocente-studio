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

from build_ciencias_8basico import generar_ciencias_8basico_completa
from build_temarios import generar_todos_los_temarios
from build_rubricas_musica import generar_rubricas_musica
from build_orientacion_5basico import generar_orientacion_completa
from crear_excel_entrevistas import generar_cronograma_excel
from perfect_symbols import *  # Circuit symbols generator
from generate_matter_states import *  # Matter states generator
from generate_atom_assets import generate_student_atom_skeleton, generate_teacher_atom_solved

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

def run_ciencias_8basico():
    generar_ciencias_8basico_completa()

def run_musica():
    generar_rubricas_musica()

def run_orientacion():
    generar_orientacion_completa()

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
    generate_student_atom_skeleton()
    generate_teacher_atom_solved()
    print("[COMPLETADO] Símbolos eléctricos, partículas y modelos atómicos regenerados en assets/.")

def run_all():
    print("\n=======================================================")
    print(" EJECUTANDO GENERACIÓN COMPLETA DE RECURSOS DOCENTES ")
    print("=======================================================")
    run_assets()
    run_ciencias_5basico()
    run_ciencias_6basico()
    run_ciencias_8basico()
    run_musica()
    run_orientacion()
    run_temarios()
    run_excel_entrevistas()
    print("\n=======================================================")
    print(" ¡TODO EL MATERIAL FUE GENERADO Y ACTUALIZADO CON ÉXITO! ")
    print("=======================================================\n")

def run_ai():
    from ai_connector import interactive_ai_session
    interactive_ai_session()

def run_wizard():
    from wizard import interactive_wizard
    interactive_wizard()

def menu_interactivo():
    while True:
        print("\n" + "=" * 66)
        print("     SISTEMA DOCENTE LUIS PASTEUR — PROFESORA MARGARITA       ")
        print("=" * 66)
        print(" 1. Generar 5° Básico (Prueba y Pauta de Ciencias Naturales)")
        print(" 2. Generar 6° Básico (Prueba y Pauta de Ciencias Naturales)")
        print(" 3. Generar 8° Básico (Prueba y Pauta de Ciencias Naturales)")
        print(" 4. Generar Rúbricas de Música (1° Básico A y 2° Básico A)")
        print(" 5. Generar Evaluación y Pauta de Orientación (5° Básico A)")
        print(" 6. Generar TODOS los Temarios por Curso (00 - Temarios)")
        print(" 7. Generar Planilla Excel de Entrevistas a Apoderados")
        print(" 8. Regenerar Recursos Gráficos (Circuitos, Partículas, Átomo)")
        print(" 9. Generar y Desplegar TODO a Descargas")
        print("10. [IA] Crear Nueva Evaluación Asistida por IA (Gemini/Claude/GPT/DeepSeek)")
        print("11. [WIZARD] Asistente Paso a Paso Manual")
        print(" 0. Salir")
        print("-" * 66)
        opcion = input("Selecciona una opción (0-11): ").strip()

        if opcion == "1":
            run_ciencias_5basico()
        elif opcion == "2":
            run_ciencias_6basico()
        elif opcion == "3":
            run_ciencias_8basico()
        elif opcion == "4":
            run_musica()
        elif opcion == "5":
            run_orientacion()
        elif opcion == "6":
            run_temarios()
        elif opcion == "7":
            run_excel_entrevistas()
        elif opcion == "8":
            run_assets()
        elif opcion == "9":
            run_all()
        elif opcion == "10":
            run_ai()
        elif opcion == "11":
            run_wizard()
        elif opcion == "0":
            print("Cerrando el sistema docente. ¡Hasta pronto!")
            break
        else:
            print("Opción inválida. Por favor, ingresa un número del 0 al 11.")

def main():
    parser = argparse.ArgumentParser(description="Generador de Material Docente Colegio Luis Pasteur")
    parser.add_argument("--all", action="store_true", help="Genera todas las evaluaciones, pautas, temarios y Excel")
    parser.add_argument("--ciencias5", action="store_true", help="Genera material de Ciencias 5to básico")
    parser.add_argument("--ciencias6", action="store_true", help="Genera material de Ciencias 6to básico")
    parser.add_argument("--ciencias8", action="store_true", help="Genera material de Ciencias 8vo básico")
    parser.add_argument("--musica", action="store_true", help="Genera rúbricas de música 1°A y 2°A")
    parser.add_argument("--orientacion", action="store_true", help="Genera evaluación y pauta de Orientación 5°A")
    parser.add_argument("--temarios", action="store_true", help="Genera todos los temarios por curso (00 - Temarios)")
    parser.add_argument("--excel", action="store_true", help="Genera el Excel de entrevistas de apoderados")
    parser.add_argument("--assets", action="store_true", help="Regenera las imágenes y símbolos gráficos")
    parser.add_argument("--ai", action="store_true", help="Inicia sesión interactiva con el Conector Multi-IA")
    parser.add_argument("--wizard", action="store_true", help="Inicia el asistente paso a paso manual")

    args = parser.parse_args()

    if args.all:
        run_all()
    elif args.ciencias5:
        run_ciencias_5basico()
    elif args.ciencias6:
        run_ciencias_6basico()
    elif args.ciencias8:
        run_ciencias_8basico()
    elif args.musica:
        run_musica()
    elif args.orientacion:
        run_orientacion()
    elif args.temarios:
        run_temarios()
    elif args.excel:
        run_excel_entrevistas()
    elif args.assets:
        run_assets()
    elif args.ai:
        run_ai()
    elif args.wizard:
        run_wizard()
    else:
        menu_interactivo()

if __name__ == "__main__":
    main()
