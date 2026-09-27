"""
EduDocente Wizard: Asistente Interactivo de Creación de Evaluaciones
Colegio Luis Pasteur Anexo - Framework Estandarizado
"""

import os
import sys
import json
from engine import DocenteEngine

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def crear_evaluacion_interactiva():
    print("\n" + "=" * 65)
    print("     ASISTENTE DE CREACIÓN DE EVALUACIONES ESCOLARES      ")
    print("      Estructura Estandarizada: Ítem I + Ítem II + Ítem III")
    print("=" * 65)

    print("\n[1] Ingresa los datos generales de la evaluación:")
    colegio = input("  • Nombre del Colegio [Por defecto: COLEGIO LUIS PASTEUR ANEXO]: ").strip() or "COLEGIO LUIS PASTEUR ANEXO"
    asignatura = input("  • Asignatura (ej: CIENCIAS NATURALES, HISTORIA, LENGUAJE): ").strip() or "CIENCIAS NATURALES"
    curso = input("  • Curso (ej: 7° Básico A): ").strip() or "7° Básico A"
    titulo = input("  • Título de la evaluación (ej: EVALUACIÓN FINAL DE LA MATERIA): ").strip() or f"EVALUACIÓN FINAL DE {asignatura.upper()}"
    oa = input("  • Objetivo de Aprendizaje (OA): ").strip()
    contenidos = input("  • Contenidos evaluados (y páginas del libro): ").strip()

    config = {
        "colegio": colegio,
        "asignatura": asignatura,
        "curso": curso,
        "titulo": titulo,
        "oa": oa,
        "contenidos": contenidos,
        "puntaje_total": 25,
        "docente": "Profesora Margarita Miranda B.",
        "email_docente": "profesora.margaritamiranda@cepluispasteur.cl",
        "item1_seleccion_multiple": [],
        "item2_verdadero_falso": [],
        "item3_aplicacion": {}
    }

    print("\n" + "-" * 65)
    print("[2] ÍTEM I: SELECCIÓN MÚLTIPLE (Formato A, B, C, D)")
    print("    ¿Cuántas preguntas de selección múltiple deseas ingresar? (Enter para 3 por defecto)")
    n_i1_str = input("    Cantidad: ").strip()
    n_i1 = int(n_i1_str) if n_i1_str.isdigit() else 3

    for i in range(1, n_i1 + 1):
        print(f"\n  --- Pregunta {i} ---")
        q = input(f"  Enunciado de la pregunta {i}: ").strip() or f"Pregunta modelo número {i} de la evaluación."
        a = input("    Alternativa A: ").strip() or "Opción modelo A"
        b = input("    Alternativa B: ").strip() or "Opción modelo B"
        c = input("    Alternativa C: ").strip() or "Opción modelo C"
        d = input("    Alternativa D: ").strip() or "Opción modelo D"
        corr = input("    ¿Cuál es la letra correcta? (A/B/C/D): ").strip().upper() or "A"
        just = input("    Justificación para la pauta docente: ").strip() or "Respuesta correcta conforme a los contenidos del texto de estudio."

        config["item1_seleccion_multiple"].append({
            "pregunta": f"{i}. {q}",
            "alternativas": [["A", a], ["B", b], ["C", c], ["D", d]],
            "correcta": f"{corr}) {dict(A=a, B=b, C=c, D=d).get(corr, a)}",
            "justificacion": just
        })

    print("\n" + "-" * 65)
    print("[3] ÍTEM II: VERDADERO O FALSO (con justificación en pauta)")
    print("    ¿Cuántas afirmaciones V/F deseas ingresar? (Enter para 4 por defecto)")
    n_i2_str = input("    Cantidad: ").strip()
    n_i2 = int(n_i2_str) if n_i2_str.isdigit() else 4

    for i in range(1, n_i2 + 1):
        print(f"\n  --- Afirmación {i} ---")
        oracion = input(f"  Oración {i}: ").strip() or f"Afirmación curricular evaluativa número {i}."
        resp = input("    ¿Es Verdadera (V) o Falsa (F)?: ").strip().upper() or "V"
        just = input("    Justificación pedagógica: ").strip() or ("Es verdadera conforme a la materia vista en clases." if resp == "V" else "Es falsa, lo correcto es lo explicado en el texto escolar.")

        config["item2_verdadero_falso"].append({
            "oracion": oracion,
            "resp": resp,
            "justificacion": just
        })

    print("\n" + "-" * 65)
    print("[4] ÍTEM III: DIBUJO, ESQUEMA O APLICACIÓN PRÁCTICA")
    inst3 = input("  Instrucción del ítem (ej: Dibuja y rotula las partes...): ").strip() or "Dibuja en cada rectángulo una situación o concepto trabajado en clases y describe tu dibujo:"
    caja1 = input("  Título del primer recuadro de dibujo: ").strip() or "1. Situación o concepto principal estudiado:"
    pauta1 = input("  Respuesta o ejemplo esperado para la pauta docente: ").strip() or "Representación gráfica clara del contenido con elementos centrales."

    config["item3_aplicacion"] = {
        "titulo": "ÍTEM III: DIBUJO Y APLICACIÓN PRÁCTICA",
        "puntaje": 5,
        "tipo": "drawing_boxes",
        "instruccion": inst3,
        "cajas": [
            {"titulo": caja1, "ejemplo_pauta": pauta1}
        ]
    }

    # Guardar configuración temporal en JSON
    json_path = os.path.join(BASE_DIR, "output", f"config_temp_{asignatura[:4]}_{curso[:3]}.json")
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    print(f"\n[OK] Configuración guardada en: {json_path}")

    # Compilar con el motor
    print("\n[COMPILANDO DOCUMENTOS DOCENTES...]")
    engine = DocenteEngine(config)
    p_est = engine.build_evaluacion_estudiante()
    p_pau = engine.build_pauta_correccion()

    print("\n" + "=" * 65)
    print(" ¡EVALUACIÓN Y PAUTA GENERADAS EXITOSAMENTE!")
    print(f"  • Prueba Estudiante: {p_est}")
    print(f"  • Pauta Docente:     {p_pau}")
    print("=" * 65 + "\n")

def menu_wizard():
    print("\n" + "=" * 60)
    print("     EDUDOCENTE STUDIO — GENERADOR AUTOMÁTICO DE PRUEBAS     ")
    print("=" * 60)
    print("1. Crear una nueva evaluación paso a paso (Asistente Interactivo)")
    print("2. Compilar evaluación desde un archivo JSON (ej: examples/evaluacion_modelo.json)")
    print("0. Volver")
    print("-" * 60)
    op = input("Selecciona una opción (0-2): ").strip()

    if op == "1":
        crear_evaluacion_interactiva()
    elif op == "2":
        path = input("Ingresa la ruta del archivo JSON [Enter para examples/evaluacion_modelo.json]: ").strip()
        if not path:
            path = os.path.join(BASE_DIR, "examples", "evaluacion_modelo.json")
        if os.path.exists(path):
            engine = DocenteEngine(path)
            engine.build_all()
        else:
            print(f"Error: no se encontró el archivo en {path}")
    elif op == "0":
        return

if __name__ == "__main__":
    menu_wizard()
