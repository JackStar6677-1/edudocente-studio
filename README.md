# 🎓 EduDocente-Studio
### *Automated Pedagogical Assessment & Curriculum Framework*

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![python-docx](https://img.shields.io/badge/Document_Engine-python--docx-2B579A?style=flat-square&logo=microsoftword&logoColor=white)](https://python-docx.readthedocs.io/)
[![OpenPyXL](https://img.shields.io/badge/Spreadsheets-openpyxl-217346?style=flat-square&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Matplotlib](https://img.shields.io/badge/Vector_Graphics-Matplotlib-11557c?style=flat-square)](https://matplotlib.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Production_Ready-brightgreen?style=flat-square)](#)

> **EduDocente-Studio** es un motor de software declarativo diseñado para transformar objetivos curriculares (Mineduc) y contenidos pedagógicos en **paquetes evaluativos completos y listos para imprimir en Word (.docx) y Excel (.xlsx)**, bajo rigurosos estándares de diseño institucional, lenguaje formativo y consistencia visual.

---

## 💡 El Problema y la Solución

### El Desafío
Los docentes de educación básica y media dedican decenas de horas semanales a diseñar pruebas, elaborar solucionarios con justificaciones, redactar temarios para apoderados y cuadrar planillas de entrevistas. La mayoría de estas tareas se realizan en procesadores de texto manuales, provocando:
* Desalineaciones tipográficas y estéticas.
* Filtración involuntaria de respuestas o pistas en los enunciados.
* Falta de justificaciones pedagógicas claras para la retroalimentación.
* Duplicación innecesaria de trabajo entre la versión del alumno, la pauta docente y el temario a familias.

### Nuestra Solución
**EduDocente-Studio** estandariza este flujo de trabajo mediante una **arquitectura de 3 ítems universales**:
1. **Ítem I (Selección Múltiple):** Alternativas limpias `A)`, `B)`, `C)`, `D)` sin casillas distractoras.
2. **Ítem II (Verdadero o Falso):** Casillas `(   )` para el alumno y tabla de justificaciones pedagógicas con respuestas en verde (`#1E7E34`) para el profesor.
3. **Ítem III (Aplicación Práctica / Dibujo / Esquema Técnico):** Marcos delimitados para dibujar con líneas de explicación, o láminas vectoriales tipo esqueleto para rotular y pintar.

---

## 🏛️ Arquitectura del Sistema

```mermaid
flowchart TD
    subgraph INPUT ["Entrada Curricular"]
        JSON["Archivo Declarativo (.json)"]
        WIZ["Asistente CLI Interactivo (wizard.py)"]
        PROMPT["Temas y Objetivos Mineduc (OA)"]
    end

    subgraph ENGINE ["EduDocente Core Engine"]
        DE["DocenteEngine (engine.py)"]
        GC["Motor de Estilos (generator_core.py)"]
        VEC["Generador Vectorial (Matplotlib Assets)"]
    end

    subgraph DELIVERABLES ["Entregables Institucionales"]
        EST["Prueba del Estudiante (.docx)\n- Sin pistas ni escalas de conversión\n- Espacios limpios de dibujo"]
        PAU["Pauta Oficial Docente (.docx)\n- Solucionario destacado en verde\n- Justificaciones y modelos resueltos"]
        TEM["Temario e Informativo a Familias (.docx)\n- Páginas del texto escolar\n- Sugerencias de estudio en el hogar"]
        XLS["Planilla de Gestión Escolar (.xlsx)\n- Cronograma de entrevistas\n- Formato condicional y anchos óptimos"]
    end

    JSON --> DE
    WIZ --> DE
    PROMPT --> DE
    VEC --> GC
    GC --> DE

    DE --> EST
    DE --> PAU
    DE --> TEM
    DE --> XLS
```

---

## ✨ Características Principales

* 🎨 **Línea Gráfica Institucional Impecable:**
  - Tipografía unificada `Arial`.
  - Paleta de color armónica: Azul Marino Institucional (`#173F73`), Azul Cielo (`#EAF3FB`), Fondo Tarjetas (`#F6FAFE`) y Borde Acero (`#B0C4DE`).
  - Escudo institucional integrado de alta fidelidad.
* ⚖️ **Rigor Pedagógico Integrado:**
  - Eliminación de etiquetas distractoras (`SUMATIVA`, tiempos límite que generan ansiedad, porcentajes de exigencia que no corresponden al alumno).
  - Respuestas docentes 100% justificadas con citas al texto escolar oficial.
* 📐 **Generación Vectorial de Recursos Didácticos:**
  - **Circuitos Eléctricos:** Símbolos esquemáticos normalizados (fuente/pila, interruptor abierto/cerrado, ampolleta circular con cruz, cable).
  - **Modelo Corpuscular:** Distribución de partículas en estados sólido, líquido y gaseoso.
  - **Estructura Atómica:** Modelo atómico con órbitas elípticas, corteza, núcleo, protones, neutrones y electrones.
* 📦 **Despliegue Multi-Directorio:** Salida sincronizada localmente en el repositorio Git y en las carpetas de gestión docente de Descargas.

---

## 📁 Estructura del Proyecto

```text
evaluaciones_pasteur/
│
├── assets/                               # Recursos gráficos vectoriales e insignias
│   ├── logo_colegio.png                  # Escudo oficial de la institución
│   ├── circuit_components/               # Símbolos esquemáticos de circuitos eléctricos
│   ├── matter_states/                    # Modelos corpusculares de la materia
│   └── atom_model/                       # Esqueleto y modelo atómico resuelto
│
├── output/                               # Entregables compilados (.docx y .xlsx)
│   ├── temarios/                         # Temarios oficiales para enviar por correo
│   ├── Evaluacion_Final_Ciencias_5Basico.docx
│   ├── Pauta_Correccion_Ciencias_5Basico.docx
│   ├── Evaluacion_Final_Ciencias_6Basico.docx
│   ├── Pauta_Correccion_Ciencias_6Basico.docx
│   ├── Evaluacion_Final_Ciencias_8Basico.docx
│   ├── Pauta_Correccion_Ciencias_8Basico.docx
│   ├── Rubrica_Evaluacion_Final_Musica_1Basico.docx
│   ├── Rubrica_Evaluacion_Final_Musica_2Basico.docx
│   ├── Evaluacion_Final_Orientacion_5Basico.docx
│   ├── Pauta_Correccion_Orientacion_5Basico.docx
│   └── Cronograma_Entrevistas_Apoderados_2026.xlsx
│
├── engine.py                             # Motor declarativo central (DocenteEngine)
├── wizard.py                             # Asistente CLI interactivo para crear nuevas pruebas
├── generator_core.py                     # Motor base de renderizado XML y estilos python-docx
├── main.py                               # Orquestador del sistema con menú interactivo y flags
│
├── build_temarios.py                     # Compilador unificado de los 6 temarios a apoderados
├── build_student_test.py                 # Generador Evaluación 5° Básico (OA 11 - 28 pts)
├── build_teacher_answer_key.py           # Generador Pauta 5° Básico
├── build_student_test_6basico.py         # Generador Evaluación 6° Básico (OA 13 - 26 pts)
├── build_teacher_answer_key_6basico.py   # Generador Pauta 6° Básico
├── build_ciencias_8basico.py             # Generador Evaluación y Pauta 8° Básico (OA 3 - 25 pts)
├── build_rubricas_musica.py              # Generador Rúbricas Música 1°A y 2°A (25 pts c/u)
├── build_orientacion_5basico.py          # Generador Evaluación y Pauta Orientación 5°A (OA 5 - 25 pts)
├── crear_excel_entrevistas.py            # Generador de Planilla Excel de Entrevistas
│
├── generate_atom_assets.py               # Renderizador gráfico del átomo (Matplotlib)
├── generate_matter_states.py             # Renderizador de estados de la materia
├── perfect_symbols.py                    # Renderizador de circuitos normalizados
│
├── examples/                             # Plantillas JSON de evaluaciones declarativas
│   └── evaluacion_modelo.json            # Plantilla base para nuevas evaluaciones
├── requirements.txt                      # Dependencias de Python
├── LICENSE                               # Licencia MIT
└── README.md                             # Documentación del proyecto
```

---

## 🚀 Instalación y Puesta en Marcha

### 1. Clonar el repositorio
```bash
git clone https://github.com/tu-usuario/edudocente-studio.git
cd edudocente-studio
```

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

---

## 💻 Modos de Uso

### Opción 1: Asistente Interactivo de Creación (Wizard)
Permite construir una evaluación desde cero guiado por preguntas en la terminal:
```bash
python wizard.py
```

### Opción 2: Compilación Declarativa desde JSON
Define tus preguntas, alternativas, oraciones V/F y actividades en un archivo `.json`:
```python
from engine import DocenteEngine

engine = DocenteEngine("examples/evaluacion_modelo.json")
prueba, pauta = engine.build_all()
```

### Opción 3: Orquestador General del Colegio
Genera cualquier evaluación preconfigurada o todas las materias en lote:
```bash
# Menú interactivo en pantalla
python main.py

# O mediante banderas directas por consola
python main.py --all            # Genera todo el material docente
python main.py --ciencias8      # Genera evaluación y pauta de Ciencias 8° Básico
python main.py --musica         # Genera rúbricas prácticas de Música 1° y 2° Básico
python main.py --orientacion    # Genera prueba y pauta de Orientación 5° Básico
python main.py --temarios       # Genera los 6 temarios para enviar a apoderados
python main.py --excel          # Genera la planilla Excel de entrevistas
```

---

## 📋 Estructura de Configuración Declarativa (JSON)

```json
{
  "colegio": "COLEGIO LUIS PASTEUR ANEXO",
  "asignatura": "CIENCIAS NATURALES",
  "curso": "7° Básico A",
  "titulo": "EVALUACIÓN FINAL: MICROORGANISMOS Y BACTERIAS",
  "oa": "OA 7 — Investigar y explicar las características de virus y bacterias.",
  "contenidos": "Microorganismos patógenos y benéficos, estructura y prevención.",
  "puntaje_total": 25,
  
  "item1_seleccion_multiple": [
    {
      "pregunta": "1. ¿Qué estructura celular es propia de las bacterias?",
      "alternativas": [
        ["A", "Pared celular y material genético libre en el citoplasma."],
        ["B", "Núcleo delimitado por membrana carioteca."],
        ["C", "Cápsula de cristal inorgánico."],
        ["D", "Ausencia total de ribosomas."]
      ],
      "correcta": "A) Pared celular y material genético libre en el citoplasma.",
      "justificacion": "Las bacterias son organismos procariontes sin núcleo organizado."
    }
  ],

  "item2_verdadero_falso": [
    {
      "oracion": "Los virus son considerados células vivas con metabolismo independiente.",
      "resp": "F",
      "justificacion": "Los virus son agentes acelulares que requieren una célula hospedera para replicarse."
    }
  ],

  "item3_aplicacion": {
    "titulo": "ÍTEM III: DIBUJO Y EXPLICACIÓN DE MEDIDAS PREVENTIVAS",
    "puntaje": 5,
    "tipo": "drawing_boxes",
    "instruccion": "Dibuja en el recuadro una medida de prevención sanitaria y descríbela:",
    "cajas": [
      {
        "titulo": "Medida: Lavado adecuado de manos con jabón",
        "ejemplo_pauta": "Estudiante lavando manos con agua y jabón disolviendo la cubierta viral."
      }
    ]
  }
}
```

---

## 🏆 Materiales Docentes Listos para Producción

El proyecto incluye casos reales listos para aula generados para el **Colegio Luis Pasteur Anexo**:

| Asignatura | Curso | Instrumento Evaluativo | Puntaje | Entregables |
| :--- | :--- | :--- | :---: | :--- |
| **Ciencias Naturales** | 5° Básico A-B | Energía eléctrica y circuitos | 28 pts | Prueba, Pauta y Temario |
| **Ciencias Naturales** | 6° Básico A-B | Cambios de estado y partículas | 26 pts | Prueba, Pauta y Temario |
| **Ciencias Naturales** | 8° Básico A | Teoría atómica de Dalton y átomo | 25 pts | Prueba, Pauta con modelo resuelto y Temario |
| **Música** | 1° Básico A | Canto al unísono y percusión (*Estrellita*) | 25 pts | Rúbrica práctica de aula y Temario |
| **Música** | 2° Básico A | Interpretación coral y metalófono | 25 pts | Rúbrica práctica de aula y Temario |
| **Orientación** | 5° Básico A | Prevención de drogas y autocuidado | 25 pts | Prueba, Pauta con justificaciones y Temario |
| **Gestión Docente** | Jefatura / Asignatura | Cronograma de entrevistas a apoderados | — | Planilla Excel automatizada (16 apoderados) |

---

## 👨‍💻 Autor y Reconocimientos

- **Desarrollo y Arquitectura de Software:** Jack ([GitHub Profile](https://github.com))
- **Asesoría y Validación Pedagógica:** Profesora Margarita Miranda B. (*C.E.P. Luis Pasteur Anexo*)
- **Contacto:** `profesora.margaritamiranda@cepluispasteur.cl`

---

## 📄 Licencia

Este proyecto está liberado bajo la [Licencia MIT](LICENSE). Siéntete libre de utilizarlo, bifurcarlo y adaptarlo para tu propia institución educativa.
