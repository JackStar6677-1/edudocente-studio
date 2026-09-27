# 📚 Sistema de Generación Docente y Evaluaciones
### Colegio Luis Pasteur Anexo — Profesora Margarita Miranda B.

Repositorio automatizado en Python para el diseño, generación y despliegue de **Evaluaciones Finales**, **Pautas de Corrección**, **Temarios/Informativos a Apoderados** y **Planillas Excel de Gestión y Entrevistas**, bajo estrictos estándares pedagógicos y diseño institucional.

---

## 🏛️ Estándar de Diseño Institucional

Todos los documentos generados cumplen rigurosamente con la línea gráfica del Colegio:
- **Tipografía**: `Arial` uniforme en todos los textos, títulos y tablas.
- **Paleta de Colores**:
  - `Azul Marino Institucional`: `#173F73` (Encabezados, títulos principales y líneas maestras).
  - `Azul Cielo Suave`: `#EAF3FB` (Banners de datos del estudiante y subtítulos).
  - `Fondo Tarjetas / Cuadros`: `#F6FAFE` (Cuadros de instrucciones y tablas).
  - `Bordes`: `#B0C4DE` (Líneas finas y suaves de separación).
  - `Verde Solucionario`: `#1E7E34` (Respuestas destacadas en pautas de corrección).
- **Insignia Oficial**: Logotipo institucional nítido incorporado en el encabezado izquierdo.
- **Reglas Pedagógicas Aplicadas**:
  - Sin etiquetas innecesarias como `"SUMATIVA"`.
  - Sin escalas de conversión de notas ni porcentajes de exigencia en las pruebas y temarios.
  - Sin tiempos estimados ni advertencias de "orden y limpieza".
  - En alternativas: formato directo `A) `, `B) `, `C) `, `D) ` (sin casillas `[ ]`).
  - Sin pistas ni nombres dados en ítems de dibujo o identificación.
  - Lenguaje positivo y formativo en comunicados a apoderados.

---

## 📁 Estructura del Repositorio

```text
evaluaciones_pasteur/
│
├── assets/                               # Recursos gráficos institucionales
│   ├── logo_colegio.png                  # Escudo oficial del colegio
│   ├── circuit_components/               # Símbolos vectoriales de circuitos (5° Básico)
│   ├── matter_states/                    # Modelos corpusculares de partículas (6° Básico)
│   └── atom_model/                       # Modelos atómicos y esqueleto para rotular (8° Básico)
│
├── output/                               # Salida local de documentos generados
│   ├── temarios/                         # Temarios oficiales por curso (.docx)
│   │   ├── Temario_Ciencias_5Basico_A_B.docx
│   │   ├── Temario_Ciencias_6Basico_A_B.docx
│   │   ├── Temario_Ciencias_8Basico_A.docx
│   │   ├── Temario_Musica_1Basico_A.docx
│   │   ├── Temario_Musica_2Basico_A.docx
│   │   └── Temario_Orientacion_5Basico_A.docx
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
├── main.py                               # Menú interactivo y orquestador CLI
├── generator_core.py                     # Motor base de estilos Word (python-docx)
├── build_temarios.py                     # Generador unificado de los 6 temarios para apoderados
│
├── build_student_test.py                 # Generador Evaluación 5° Básico Ciencias (OA 11 - 28 pts)
├── build_teacher_answer_key.py           # Generador Pauta 5° Básico Ciencias
│
├── build_student_test_6basico.py         # Generador Evaluación 6° Básico Ciencias (OA 13 - 26 pts)
├── build_teacher_answer_key_6basico.py   # Generador Pauta 6° Básico Ciencias
│
├── build_ciencias_8basico.py             # Generador Evaluación y Pauta 8° Básico Ciencias (OA 3 - 25 pts)
├── build_rubricas_musica.py              # Generador Rúbricas Música 1°A y 2°A (25 pts c/u)
├── build_orientacion_5basico.py          # Generador Evaluación y Pauta Orientación 5°A (OA 5 - 25 pts)
│
├── crear_excel_entrevistas.py            # Generador Planilla Excel de Entrevistas
├── perfect_symbols.py                    # Generador de símbolos eléctricos con matplotlib
├── generate_matter_states.py             # Generador de estados de la materia con matplotlib
├── generate_atom_assets.py               # Generador de modelo y esqueleto del átomo con matplotlib
├── organizar_descargas.py                # Script de organización de carpetas
│
├── requirements.txt                      # Dependencias de Python
├── .gitignore                            # Exclusión de temporales y locks de Office
└── README.md                             # Documentación del proyecto
```

---

## 🚀 Requisitos e Instalación

Este proyecto utiliza librerías nativas de Python y genera archivos Word (`.docx`) y Excel (`.xlsx`) sin necesidad de instalar suites ofimáticas externas en la terminal.

1. **Clonar o abrir este directorio**:
   ```bash
   cd C:\Users\Jack\.gemini\antigravity\scratch\evaluaciones_pasteur
   ```

2. **Instalar dependencias**:
   ```bash
   pip install -r requirements.txt
   ```

---

## 💻 Uso y Ejecución

### 1. Menú Interactivo
Simplemente ejecuta:
```bash
python main.py
```
Aparecerá un menú con opciones numéricas para generar material específico o todo el conjunto.

### 2. Modo Línea de Comandos (CLI)
Puedes pasar argumentos directos:
- **Generar TODO y desplegar a Descargas**:
  ```bash
  python main.py --all
  ```
- **Solo 5° Básico**:
  ```bash
  python main.py --ciencias5
  ```
- **Solo 6° Básico**:
  ```bash
  python main.py --ciencias6
  ```
- **Solo Planilla de Entrevistas**:
  ```bash
  python main.py --excel
  ```
- **Regenerar Gráficos y Símbolos**:
  ```bash
  python main.py --assets
  ```

---

## 🗂️ Despliegue Automático

Cada vez que se ejecutan los scripts, los archivos no solo se guardan en la carpeta local `output/`, sino que también se actualizan automáticamente en la estructura ordenada de la Profesora Margarita en su carpeta de Descargas:

`C:\Users\Jack\Downloads\Evaluaciones Finales Profesora Margarita\`
- `01 - Ciencias Naturales\`
  - `5° Básico (5°A y 5°B)\`: Prueba, Pauta y Temario de 5° Básico.
  - `6° Básico (6°A y 6°B)\`: Prueba, Pauta y Temario de 6° Básico.
- `03 - Registro y Gestión Docente\`
  - `Nóminas y Listas\`: Cronograma de Entrevistas a Apoderados 2026.

---

## ✍️ Información de Autoría y Contacto Docente

- **Docente:** Profesora Margarita Miranda B.
- **Institución:** C.E.P. Luis Pasteur Anexo
- **Correo Institucional:** `profesora.margaritamiranda@cepluispasteur.cl`
