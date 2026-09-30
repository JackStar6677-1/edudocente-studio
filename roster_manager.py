"""
EduDocente-Studio: Gestor Especializado de Nóminas y Cronogramas de Entrevistas en Excel
Genera planillas openpyxl configurables, actualizables y con identidad institucional dinámica.
Permite registrar estudiantes, apoderados, cursos, fechas, horarios y estados de citación.
"""

import os
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
ROSTER_FILE = os.path.join(CONFIG_DIR, "roster_data.json")

os.makedirs(CONFIG_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

from institution_manager import institution_manager

DEFAULT_INTERVIEWS = [
    # GRUPO 1: Miércoles 07 de octubre
    {"id": 1, "apoderado": "Luce Eliset Gajardo", "email": "luceelisetgajardo30@gmail.com", "estudiante": "Gajardo (Estudiante)", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Revisión rendimiento mensual"},
    {"id": 2, "apoderado": "María Reyes", "email": "mariareyessanti.14@gmail.com", "estudiante": "Reyes Cabrera, Catalina Amanda", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Situación académica"},
    {"id": 3, "apoderado": "Aurora Celin / Familia Celin", "email": "suacerfer2805.a@gmail.com", "estudiante": "Mosquera Celin, Aurora Susana", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Pendiente", "obs": "Esperando confirmación"},
    {"id": 4, "apoderado": "Cecilia Riquelme Cartagena", "email": "ceci.riquelme.cartagena@gmail.com", "estudiante": "Segovia Riquelme, Marcelo Leonardo", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Felicitaciones por avance"},
    {"id": 5, "apoderado": "Victoria Sepúlveda Fonseca", "email": "victoria.fonseca.7@gmail.com", "estudiante": "Sepúlveda Sepúlveda, Pedro Nahuel", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Plan de refuerzo"},
    {"id": 6, "apoderado": "Luz Elena Guerrero", "email": "luzeg782511@gmail.com", "estudiante": "Restrepo Guerrero, Luis David", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Pendiente", "obs": "Reiterar citación"},
    {"id": 7, "apoderado": "Paulina Encina", "email": "encinavich19@gmail.com", "estudiante": "Muñoz Encina, Simón Ignacio", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "07 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Pauta de conducta"},
    # GRUPO 2: Miércoles 14 de octubre
    {"id": 8, "apoderado": "Mella Navarrete (Sra. Mella)", "email": "mella.navarrete.m@gmail.com", "estudiante": "Aravena Mella, Martín Andree", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Situación conductual"},
    {"id": 9, "apoderado": "Marcia Elizabeth Pacheco", "email": "marciaelizabethpachecopacheco@gmail.com", "estudiante": "Flores Pacheco, Hamandha Sophia", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Evaluación diferenciada"},
    {"id": 10, "apoderado": "Osvaldo González Correa", "email": "osvaldo.gonzalez.correa@gmail.com", "estudiante": "González Ibáñez, Amanda Trinidad", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Apoyo pedagógico"},
    {"id": 11, "apoderado": "David Ignacio Herrera Rojas", "email": "davidignacioherrerarojas@gmail.com", "estudiante": "Herrera Díaz, Paz Anaís", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Pendiente", "obs": "Contactar por teléfono"},
    {"id": 12, "apoderado": "Valentina Constanza Aguayo", "email": "valentinacaguayo7@gmail.com", "estudiante": "Leyton Aguayo, Gabriel Andrés", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Revisión de cuadernos"},
    {"id": 13, "apoderado": "Familia Morales Silva", "email": "clarodeluna_1980@htmail.com", "estudiante": "Morales Silva, Danilo Cristián", "curso": "5° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Informe de notas"},
    {"id": 14, "apoderado": "Fabiola Mora C.", "email": "fabiola.mora.c@gmail.com", "estudiante": "Ibarra Oyarzún, Benjamín Ricardo", "curso": "8° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Postulación enseñanza media"},
    {"id": 15, "apoderado": "Verónica Urrutia", "email": "veritoeleniz2013@gmail.com", "estudiante": "Lobos Urrutia, Eleniz Javiera Lía", "curso": "8° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Orientación vocacional"},
    {"id": 16, "apoderado": "Carana M. / Familia Arana", "email": "caranamza@gmail.com", "estudiante": "Oliva Arana, Thiago Yared", "curso": "8° Básico A", "dia": "Miércoles", "fecha": "14 de octubre", "hora": "13:55 a 14:40 hrs", "estado": "Confirmado", "obs": "Asistencia y puntualidad"}
]

class RosterManager:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(RosterManager, cls).__new__(cls)
            cls._instance._load()
        return cls._instance

    def _load(self):
        if os.path.exists(ROSTER_FILE):
            try:
                with open(ROSTER_FILE, "r", encoding="utf-8") as f:
                    self.data = json.load(f)
            except Exception:
                self.data = {"interviews": DEFAULT_INTERVIEWS}
                self._save()
        else:
            self.data = {"interviews": DEFAULT_INTERVIEWS}
            self._save()

    def _save(self):
        with open(ROSTER_FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, indent=2, ensure_ascii=False)

    def get_all(self):
        return self.data.get("interviews", [])

    def add_interview(self, item):
        interviews = self.data.get("interviews", [])
        new_id = max([x.get("id", 0) for x in interviews] + [0]) + 1
        item["id"] = new_id
        interviews.append(item)
        self.data["interviews"] = interviews
        self._save()
        return item

    def update_interview(self, item_id, updated_fields):
        interviews = self.data.get("interviews", [])
        for idx, it in enumerate(interviews):
            if it.get("id") == item_id:
                interviews[idx].update(updated_fields)
                self.data["interviews"] = interviews
                self._save()
                return interviews[idx]
        return None

    def delete_interview(self, item_id):
        interviews = self.data.get("interviews", [])
        self.data["interviews"] = [x for x in interviews if x.get("id") != item_id]
        self._save()
        return True

    def generate_excel(self, filename="Cronograma_Entrevistas_Apoderados_2026.xlsx"):
        """Genera el libro Excel institucional con formato condicional y paleta de la institución activa"""
        inst = institution_manager.get_active()
        theme = institution_manager.get_active_theme_bundle()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Cronograma Entrevistas"
        ws.views.sheetView[0].showGridLines = True

        p_color = theme["primary_hex"]
        s_color = theme["secondary_hex"]
        b_color = theme["border_hex"]

        # Encabezado 1: Nombre de la Institución y Título
        ws.merge_cells("A1:I1")
        school_name = inst.get("name", "COLEGIO CASTELGANDOLFO")
        ws["A1"] = f"{school_name.upper()} — CRONOGRAMA DE ENTREVISTAS DE APODERADOS 2026"
        ws["A1"].font = Font(name=inst.get("font_family", "Arial"), size=12, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill(start_color=p_color, end_color=p_color, fill_type="solid")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 30

        # Encabezado 2: Sub-encabezado / Profesor / Asignatura
        ws.merge_cells("A2:I2")
        ws["A2"] = f"Gestión de Asignatura y Jefatura de Curso | Profesora Margarita Miranda B. | Sistema EduDocente"
        ws["A2"].font = Font(name=inst.get("font_family", "Arial"), size=10, italic=True, color=p_color)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22

        # Fila 3 en blanco
        ws.row_dimensions[3].height = 10

        # Encabezados de Columnas (Fila 4)
        headers = [
            ("N°", 6),
            ("Nombre del apoderado/a", 34),
            ("Correo electrónico", 38),
            ("Nombre del/la estudiante", 34),
            ("Curso", 14),
            ("Día", 14),
            ("Fecha entrevista", 22),
            ("Hora de entrevista", 20),
            ("Estado", 16)
        ]

        ws.row_dimensions[4].height = 26
        for col_idx, (h_text, width) in enumerate(headers, 1):
            cell = ws.cell(4, col_idx, h_text)
            cell.font = Font(name=inst.get("font_family", "Arial"), size=9.5, bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color=p_color, end_color=p_color, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # Borde delgado institucional
        cell_border = Border(
            left=Side(style='thin', color=b_color),
            right=Side(style='thin', color=b_color),
            top=Side(style='thin', color=b_color),
            bottom=Side(style='thin', color=b_color)
        )

        # Rellenar datos
        interviews = self.get_all()
        row_cur = 5
        for it in interviews:
            ws.row_dimensions[row_cur].height = 21

            # Alternar color zebra con tinte institucional
            is_zebra = (row_cur % 2 == 0)
            fill_hex = s_color if is_zebra else "FFFFFF"

            row_values = [
                it.get("id", row_cur - 4),
                it.get("apoderado", ""),
                it.get("email", ""),
                it.get("estudiante", ""),
                it.get("curso", ""),
                it.get("dia", ""),
                it.get("fecha", ""),
                it.get("hora", ""),
                it.get("estado", "Pendiente")
            ]

            for col_idx, val in enumerate(row_values, 1):
                c = ws.cell(row_cur, col_idx, val)
                c.font = Font(name=inst.get("font_family", "Arial"), size=9)
                c.border = cell_border
                c.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")

                # Alineación
                if col_idx in [1, 5, 6, 7, 8, 9]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")

                # Destacar estado
                if col_idx == 9:
                    if val == "Confirmado":
                        c.font = Font(name=inst.get("font_family", "Arial"), size=9, bold=True, color="1E7E34")
                    elif val == "Pendiente":
                        c.font = Font(name=inst.get("font_family", "Arial"), size=9, bold=True, color="D97706")

            row_cur += 1

        # Validación de datos para columna Estado (dropdown)
        dv = DataValidation(type="list", formula1='"Confirmado,Pendiente,Reprogramado,Inasistencia"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"I5:I{row_cur + 50}")

        # Guardar en output/
        out_path = os.path.join(OUTPUT_DIR, filename)
        wb.save(out_path)

        # Si existe carpeta en Descargas, actualizarla también
        downloads_dir = r"C:\Users\Jack\Downloads\Evaluaciones Finales Profesora Margarita\03 - Registro y Gestión Docente\Nóminas y Listas"
        if os.path.exists(downloads_dir):
            try:
                wb.save(os.path.join(downloads_dir, filename))
            except Exception:
                pass

        print(f"[ROSTER] Planilla Excel generada en: {out_path}")
        return out_path

roster_manager = RosterManager()
