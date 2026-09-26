import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Entrevistas Apoderados"

# Encabezado institucional
ws.merge_cells("A1:G1")
ws["A1"] = "COLEGIO LUIS PASTEUR ANEXO - CRONOGRAMA DE ENTREVISTAS DE APODERADOS 2026"
ws["A1"].font = Font(name="Arial", size=12, bold=True, color="FFFFFF")
ws["A1"].fill = PatternFill(start_color="173F73", end_color="173F73", fill_type="solid")
ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:G2")
ws["A2"] = "Profesora Margarita Miranda B. | Jefatura 5° Básico A y Asignaturas"
ws["A2"].font = Font(name="Arial", size=10, italic=True, color="173F73")
ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
ws.row_dimensions[2].height = 20

# Encabezados de tabla
headers = [
    "N°", "Nombre del apoderado/a", "Correo electrónico",
    "Nombre del/la estudiante", "Curso", "Fecha entrevista", "Hora de entrevista"
]

ws.row_dimensions[4].height = 25
for col_idx, h in enumerate(headers, 1):
    cell = ws.cell(4, col_idx, h)
    cell.font = Font(name="Arial", size=9.5, bold=True, color="FFFFFF")
    cell.fill = PatternFill(start_color="173F73", end_color="173F73", fill_type="solid")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

data = [
    # GRUPO 1: Miércoles 07 de octubre
    (1, "Luce Eliset Gajardo", "luceelisetgajardo30@gmail.com", "Gajardo (Estudiante)", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    (2, "María Reyes", "mariareyessanti.14@gmail.com", "Reyes Cabrera, Catalina Amanda", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    (3, "Aurora Celin / Familia Celin Mosquera", "suacerfer2805.a@gmail.com", "Mosquera Celin, Aurora Susana", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    (4, "Cecilia Riquelme Cartagena", "ceci.riquelme.cartagena@gmail.com", "Segovia Riquelme, Marcelo Leonardo", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    (5, "Victoria Sepúlveda Fonseca", "victoria.fonseca.7@gmail.com", "Sepúlveda Sepúlveda, Pedro Nahuel", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    (6, "Luz Elena Guerrero", "luzeg782511@gmail.com", "Restrepo Guerrero, Luis David", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    (7, "Paulina Encina", "encinavich19@gmail.com", "Muñoz Encina, Simón Ignacio", "5° Básico A", "Miércoles 07 de octubre", "13:55 a 14:40 hrs"),
    # GRUPO 2: Miércoles 14 de octubre
    (8, "Mella Navarrete (Sra. Mella)", "mella.navarrete.m@gmail.com", "Aravena Mella, Martín Andree", "5° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (9, "Marcia Elizabeth Pacheco Pacheco", "marciaelizabethpachecopacheco@gmail.com", "Flores Pacheco, Hamandha Sophia", "5° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (10, "Osvaldo González Correa", "osvaldo.gonzalez.correa@gmail.com", "González Ibáñez, Amanda Trinidad", "5° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (11, "David Ignacio Herrera Rojas", "davidignacioherrerarojas@gmail.com", "Herrera Díaz, Paz Anaís", "5° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (12, "Valentina Constanza Aguayo Trejo", "valentinacaguayo7@gmail.com", "Leyton Aguayo, Gabriel Andrés", "5° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (13, "Familia Morales Silva (Claro de Luna)", "clarodeluna_1980@htmail.com", "Morales Silva, Danilo Cristián", "5° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (14, "Fabiola Mora C.", "fabiola.mora.c@gmail.com", "Ibarra Oyarzún, Benjamín Ricardo", "8° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (15, "Verónica Urrutia (Verito Eleniz)", "veritoeleniz2013@gmail.com", "Lobos Urrutia, Eleniz Javiera Lía", "8° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs"),
    (16, "Carana M. / Familia Arana", "caranamza@gmail.com", "Oliva Arana, Thiago Yared", "8° Básico A", "Miércoles 14 de octubre", "13:55 a 14:40 hrs")
]

thin_border = Border(
    left=Side(style='thin', color='B0C4DE'),
    right=Side(style='thin', color='B0C4DE'),
    top=Side(style='thin', color='B0C4DE'),
    bottom=Side(style='thin', color='B0C4DE')
)

row_start = 5
for r_data in data:
    ws.row_dimensions[row_start].height = 20
    is_grupo1 = "07 de octubre" in r_data[5]
    fill_bg = "F6FAFE" if is_grupo1 else "FFFFFF"

    for c_i, val in enumerate(r_data, 1):
        cell = ws.cell(row_start, c_i, val)
        cell.font = Font(name="Arial", size=9)
        cell.border = thin_border
        cell.fill = PatternFill(start_color=fill_bg, end_color=fill_bg, fill_type="solid")
        if c_i in [1, 5, 6, 7]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    row_start += 1

def generar_cronograma_excel():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(base_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    # Ajustar anchos de columnas
    col_widths = {1: 6, 2: 34, 3: 38, 4: 36, 5: 14, 6: 24, 7: 20}
    for col_idx, width in col_widths.items():
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    local_path = os.path.join(output_dir, "Cronograma_Entrevistas_Apoderados_2026.xlsx")
    wb.save(local_path)
    print(f"Excel guardado localmente en: {local_path}")

    # Si existe la carpeta de descargas de la profesora, guardar copia también
    downloads_path = r"C:\Users\Jack\Downloads\Evaluaciones Finales Profesora Margarita\03 - Registro y Gestión Docente\Nóminas y Listas"
    if os.path.exists(downloads_path):
        target_path = os.path.join(downloads_path, "Cronograma_Entrevistas_Apoderados_2026.xlsx")
        try:
            wb.save(target_path)
            print(f"Copia actualizada en Descargas: {target_path}")
        except Exception as e:
            print(f"Aviso al guardar en Descargas: {e}")

    return local_path

if __name__ == "__main__":
    generar_cronograma_excel()
