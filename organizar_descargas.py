import os
import shutil

downloads_dir = r"C:\Users\Jack\Downloads"
margarita_dir = os.path.join(downloads_dir, "Evaluaciones Finales Profesora Margarita")

# Estructura docente de Profesora Margarita
prof_structure = {
    "01 - Ciencias Naturales": [
        "5° Básico (5°A y 5°B)",
        "6° Básico (6°A y 6°B)",
        "7° Básico (7°A)",
        "8° Básico (8°A)"
    ],
    "02 - Música": [
        "1° Básico (1°A)",
        "2° Básico (2°A)",
        "5° y 6° Básico"
    ],
    "03 - Registro y Gestión Docente": [
        "Nóminas y Listas",
        "Actas y Evidencias",
        "Plantillas y Banners"
    ]
}

# Crear estructura de Profesora Margarita
for sub, courses in prof_structure.items():
    for c in courses:
        os.makedirs(os.path.join(margarita_dir, sub, c), exist_ok=True)

print("Estructura de Profesora Margarita creada con éxito.")

# Carpetas estándar para organizar Downloads general
general_folders = {
    "_Programas e Instaladores": [
        "Antigravity-x64.exe", "Antigravity.tar.gz", "CurseForge Windows - Installer.exe",
        "OfficeSetup.exe", "RaiDrive.Mount_2026.6.25_x64.msi", "tailscale-setup-1.102.4.exe",
        "utorrent_installer.exe", "winrar-x64-723es.exe", "Antigravity-x64"
    ],
    "_Comprobantes y Finanzas": [
        "comprobante_pago_cuentas_9468961_20260917104435.pdf",
        "ComprobantePago (1).pdf", "ComprobantePago (2).pdf", "ComprobantePago.pdf",
        "ficha_ideapro_ConstruMetálicas_S.A.S.pdf"
    ],
    "_Imágenes y Gráficos": [
        "26c13d86-071e-4d47-960a-df14866335ea_11zon (1).png",
        "26c13d86-071e-4d47-960a-df14866335ea_11zon.png",
        "26c13d86-071e-4d47-960a-df14866335ea.png",
        "Mapa de árbol del problema.png",
        "drakescraft_banner_animated.gif",
        "drakescraft_banner_animated.svg"
    ],
    "_Videos y Multimedia": [
        "drakescraft-tu-legado-final.mp4",
        "ssstik.io_@chiligames22_1788658049310.mp4",
        "WhatsApp Video 2026-09-18 at 9.05.45 PM.mp4",
        "evidenciarobo"
    ],
    "_Juegos y Modpacks": [
        "amd_ags_x64.zip", "bot de lauti.zip", "Dra0gonBSpa0rkingZE0RO-29.07.2026-elamigosgames.net.torrent",
        "DrakesCraft_banner_pack", "DrakesCraft_banner_pack.zip", "DrakesCraft-ModPack-v67.zip",
        "forzahorizon6-pe1778859852461.torrent", "GokuAndVegetaAllForm 1498 2 2026-09-21T17-44Z KVJvJHdhW.rar",
        "Imperialpalace.schem", "Marvels.Spider.Man.2.elamigos.torrent", "Microsoft.Services.Store.winmd",
        "org.waywallen.open-wallpaper-engine-0.2.8-linux-x86_64.zip", "Savegame_100_V2.2-16-2-2-0-1745882070.rar",
        "Slimefun 1.21.5+ v2.4 (1).zip", "waywallen-kde-0.3.3-x86-64-embed.zip"
    ],
    "_Documentos Personales": [
        "auditoria-castel-2026-09-21.md",
        "CATEDRA N°1. Ajustes y validación del Proyecto.docx"
    ]
}

for folder_name, file_list in general_folders.items():
    dest_path = os.path.join(downloads_dir, folder_name)
    os.makedirs(dest_path, exist_ok=True)
    for fname in file_list:
        src = os.path.join(downloads_dir, fname)
        if os.path.exists(src):
            dst = os.path.join(dest_path, fname)
            try:
                shutil.move(src, dst)
                print(f"Movido: {fname} -> {folder_name}")
            except Exception as e:
                print(f"Error moviendo {fname}: {e}")

# Mover y organizar archivos de Profesora Margarita
# 1. Ciencias 5° Básico
c5_dir = os.path.join(margarita_dir, "01 - Ciencias Naturales", "5° Básico (5°A y 5°B)")
for f in ["Evaluacion_Final_Ciencias_5Basico.docx", "Pauta_Correccion_Ciencias_5Basico.docx", "Temario_Evaluacion_Final_Ciencias_5Basico.docx"]:
    # Check in Evaluaciones_Pasteur_5Basico or scratch output
    src_pasteur = os.path.join(downloads_dir, "Evaluaciones_Pasteur_5Basico", f)
    src_scratch = os.path.join(r"C:\Users\Jack\.gemini\antigravity\scratch\evaluaciones_pasteur\output", f)
    dst = os.path.join(c5_dir, f)
    if os.path.exists(src_pasteur):
        shutil.copyfile(src_pasteur, dst)
        print(f"Instalado en Ciencias 5°: {f}")
    elif os.path.exists(src_scratch):
        shutil.copyfile(src_scratch, dst)
        print(f"Instalado en Ciencias 5°: {f}")

# 2. Ciencias 6° Básico
c6_dir = os.path.join(margarita_dir, "01 - Ciencias Naturales", "6° Básico (6°A y 6°B)")
guia_6 = os.path.join(downloads_dir, "Guia - La Energia se Transforma - 6Basico.docx")
if os.path.exists(guia_6):
    shutil.move(guia_6, os.path.join(c6_dir, "Guia - La Energia se Transforma - 6Basico.docx"))
    print("Movido a Ciencias 6°: Guia - La Energia se Transforma - 6Basico.docx")

modelo_6 = os.path.join(r"C:\Users\Jack\.gemini\antigravity\scratch\evaluaciones_pasteur\output", "Evaluacion_Energia_6Basico_Modelo.docx")
if os.path.exists(modelo_6):
    shutil.copyfile(modelo_6, os.path.join(c6_dir, "Evaluacion_Energia_6Basico_Modelo.docx"))
    print("Copiado a Ciencias 6°: Evaluacion_Energia_6Basico_Modelo.docx")

# 3. Música: Copiar diagnósticos y rúbricas disponibles de la docente
musica_1_dir = os.path.join(margarita_dir, "02 - Música", "1° Básico (1°A)")
musica_2_dir = os.path.join(margarita_dir, "02 - Música", "2° Básico (2°A)")

doc_diag = r"C:\Users\Jack\Documents\Profesora\Año Escolar 2026\04 - Evaluaciones\Diagnosticos"
for d_type in ["Diagnosticos", "Pautas"]:
    src_folder = os.path.join(doc_diag, d_type)
    if os.path.exists(src_folder):
        for item in os.listdir(src_folder):
            if "musica 1" in item.lower():
                shutil.copyfile(os.path.join(src_folder, item), os.path.join(musica_1_dir, item))
                print(f"Copiado a Música 1°: {item}")
            elif "musica 2" in item.lower():
                shutil.copyfile(os.path.join(src_folder, item), os.path.join(musica_2_dir, item))
                print(f"Copiado a Música 2°: {item}")

# 4. Registro y Gestión Docente
nomina_src = os.path.join(downloads_dir, "NÓMINA DE ESTUDIANTES 2026.xlsx")
if os.path.exists(nomina_src):
    shutil.move(nomina_src, os.path.join(margarita_dir, "03 - Registro y Gestión Docente", "Nóminas y Listas", "NÓMINA DE ESTUDIANTES 2026.xlsx"))
    print("Movido: NÓMINA DE ESTUDIANTES 2026.xlsx")

acta_src = os.path.join(downloads_dir, "Acta")
if os.path.exists(acta_src):
    acta_dst = os.path.join(margarita_dir, "03 - Registro y Gestión Docente", "Actas y Evidencias", "Evidencias de Clase y Actas")
    if not os.path.exists(acta_dst):
        shutil.move(acta_src, acta_dst)
        print("Movido: Carpeta Acta -> Actas y Evidencias")

for extra_f in ["titulos 5 copias.docx", "banner.docx"]:
    f_src = os.path.join(downloads_dir, extra_f)
    if os.path.exists(f_src):
        shutil.move(f_src, os.path.join(margarita_dir, "03 - Registro y Gestión Docente", "Plantillas y Banners", extra_f))
        print(f"Movido: {extra_f}")

# Limpiar carpeta temporal si ya fue migrada
temp_folder = os.path.join(downloads_dir, "Evaluaciones_Pasteur_5Basico")
if os.path.exists(temp_folder):
    try:
        shutil.rmtree(temp_folder)
        print("Carpeta temporal Evaluaciones_Pasteur_5Basico limpiada.")
    except Exception as e:
        print(f"Nota limpieza: {e}")

print("\n--- ORGANIZACIÓN COMPLETADA CON ÉXITO ---")
