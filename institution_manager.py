"""
EduDocente-Studio: Gestor de Identidad Institucional y Memoria Multi-Colegio
Permite registrar y alternar entre múltiples instituciones educativas con sus propios:
  - Nombres oficiales y dependencias (UTP, Departamentos)
  - Logos de alta resolución y con fondo transparente
  - Paletas de colores corporativas (primario, secundario, tarjetas, bordes, corrección)
  - Tipografías institucionales
"""

import os
import json
from docx.shared import RGBColor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
LOGOS_DIR = os.path.join(BASE_DIR, "assets", "logos")
os.makedirs(CONFIG_DIR, exist_ok=True)
os.makedirs(LOGOS_DIR, exist_ok=True)

INSTITUTIONS_FILE = os.path.join(CONFIG_DIR, "institutions.json")

# Perfiles de referencia predeterminados
DEFAULT_INSTITUTIONS = [
    {
        "id": "pasteur",
        "name": "COLEGIO LUIS PASTEUR ANEXO",
        "sub_header": "DEPARTAMENTO DE CIENCIAS Y EVALUACIÓN DOCENTE",
        "motto": "Ciencia, Esfuerzo y Futuro",
        "logo_path": "assets/logo_colegio.png",
        "font_family": "Arial",
        "colors": {
            "primary_hex": "173F73",       # Azul Marino Institucional
            "secondary_hex": "EAF3FB",     # Azul Cielo Suave
            "card_bg_hex": "F6FAFE",      # Fondo Tarjetas/Destacados
            "border_hex": "B0C4DE",       # Borde Azul Acero
            "teacher_correct_hex": "1E7E34" # Verde Pauta Docente
        },
        "is_active": True
    },
    {
        "id": "liceo_bicentenario",
        "name": "LICEO BICENTENARIO DE EXCELENCIA",
        "sub_header": "UNIDAD TÉCNICO PEDAGÓGICA (UTP)",
        "motto": "Compromiso, Mérito y Superación",
        "logo_path": "assets/logo_colegio.png",
        "font_family": "Arial",
        "colors": {
            "primary_hex": "0D5C3A",       # Verde Bosque Bicentenario
            "secondary_hex": "E8F5E9",     # Verde Claro Suave
            "card_bg_hex": "F1F8E9",      # Fondo Tarjetas Verde
            "border_hex": "A5D6A7",       # Borde Verde Acero
            "teacher_correct_hex": "1E7E34"
        },
        "is_active": False
    },
    {
        "id": "colegio_santa_maria",
        "name": "COLEGIO SANTA MARÍA",
        "sub_header": "COORDINACIÓN ACADÉMICA Y DE ASIGNATURA",
        "motto": "Virtud y Sabiduría",
        "logo_path": "assets/logo_colegio.png",
        "font_family": "Calibri",
        "colors": {
            "primary_hex": "800020",       # Borgoña / Burdeos Clásico
            "secondary_hex": "FCE4EC",     # Rosa Suave
            "card_bg_hex": "FFF0F5",      # Lavanda Cálido
            "border_hex": "F8BBD0",       # Borde Suave
            "teacher_correct_hex": "1E7E34"
        },
        "is_active": False
    }
]

class InstitutionManager:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(InstitutionManager, cls).__new__(cls)
            cls._instance._load()
        return cls._instance

    def _load(self):
        if os.path.exists(INSTITUTIONS_FILE):
            try:
                with open(INSTITUTIONS_FILE, "r", encoding="utf-8") as f:
                    self.institutions = json.load(f)
            except Exception:
                self.institutions = DEFAULT_INSTITUTIONS
                self._save()
        else:
            self.institutions = DEFAULT_INSTITUTIONS
            self._save()

    def _save(self):
        with open(INSTITUTIONS_FILE, "w", encoding="utf-8") as f:
            json.dump(self.institutions, f, indent=2, ensure_ascii=False)

    def get_all(self):
        return self.institutions

    def get_active(self):
        for inst in self.institutions:
            if inst.get("is_active"):
                return inst
        # Predeterminado primer elemento
        if self.institutions:
            self.institutions[0]["is_active"] = True
            return self.institutions[0]
        return DEFAULT_INSTITUTIONS[0]

    def set_active(self, inst_id):
        found = False
        for inst in self.institutions:
            if inst["id"] == inst_id:
                inst["is_active"] = True
                found = True
            else:
                inst["is_active"] = False
        if found:
            self._save()
        return self.get_active()

    def save_institution(self, data):
        inst_id = data.get("id") or data.get("name", "inst").lower().replace(" ", "_")[:20]
        data["id"] = inst_id

        # Asegurar estructura de colores
        colors = data.get("colors", {})
        data["colors"] = {
            "primary_hex": colors.get("primary_hex", "173F73").replace("#", "").upper(),
            "secondary_hex": colors.get("secondary_hex", "EAF3FB").replace("#", "").upper(),
            "card_bg_hex": colors.get("card_bg_hex", "F6FAFE").replace("#", "").upper(),
            "border_hex": colors.get("border_hex", "B0C4DE").replace("#", "").upper(),
            "teacher_correct_hex": colors.get("teacher_correct_hex", "1E7E34").replace("#", "").upper(),
        }

        # Comprobar si ya existe
        exists = False
        for idx, inst in enumerate(self.institutions):
            if inst["id"] == inst_id:
                self.institutions[idx] = data
                exists = True
                break
        if not exists:
            self.institutions.append(data)

        self._save()
        return data

    def hex_to_rgb(self, hex_str):
        hex_clean = hex_str.replace("#", "")
        if len(hex_clean) != 6:
            hex_clean = "173F73"
        r = int(hex_clean[0:2], 16)
        g = int(hex_clean[2:4], 16)
        b = int(hex_clean[4:6], 16)
        return RGBColor(r, g, b)

    def get_active_theme_bundle(self):
        """Retorna el paquete de estilos activos listos para python-docx y openpyxl"""
        inst = self.get_active()
        colors = inst["colors"]

        p_hex = colors["primary_hex"]
        s_hex = colors["secondary_hex"]
        c_hex = colors["card_bg_hex"]
        b_hex = colors["border_hex"]
        t_hex = colors["teacher_correct_hex"]

        logo_rel = inst.get("logo_path", "assets/logo_colegio.png")
        full_logo_path = os.path.join(BASE_DIR, logo_rel) if not os.path.isabs(logo_rel) else logo_rel
        if not os.path.exists(full_logo_path):
            full_logo_path = os.path.join(BASE_DIR, "assets", "logo_colegio.png")

        return {
            "institution_id": inst["id"],
            "institution_name": inst["name"],
            "sub_header": inst.get("sub_header", "DEPARTAMENTO DE EVALUACIÓN"),
            "motto": inst.get("motto", ""),
            "logo_path": full_logo_path,
            "font_family": inst.get("font_family", "Arial"),
            "primary_hex": p_hex,
            "secondary_hex": s_hex,
            "card_bg_hex": c_hex,
            "border_hex": b_hex,
            "teacher_correct_hex": t_hex,
            "primary_rgb": self.hex_to_rgb(p_hex),
            "teacher_correct_rgb": self.hex_to_rgb(t_hex),
            "white_rgb": RGBColor(0xFF, 0xFF, 0xFF),
            "dark_rgb": RGBColor(0x22, 0x22, 0x22),
            "gray_rgb": RGBColor(0x55, 0x55, 0x55),
        }

institution_manager = InstitutionManager()
