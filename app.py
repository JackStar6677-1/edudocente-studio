"""
EduDocente-Studio: Servidor Web Local y API REST
Proporciona la interfaz gráfica de usuario (UI), autenticación OAuth 2.0 y compilación asistida por IA.
No requiere dependencias externas complejas; utiliza el servidor HTTP nativo multi-hilo de Python.
"""

import os
import sys
import json
import re
import secrets
import shutil
import urllib.parse
import webbrowser
import threading
from pathlib import Path
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

# Asegurar codificación utf-8 en Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, "web")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

from sso_auth import (
    create_session, get_session, delete_session, consume_sso_token,
    grant_download, can_download, SESSION_SECONDS,
)
from ai_connector import AIConnector
from engine import DocenteEngine
from institution_manager import institution_manager
from roster_manager import roster_manager
from generator_core import reload_active_theme

class EduDocenteRequestHandler(BaseHTTPRequestHandler):

    CENTRAL_SSO_URL = "https://colegiocastelgandolfo.cl/admin/sso_jump.php?target=edudocente"

    def _session_user(self):
        """Resuelve la identidad desde una cookie de servidor, nunca desde el cliente."""
        return get_session(self.headers.get("Cookie", ""))

    def _require_session(self, path):
        """Redirige páginas y rechaza API/descargas sin sesión institucional."""
        if self._session_user():
            return True
        if path.startswith("/api/") or path.startswith("/download/"):
            self._send_json(401, {"error": "Sesión institucional requerida"})
        else:
            self.send_response(302)
            self.send_header("Location", self.CENTRAL_SSO_URL)
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
        return False

    def _require_admin(self):
        """Reserva nóminas y configuración institucional para administradores."""
        if self._session_user().get("role") == "admin":
            return True
        self._send_json(403, {"error": "Se requiere rol administrativo"})
        return False

    def _handle_sso(self, query):
        """Canjea un pase firmado de un solo uso por cookie HttpOnly."""
        try:
            user = consume_sso_token(query.get("token", [""])[0])
        except RuntimeError:
            return self._send_json(503, {"error": "SSO institucional no configurado"})
        except ValueError:
            return self._send_json(403, {"error": "Pase SSO inválido o vencido"})
        session = create_session(user)
        secure = "; Secure" if os.environ.get("EDUDOCENTE_COOKIE_SECURE", "1") != "0" else ""
        self.send_response(303)
        self.send_header("Location", "/")
        self.send_header("Set-Cookie", f"edu_session={session}; Max-Age={SESSION_SECONDS}; Path=/; HttpOnly; SameSite=Lax{secure}")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Referrer-Policy", "no-referrer")
        self.end_headers()

    def log_message(self, format, *args):
        """Registra la ruta sin persistir el pase SSO en texto plano."""
        safe_path = urllib.parse.urlparse(self.path).path
        detail = re.sub(r"([?&]token=)[^&\s]+", r"\1[REDACTED]", format % args)
        sys.stderr.write(f"[HTTP] {self.command} {safe_path} - {detail}\n")

    def _send_json(self, status_code, data):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def do_HEAD(self):
        self.send_response(200 if self._session_user() else 401)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

    def do_OPTIONS(self):
        self.send_response(204 if self._session_user() else 401)
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path == "/sso":
            return self._handle_sso(query)
        if path == "/api/health":
            return self._send_json(200, {"status": "ok"})
        if not self._require_session(path):
            return
        user = self._session_user()

        # 1. Página principal Web SPA
        if path == "/" or path == "/index.html":
            index_file = os.path.join(WEB_DIR, "index.html")
            if os.path.exists(index_file):
                with open(index_file, "rb") as f:
                    content = f.read()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                self.wfile.write(content)
                return
            else:
                self.send_error(404, "web/index.html no encontrado")
                return

        # 2. Servir recursos gráficos estáticos (/assets/)
        if path.startswith("/assets/"):
            assets_root = (Path(BASE_DIR) / "assets").resolve()
            asset_file = (assets_root / urllib.parse.unquote(path[8:])).resolve()
            if asset_file.is_relative_to(assets_root) and asset_file.is_file():
                with open(asset_file, "rb") as f:
                    asset_data = f.read()
                self.send_response(200)
                if asset_file.suffix == ".svg":
                    self.send_header("Content-Type", "image/svg+xml")
                elif asset_file.suffix == ".png":
                    self.send_header("Content-Type", "image/png")
                elif asset_file.suffix in (".jpg", ".jpeg"):
                    self.send_header("Content-Type", "image/jpeg")
                else:
                    self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Length", str(len(asset_data)))
                self.send_header("Cache-Control", "public, max-age=3600")
                self.end_headers()
                self.wfile.write(asset_data)
                return
            else:
                self.send_error(404, "Recurso gráfico no encontrado")
                return

        # 3. Descargas de documentos generados
        if path.startswith("/download/"):
            parts = path.split("/")
            if (len(parts) != 4 or not re.fullmatch(r"[0-9a-f]{32}", parts[2])
                    or parts[3] not in ("Evaluacion.docx", "Pauta.docx", "Entrevistas.xlsx")
                    or not can_download(self.headers.get("Cookie", ""), parts[2])):
                return self._send_json(404, {"error": "Documento no disponible"})
            filename = parts[3]
            filepath = Path(OUTPUT_DIR) / parts[2] / filename
            if filepath.is_file():
                with open(filepath, "rb") as f:
                    file_data = f.read()
                self.send_response(200)
                if filename.endswith(".docx"):
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                elif filename.endswith(".xlsx"):
                    self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    self.send_header("Content-Type", "application/octet-stream")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(len(file_data)))
                self.send_header("Cache-Control", "private, no-store")
                self.end_headers()
                self.wfile.write(file_data)
                return
            else:
                self.send_error(404, f"Archivo no encontrado: {filename}")
                return

        # 3. API: Estado general del sistema y usuario
        if path == "/api/status":
            self._send_json(200, {
                "status": "online",
                "version": "1.0.0",
                "user": user,
                "oauth_providers": {}
            })
            return

        if path.startswith("/api/auth/"):
            return self._send_json(410, {"error": "Acceso docente disponible solo desde /admin/"})

        # 6. API: Instituciones Educativas y Memoria de Branding
        if path == "/api/institutions":
            insts = institution_manager.get_all()
            active = institution_manager.get_active()
            self._send_json(200, {"institutions": insts, "active": active})
            return

        # 7. API: Nómina de Entrevistas y Estudiantes
        if path == "/api/roster":
            if not self._require_admin():
                return
            interviews = roster_manager.get_all()
            self._send_json(200, {"interviews": interviews})
            return

        # 8. API: Proveedores de IA con información de cuotas gratuitas/semanales
        if path == "/api/ai/providers":
            from ai_connector import AI_PROVIDERS
            providers_list = []
            for pid, info in AI_PROVIDERS.items():
                providers_list.append({
                    "id": pid,
                    "name": info.get("name", pid),
                    "badge": info.get("badge", ""),
                    "default_model": info.get("default_model", ""),
                    "help_url": info.get("help_url", "#"),
                    "free_tier_info": info.get("free_tier_info", "")
                })
            self._send_json(200, {"providers": providers_list})
            return

        # 9. API: Catálogo de Objetivos de Aprendizaje Priorizados Mineduc
        if path == "/api/curriculum/oa":
            oa_path = os.path.join(BASE_DIR, "config", "mineduc_oa_priorizados.json")
            if os.path.exists(oa_path):
                try:
                    with open(oa_path, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    self._send_json(200, data)
                    return
                except Exception as e:
                    self._send_json(500, {"error": f"Error leyendo OA: {e}"})
                    return
            self._send_json(404, {"error": "Catálogo OA no encontrado"})
            return

        self.send_error(404, "Ruta no encontrada")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        if not self._require_session(path):
            return

        if path == "/api/auth/logout":
            delete_session(self.headers.get("Cookie", ""))
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Set-Cookie", "edu_session=; Max-Age=0; Path=/; HttpOnly; SameSite=Lax; Secure")
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(b'{"success":true}')
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
        except (TypeError, ValueError):
            return self._send_json(400, {"error": "Content-Length inválido"})
        if content_length < 0 or content_length > 2 * 1024 * 1024:
            self.close_connection = True
            return self._send_json(413, {"error": "Solicitud demasiado grande"})
        body = self.rfile.read(content_length).decode("utf-8") if content_length > 0 else "{}"
        try:
            payload = json.loads(body) if body else {}
        except json.JSONDecodeError:
            payload = {}

        # 2. API: Generar evaluación completa con IA y DocenteEngine
        if path == "/api/generate":
            provider = payload.get("provider", "antigravity")
            api_key = payload.get("api_key", None)
            model = payload.get("model", None)
            subject = payload.get("subject", "CIENCIAS NATURALES")
            grade = payload.get("grade", "6° Básico A")
            topic = payload.get("topic", "Cambios del estado de la materia")
            oa = payload.get("oa", "OA 13 — Demostrar cambios de estado")
            textbook_pages = payload.get("textbook_pages", "Páginas 170 y 173")
            job_dir = None
            try:
                points = int(payload.get("points", 25))
            except (ValueError, TypeError):
                points = 25
            teacher = self._session_user()

            print(f"\n[GENERACIÓN WEB] Solicitud recibida:")
            print(f"  • Proveedor: {provider.upper()}")
            print(f"  • API Key personalizada: {'Sí (Proporcionada por Docente)' if api_key else 'No (Usando Buffer/Sistema)'}")
            print(f"  • Asignatura: {subject} | Curso: {grade} | Puntaje: {points} pts")
            print(f"  • Contenidos: {topic}")

            try:
                # 1. Ejecutar Conector de IA con la clave del profesor o buffer
                connector = AIConnector(provider=provider, api_key=api_key, model=model)
                assessment_data = connector.generate_assessment_json(
                    topic=topic,
                    grade=grade,
                    oa=oa,
                    textbook_pages=textbook_pages,
                    subject=subject,
                    points=points
                )
                assessment_data["docente"] = teacher["name"]
                assessment_data["email_docente"] = teacher["email"]

                # 2. Compilar documentos con DocenteEngine
                engine = DocenteEngine(assessment_data)
                job_id = secrets.token_hex(16)
                new_job_dir = Path(OUTPUT_DIR) / job_id
                new_job_dir.mkdir(mode=0o700)
                job_dir = new_job_dir
                student_path = engine.build_evaluacion_estudiante(str(job_dir / "Evaluacion.docx"))
                teacher_path = engine.build_pauta_correccion(str(job_dir / "Pauta.docx"))
                if not Path(student_path).is_file() or not Path(teacher_path).is_file():
                    raise RuntimeError("No se pudieron guardar ambos documentos")
                grant_download(self.headers.get("Cookie", ""), job_id)

                self._send_json(200, {
                    "success": True,
                    "student_file": "Evaluacion.docx",
                    "teacher_file": "Pauta.docx",
                    "student_url": f"/download/{job_id}/Evaluacion.docx",
                    "teacher_url": f"/download/{job_id}/Pauta.docx",
                    "assessment": assessment_data
                })
            except Exception as e:
                import traceback
                traceback.print_exc()
                if job_dir is not None and job_dir.is_dir():
                    shutil.rmtree(job_dir, ignore_errors=True)
                self._send_json(500, {"success": False, "error": str(e)})
            return

        # 3. API: Cambiar Institución Activa
        if path == "/api/institutions/active":
            if not self._require_admin():
                return
            inst_id = payload.get("id")
            if inst_id:
                active = institution_manager.set_active(inst_id)
                reload_active_theme()
                self._send_json(200, {"success": True, "active": active})
            else:
                self._send_json(400, {"error": "Falta id de institución"})
            return

        # 4. API: Guardar / Actualizar Institución y Paleta
        if path == "/api/institutions/save":
            if not self._require_admin():
                return
            saved = institution_manager.save_institution(payload)
            reload_active_theme()
            self._send_json(200, {"success": True, "institution": saved})
            return

        # 5. API: Agregar fila a Nómina de Entrevistas
        if path == "/api/roster/add":
            if not self._require_admin():
                return
            item = roster_manager.add_interview(payload)
            roster_manager.generate_excel()
            self._send_json(200, {"success": True, "item": item})
            return

        # 6. API: Actualizar fila de Nómina de Entrevistas
        if path == "/api/roster/update":
            if not self._require_admin():
                return
            item_id = payload.get("id")
            updated = roster_manager.update_interview(item_id, payload)
            roster_manager.generate_excel()
            self._send_json(200, {"success": True, "item": updated})
            return

        # 7. API: Regenerar Planilla Excel de Nóminas
        if path == "/api/roster/generate":
            if not self._require_admin():
                return
            job_dir = None
            try:
                out_path = roster_manager.generate_excel()
                job_id = secrets.token_hex(16)
                new_job_dir = Path(OUTPUT_DIR) / job_id
                new_job_dir.mkdir(mode=0o700)
                job_dir = new_job_dir
                shutil.copy2(out_path, job_dir / "Entrevistas.xlsx")
                grant_download(self.headers.get("Cookie", ""), job_id)
                self._send_json(200, {
                    "success": True,
                    "filename": "Entrevistas.xlsx",
                    "download_url": f"/download/{job_id}/Entrevistas.xlsx"
                })
            except Exception as error:
                if job_dir is not None and job_dir.is_dir():
                    shutil.rmtree(job_dir, ignore_errors=True)
                self._send_json(500, {"success": False, "error": "No se pudo generar la planilla"})
            return

        self.send_error(404, "Ruta POST no encontrada")

def start_server(port=8080, open_browser=True):
    server_address = ("", port)
    httpd = ThreadingHTTPServer(server_address, EduDocenteRequestHandler)
    url = f"http://localhost:{port}"

    print("\n" + "=" * 66)
    print("      🎓 EDUDOCENTE-STUDIO — SERVIDOR WEB ACTIVO      ")
    print("=" * 66)
    print(f"  • Interfaz Web:  {url}")
    print(f"  • API Endpoints: {url}/api/status")
    print(f"  • Directorio:    {BASE_DIR}")
    print("=" * 66)
    print("Presiona Ctrl+C en cualquier momento para detener el servidor.\n")

    if open_browser:
        def _open():
            import time
            time.sleep(0.8)
            webbrowser.open(url)
        threading.Thread(target=_open, daemon=True).start()

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nDeteniendo servidor web...")
        httpd.server_close()
        print("Servidor detenido correctamente.")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Servidor Web Local EduDocente-Studio")
    default_port = int(os.environ.get("EDUDOCENTE_PORT", os.environ.get("PORT", 8080)))
    parser.add_argument("--port", type=int, default=default_port, help=f"Puerto HTTP (predeterminado: {default_port})")
    parser.add_argument("--no-browser", action="store_true", help="No abrir automáticamente el navegador web")
    args = parser.parse_args()

    # Si estamos en servidor headless/linux daemon, no intentar abrir navegador
    is_headless = not os.environ.get("DISPLAY") and not sys.platform.startswith("win")
    should_open_browser = not args.no_browser and not is_headless
    start_server(port=args.port, open_browser=should_open_browser)
