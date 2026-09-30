"""
EduDocente-Studio: Servidor Web Local y API REST
Proporciona la interfaz gráfica de usuario (UI), autenticación OAuth 2.0 y compilación asistida por IA.
No requiere dependencias externas complejas; utiliza el servidor HTTP nativo multi-hilo de Python.
"""

import os
import sys
import json
import urllib.parse
import webbrowser
import threading
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler

# Asegurar codificación utf-8 en Windows
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, "web")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

from oauth_manager import auth_manager
from ai_connector import AIConnector
from engine import DocenteEngine
from institution_manager import institution_manager
from roster_manager import roster_manager
from generator_core import reload_active_theme

class EduDocenteRequestHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        # Log simplificado y limpio en consola
        sys.stderr.write(f"[HTTP] {self.command} {self.path} - {format % args}\n")

    def _send_json(self, status_code, data):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

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
            rel_path = urllib.parse.unquote(path[8:])
            asset_file = os.path.join(BASE_DIR, "assets", rel_path)
            if os.path.exists(asset_file) and os.path.isfile(asset_file):
                with open(asset_file, "rb") as f:
                    asset_data = f.read()
                self.send_response(200)
                if asset_file.endswith(".svg"):
                    self.send_header("Content-Type", "image/svg+xml")
                elif asset_file.endswith(".png"):
                    self.send_header("Content-Type", "image/png")
                elif asset_file.endswith(".jpg") or asset_file.endswith(".jpeg"):
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
            filename = os.path.basename(urllib.parse.unquote(path[10:]))
            filepath = os.path.join(OUTPUT_DIR, filename)
            if os.path.exists(filepath):
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
                self.end_headers()
                self.wfile.write(file_data)
                return
            else:
                self.send_error(404, f"Archivo no encontrado: {filename}")
                return

        # 3. API: Estado general del sistema y usuario
        if path == "/api/status":
            user = auth_manager.get_current_user()
            providers = auth_manager.get_providers_status()
            self._send_json(200, {
                "status": "online",
                "version": "1.0.0",
                "user": user,
                "oauth_providers": providers
            })
            return

        # 4. API: OAuth Login Flow
        if path.startswith("/api/auth/") and path.endswith("/login"):
            parts = path.strip("/").split("/")
            if len(parts) >= 3:
                provider = parts[2]
                host = self.headers.get("Host", "localhost:8080")
                redirect_uri = f"http://{host}/api/auth/{provider}/callback"
                try:
                    url = auth_manager.generate_auth_url(provider, redirect_uri)
                    self._send_json(200, {"redirect_url": url, "provider": provider})
                    return
                except Exception as e:
                    self._send_json(400, {"error": str(e)})
                    return

        # 5. API: OAuth Callback Flow
        if path.startswith("/api/auth/") and path.endswith("/callback"):
            parts = path.strip("/").split("/")
            if len(parts) >= 3:
                provider = parts[2]
                code = query.get("code", ["sandbox"])[0]
                state = query.get("state", [""])[0]
                is_sandbox = "sandbox" in query or code == "sandbox"
                user = auth_manager.handle_callback(provider, code, state, is_sandbox=is_sandbox)
                self._send_json(200, user)
                return

        # 6. API: Instituciones Educativas y Memoria de Branding
        if path == "/api/institutions":
            insts = institution_manager.get_all()
            active = institution_manager.get_active()
            self._send_json(200, {"institutions": insts, "active": active})
            return

        # 7. API: Nómina de Entrevistas y Estudiantes
        if path == "/api/roster":
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

        self.send_error(404, "Ruta no encontrada")

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        content_length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_length).decode("utf-8") if content_length > 0 else "{}"
        try:
            payload = json.loads(body) if body else {}
        except json.JSONDecodeError:
            payload = {}

        # 1. API: Logout
        if path == "/api/auth/logout":
            user = auth_manager.logout()
            self._send_json(200, {"success": True, "user": user})
            return

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
            teacher = payload.get("teacher", None)

            print(f"\n[GENERACIÓN WEB] Solicitud recibida:")
            print(f"  • Proveedor: {provider.upper()}")
            print(f"  • API Key personalizada: {'Sí (Proporcionada por Docente)' if api_key else 'No (Usando Buffer/Sistema)'}")
            print(f"  • Asignatura: {subject} | Curso: {grade}")
            print(f"  • Contenidos: {topic}")

            try:
                # 1. Ejecutar Conector de IA con la clave del profesor o buffer
                connector = AIConnector(provider=provider, api_key=api_key, model=model)
                assessment_data = connector.generate_assessment_json(
                    topic=topic,
                    grade=grade,
                    oa=oa,
                    textbook_pages=textbook_pages
                )
                if teacher:
                    assessment_data["docente"] = teacher

                # 2. Compilar documentos con DocenteEngine
                engine = DocenteEngine(assessment_data)
                student_path, teacher_path = engine.build_all()

                student_file = os.path.basename(student_path)
                teacher_file = os.path.basename(teacher_path)

                self._send_json(200, {
                    "success": True,
                    "student_file": student_file,
                    "teacher_file": teacher_file,
                    "student_path": student_path,
                    "teacher_path": teacher_path,
                    "assessment": assessment_data
                })
            except Exception as e:
                import traceback
                traceback.print_exc()
                self._send_json(500, {"success": False, "error": str(e)})
            return

        # 3. API: Cambiar Institución Activa
        if path == "/api/institutions/active":
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
            saved = institution_manager.save_institution(payload)
            reload_active_theme()
            self._send_json(200, {"success": True, "institution": saved})
            return

        # 5. API: Agregar fila a Nómina de Entrevistas
        if path == "/api/roster/add":
            item = roster_manager.add_interview(payload)
            roster_manager.generate_excel()
            self._send_json(200, {"success": True, "item": item})
            return

        # 6. API: Actualizar fila de Nómina de Entrevistas
        if path == "/api/roster/update":
            item_id = payload.get("id")
            updated = roster_manager.update_interview(item_id, payload)
            roster_manager.generate_excel()
            self._send_json(200, {"success": True, "item": updated})
            return

        # 7. API: Regenerar Planilla Excel de Nóminas
        if path == "/api/roster/generate":
            out_path = roster_manager.generate_excel()
            filename = os.path.basename(out_path)
            self._send_json(200, {
                "success": True,
                "filename": filename,
                "download_url": f"/download/{filename}"
            })
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
    parser.add_argument("--port", type=int, default=8080, help="Puerto HTTP (predeterminado: 8080)")
    parser.add_argument("--no-browser", action="store_true", help="No abrir automáticamente el navegador web")
    args = parser.parse_args()

    start_server(port=args.port, open_browser=not args.no_browser)
