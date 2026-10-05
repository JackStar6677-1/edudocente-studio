"""
EduDocente-Studio: Gestor de Autenticación y OAuth 2.0 Empresarial
Soporta:
  - Google Workspace for Education / Google Classroom OAuth 2.0
  - GitHub OAuth
  - Microsoft 365 Education (Entra ID)
  - Modo Sandbox / Simulación Educativa para desarrollo y demostración offline
"""

import os
import json
import time
import secrets
import urllib.parse
import urllib.request
import urllib.error

# Configuración de Proveedores OAuth
OAUTH_CONFIGS = {
    "google": {
        "name": "Google Workspace / Classroom",
        "auth_url": "https://accounts.google.com/o/oauth2/v2/auth",
        "token_url": "https://oauth2.googleapis.com/token",
        "userinfo_url": "https://www.googleapis.com/oauth2/v2/userinfo",
        "client_id_env": "GOOGLE_CLIENT_ID",
        "client_secret_env": "GOOGLE_CLIENT_SECRET",
        "default_scopes": [
            "openid",
            "email",
            "profile",
            "https://www.googleapis.com/auth/classroom.courses.readonly"
        ],
        "icon": "google"
    },
    "github": {
        "name": "GitHub Academic",
        "auth_url": "https://github.com/login/oauth/authorize",
        "token_url": "https://github.com/login/oauth/access_token",
        "userinfo_url": "https://api.github.com/user",
        "client_id_env": "GITHUB_CLIENT_ID",
        "client_secret_env": "GITHUB_CLIENT_SECRET",
        "default_scopes": ["read:user", "user:email"],
        "icon": "github"
    },
    "microsoft": {
        "name": "Microsoft 365 Education",
        "auth_url": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
        "token_url": "https://login.microsoftonline.com/common/oauth2/v2.0/token",
        "userinfo_url": "https://graph.microsoft.com/v1.0/me",
        "client_id_env": "MICROSOFT_CLIENT_ID",
        "client_secret_env": "MICROSOFT_CLIENT_SECRET",
        "default_scopes": ["User.Read", "openid", "email"],
        "icon": "microsoft"
    }
}

class AuthManager:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(AuthManager, cls).__new__(cls)
            cls._instance._init_state()
        return cls._instance

    def _init_state(self):
        self.active_sessions = {}
        self.state_tokens = {}
        self.current_user = {
            "authenticated": False,
            "provider": None,
            "name": "",
            "email": "",
            "institution": "",
            "role": "",
            "avatar": "",
            "classroom_connected": False,
            "connected_at": ""
        }

    def get_providers_status(self):
        """Retorna el estado de configuración de los proveedores OAuth"""
        status = {}
        for key, conf in OAUTH_CONFIGS.items():
            cid = os.getenv(conf["client_id_env"], "")
            csec = os.getenv(conf["client_secret_env"], "")
            configured = bool(cid and csec)
            status[key] = {
                "name": conf["name"],
                "configured": configured,
                "sandbox_available": os.getenv("EDUDOCENTE_SANDBOX", "0") == "1",
                "icon": conf["icon"]
            }
        return status

    def get_current_user(self):
        return self.current_user

    def generate_auth_url(self, provider, redirect_uri):
        """Genera la URL de autorización para el proveedor seleccionado"""
        if provider not in OAUTH_CONFIGS:
            raise ValueError(f"Proveedor desconocido: {provider}")

        conf = OAUTH_CONFIGS[provider]
        state = secrets.token_urlsafe(16)
        self.state_tokens[state] = {
            "provider": provider,
            "created_at": time.time(),
            "redirect_uri": redirect_uri
        }

        cid = os.getenv(conf["client_id_env"], "")
        if not cid:
            raise RuntimeError("Proveedor OAuth no configurado")

        params = {
            "client_id": cid,
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": " ".join(conf["default_scopes"]),
            "state": state,
            "access_type": "offline",
            "prompt": "consent"
        }
        return f"{conf['auth_url']}?{urllib.parse.urlencode(params)}"

    def handle_callback(self, provider, code, state, is_sandbox=False, custom_email=None, custom_name=None):
        """Procesa el callback de OAuth y registra la sesión del docente"""
        token_state = self.state_tokens.pop(state, None)
        if (not token_state or token_state["provider"] != provider
                or time.time() - token_state["created_at"] > 300):
            raise ValueError("Estado OAuth inválido o expirado")
        if is_sandbox or code == "sandbox":
            if os.getenv("EDUDOCENTE_SANDBOX", "0") != "1":
                raise ValueError("Modo sandbox deshabilitado")
            # Autenticación Sandbox para demostraciones y pruebas rápidas
            names = {
                "google": ("Docente Google Workspace", custom_email or "docente@colegiocastelgandolfo.cl"),
                "github": ("Docente GitHub Academic", custom_email or "docente@colegiocastelgandolfo.cl"),
                "microsoft": ("Docente Microsoft 365", custom_email or "docente@colegiocastelgandolfo.cl")
            }
            def_name, def_email = names.get(provider, ("Docente Castelgandolfo", "docente@colegiocastelgandolfo.cl"))
            name = custom_name or def_name
            email = custom_email or def_email

            self.current_user = {
                "authenticated": True,
                "provider": provider,
                "name": name,
                "email": email,
                "institution": "Colegio Castelgandolfo",
                "role": "Docente / Evaluador Institucional",
                "avatar": f"https://api.dicebear.com/7.x/bottts/svg?seed={email}",
                "classroom_connected": (provider == "google"),
                "connected_at": time.strftime("%Y-%m-%d %H:%M:%S")
            }
            return self.current_user

        # Flujo real de intercambio de código con el proveedor
        conf = OAUTH_CONFIGS.get(provider)
        if not conf:
            raise ValueError("Proveedor inválido")

        cid = os.getenv(conf["client_id_env"], "")
        csec = os.getenv(conf["client_secret_env"], "")

        token_data = {
            "client_id": cid,
            "client_secret": csec,
            "code": code,
            "grant_type": "authorization_code"
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"}
        req = urllib.request.Request(conf["token_url"], data=urllib.parse.urlencode(token_data).encode("utf-8"), headers=headers)
        
        try:
            with urllib.request.urlopen(req, timeout=15) as res:
                tokens = json.loads(res.read().decode("utf-8"))
                access_token = tokens.get("access_token")
        except Exception as e:
            raise RuntimeError(f"Error conectando a {provider}") from e

        # Obtener perfil del usuario
        user_headers = {"Authorization": f"Bearer {access_token}", "Accept": "application/json"}
        user_req = urllib.request.Request(conf["userinfo_url"], headers=user_headers)
        try:
            with urllib.request.urlopen(user_req, timeout=15) as res:
                info = json.loads(res.read().decode("utf-8"))
                self.current_user = {
                    "authenticated": True,
                    "provider": provider,
                    "name": info.get("name") or info.get("login") or "Docente Autenticado",
                    "email": info.get("email") or "docente@institucion.cl",
                    "institution": "Red Educativa Mineduc",
                    "role": "Profesor Titular",
                    "avatar": info.get("picture") or info.get("avatar_url") or "",
                    "classroom_connected": (provider == "google"),
                    "connected_at": time.strftime("%Y-%m-%d %H:%M:%S")
                }
                return self.current_user
        except Exception as e:
            raise RuntimeError(f"Error obteniendo perfil de {provider}") from e

    def logout(self):
        self.current_user = {
            "authenticated": False,
            "provider": None,
            "name": "Invitado",
            "email": "",
            "institution": "",
            "role": "",
            "avatar": "",
            "classroom_connected": False,
            "connected_at": ""
        }
        return self.current_user

auth_manager = AuthManager()
