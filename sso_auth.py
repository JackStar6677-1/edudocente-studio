"""Sesiones docentes emitidas exclusivamente por el panel institucional."""

import base64
import binascii
import hashlib
import hmac
import json
import os
import re
import secrets
import threading
import time
from http.cookies import SimpleCookie

SIGNING_KEY = os.environ.get("EDUDOCENTE_SSO_KEY", "")
SESSION_SECONDS = 8 * 60 * 60
TOKEN_SECONDS = 120
_lock = threading.RLock()
_sessions = {}
_used_nonces = {}
_downloads = {}


def _cookie_token(cookie_header):
    """Extrae el identificador opaco de sesión sin confiar en otros campos."""
    try:
        cookies = SimpleCookie(cookie_header or "")
        return cookies["edu_session"].value if "edu_session" in cookies else ""
    except Exception:
        return ""


def _signed_message(fields):
    """Codifica longitudes UTF-8 igual que el emisor PHP."""
    return b"".join(
        str(len(value.encode("utf-8"))).encode("ascii") + b":" + value.encode("utf-8")
        for value in fields
    )


def consume_sso_token(encoded):
    """Valida firma, audiencia y vigencia antes de consumir el nonce una vez."""
    if len(SIGNING_KEY) < 32:
        raise RuntimeError("SSO institucional no configurado")
    if not encoded or len(encoded) > 2048:
        raise ValueError("Token SSO ausente o demasiado largo")
    try:
        padded = encoded.replace("-", "+").replace("_", "/")
        padded += "=" * ((4 - len(padded) % 4) % 4)
        data = json.loads(base64.b64decode(padded, validate=True).decode("utf-8"))
        if not isinstance(data, dict):
            raise ValueError("Payload inválido")
        email = str(data.get("email", "")).strip().lower()
        role = str(data.get("role", "")).strip().lower()
        name = str(data.get("name", "")).strip()
        timestamp = int(data.get("ts", 0))
        nonce = str(data.get("nonce", ""))
        signature = str(data.get("sig", ""))
    except (ValueError, TypeError, UnicodeError, binascii.Error, json.JSONDecodeError) as exc:
        raise ValueError("Token SSO inválido") from exc

    if (data.get("ver") != 2 or data.get("aud") != "edudocente"
            or role not in ("admin", "docente")
            or not email.endswith("@colegiocastelgandolfo.cl")
            or not re.fullmatch(r"[0-9a-f]{32}", nonce)
            or not re.fullmatch(r"[0-9a-f]{64}", signature)
            or abs(time.time() - timestamp) > TOKEN_SECONDS):
        raise ValueError("Token SSO no autorizado o expirado")

    message = _signed_message((email, role, name, str(timestamp), nonce, "edudocente"))
    expected = hmac.new(SIGNING_KEY.encode("utf-8"), message, hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, signature):
        raise ValueError("Firma SSO inválida")

    with _lock:
        cutoff = time.time() - TOKEN_SECONDS * 2
        for prior, used_at in list(_used_nonces.items()):
            if used_at < cutoff:
                del _used_nonces[prior]
        if nonce in _used_nonces:
            raise ValueError("Token SSO ya utilizado")
        _used_nonces[nonce] = time.time()
    return {"email": email, "name": name, "role": role, "authenticated": True}


def create_session(user):
    """Crea una sesión independiente de la URL del pase SSO."""
    token = secrets.token_urlsafe(32)
    with _lock:
        _sessions[token] = (dict(user), time.time() + SESSION_SECONDS)
    return token


def get_session(cookie_header):
    """Resuelve la cookie HttpOnly y elimina sesiones expiradas."""
    token = _cookie_token(cookie_header)
    with _lock:
        record = _sessions.get(token)
        if not record:
            return None
        user, expires = record
        if expires < time.time():
            del _sessions[token]
            return None
        return dict(user)


def delete_session(cookie_header):
    """Revoca la sesión activa al cerrar sesión."""
    token = _cookie_token(cookie_header)
    with _lock:
        _sessions.pop(token, None)
        for key in list(_downloads):
            if key[0] == token:
                del _downloads[key]


def grant_download(cookie_header, job_id):
    """Asocia una generación con la sesión que la solicitó."""
    token = _cookie_token(cookie_header)
    with _lock:
        if not token or token not in _sessions:
            raise ValueError("Sesión no disponible")
        _downloads[(token, job_id)] = min(_sessions[token][1], time.time() + SESSION_SECONDS)


def can_download(cookie_header, job_id):
    """Comprueba que la generación pertenece a la sesión vigente."""
    token = _cookie_token(cookie_header)
    with _lock:
        expiry = _downloads.get((token, job_id), 0)
        return bool(token and expiry > time.time() and token in _sessions)
