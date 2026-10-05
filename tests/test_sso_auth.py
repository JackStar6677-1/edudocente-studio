"""Regresiones del pase institucional, sin servidor ni datos reales."""

import base64
import hashlib
import hmac
import json
import shutil
import subprocess
import time
import unittest
from http.cookies import SimpleCookie

import sso_auth


class SsoAuthTests(unittest.TestCase):
    """El pase debe ser firmado, de una sola audiencia y de un solo uso."""

    def setUp(self):
        self.previous_key = sso_auth.SIGNING_KEY
        sso_auth.SIGNING_KEY = "test-key-" * 8
        with sso_auth._lock:
            sso_auth._used_nonces.clear()
            sso_auth._sessions.clear()

    def tearDown(self):
        sso_auth.SIGNING_KEY = self.previous_key

    def token(self, **changes):
        """Firma un payload sintético con el protocolo PHP v2."""
        payload = {
            "email": "test@colegiocastelgandolfo.cl",
            "role": "docente",
            "name": "Docente Ñ",
            "ts": int(time.time()),
            "nonce": "a" * 32,
            "aud": "edudocente",
            "ver": 2,
        }
        payload.update(changes)
        fields = tuple(str(payload[key]) for key in ("email", "role", "name", "ts", "nonce", "aud"))
        payload["sig"] = hmac.new(
            sso_auth.SIGNING_KEY.encode(), sso_auth._signed_message(fields), hashlib.sha256
        ).hexdigest()
        return base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")

    def test_pase_valido_y_replay(self):
        token = self.token()
        self.assertEqual(sso_auth.consume_sso_token(token)["email"], "test@colegiocastelgandolfo.cl")
        with self.assertRaises(ValueError):
            sso_auth.consume_sso_token(token)

    def test_firma_invalida_no_quema_nonce(self):
        token = self.token()
        decoded = json.loads(base64.urlsafe_b64decode(token + "=="))
        decoded["sig"] = "0" * 64
        invalid = base64.urlsafe_b64encode(json.dumps(decoded).encode()).decode().rstrip("=")
        with self.assertRaises(ValueError):
            sso_auth.consume_sso_token(invalid)
        self.assertTrue(sso_auth.consume_sso_token(token)["authenticated"])

    def test_audiencia_y_vigencia(self):
        for token in (self.token(aud="castelboard"), self.token(ts=int(time.time()) - 300)):
            with self.assertRaises(ValueError):
                sso_auth.consume_sso_token(token)

    def test_sesion_solo_con_cookie(self):
        token = sso_auth.create_session({"email": "test@colegiocastelgandolfo.cl"})
        cookie = SimpleCookie()
        cookie["edu_session"] = token
        self.assertIsNone(sso_auth.get_session(""))
        self.assertEqual(sso_auth.get_session(cookie.output(header=""))["email"], "test@colegiocastelgandolfo.cl")
        sso_auth.delete_session(cookie.output(header=""))
        self.assertIsNone(sso_auth.get_session(cookie.output(header="")))

    @unittest.skipUnless(shutil.which("php"), "PHP no disponible")
    def test_pase_emitido_por_php(self):
        """Comprueba longitud UTF-8 y HMAC del emisor PHP real."""
        script = '''$key="test-key-" . str_repeat("test-key-", 7);
        $p=["email"=>"test@colegiocastelgandolfo.cl","role"=>"docente","name"=>"Docente Ñ","ts"=>time(),"nonce"=>str_repeat("b",32),"aud"=>"edudocente","ver"=>2];
        $message=""; foreach (["email","role","name","ts","nonce","aud"] as $field) { $value=(string)$p[$field]; $message.=strlen($value).":".$value; }
        $p["sig"]=hash_hmac("sha256",$message,$key);
        echo rtrim(strtr(base64_encode(json_encode($p,JSON_UNESCAPED_UNICODE)),"+/","-_"),"=");'''
        result = subprocess.run(["php", "-r", script], capture_output=True, text=True, check=True, timeout=5)
        self.assertEqual(sso_auth.consume_sso_token(result.stdout)["name"], "Docente Ñ")


if __name__ == "__main__":
    unittest.main()
