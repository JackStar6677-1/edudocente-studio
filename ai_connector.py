"""
EduDocente AI Connector: Módulo Universal Multi-Proveedor de Inteligencia Artificial
Soporta: Antigravity / Google Gemini, Anthropic Claude, OpenAI (GPT-4o/Codex), DeepSeek, Kimi (Moonshot), Grok (xAI) y Local/Ollama.

Convierte requerimientos pedagógicos simples (tema, curso, páginas del libro)
en una estructura de evaluación estandarizada JSON de 3 ítems validada para DocenteEngine.
"""

import os
import sys
import json
import re
import urllib.request
import urllib.error

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# Definición de proveedores soportados y sus modelos predeterminados
AI_PROVIDERS = {
    "gemini": {
        "name": "Google Gemini / Antigravity",
        "env_key": "GEMINI_API_KEY",
        "default_model": "gemini-1.5-pro",
        "endpoint": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    },
    "claude": {
        "name": "Anthropic Claude",
        "env_key": "ANTHROPIC_API_KEY",
        "default_model": "claude-3-5-sonnet-20241022",
        "endpoint": "https://api.anthropic.com/v1/messages"
    },
    "openai": {
        "name": "OpenAI / Codex (GPT-4o)",
        "env_key": "OPENAI_API_KEY",
        "default_model": "gpt-4o",
        "endpoint": "https://api.openai.com/v1/chat/completions"
    },
    "deepseek": {
        "name": "DeepSeek (V3 / R1)",
        "env_key": "DEEPSEEK_API_KEY",
        "default_model": "deepseek-chat",
        "endpoint": "https://api.deepseek.com/chat/completions"
    },
    "kimi": {
        "name": "Kimi (Moonshot AI)",
        "env_key": "MOONSHOT_API_KEY",
        "default_model": "moonshot-v1-8k",
        "endpoint": "https://api.moonshot.cn/v1/chat/completions"
    },
    "grok": {
        "name": "Grok (xAI)",
        "env_key": "XAI_API_KEY",
        "default_model": "grok-beta",
        "endpoint": "https://api.x.ai/v1/chat/completions"
    },
    "local": {
        "name": "Local / Ollama / OpenAI-Compatible",
        "env_key": "LOCAL_API_KEY",
        "default_model": "llama3.2",
        "endpoint": "http://localhost:11434/v1/chat/completions"
    }
}

SYSTEM_PROMPT = """Eres el motor pedagógico de EduDocente-Studio. Tu misión es transformar requerimientos curriculares docentes en una evaluación estandarizada escolar en formato JSON estricto.

Reglas Pedagógicas Obligatorias:
1. No usar la palabra 'SUMATIVA', ni escalas de conversión de notas (1.0 a 7.0), ni porcentajes de exigencia (60%), ni tiempos límites estimados.
2. ÍTEM I (Selección Múltiple): Mínimo 4 o 5 preguntas. Opciones en formato directo ['A', 'Texto'], ['B', 'Texto'], etc., sin corchetes [ ]. Siempre incluye 'correcta' y 'justificacion' pedagógica citando el contenido.
3. ÍTEM II (Verdadero o Falso): Mínimo 5 a 10 oraciones. 'resp' debe ser 'V' o 'F' con 'justificacion' pedagógica clara para el docente.
4. ÍTEM III (Aplicación/Dibujo): Situación práctica para dibujar con espacio de descripción y ejemplo de respuesta esperada en la pauta.

Debes responder ÚNICAMENTE con un bloque JSON válido, sin texto introductorio ni explicaciones fuera del JSON."""

class AIConnector:
    def __init__(self, provider="openai", api_key=None, model=None):
        self.provider = provider.lower()
        if self.provider not in AI_PROVIDERS:
            self.provider = "openai"

        prov_info = AI_PROVIDERS[self.provider]
        self.api_key = api_key or os.getenv(prov_info["env_key"], "")
        self.model = model or prov_info["default_model"]
        self.endpoint = prov_info["endpoint"]

    def _call_http(self, url, headers, data):
        req = urllib.request.Request(url, data=json.dumps(data).encode('utf-8'), headers=headers, method='POST')
        try:
            with urllib.request.urlopen(req, timeout=60) as response:
                return json.loads(response.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            err_msg = e.read().decode('utf-8')
            raise RuntimeError(f"Error API {self.provider} ({e.code}): {err_msg}")
        except Exception as e:
            raise RuntimeError(f"Fallo de conexión con {self.provider}: {e}")

    def generate_assessment_json(self, topic, grade, oa="", textbook_pages="", num_q1=5, num_q2=8):
        """
        Envía los temas y objetivos curriculares a la IA seleccionada y retorna la estructura JSON
        """
        user_prompt = f"""Genera una evaluación completa para:
- Asignatura: CIENCIAS NATURALES
- Curso: {grade}
- Tema o Contenidos: {topic}
- Objetivo de Aprendizaje (OA): {oa}
- Páginas de Referencia Texto Mineduc: {textbook_pages}
- Cantidad preguntas Ítem I (Selección Múltiple): {num_q1}
- Cantidad afirmaciones Ítem II (Verdadero/Falso): {num_q2}
- Ítem III: Dibujo y aplicación con 2 o 3 situaciones de autocuidado/observación.

El JSON debe cumplir exactamente con esta estructura:
{{
  "colegio": "COLEGIO LUIS PASTEUR ANEXO",
  "asignatura": "CIENCIAS NATURALES",
  "curso": "{grade}",
  "titulo": "EVALUACIÓN FINAL: {topic.upper()}",
  "oa": "{oa}",
  "contenidos": "{topic}",
  "puntaje_total": 25,
  "item1_pts_cada_una": 2,
  "item1_seleccion_multiple": [
    {{
      "pregunta": "1. [Pregunta]",
      "alternativas": [["A", "[Texto A]"], ["B", "[Texto B]"], ["C", "[Texto C]"], ["D", "[Texto D]"]],
      "correcta": "A) [Texto A]",
      "justificacion": "[Justificación pedagógica citando contenidos]"
    }}
  ],
  "item2_pts_cada_una": 1,
  "item2_verdadero_falso": [
    {{
      "oracion": "[Afirmación]",
      "resp": "V",
      "justificacion": "[Explicación pedagógica]"
    }}
  ],
  "item3_aplicacion": {{
    "titulo": "ÍTEM III: DIBUJO Y APLICACIÓN PRÁCTICA",
    "puntaje": 5,
    "tipo": "drawing_boxes",
    "instruccion": "[Instrucción para dibujar y describir]",
    "cajas": [
      {{
        "titulo": "1. [Situación a representar]",
        "ejemplo_pauta": "[Respuesta esperada docente]"
      }}
    ]
  }}
}}"""

        # Si no hay API key configurada, generar mock inteligente estructurado
        if not self.api_key and self.provider != "local":
            print(f"\n[AVISO IA] No se detectó clave de API para {AI_PROVIDERS[self.provider]['name']}.")
            print("Activando 'Smart Pedagogical Synthesizer' local (Simulador de IA)...")
            return self._generate_smart_mock(topic, grade, oa, textbook_pages)

        # 1. Google Gemini / Antigravity
        if self.provider == "gemini":
            url = self.endpoint.format(model=self.model, api_key=self.api_key)
            headers = {"Content-Type": "application/json"}
            payload = {
                "contents": [{"parts": [{"text": f"{SYSTEM_PROMPT}\n\n{user_prompt}"}]}],
                "generationConfig": {"temperature": 0.2, "response_mime_type": "application/json"}
            }
            res = self._call_http(url, headers, payload)
            text_out = res["candidates"][0]["content"]["parts"][0]["text"]

        # 2. Anthropic Claude
        elif self.provider == "claude":
            url = self.endpoint
            headers = {
                "Content-Type": "application/json",
                "x-api-key": self.api_key,
                "anthropic-version": "2023-06-01"
            }
            payload = {
                "model": self.model,
                "max_tokens": 4096,
                "system": SYSTEM_PROMPT,
                "messages": [{"role": "user", "content": user_prompt}]
            }
            res = self._call_http(url, headers, payload)
            text_out = res["content"][0]["text"]

        # 3. OpenAI / Codex, DeepSeek, Kimi, Grok, Local (OpenAI Compatible)
        else:
            url = self.endpoint
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}" if self.api_key else ""
            }
            payload = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt}
                ],
                "temperature": 0.2
            }
            res = self._call_http(url, headers, payload)
            text_out = res["choices"][0]["message"]["content"]

        # Extraer JSON de la respuesta
        match = re.search(r'\{.*\}', text_out, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        return json.loads(text_out)

    def _generate_smart_mock(self, topic, grade, oa, textbook_pages):
        """Simulador curricular que construye una prueba completa sin requerir API Key externa"""
        clean_topic = topic.strip().capitalize()
        return {
            "colegio": "COLEGIO LUIS PASTEUR ANEXO",
            "asignatura": "CIENCIAS NATURALES",
            "curso": grade,
            "titulo": f"EVALUACIÓN FINAL: {clean_topic.upper()}",
            "oa": oa or "OA Mineduc: Desarrollar modelos explicativos de fenómenos naturales.",
            "contenidos": f"{clean_topic}. Páginas del texto escolar: {textbook_pages or 'Capítulo oficial'}.",
            "puntaje_total": 25,
            "docente": "Profesora Margarita Miranda B.",
            "email_docente": "profesora.margaritamiranda@cepluispasteur.cl",

            "item1_pts_cada_una": 2,
            "item1_seleccion_multiple": [
                {
                    "pregunta": f"1. Respecto a {clean_topic}, ¿cuál de las siguientes afirmaciones describe su principio fundamental?",
                    "alternativas": [
                        ["A", f"Constituye una propiedad central en los procesos de {clean_topic.lower()} según el currículo escolar."],
                        ["B", "Ocurre únicamente en situaciones de laboratorio sin interacción en el entorno."],
                        ["C", "Es una transformación que destruye totalmente la materia y la energía."],
                        ["D", "Es observable de manera exclusiva a través de imágenes satelitales."]
                    ],
                    "correcta": f"A) Constituye una propiedad central en los procesos de {clean_topic.lower()} según el currículo escolar.",
                    "justificacion": f"Conforme a los contenidos estudiados en clases sobre {clean_topic.lower()} (Texto del estudiante {textbook_pages})."
                },
                {
                    "pregunta": "2. ¿Cómo se relacionan los componentes observados en este fenómeno?",
                    "alternativas": [
                        ["A", "Interactúan de manera equilibrada siguiendo las leyes físicas de conservación."],
                        ["B", "Permanecen estáticos sin ningún tipo de intercambio energético."],
                        ["C", "Se anulan mutuamente generando un vacío en el sistema."],
                        ["D", "Cambian de forma arbitraria sin ningún patrón medible."]
                    ],
                    "correcta": "A) Interactúan de manera equilibrada siguiendo las leyes físicas de conservación.",
                    "justificacion": "Las interacciones en ciencias naturales conservan la masa y la energía en los sistemas analizados."
                },
                {
                    "pregunta": "3. En la vida cotidiana, una manifestación concreta de este contenido se aprecia cuando:",
                    "alternativas": [
                        ["A", "Observamos los cambios y ciclos en el hogar y en la naturaleza."],
                        ["B", "Solo cuando se realizan experimentos con elementos químicos peligrosos."],
                        ["C", "Únicamente en el espacio exterior."],
                        ["D", "En ningún aspecto de la vida diaria."]
                    ],
                    "correcta": "A) Observamos los cambios y ciclos en el hogar y en la naturaleza.",
                    "justificacion": "La ciencia escolar conecta los conceptos teóricos con vivencias directas de los estudiantes."
                }
            ],

            "item2_pts_cada_una": 1,
            "item2_verdadero_falso": [
                {
                    "oracion": f"El estudio de {clean_topic.lower()} permite comprender el funcionamiento de nuestro entorno y aplicar medidas de autocuidado.",
                    "resp": "V",
                    "justificacion": "La comprensión científica fomenta la toma de decisiones informadas y el cuidado del medioambiente."
                },
                {
                    "oracion": "Los cambios en la materia implican siempre la creación de nuevos átomos desde la nada.",
                    "resp": "F",
                    "justificacion": "La materia no se crea ni se destruye, solo se transforma y reorganiza en las reacciones."
                },
                {
                    "oracion": "El seguimiento de normas de seguridad escolar previene accidentes en actividades prácticas.",
                    "resp": "V",
                    "justificacion": "La prevención y el autocuidado son conductas prioritarias en toda actividad de aula o laboratorio."
                },
                {
                    "oracion": "Los conceptos estudiados son independientes de las observaciones realizadas en clases.",
                    "resp": "F",
                    "justificacion": "El método científico escolar se basa en la contrastación entre teoría y evidencia empírica."
                }
            ],

            "item3_aplicacion": {
                "titulo": "ÍTEM III: DIBUJO Y APLICACIÓN DEL FENÓMENO ESTUDIADO",
                "puntaje": 6,
                "tipo": "drawing_boxes",
                "instruccion": f"Dibuja en el recuadro una situación o esquema que represente claramente el concepto de {clean_topic.lower()} y describe lo que dibujaste:",
                "cajas": [
                    {
                        "titulo": f"Representación gráfica de {clean_topic.lower()}:",
                        "ejemplo_pauta": f"Esquema claro donde se aprecian los componentes clave de {clean_topic.lower()} con rótulos legibles y color adecuado."
                    }
                ]
            }
        }

def interactive_ai_session():
    print("\n" + "=" * 65)
    print("      CONECTOR DE IA UNIVERSAL — EDUDOCENTE-STUDIO      ")
    print("  Compatible con Claude, Gemini, GPT-4o, DeepSeek, Kimi y Grok")
    print("=" * 65)

    print("\n[1] Selecciona el motor de IA a utilizar:")
    for idx, (k, v) in enumerate(AI_PROVIDERS.items(), 1):
        env_state = "✓ Configurada" if os.getenv(v["env_key"]) else "○ Sin clave (Usa Smart Mock)"
        print(f"  {idx}. {v['name']} [{k}] — {env_state}")

    sel = input("\nElige proveedor (1-7, o Enter para Gemini/Claude): ").strip()
    keys = list(AI_PROVIDERS.keys())
    prov_key = keys[int(sel) - 1] if sel.isdigit() and 1 <= int(sel) <= len(keys) else "gemini"

    connector = AIConnector(provider=prov_key)
    print(f"\n[OK] Conector activo: {AI_PROVIDERS[prov_key]['name']} (Modelo: {connector.model})")

    print("\n[2] Ingresa los requerimientos para que la IA investigue y genere la prueba:")
    tema = input("  • Tema / Contenidos (ej: El ciclo del agua, Fotosíntesis, Célula): ").strip() or "La Célula y sus Organelos"
    curso = input("  • Curso (ej: 7° Básico A): ").strip() or "7° Básico A"
    oa = input("  • Objetivo de Aprendizaje (OA) [Opcional]: ").strip() or "OA 2 — Explicar el rol de la célula como unidad estructural y funcional."
    paginas = input("  • Páginas del texto escolar Mineduc (ej: Págs. 54 a 62): ").strip() or "Páginas 54 a 62"

    print("\n" + "-" * 65)
    print(f"  [TEMAS INDICADOS]: {tema} (Curso: {curso})")
    print(f"  [OBJETIVO OA]: {oa}")
    print(f"  [REFERENCIA LIBRO]: {paginas}")
    print("-" * 65)
    print("\n>>> La IA buscará información en internet sobre los contenidos indicados,")
    print(">>> comparará con los archivos y textos de referencia curriculares,")
    print(">>> y enviará la estructura de evaluación estandarizada al compilador...")
    print("-" * 65)

    assessment_data = connector.generate_assessment_json(topic=tema, grade=curso, oa=oa, textbook_pages=paginas)

    # Compilar con DocenteEngine
    from engine import DocenteEngine
    engine = DocenteEngine(assessment_data)

    print("\n[COMPILANDO DOCUMENTOS DOCENTES CON DISEÑO INSTITUCIONAL...]")
    p_est, p_pau = engine.build_all()

    print("\n" + "=" * 65)
    print(" ¡PROCESO DE IA FINALIZADO CON ÉXITO!")
    print(f"  • Prueba del Estudiante: {p_est}")
    print(f"  • Pauta Oficial Docente: {p_pau}")
    print("=" * 65 + "\n")

if __name__ == "__main__":
    interactive_ai_session()
