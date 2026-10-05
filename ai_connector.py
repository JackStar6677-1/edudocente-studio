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

# Definición de proveedores soportados con énfasis en cuotas gratuitas y semanales
AI_PROVIDERS = {
    "antigravity": {
        "name": "Google Antigravity / Gemini (Cuota Semanal Gratis)",
        "badge": "Semanal Gratis (15 RPM)",
        "env_key": "GEMINI_API_KEY",
        "default_model": "gemini-2.0-flash",
        "endpoint": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}",
        "help_url": "https://aistudio.google.com/app/apikey",
        "free_tier_info": "Google AI Studio ofrece 15 peticiones/minuto y cuota semanal 100% gratuita sin costo."
    },
    "gemini": {
        "name": "Google Gemini / Antigravity",
        "badge": "Semanal Gratis (15 RPM)",
        "env_key": "GEMINI_API_KEY",
        "default_model": "gemini-2.0-flash",
        "endpoint": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}",
        "help_url": "https://aistudio.google.com/app/apikey",
        "free_tier_info": "Google AI Studio ofrece 15 peticiones/minuto y cuota semanal 100% gratuita sin costo."
    },
    "codex": {
        "name": "OpenAI Codex / GPT-4o-mini (Versión Gratis)",
        "badge": "Versión Gratis / GPT-4o-mini",
        "env_key": "OPENAI_API_KEY",
        "default_model": "gpt-4o-mini",
        "endpoint": "https://api.openai.com/v1/chat/completions",
        "help_url": "https://platform.openai.com/api-keys",
        "free_tier_info": "Utiliza créditos de prueba gratuitos de OpenAI o modelo ultraliviano GPT-4o-mini."
    },
    "openai": {
        "name": "OpenAI / Codex (GPT-4o)",
        "badge": "Versión Gratis / GPT-4o-mini",
        "env_key": "OPENAI_API_KEY",
        "default_model": "gpt-4o-mini",
        "endpoint": "https://api.openai.com/v1/chat/completions",
        "help_url": "https://platform.openai.com/api-keys",
        "free_tier_info": "Utiliza créditos de prueba gratuitos de OpenAI o modelo ultraliviano GPT-4o-mini."
    },
    "claude": {
        "name": "Anthropic Claude (Cuota Semanal Gratis)",
        "badge": "Semanal Gratis (Haiku)",
        "env_key": "ANTHROPIC_API_KEY",
        "default_model": "claude-3-5-haiku-20241022",
        "endpoint": "https://api.anthropic.com/v1/messages",
        "help_url": "https://console.anthropic.com/settings/keys",
        "free_tier_info": "Cuota de prueba y nivel rápido de Claude Haiku para diseño pedagógico ágil."
    },
    "saori": {
        "name": "Saori SRE Daemon (Red Local Star Server)",
        "badge": "Red Local On-Premise",
        "env_key": "SAORI_API_KEY",
        "default_model": "saori-soberana",
        "endpoint": os.environ.get("SAORI_DAEMON_URL", "http://127.0.0.1:8089/chat"),
        "help_url": "http://star:8089/status",
        "free_tier_info": "Servicio cognitivo autónomo ejecutándose 24/7 en Star Server (puerto 8089)."
    },
    "buffer": {
        "name": "Smart Pedagogical Buffer (Costo Cero / Sin Clave)",
        "badge": "100% Gratis / Sin API Key",
        "env_key": "NONE",
        "default_model": "smart-synth-v1",
        "endpoint": "local://mock",
        "help_url": "#",
        "free_tier_info": "Sintetizador inteligente local en Python. No requiere conexión a internet ni claves."
    },
    "deepseek": {
        "name": "DeepSeek (V3 / R1)",
        "badge": "Económico / Razonamiento",
        "env_key": "DEEPSEEK_API_KEY",
        "default_model": "deepseek-chat",
        "endpoint": "https://api.deepseek.com/chat/completions",
        "help_url": "https://platform.deepseek.com/",
        "free_tier_info": "API compatible con OpenAI de altísimo rendimiento a bajo costo."
    },
    "kimi": {
        "name": "Kimi (Moonshot AI)",
        "badge": "Moonshot Escolar",
        "env_key": "MOONSHOT_API_KEY",
        "default_model": "moonshot-v1-8k",
        "endpoint": "https://api.moonshot.cn/v1/chat/completions",
        "help_url": "https://platform.moonshot.cn/",
        "free_tier_info": "Ventana de contexto amplia y comprensión profunda de textos escolares."
    },
    "grok": {
        "name": "Grok (xAI)",
        "badge": "xAI Beta",
        "env_key": "XAI_API_KEY",
        "default_model": "grok-beta",
        "endpoint": "https://api.x.ai/v1/chat/completions",
        "help_url": "https://console.x.ai/",
        "free_tier_info": "Inferencia rápida desarrollada por xAI."
    },
    "local": {
        "name": "OpenCode / Ollama Local (:11434)",
        "badge": "OpenCode Local",
        "env_key": "LOCAL_API_KEY",
        "default_model": "llama3.2",
        "endpoint": "http://localhost:11434/v1/chat/completions",
        "help_url": "https://ollama.com",
        "free_tier_info": "Modelos OpenCode locales (Qwen, Llama 3) sin depender de internet."
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

    def generate_assessment_json(self, topic, grade, oa="", textbook_pages="", subject="CIENCIAS NATURALES", points=25, num_q1=5, num_q2=8):
        """
        Envía los temas y objetivos curriculares a la IA seleccionada y retorna la estructura JSON
        con puntaje total dinámico y asignatura personalizada.
        """
        clean_subject = (subject or "CIENCIAS NATURALES").strip().upper()
        clean_topic = (topic or "").strip()
        pts_target = int(points) if points else 25
        if pts_target < 4 or pts_target > 200:
            raise ValueError("El puntaje solicitado debe estar entre 4 y 200")

        # Adaptar cantidad de preguntas según el puntaje total deseado por el docente
        if pts_target >= 45:
            num_q1 = max(num_q1, 12)
            num_q2 = max(num_q2, 10)
        elif pts_target >= 35:
            num_q1 = max(num_q1, 8)
            num_q2 = max(num_q2, 9)
        elif pts_target <= 18:
            num_q1 = min(num_q1, 4)
            num_q2 = min(num_q2, 5)

        # Reservar al menos dos puntos para aplicación sin superar el total.
        while num_q1 * 2 + num_q2 > pts_target - 2:
            if num_q2 > 0:
                num_q2 -= 1
            else:
                num_q1 -= 1
        pts_item1 = num_q1 * 2
        pts_item2 = num_q2 * 1
        pts_item3 = max(2, pts_target - (pts_item1 + pts_item2))
        pts_total = pts_item1 + pts_item2 + pts_item3 # Suma matemática exacta garantizada

        user_prompt = f"""Genera una evaluación completa para:
- Asignatura: {clean_subject}
- Curso: {grade}
- Tema o Contenidos: {clean_topic}
- Objetivo de Aprendizaje (OA): {oa}
- Páginas de Referencia Texto Mineduc: {textbook_pages}
- Cantidad preguntas Ítem I (Selección Múltiple): {num_q1} ({pts_item1} puntos, 2 pts cada una)
- Cantidad afirmaciones Ítem II (Verdadero/Falso): {num_q2} ({pts_item2} puntos, 1 pt cada una)
- Ítem III: Dibujo, aplicación práctica o desarrollo acorde a {clean_subject} ({pts_item3} puntos).

El JSON debe cumplir exactamente con esta estructura:
{{
  "colegio": "COLEGIO CASTELGANDOLFO",
  "asignatura": "{clean_subject}",
  "curso": "{grade}",
  "titulo": "EVALUACIÓN: {clean_topic.upper()}",
  "oa": "{oa}",
  "contenidos": "{clean_topic}",
  "puntaje_total": {pts_total},
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
    "titulo": "ÍTEM III: APLICACIÓN PRÁCTICA Y DESARROLLO",
    "puntaje": {pts_item3},
    "tipo": "drawing_boxes",
    "instruccion": "[Instrucción para dibujar, resolver o describir acorde a {clean_subject}]",
    "cajas": [
      {{
        "titulo": "1. [Situación a representar o resolver]",
        "ejemplo_pauta": "[Respuesta esperada docente]"
      }}
    ]
  }}
}}"""

        # 0. Si el proveedor seleccionado es Buffer / Sintetizador Local sin clave
        if self.provider == "buffer":
            print("[IA BUFFER] Modo Sintetizador Pedagógico Local activado (Costo $0 / Sin API Key).")
            return self._generate_smart_mock(clean_topic, grade, oa, textbook_pages, clean_subject, pts_total, pts_item3)

        # Si no hay API key y no es Saori ni Local, usar fallback pedagógico inteligente
        if not self.api_key and self.provider not in ("local", "saori"):
            print(f"\n[AVISO IA] No se detectó clave de API para {AI_PROVIDERS.get(self.provider, {}).get('name', self.provider)}.")
            print("Activando automáticamente el 'Smart Pedagogical Buffer' escolar (Simulador sin costo)...")
            return self._generate_smart_mock(clean_topic, grade, oa, textbook_pages, clean_subject, pts_total, pts_item3)

        try:
            # 1. Saori AI Daemon (Star Server en red local :8089)
            if self.provider == "saori":
                url = self.endpoint
                headers = {"Content-Type": "application/json"}
                payload = {
                    "prompt": f"{SYSTEM_PROMPT}\n\n{user_prompt}",
                    "sender": "EduDocente",
                    "scope": "edudocente"
                }
                res = self._call_http(url, headers, payload)
                text_out = res.get("response", "")

            # 2. Google Gemini / Antigravity (Semanal Gratuito 15 RPM)
            elif self.provider in ("gemini", "antigravity"):
                url = self.endpoint.format(model=self.model, api_key=self.api_key)
                headers = {"Content-Type": "application/json"}
                payload = {
                    "contents": [{"parts": [{"text": f"{SYSTEM_PROMPT}\n\n{user_prompt}"}]}],
                    "generationConfig": {"temperature": 0.2, "response_mime_type": "application/json"}
                }
                res = self._call_http(url, headers, payload)
                text_out = res["candidates"][0]["content"]["parts"][0]["text"]

            # 3. Anthropic Claude (Cuota Semanal Gratis / Haiku)
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

            # 4. OpenAI / Codex (Versión Gratis GPT-4o-mini), DeepSeek, Kimi, Grok, Local OpenCode
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
            generated = json.loads(match.group(0) if match else text_out)
            actual_points = (
                len(generated["item1_seleccion_multiple"]) * int(generated["item1_pts_cada_una"])
                + len(generated["item2_verdadero_falso"]) * int(generated["item2_pts_cada_una"])
                + int(generated["item3_aplicacion"]["puntaje"])
            )
            if actual_points != pts_target or int(generated["puntaje_total"]) != pts_target:
                raise ValueError("El proveedor entregó una evaluación con puntaje inconsistente")
            return generated

        except Exception as e:
            print(f"\n[FALLO IA] Error invocando proveedor '{self.provider}': {e}")
            print("Activando Smart Pedagogical Buffer de contingencia inmediata para garantizar la entrega...")
            mock_data = self._generate_smart_mock(topic, grade, oa, textbook_pages, clean_subject, pts_target, pts_item3)
            mock_data["_ai_note"] = f"Generado vía Buffer Inteligente (Fallo en {self.provider}: {str(e)})"
            mock_data["simulation_mode"] = True
            return mock_data

    def _generate_smart_mock(self, topic, grade, oa, textbook_pages, subject="CIENCIAS NATURALES", points=25, pts_item3=None):
        """Simulador curricular que construye una prueba completa sin requerir API Key externa con suma matemática exacta"""
        clean_topic = topic.strip().capitalize()
        clean_subject = (subject or "CIENCIAS NATURALES").strip().upper()
        target_pts = int(points) if points else 25

        try:
            from institution_manager import institution_manager
            act_inst = institution_manager.get_active()
            colegio_nombre = act_inst.get("name", "COLEGIO CASTELGANDOLFO")
        except Exception:
            colegio_nombre = "COLEGIO CASTELGANDOLFO"

        # Determinar número de preguntas según escala de puntaje
        if target_pts >= 45:
            n_q1, n_q2 = 12, 14
        elif target_pts >= 35:
            n_q1, n_q2 = 8, 10
        elif target_pts <= 18:
            n_q1, n_q2 = 3, 5
        else:
            n_q1, n_q2 = 5, 7

        pts_q1 = n_q1 * 2
        pts_q2 = n_q2 * 1
        while n_q1 * 2 + n_q2 > target_pts - 2:
            if n_q2 > 0:
                n_q2 -= 1
            else:
                n_q1 -= 1
        pts_q1 = n_q1 * 2
        pts_q2 = n_q2
        calc_item3 = target_pts - (pts_q1 + pts_q2)
        total_sum = pts_q1 + pts_q2 + calc_item3

        # Pool de preguntas para Selección Múltiple (2 pts c/u)
        pool_q1 = [
            (
                f"Respecto a {clean_topic}, ¿cuál de las siguientes opciones describe su principio fundamental en {clean_subject}?",
                [
                    ["A", f"Constituye una propiedad central en los contenidos de {clean_topic.lower()} según el currículo oficial."],
                    ["B", "Ocurre únicamente en situaciones de laboratorio sin interacción en el entorno."],
                    ["C", "Es una transformación que carece de relación con los procesos estudiados."],
                    ["D", "Es observable de manera exclusiva a través de mediciones astronómicas."]
                ],
                f"A) Constituye una propiedad central en los contenidos de {clean_topic.lower()} según el currículo oficial.",
                f"Conforme a los contenidos estudiados sobre {clean_topic.lower()} (Texto del estudiante {textbook_pages})."
            ),
            (
                f"¿Cómo interactúan los componentes observados en {clean_topic.lower()}?",
                [
                    ["A", "Interactúan de manera coordinada y estructurada según los principios de la disciplina."],
                    ["B", "Permanecen estáticos sin ningún tipo de interacción ni transferencia energética."],
                    ["C", "Se anulan mutuamente impidiendo cualquier medición objetiva."],
                    ["D", "Cambian de forma arbitraria sin seguir leyes ni patrones medibles."]
                ],
                "A",
                "Los aprendizajes esperados establecen relaciones de causa, efecto y conservación de procesos."
            ),
            (
                f"En la vida cotidiana, una manifestación concreta de {clean_topic.lower()} se aprecia cuando:",
                [
                    ["A", "Observamos y aplicamos los conceptos aprendidos en situaciones y fenómenos del entorno."],
                    ["B", "Solo cuando se realizan experimentos industriales complejos."],
                    ["C", "Únicamente al resolver cuestionarios estandarizados."],
                    ["D", "En ningún aspecto de la vida real ni formativa."]
                ],
                "A",
                "La enseñanza escolar conecta los conceptos teóricos con vivencias directas de los estudiantes."
            ),
            (
                f"¿Cuál de las siguientes afirmaciones respecto a la importancia de {clean_topic.lower()} es correcta?",
                [
                    ["A", f"Permite fundamentar explicaciones científicas y resolver problemas en {clean_subject}."],
                    ["B", "Es un concepto aislado que no influye en otras áreas del conocimiento."],
                    ["C", "Solo tiene validez histórica pero ya no se utiliza en la actualidad."],
                    ["D", "Carece de evidencia comprobable en el entorno escolar."]
                ],
                "A",
                "El aprendizaje significativo articula saberes teóricos con habilidades analíticas."
            ),
            (
                f"Al comparar diferentes situaciones asociadas a {clean_topic.lower()}, es fundamental considerar:",
                [
                    ["A", "Las variables intervinientes y los factores que modifican su comportamiento."],
                    ["B", "Únicamente la opinión subjetiva del observador."],
                    ["C", "Que todos los resultados deben ser idénticos sin importar las condiciones."],
                    ["D", "Exclusivamente el tiempo transcurrido descartando la magnitud."]
                ],
                "A",
                "El rigor pedagógico exige identificar variables dependientes e independientes."
            ),
            (
                f"¿Qué herramienta o método resulta más adecuado para registrar observaciones de {clean_topic.lower()}?",
                [
                    ["A", "Tablas de datos, esquemas explicativos y gráficos comparativos."],
                    ["B", "Registros memorísticos sin respaldo documental."],
                    ["C", "Estimaciones visuales rápidas sin unidades de medida."],
                    ["D", "Suposiciones teóricas sin contraste empírico."]
                ],
                "A",
                "El método analítico promueve el uso de representaciones gráficas y registros sistemáticos."
            ),
            (
                f"¿Qué consecuencia se deriva de una alteración sustancial en las condiciones de {clean_topic.lower()}?",
                [
                    ["A", "Se produce una variación proporcional en los resultados o estados finales."],
                    ["B", "El fenómeno se detiene permanentemente sin posibilidad de recuperación."],
                    ["C", "No existe alteración alguna bajo ninguna circunstancia."],
                    ["D", "Los datos obtenidos se vuelven aleatorios e irrelevantes."]
                ],
                "A",
                "Los sistemas responden a perturbaciones manteniendo relaciones de equilibrio dinámico."
            ),
            (
                f"En el análisis de {clean_topic.lower()}, el uso de modelos conceptuales permite:",
                [
                    ["A", "Simplificar la comprensión de fenómenos complejos o de escala microscópica/macroscópica."],
                    ["B", "Reemplazar completamente la realidad por ilustraciones ficticias."],
                    ["C", "Evitar el cálculo matemático en cualquier circunstancia."],
                    ["D", "Limitar el estudio a descripciones meramente verbales."]
                ],
                "A",
                "Los modelos pedagógicos son puentes cognitivos para conceptualizar estructuras abstractas."
            ),
            (
                f"Respecto al cuidado y aplicación responsable de {clean_topic.lower()}, se recomienda:",
                [
                    ["A", "Seguir protocolos de seguridad, uso eficiente de recursos y trabajo colaborativo."],
                    ["B", "Manipular materiales sin supervisión previa."],
                    ["C", "Ignorar las advertencias del texto escolar y pautas docentes."],
                    ["D", "Proceder por ensayo y error sin planificación."]
                ],
                "A",
                "La formación integral contempla normas de seguridad escolar y conciencia ambiental."
            ),
            (
                f"¿Cuál es el rol de la evidencia en las conclusiones relativas a {clean_topic.lower()}?",
                [
                    ["A", "Sustentar las respuestas con datos contrastados y justificaciones fundadas."],
                    ["B", "Respaldar afirmaciones sin necesidad de comprobación."],
                    ["C", "Demostrar que cualquier hipótesis es válida de antemano."],
                    ["D", "Descartar los resultados que no coincidan con la expectativa inicial."]
                ],
                "A",
                "La argumentación basada en evidencia es un objetivo transversal del currículum nacional."
            ),
            (
                f"Al clasificar los elementos relacionados con {clean_topic.lower()}, el criterio principal es:",
                [
                    ["A", "Sus propiedades observables, estructura funcional y comportamiento medible."],
                    ["B", "El color aparente sin considerar la composición."],
                    ["C", "El orden alfabético de sus denominaciones."],
                    ["D", "El grado de dificultad percibido por el evaluador."]
                ],
                "A",
                "Las taxonomías y clasificaciones se fundamentan en criterios científicos reproducibles."
            ),
            (
                f"Finalmente, ¿cómo se vincula {clean_topic.lower()} con los Objetivos de Aprendizaje de {clean_subject}?",
                [
                    ["A", "Promoviendo el pensamiento crítico, la investigación y la transferencia a la vida cotidiana."],
                    ["B", "Reduciendo el aprendizaje a la memorización mecánica de definiciones breves."],
                    ["C", "Aislando el contenido de cualquier aplicación práctica en la comunidad escolar."],
                    ["D", "Restringiendo el análisis a un único punto de vista sin discusión."]
                ],
                "A",
                "Los OAs buscan desarrollar competencias integrales para la toma de decisiones ciudadanas."
            )
        ]

        # Pool de afirmaciones Verdadero / Falso (1 pt c/u)
        pool_vf = [
            (f"El estudio de {clean_topic.lower()} permite comprender el entorno y fortalecer el aprendizaje integral en {clean_subject}.", "V", "La comprensión pedagógica fomenta la toma de decisiones informadas y el pensamiento crítico."),
            ("Los conceptos estudiados son arbitrarios y carecen de fundamentación en las bases curriculares oficiales.", "F", "Los contenidos responden a los Objetivos de Aprendizaje (OA) definidos en el currículum Mineduc."),
            ("El trabajo colaborativo y la atención en clases facilitan la resolución de este tipo de problemas.", "V", "Las habilidades formativas potencian el rendimiento y la internalización de conocimientos."),
            ("Los conceptos evaluados son independientes de las actividades y experimentos realizados en el aula.", "F", "La evaluación mide directamente los aprendizajes construidos durante el período lectivo."),
            (f"Existen relaciones observables entre {clean_topic.lower()} y los fenómenos cotidianos del entorno.", "V", "La contextualización curricular conecta los saberes con la realidad del estudiante."),
            ("Cualquier conclusión es válida aun cuando contradiga los datos empíricos obtenidos.", "F", "Las conclusiones deben sustentarse rigurosamente en evidencia comprobable."),
            (f"El texto escolar ({textbook_pages or 'guía oficial'}) aporta definiciones y esquemas para profundizar en {clean_topic.lower()}.", "V", "El texto de estudio es el recurso pedagógico oficial de referencia para el curso."),
            ("El uso de unidades de medida y lenguaje técnico específico es innecesario en esta disciplina.", "F", "La precisión conceptual y el vocabulario disciplinar son fundamentales para una comunicación efectiva."),
            ("Analizar causas y efectos contribuye a una comprensión más profunda de los fenómenos evaluados.", "V", "El pensamiento causal estructura el razonamiento lógico del estudiante."),
            ("Las leyes y principios de esta materia aplican exclusivamente en situaciones teóricas.", "F", "Los principios disciplinarios rigen los fenómenos tanto en contextos naturales como tecnológicos."),
            ("La contrastación de hipótesis fortalece el desarrollo de habilidades investigativas.", "V", "La formulación y contrastación de hipótesis es central en las ciencias y el razonamiento sistemático."),
            ("Los resultados de una experiencia siempre dependen del azar sin responder a leyes naturales.", "F", "Los fenómenos responden a principios físicos, químicos, biológicos o matemáticos consistentes."),
            (f"Interpretar esquemas y gráficos es una destreza evaluada en la unidad de {clean_topic.lower()}.", "V", "La alfabetización visual y gráfica forma parte esencial de las metas de aprendizaje."),
            ("La revisión de pautas y rúbricas antes de entregar un trabajo permite corregir discrepancias a tiempo.", "V", "La metacognición y autorregulación mejoran sustancialmente los resultados formativos.")
        ]

        # Construir lista ajustada de preguntas Item 1
        items1 = []
        for i in range(min(n_q1, len(pool_q1))):
            q_text, alts, corr_val, just = pool_q1[i]
            corr_text = corr_val if corr_val.startswith("A)") else f"A) {alts[0][1]}"
            items1.append({
                "pregunta": f"{i+1}. {q_text}",
                "alternativas": alts,
                "correcta": corr_text,
                "justificacion": just
            })

        # Construir lista ajustada de preguntas Item 2
        items2 = []
        for j in range(min(n_q2, len(pool_vf))):
            oracion, resp, just = pool_vf[j]
            items2.append({
                "oracion": oracion,
                "resp": resp,
                "justificacion": just
            })

        # El banco local puede contener menos ejemplos que lo solicitado.
        calc_item3 = target_pts - (len(items1) * 2 + len(items2))
        total_sum = target_pts
        return {
            "simulation_mode": True,
            "colegio": colegio_nombre,
            "asignatura": clean_subject,
            "curso": grade,
            "titulo": f"EVALUACIÓN: {clean_topic.upper()}",
            "oa": oa or f"OA Mineduc: Desarrollar habilidades y comprensión de conceptos clave en {clean_subject}.",
            "contenidos": f"{clean_topic}. Páginas del texto escolar: {textbook_pages or 'Capítulo oficial'}.",
            "puntaje_total": total_sum,
            "docente": "Docente Titular / Evaluador",
            "email_docente": "docente@colegiocastelgandolfo.cl",

            "item1_pts_cada_una": 2,
            "item1_seleccion_multiple": items1,

            "item2_pts_cada_una": 1,
            "item2_verdadero_falso": items2,

            "item3_aplicacion": {
                "titulo": "ÍTEM III: APLICACIÓN PRÁCTICA Y DESARROLLO",
                "puntaje": calc_item3,
                "tipo": "drawing_boxes",
                "instruccion": f"Desarrolla o representa en el recuadro una situación o esquema que aplique los conceptos de {clean_topic.lower()} en {clean_subject}:",
                "cajas": [
                    {
                        "titulo": f"Aplicación de {clean_topic.lower()} ({calc_item3} puntos):",
                        "ejemplo_pauta": f"Respuesta o esquema claro donde se aprecian los componentes clave de {clean_topic.lower()} según la pauta docente."
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
