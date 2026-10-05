"""
Détection des champs d'un amendement par IA (modèle de langage gratuit).

Fournisseurs supportés (tous via l'API « OpenAI-compatible » /chat/completions) :
  - huggingface : bibliothèque officielle huggingface_hub (InferenceClient.chat_completion)
                  jeton gratuit : https://huggingface.co/settings/tokens
  - groq        : api.groq.com             (clé gratuite : https://console.groq.com/keys)
  - openrouter  : openrouter.ai            (modèles « :free »)
  - ollama      : modèle local, 100 % gratuit et hors-ligne (https://ollama.com)
  - none        : désactive l'IA, seules les règles sont utilisées

Configuration dans le fichier .env (voir .env.example) :
  AI_PROVIDER=huggingface
  AI_API_KEY=hf_xxx
  AI_MODEL=            (optionnel, sinon modèle par défaut du fournisseur)

Si l'IA ne répond pas (pas d'internet, quota, mauvaise clé…), la détection
par règles (amendements/extraction.py) prend le relais automatiquement.
"""
import hashlib
import json
import logging
import os
import re

import requests

try:  # bibliothèque officielle Hugging Face (pip install huggingface_hub)
    from huggingface_hub import InferenceClient
except ImportError:  # repli : appel HTTP direct
    InferenceClient = None

from .extraction import META_FIELDS, detecter_par_regles, fusionner, normaliser

logger = logging.getLogger(__name__)

PROVIDERS = {
    "huggingface": {
        "url": "https://router.huggingface.co/v1/chat/completions",
        "model": "Qwen/Qwen2.5-72B-Instruct",
        "key_env": ["AI_API_KEY", "HF_TOKEN", "HUGGINGFACE_API_KEY"],
    },
    "groq": {
        "url": "https://api.groq.com/openai/v1/chat/completions",
        "model": "llama-3.3-70b-versatile",
        "key_env": ["AI_API_KEY", "GROQ_API_KEY"],
    },
    "openrouter": {
        "url": "https://openrouter.ai/api/v1/chat/completions",
        "model": "meta-llama/llama-3.3-70b-instruct:free",
        "key_env": ["AI_API_KEY", "OPENROUTER_API_KEY"],
    },
    "ollama": {
        "url": "http://localhost:11434/v1/chat/completions",
        "model": "qwen2.5:7b",
        "key_env": [],
    },
}

SYSTEM_PROMPT = """Tu es un assistant qui extrait des informations structurées d'amendements \
législatifs marocains rédigés en arabe (projet de loi de finances).

On te donne l'EN-TÊTE d'un amendement (et le début du texte original). Renvoie UNIQUEMENT \
un objet JSON avec exactement ces clés (valeurs = chaînes, "" si absent, ne rien inventer) :

- "numero_modification" : numéro de l'amendement (après « التعديل رقم »), chiffres seulement
- "numero_groupe"       : numéro du groupe parlementaire (après « الفريق »), chiffres seulement
- "type_modification"   : type s'il est écrit explicitement (ex. « إضافة مادة جديدة », « حذف »), sinon ""
- "nom_loi"             : nom complet du code / de la loi / du tableau visé \
(ex. « مدونة الجمارك والضرائب غير المباشرة », « المدونة العامة للضرائب », « جدول التعريفة الجمركية »)
- "article"             : numéro de la PREMIÈRE « المادة » qui suit le nom de la loi \
(article du projet de loi de finances), chiffres seulement
- "alinea"              : valeur après « البند » (ex. « I », « II », « 3 »)
- "numero_article"      : article du code modifié : valeur après « الفصل », ou la DEUXIÈME « المادة » \
s'il n'y a pas de « الفصل ». Garder les mots « المكرر », « مكرر مرتين » s'ils suivent le numéro.

Exemple
EN-TÊTE: التعديل رقم : 4 الفريق 1 إضافة مادة جديدة مدونة الجمارك والضرائب غير المباشرة المادة 3 البند I الفصل 164 المكرر
JSON: {"numero_modification":"4","numero_groupe":"1","type_modification":"إضافة مادة جديدة",\
"nom_loi":"مدونة الجمارك والضرائب غير المباشرة","article":"3","alinea":"I","numero_article":"164 المكرر"}

EN-TÊTE: التعديل رقم : 7 الفريق 1 المدونة العامة للضرائب المادة 8 البند I الإعفاءات المادة 6
JSON: {"numero_modification":"7","numero_groupe":"1","type_modification":"",\
"nom_loi":"المدونة العامة للضرائب","article":"8","alinea":"I","numero_article":"6"}
"""


def _env(name, default=""):
    return (os.environ.get(name) or default).strip()


class HuggingFaceFieldDetector:
    """Détecteur IA + règles. Le nom de classe est conservé pour compatibilité."""

    def __init__(self):
        self.provider = _env("AI_PROVIDER", "huggingface").lower()
        conf = PROVIDERS.get(self.provider, {})
        self.api_url = _env("AI_API_URL", conf.get("url", ""))
        self.model = _env("AI_MODEL", conf.get("model", ""))
        self.api_key = ""
        for k in conf.get("key_env", ["AI_API_KEY"]):
            if _env(k):
                self.api_key = _env(k)
                break
        self.timeout = float(_env("AI_TIMEOUT", "40"))
        self.hf_provider = _env("AI_HF_PROVIDER", "auto")
        self._cache = {}
        self.derniere_erreur = ""

    # -- état -------------------------------------------------------------
    @property
    def ia_active(self):
        if self.provider == "none" or not self.api_url or not self.model:
            return False
        if self.provider != "ollama" and not self.api_key:
            return False
        return True

    def description(self):
        if not self.ia_active:
            return "règles uniquement (IA désactivée ou clé manquante)"
        via = " (InferenceClient)" if self.utilise_inference_client else ""
        return f"{self.provider}{via} / {self.model}"

    # -- appel au modèle ----------------------------------------------------
    def _appeler_llm(self, entete, debut_texte=""):
        contenu = f"EN-TÊTE: {entete}"
        if debut_texte:
            contenu += f"\nDÉBUT DU TEXTE ORIGINAL: {debut_texte[:300]}"
        contenu += "\nJSON:"

        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": contenu},
        ]
        if self.utilise_inference_client:
            texte = self._appel_inference_client(messages)
        else:
            texte = self._appel_http(messages)
        return self._parser_json(texte)

    @property
    def utilise_inference_client(self):
        """Hugging Face via la bibliothèque officielle huggingface_hub (InferenceClient)."""
        return (self.provider == "huggingface" and InferenceClient is not None
                and not _env("AI_API_URL"))

    def _appel_inference_client(self, messages):
        # Doc : https://huggingface.co/docs/huggingface_hub/guides/inference
        client = InferenceClient(
            provider=self.hf_provider,  # "auto" = fournisseur choisi par Hugging Face
            api_key=self.api_key,
            timeout=self.timeout,
        )
        rep = client.chat_completion(
            messages=messages,
            model=self.model,
            max_tokens=300,
            temperature=0.1,
        )
        return rep.choices[0].message.content

    def _appel_http(self, messages):
        """Appel direct à une API compatible OpenAI (Groq, OpenRouter, Ollama, ou HF sans la lib)."""
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        payload = {"model": self.model, "messages": messages, "temperature": 0.1, "max_tokens": 300}
        r = requests.post(self.api_url, headers=headers, json=payload, timeout=self.timeout)
        if r.status_code != 200:
            raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
        return r.json()["choices"][0]["message"]["content"]

    @staticmethod
    def _parser_json(texte):
        texte = re.sub(r"^```(?:json)?|```$", "", texte.strip(), flags=re.M).strip()
        m = re.search(r"\{.*\}", texte, re.S)
        if not m:
            raise ValueError(f"Réponse IA sans JSON : {texte[:200]}")
        data = json.loads(m.group(0))
        return {f: normaliser(str(data.get(f) or "")) for f in META_FIELDS}

    # -- API publique ---------------------------------------------------------
    def detecter(self, entete, texte_original=""):
        """
        Renvoie (valeurs, champs_detectes) :
          valeurs         -> {champ: valeur} à enregistrer dans le modèle Amendement
          champs_detectes -> {champ: {value, confidence, method}} pour l'affichage
        """
        entete = normaliser(entete)
        regles = detecter_par_regles(entete)
        ia = None
        if self.ia_active and entete:
            cle = hashlib.sha1(entete.encode("utf-8")).hexdigest()
            if cle in self._cache:
                ia = self._cache[cle]
            else:
                try:
                    ia = self._appeler_llm(entete, normaliser(texte_original))
                    self._cache[cle] = ia
                    self.derniere_erreur = ""
                except Exception as e:  # réseau, quota, JSON invalide…
                    self.derniere_erreur = str(e)
                    logger.warning("Détection IA indisponible, repli sur les règles : %s", e)
        return fusionner(ia, regles)

    def extract_structured_fields(self, text):
        """Compatibilité avec l'ancien code : renvoie seulement champs_detectes."""
        return self.detecter(text)[1]


field_detector = HuggingFaceFieldDetector()
