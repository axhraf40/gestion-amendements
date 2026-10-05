"""
Tests de l'extraction des amendements et de la détection des champs (IA + règles).

Lancer :   python manage.py test amendements
Test réel de l'IA (optionnel, utilise ta clé du .env) :
           PowerShell : $env:AI_LIVE_TEST="1"  ·  bash : AI_LIVE_TEST=1   puis   python manage.py test amendements
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from unittest import mock, skipUnless

from django.contrib.auth.models import User
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from django.utils.crypto import get_random_string

from .extraction import detecter_par_regles, extraire_blocs, fusionner
from .huggingface_utils import HuggingFaceFieldDetector
from .models import Amendement
from .sample_docx import AMENDEMENTS as AMENDEMENTS_FICTIFS, enregistrer as enregistrer_docx_fictif
from .views import get_soffice_path

DOUANE = "مدونة الجمارك والضرائب غير المباشرة"
ADD = "إضافة مادة جديدة"

# Documents de test FICTIFS, générés à la volée (aucun document réel dans le dépôt).
# exemple_court.docx : les 3 premiers amendements seulement (2e fichier différent).
DATA = Path(tempfile.mkdtemp(prefix="amendements_tests_"))
FICHIERS = {
    "exemple.docx": AMENDEMENTS_FICTIFS,
    "exemple_court.docx": AMENDEMENTS_FICTIFS[:3],
}
for _nom, _amendements in FICHIERS.items():
    enregistrer_docx_fictif(DATA / _nom, _amendements)
ATTENDU = {nom: [a["attendu"] for a in amds] for nom, amds in FICHIERS.items()}
CHAMPS = ["numero_modification", "numero_groupe", "type_modification", "nom_loi",
          "article", "alinea", "numero_article"]


def _soffice_disponible():
    p = get_soffice_path()
    return os.path.exists(p) or shutil.which(p) is not None


def _docx_vers_html(nom, dossier):
    src = Path(dossier) / nom
    shutil.copy(DATA / nom, src)
    subprocess.run([get_soffice_path(), "--headless", "--convert-to", "html",
                    "--outdir", str(dossier), str(src)], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return (Path(dossier) / (src.stem + ".html")).read_text(encoding="utf-8")


def _faux_llm(reponse_json=None, erreur=None):
    """Simule la réponse d'un modèle OpenAI-compatible (HF router, Groq, Ollama...)."""
    def _post(url, headers=None, json=None, timeout=None):
        if erreur:
            raise erreur
        entete = json["messages"][-1]["content"].split("EN-TÊTE: ", 1)[1].split("\n")[0]
        contenu = reponse_json or __import__("json").dumps(detecter_par_regles(entete), ensure_ascii=False)
        r = mock.Mock(status_code=200)
        r.json.return_value = {"choices": [{"message": {"content": "```json\n" + contenu + "\n```"}}]}
        return r
    return _post


def _faux_inference_client(reponse_json=None, erreur=None, appels=None):
    """Simule huggingface_hub.InferenceClient (chat_completion)."""
    post = _faux_llm(reponse_json, erreur)

    class FauxClient:
        def __init__(self, **kwargs):
            self.kwargs = kwargs
            if appels is not None:
                appels.append(kwargs)

        def chat_completion(self, messages, model=None, **kw):
            if appels is not None:
                appels[-1]["model"] = model
            data = post(None, json={"messages": messages}).json()
            msg = mock.Mock(content=data["choices"][0]["message"]["content"])
            return mock.Mock(choices=[mock.Mock(message=msg)])
    return FauxClient


# ---------------------------------------------------------------------------
class ReglesTests(TestCase):
    """Détection par règles sur des en-têtes typiques."""

    def test_entete_complet(self):
        r = detecter_par_regles(
            "التعديل رقم : 4 الفريق 1 إضافة مادة جديدة مدونة الجمارك والضرائب غير المباشرة "
            "المادة 3 البند I الفصل 164 المكرر")
        self.assertEqual(r, {"numero_modification": "4", "numero_groupe": "1", "type_modification": ADD,
                             "nom_loi": DOUANE, "article": "3", "alinea": "I",
                             "numero_article": "164 المكرر"})

    def test_deux_madda_et_tashkeel(self):
        r = detecter_par_regles("التعديل رقم : 13 الفريق 2 المدونة العامة للضرائب المادة 8 البند I "
                                "المادة 42 المكرر مرّتين")
        self.assertEqual((r["article"], r["numero_article"]), ("8", "42 المكرر مرتين"))

    def test_loi_sans_mot_cle_madawana(self):
        r = detecter_par_regles("التعديل رقم : 2 الفريق 2 الضرائب الداخلية على الاستهلاك المادة 5 البند I الفصل 9")
        self.assertEqual(r["nom_loi"], "الضرائب الداخلية على الاستهلاك")

    def test_chiffres_arabes_indiens(self):
        r = detecter_par_regles("التعديل رقم : ١٢ الفريق ٣")
        self.assertEqual((r["numero_modification"], r["numero_groupe"]), ("12", "3"))

    def test_ne_lit_pas_le_texte_de_justification(self):
        # ancien bug : « التعديل رقم1 » dans la justification était pris pour le numéro
        r = detecter_par_regles("التعديل رقم : 8 الفريق 1 هذا تعليل التعديل رقم1")
        self.assertEqual(r["numero_modification"], "8")


class FusionTests(TestCase):
    def test_accord_ia_regles(self):
        v, d = fusionner({"article": "3"}, {"article": "3"})
        self.assertEqual(d["article"]["method"], "ia+regles")

    def test_ia_prioritaire_mais_regles_conservees(self):
        v, d = fusionner({"article": "5"}, {"article": "3"})
        self.assertEqual(v["article"], "5")
        self.assertEqual(d["article"]["regles"], "3")

    def test_accord_malgre_diacritiques(self):
        v, d = fusionner({"numero_article": "42 المكرر مرّتين"}, {"numero_article": "42 المكرر مرتين"})
        self.assertEqual(d["numero_article"]["method"], "ia+regles")
        self.assertEqual(v["numero_article"], "42 المكرر مرتين")

    def test_repli_regles(self):
        v, d = fusionner(None, {"numero_modification": "7"})
        self.assertEqual((v["numero_modification"], d["numero_modification"]["method"]), ("7", "regles"))

    def test_nettoyage_numero_ia(self):
        v, _ = fusionner({"numero_modification": "رقم 7"}, {})
        self.assertEqual(v["numero_modification"], "7")


class DetecteurIATests(TestCase):
    ENTETE = "التعديل رقم : 1 الفريق 1 إضافة مادة جديدة مدونة الجمارك والضرائب غير المباشرة المادة 3 البند I الفصل 70 المكرر"

    def _det(self, **env):
        base = {"AI_PROVIDER": "huggingface", "AI_API_KEY": "hf_test", "AI_MODEL": "", "AI_API_URL": ""}
        base.update(env)
        with mock.patch.dict(os.environ, base):
            return HuggingFaceFieldDetector()

    def test_ia_hugging_face_inference_client(self):
        det = self._det()
        appels = []
        with mock.patch("amendements.huggingface_utils.InferenceClient",
                        _faux_inference_client(appels=appels)), \
                mock.patch("amendements.huggingface_utils.requests.post") as post:
            v, d = det.detecter(self.ENTETE)
        self.assertFalse(post.called)  # pas d'appel HTTP brut : on passe par la lib officielle
        self.assertEqual(appels[0]["api_key"], "hf_test")
        self.assertEqual(appels[0]["provider"], "auto")
        self.assertEqual(appels[0]["model"], "Qwen/Qwen2.5-72B-Instruct")
        self.assertEqual(v["numero_article"], "70 المكرر")
        self.assertTrue(all(x["method"] == "ia+regles" for x in d.values()))

    def test_ia_http_groq(self):
        det = self._det(AI_PROVIDER="groq", AI_API_KEY="gsk_test")
        with mock.patch("amendements.huggingface_utils.requests.post", side_effect=_faux_llm()) as post:
            v, d = det.detecter(self.ENTETE)
        self.assertEqual(post.call_args.kwargs["headers"]["Authorization"], "Bearer gsk_test")
        self.assertTrue(all(x["method"] == "ia+regles" for x in d.values()))

    def test_repli_si_ia_injoignable(self):
        det = self._det()
        with mock.patch("amendements.huggingface_utils.InferenceClient",
                        _faux_inference_client(erreur=ConnectionError("pas d'internet"))):
            v, d = det.detecter(self.ENTETE)
        self.assertEqual(v["numero_modification"], "1")
        self.assertIn("pas d'internet", det.derniere_erreur)
        self.assertTrue(all(x["method"] == "regles" for x in d.values()))

    def test_json_invalide(self):
        det = self._det()
        with mock.patch("amendements.huggingface_utils.InferenceClient",
                        _faux_inference_client(reponse_json="je ne sais pas")):
            v, _ = det.detecter(self.ENTETE)
        self.assertEqual(v["article"], "3")  # règles
        self.assertTrue(det.derniere_erreur)

    def test_sans_cle_pas_d_appel(self):
        det = self._det(AI_API_KEY="")
        with mock.patch("amendements.huggingface_utils.requests.post") as post, \
                mock.patch("amendements.huggingface_utils.InferenceClient") as client:
            det.detecter(self.ENTETE)
        self.assertFalse(post.called or client.called)

    def test_fournisseurs(self):
        self.assertIn("groq.com", self._det(AI_PROVIDER="groq").api_url)
        self.assertTrue(self._det(AI_PROVIDER="ollama", AI_API_KEY="").ia_active)
        self.assertFalse(self._det(AI_PROVIDER="none").ia_active)


@skipUnless(_soffice_disponible(), "LibreOffice introuvable (définir SOFFICE_PATH)")
class ExemplesTests(TestCase):
    """Extraction complète sur les documents Word fictifs générés (amendements/sample_docx.py)."""

    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.tmp = tempfile.mkdtemp()
        cls.html = {nom: _docx_vers_html(nom, cls.tmp) for nom in ATTENDU}

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)
        super().tearDownClass()

    def test_exemples(self):
        for nom, attendus in ATTENDU.items():
            with self.subTest(fichier=nom):
                blocs = extraire_blocs(self.html[nom])
                self.assertEqual(len(blocs), len(attendus))
                for bloc, attendu in zip(blocs, attendus):
                    r = detecter_par_regles(bloc["entete"])
                    self.assertEqual(tuple(r[c] for c in CHAMPS), attendu, bloc["entete"])
                    self.assertTrue(bloc["texte_original"] and bloc["texte_modifie"] and bloc["justification"])
                    self.assertNotIn("النص كما جاء في المشروع", bloc["texte_original"][:30])


    def test_textes_identiques_au_word(self):
        """Les 3 colonnes (texte original / modifié / justification) contiennent exactement le texte
        du tableau Word, dans le même ordre — vérité lue indépendamment avec python-docx."""
        import re
        from docx import Document
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

        def lettres(t):
            t = re.sub(r"[\u064B-\u0652\u0640\u200c\u200d\ufeff]", "", t or "")
            return re.sub(r"[\s.\-–:؛,،…]+", "", t)

        for nom in ATTENDU:
            with self.subTest(fichier=nom):
                doc = Document(DATA / nom)
                verite = []
                for tbl in doc.element.body.iterchildren(W + "tbl"):
                    lignes = tbl.findall(W + "tr")
                    entete = "".join(t.text or "" for t in lignes[0].iter(W + "t")) if lignes else ""
                    if "التعديل" in entete and "التعليل" in entete and "رقم" not in entete:
                        cols = ["", "", ""]
                        for tr in lignes[1:]:
                            for i, tc in enumerate(tr.findall(W + "tc")[:3]):
                                cols[i] += " ".join(t.text or "" for t in tc.iter(W + "t"))
                        verite.append(cols)
                blocs = extraire_blocs(self.html[nom])
                self.assertEqual(len(blocs), len(verite))
                for bloc, (orig, mod, just) in zip(blocs, verite):
                    self.assertEqual(lettres(bloc["texte_original"]), lettres(orig))
                    self.assertEqual(lettres(bloc["texte_modifie"]), lettres(mod))
                    self.assertEqual(lettres(bloc["justification"]), lettres(just))


@skipUnless(_soffice_disponible(), "LibreOffice introuvable (définir SOFFICE_PATH)")
class UploadTests(TestCase):
    """Upload réel via la vue Django, IA simulée : vérifie ce qui est enregistré en base."""

    def setUp(self):
        self.media = tempfile.mkdtemp()
        self.override = override_settings(MEDIA_ROOT=self.media)
        self.override.enable()
        self.user = User.objects.create_user("testeur", password=get_random_string(20))
        self.client.force_login(self.user)

    def tearDown(self):
        self.override.disable()
        shutil.rmtree(self.media, ignore_errors=True)

    def _upload(self, nom):
        fichier = SimpleUploadedFile(nom, (DATA / nom).read_bytes())
        return self.client.post("/amendements/upload/", {"fichier": fichier})

    def test_upload_exemples_avec_ia(self):
        det = HuggingFaceFieldDetector()
        det.provider, det.model, det.api_key = "huggingface", "m", "k"
        det.api_url = "https://router.huggingface.co/v1/chat/completions"
        with mock.patch("amendements.views.field_detector", det), \
                mock.patch("amendements.huggingface_utils.InferenceClient", _faux_inference_client()):
            for nom, attendus in ATTENDU.items():
                with self.subTest(fichier=nom):
                    Amendement.objects.all().delete()
                    self.assertEqual(self._upload(nom).status_code, 200)
                    rows = list(Amendement.objects.order_by("id"))
                    self.assertEqual(len(rows), len(attendus))
                    for a, attendu in zip(rows, attendus):
                        obtenu = tuple((getattr(a, c) or "") for c in CHAMPS)
                        self.assertEqual(obtenu, attendu)
                        self.assertEqual(a.champs_detectes["numero_modification"]["method"], "ia+regles")
                        self.assertTrue(a.texte_entete)

    def test_upload_sans_ia(self):
        det = HuggingFaceFieldDetector()
        det.provider = "none"
        with mock.patch("amendements.views.field_detector", det):
            self._upload("exemple.docx")
        self.assertEqual(Amendement.objects.count(), 5)
        a = Amendement.objects.get(numero_modification="12")
        self.assertEqual(a.numero_article, "42 المكرر مرتين")
        self.assertEqual(a.champs_detectes["numero_article"]["method"], "regles")

    def test_pages_detection(self):
        erreur = ConnectionError("x")
        with mock.patch("amendements.huggingface_utils.requests.post", side_effect=_faux_llm(erreur=erreur)), \
                mock.patch("amendements.huggingface_utils.InferenceClient", _faux_inference_client(erreur=erreur)):
            self._upload("exemple_court.docx")
            a = Amendement.objects.first()
            self.assertEqual(self.client.get(f"/amendements/amendements/{a.id}/champs-detectes/").status_code, 200)
            self.assertEqual(self.client.get(f"/amendements/amendements/{a.id}/re-detecter-champs/").status_code, 302)
            self.assertEqual(
                self.client.get(f"/amendements/fichiers/{a.fichier.id}/statistiques-detection/").status_code, 200)

    def test_pdf_telechargement_et_apercu(self):
        self._upload("exemple_court.docx")
        f = Amendement.objects.first().fichier
        r = self.client.get(f"/amendements/amendements/{f.id}/telecharger_pdf_libreoffice/")
        self.assertEqual(r.status_code, 200)
        self.assertIn("attachment", r["Content-Disposition"])
        self.assertTrue(b"".join(r.streaming_content).startswith(b"%PDF"))
        r = self.client.get(f"/amendements/fichiers/{f.id}/apercu_pdf/")
        self.assertEqual(r.status_code, 200)
        self.assertContains(r, "?inline=1")
        r = self.client.get(f"/amendements/amendements/{f.id}/telecharger_pdf_libreoffice/?inline=1")
        self.assertIn("inline", r["Content-Disposition"])

    def test_ancienne_page_upload_redirige(self):
        r = self.client.get("/amendements/fichiers/upload/")
        self.assertRedirects(r, "/amendements/upload/", fetch_redirect_response=False)


@override_settings(MEDIA_ROOT=tempfile.mkdtemp())
class SansLibreOfficeTests(TestCase):
    """Si LibreOffice est introuvable : message clair, jamais d'erreur 500."""

    def setUp(self):
        self.user = User.objects.create_user("u", password=get_random_string(20))
        self.client.force_login(self.user)
        self.p = mock.patch("amendements.views.get_soffice_path", return_value="/introuvable/soffice")
        self.p.start()

    def tearDown(self):
        self.p.stop()

    def test_upload_message_et_rien_enregistre(self):
        from .models import AmendementFichier
        fichier = SimpleUploadedFile("exemple.docx", (DATA / "exemple.docx").read_bytes())
        r = self.client.post("/amendements/upload/", {"fichier": fichier})
        self.assertEqual(r.status_code, 200)
        self.assertContains(r, "LibreOffice est introuvable")
        self.assertEqual(AmendementFichier.objects.count(), 0)

    def test_pdf_message(self):
        from .models import AmendementFichier
        f = AmendementFichier.objects.create(utilisateur=self.user, fichier=SimpleUploadedFile("f.docx", b"x"))
        r = self.client.get(f"/amendements/amendements/{f.id}/telecharger_pdf_libreoffice/", follow=True)
        self.assertEqual(r.status_code, 200)
        self.assertContains(r, "LibreOffice est introuvable")


@skipUnless(os.environ.get("AI_LIVE_TEST") == "1", "Test IA réel désactivé (mettre AI_LIVE_TEST=1)")
class IAReelleTests(TestCase):
    """Appelle vraiment le fournisseur configuré dans .env (Hugging Face : InferenceClient)."""

    def test_ia_reelle(self):
        det = HuggingFaceFieldDetector()
        self.assertTrue(det.ia_active, "Clé IA manquante dans .env (AI_API_KEY)")
        v, d = det.detecter("التعديل رقم : 12 الفريق 2 إضافة مادة جديدة المدونة العامة للضرائب "
                            "المادة 8 البند I المادة 40")
        self.assertEqual(det.derniere_erreur, "", det.derniere_erreur)
        self.assertEqual((v["numero_modification"], v["numero_groupe"], v["article"]), ("12", "2", "8"))
        self.assertIn("ia", d["numero_modification"]["method"])


@override_settings(MEDIA_ROOT=tempfile.mkdtemp())
class SecuriteTests(TestCase):
    """Un utilisateur ne peut ni voir, ni modifier, ni supprimer les fichiers d'un autre."""

    def setUp(self):
        from django.utils.crypto import get_random_string
        from .models import AmendementFichier
        self.proprio = User.objects.create_user("proprio", password=get_random_string(20))
        self.autre = User.objects.create_user("autre", password=get_random_string(20))
        self.fichier = AmendementFichier.objects.create(
            utilisateur=self.proprio, fichier=SimpleUploadedFile("f.docx", b"x"),
            extracted_content="<p>contenu</p>")

    def test_autre_utilisateur_bloque(self):
        self.client.force_login(self.autre)
        fid = self.fichier.id
        for url in [f"/amendements/fichier_mammoth/{fid}/", f"/amendements/fichiers/{fid}/modifier/",
                    f"/amendements/fichiers/{fid}/renommer/",
                    f"/amendements/amendements/{fid}/telecharger_pdf_libreoffice/",
                    f"/amendements/fichiers/{fid}/statistiques-detection/",
                    f"/amendements/fichiers/{fid}/apercu_pdf/"]:
            with self.subTest(url=url):
                self.assertEqual(self.client.get(url).status_code, 404)
        self.assertEqual(self.client.post(f"/amendements/fichiers/{fid}/supprimer/").status_code, 404)
        self.assertTrue(type(self.fichier).objects.filter(id=fid).exists())

    def test_suppression_uniquement_en_post(self):
        self.client.force_login(self.proprio)
        url = f"/amendements/fichiers/{self.fichier.id}/supprimer/"
        self.assertEqual(self.client.get(url).status_code, 405)
        self.assertEqual(self.client.post(url).status_code, 302)
        self.assertFalse(type(self.fichier).objects.filter(id=self.fichier.id).exists())

    def test_html_modifie_nettoye(self):
        self.client.force_login(self.proprio)
        html = '<p onclick="alert(1)">ok</p><script>alert(2)</script><a href="javascript:x()">l</a>'
        self.client.post(f"/amendements/fichiers/{self.fichier.id}/modifier/", {"extracted_content": html})
        self.fichier.refresh_from_db()
        contenu = self.fichier.extracted_content
        self.assertIn("ok", contenu)
        for interdit in ("<script", "onclick", "javascript:"):
            self.assertNotIn(interdit, contenu)

    def test_pages_protegees_sans_connexion(self):
        for url in ["/amendements/upload/", "/amendements/fichiers/", "/stats/dashboard/"]:
            with self.subTest(url=url):
                r = self.client.get(url)
                self.assertEqual(r.status_code, 302)
                self.assertIn("/accounts/login/", r["Location"])
