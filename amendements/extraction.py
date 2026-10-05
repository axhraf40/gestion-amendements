"""
Extraction des amendements depuis le HTML produit par LibreOffice.

Structure attendue d'un document (ex. الفريق_1.docx) :

    [tableau 1 ligne]  التعديل رقم : 1 | الفريق 1 | إضافة مادة جديدة
    <p> مدونة الجمارك والضرائب غير المباشرة
    <p> المادة 3  البند I
    <p> الفصل 70 المكرر
    [tableau 3 colonnes]  النص الأصلي | نص التعديل | التعليل
                          <texte>     | <texte>    | <texte>

Ce module ne dépend pas de Django : il peut être testé seul.
"""
import copy
import re

from bs4 import BeautifulSoup

# Champs « métadonnées » stockés dans le modèle Amendement
META_FIELDS = [
    "numero_modification",
    "numero_groupe",
    "type_modification",
    "nom_loi",
    "article",
    "alinea",
    "numero_article",
]

_WS = re.compile(r"[\s\u200c\u200d\u200f\u200e\ufeff]+")
_TASHKEEL = re.compile(r"[\u064B-\u0652\u0640]")
_DIGITS_AR = str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789")


def normaliser(texte):
    """Espaces multiples -> un espace, chiffres arabes-indiens -> chiffres latins."""
    if not texte:
        return ""
    return _WS.sub(" ", texte.translate(_DIGITS_AR)).strip()


def _compact(texte):
    return re.sub(r"[\s:ـ\-_]", "", texte or "")


# ---------------------------------------------------------------------------
# 1. Découpage du document en blocs (un bloc = un amendement)
# ---------------------------------------------------------------------------

def _lignes_directes(table):
    return [tr for tr in table.find_all("tr") if tr.find_parent("table") is table]


def _cellules_directes(tr):
    return [td for td in tr.find_all(["td", "th"]) if td.find_parent("tr") is tr]


def est_tableau_contenu(table):
    """Tableau à 3 colonnes : (النص الأصلي | النص كما جاء في المشروع) / نص التعديل / التعليل."""
    lignes = _lignes_directes(table)
    if not lignes:
        return False
    entetes = [_compact(c.get_text()) for c in _cellules_directes(lignes[0])]
    a_texte_modifie = any("نصالتعديل" in h or "النصالمعدل" in h for h in entetes)
    a_justification = any("التعليل" in h for h in entetes)
    return a_texte_modifie and a_justification


_TITRES_COLONNES = {
    0: ["النص الأصلي", "النص كما جاء في المشروع"],
    1: ["نص التعديل", "النص المعدل"],
    2: ["التعليل"],
}


def _texte_cellule(cell, idx_colonne):
    """
    Texte d'une cellule, dans l'ordre du document. Les sous-tableaux (ex. tarifs douaniers)
    sont rendus ligne par ligne (cellules séparées par des tabulations) à leur place exacte.
    """
    cell = copy.copy(cell)
    tableaux = []
    for sous_table in cell.find_all("table"):
        if sous_table.find_parent("td") is not cell:
            continue  # tableau imbriqué dans un sous-tableau : rendu avec son parent
        lignes = []
        for tr in _lignes_directes(sous_table):
            cols = [normaliser(td.get_text(" ")) for td in _cellules_directes(tr)]
            lignes.append("\t".join(cols))
        sous_table.replace_with(f" \x00{len(tableaux)}\x00 ")
        tableaux.append("\n".join(lignes))

    texte = normaliser(cell.get_text(" "))
    for titre in _TITRES_COLONNES.get(idx_colonne, []):
        if texte == titre:
            texte = ""
        elif texte.startswith(titre):
            texte = texte[len(titre):].lstrip(" :،-")

    morceaux = []
    for i, part in enumerate(re.split(r"\s*\x00(\d+)\x00\s*", texte)):
        part = part if i % 2 == 0 else tableaux[int(part)]
        if part.strip():
            morceaux.append(part.strip("\n "))
    return "\n".join(morceaux).strip()


def extraire_blocs(html):
    """
    Parcourt le document dans l'ordre et renvoie une liste de dicts :
        {entete, texte_original, texte_modifie, justification}
    `entete` contient tout le texte situé entre deux tableaux de contenu
    (numéro, groupe, type, loi, article, ...).
    """
    soup = BeautifulSoup(html, "html.parser")
    racine = soup.body or soup
    blocs = []
    tampon = []

    for el in racine.find_all(["p", "h1", "h2", "h3", "h4", "h5", "h6", "table"]):
        if el.find_parent("table") is not None:
            continue  # contenu interne d'un tableau : traité avec son tableau
        if el.name == "table" and est_tableau_contenu(el):
            colonnes = [[], [], []]
            for tr in _lignes_directes(el)[1:]:
                cells = _cellules_directes(tr)
                for i in range(min(3, len(cells))):
                    t = _texte_cellule(cells[i], i)
                    if t:
                        colonnes[i].append(t)
            blocs.append({
                "entete": normaliser(" ".join(tampon)),
                "texte_original": "\n".join(colonnes[0]),
                "texte_modifie": "\n".join(colonnes[1]),
                "justification": "\n".join(colonnes[2]),
            })
            tampon = []
        elif el.name == "table":
            for tr in _lignes_directes(el):
                tampon.append(" ".join(normaliser(c.get_text(" ")) for c in _cellules_directes(tr)))
        else:
            t = normaliser(el.get_text(" "))
            if t:
                tampon.append(t)
    return blocs


# ---------------------------------------------------------------------------
# 2. Détection des champs par règles (repli si l'IA n'est pas disponible)
# ---------------------------------------------------------------------------

_RE_NUMERO = re.compile(r"(?:التعديل\s*رقم|رقم\s*التعديل)\s*[:：\-]?\s*(\d+)")
_RE_GROUPE = re.compile(r"(?:الفريق|فريق)\s*[:：\-]?\s*(\d+)")
_RE_ALINEA = re.compile(r"البند\s*[:：\-]?\s*([IVXLCDM]+|\d+|[أ-ي])(?=\s|$|[.\-–])")
# Parties « connues » de l'en-tête, retirées avant de chercher le nom de la loi
_RE_PREFIXES = re.compile(
    r"الفريق\s*\d+\s*التعديلات\s*المقترحة|التعديلات\s*المقترحة|"
    r"(?:التعديل\s*رقم|رقم\s*التعديل)\s*[:：\-]?\s*\d+|(?:الفريق|فريق)\s*[:：\-]?\s*\d+|"
    r"إضافة\s*مادة\s*جديدة|إضافة\s*(?:فقرة|بند)|\bحذف\b"
)
_RE_DEBUT_REF = re.compile(r"\b(?:المادة|الفصل|البند)\b")

_MOTS_LIES = r"(?:\s+(?:المكرر(?:ة)?|مكرر(?:ة)?)(?:\s+(?:مرتين|ثلاث\s+مرات))?)?"
_RE_FASL = re.compile(r"الفصل\s*[:：\-]?\s*(\d+" + _MOTS_LIES + r")")
_RE_MADDA = re.compile(r"المادة\s*[:：\-]?\s*(\d+" + _MOTS_LIES + r")")


def detecter_par_regles(entete):
    """Détecte les champs de métadonnées dans le texte d'en-tête d'un amendement."""
    t = normaliser(_TASHKEEL.sub("", entete or ""))
    res = {f: "" for f in META_FIELDS}

    m = _RE_NUMERO.search(t)
    if m:
        res["numero_modification"] = m.group(1)
    m = _RE_GROUPE.search(t)
    if m:
        res["numero_groupe"] = m.group(1)

    if re.search(r"إضافة\s*مادة\s*جديدة", t):
        res["type_modification"] = "إضافة مادة جديدة"
    elif re.search(r"إضافة\s*(?:فقرة|بند)", t):
        res["type_modification"] = re.search(r"إضافة\s*(?:فقرة|بند)", t).group(0)
    elif re.search(r"\bحذف\b", t):
        res["type_modification"] = "حذف"

    # Nom de la loi : ce qui reste avant la première référence (المادة / الفصل / البند)
    sans_prefixes = normaliser(_RE_PREFIXES.sub(" ", t))
    m = _RE_DEBUT_REF.search(sans_prefixes)
    loi = sans_prefixes[:m.start()] if m else sans_prefixes
    loi = loi.strip(" :،-")
    if loi and not re.fullmatch(r"[\d\s]*", loi):
        res["nom_loi"] = loi
    apres_loi = sans_prefixes[m.start():] if m else ""

    # 1re « المادة N » après le nom de la loi = article du projet de loi de finances
    m_art = _RE_MADDA.search(apres_loi)
    if m_art:
        res["article"] = m_art.group(1).split()[0]
        reste = apres_loi[m_art.end():]
    else:
        reste = apres_loi

    m = _RE_ALINEA.search(reste) or _RE_ALINEA.search(t)
    if m:
        res["alinea"] = m.group(1)

    # Article du code modifié : « الفصل N [المكرر] » ou 2e « المادة N »
    m = _RE_FASL.search(reste) or _RE_MADDA.search(reste)
    if m:
        res["numero_article"] = normaliser(m.group(1))
    return res


# ---------------------------------------------------------------------------
# 3. Fusion IA + règles
# ---------------------------------------------------------------------------

def fusionner(champs_ia, champs_regles):
    """
    Combine les deux détections champ par champ.
    Renvoie (valeurs_finales, champs_detectes) où champs_detectes a le format
    {champ: {"value", "confidence", "method"}} attendu par les templates.
    """
    finales, details = {}, {}
    champs_ia = champs_ia or {}
    for f in META_FIELDS:
        # sans signes diacritiques (ّ ً ...) pour que « مرّتين » == « مرتين »
        v_ia = normaliser(_TASHKEEL.sub("", str(champs_ia.get(f) or "")))
        v_re = normaliser(_TASHKEEL.sub("", champs_regles.get(f) or ""))
        if f in ("numero_modification", "numero_groupe", "article") and v_ia and not re.fullmatch(r"\d+", v_ia):
            v_ia = re.search(r"\d+", v_ia).group(0) if re.search(r"\d+", v_ia) else ""
        if v_ia and v_re:
            if v_ia == v_re:
                val, conf, meth = v_ia, 0.98, "ia+regles"
            else:
                val, conf, meth = v_ia, 0.8, "ia"
        elif v_ia:
            val, conf, meth = v_ia, 0.85, "ia"
        elif v_re:
            val, conf, meth = v_re, 0.7, "regles"
        else:
            val, conf, meth = "", 0.0, ""
        finales[f] = val
        if val:
            details[f] = {"value": val, "confidence": conf, "method": meth}
            if v_ia and v_re and v_ia != v_re:
                details[f]["regles"] = v_re
    return finales, details
