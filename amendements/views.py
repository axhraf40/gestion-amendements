from django.shortcuts import render, redirect, get_object_or_404
from .forms import UploadWordForm, AmendementForm
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import Amendement, AmendementFichier
from .forms import AmendementFichierForm
from django.http import HttpResponseForbidden
import docx
import os
import tempfile
from django.forms import modelform_factory, formset_factory, modelformset_factory
import re
from docx.document import Document as _Document
from docx.table import Table as _Table
from docx.text.paragraph import Paragraph as _Paragraph
import unicodedata
from docx.shared import RGBColor
from lxml import etree
import subprocess
from zipfile import BadZipFile
import mammoth
from bs4 import BeautifulSoup
from django.views.decorators.http import require_http_methods, require_POST
from django.urls import reverse
from django.http import HttpResponse
from django.template.loader import render_to_string
# PDF : généré par LibreOffice (voir convert_docx_to_pdf) — WeasyPrint n'est plus nécessaire
from docx.oxml.ns import qn
import requests
from django.http import FileResponse, Http404
import logging
from .huggingface_utils import field_detector
from .extraction import extraire_blocs
import shutil

logger = logging.getLogger(__name__)

# Nouvelle extraction : tous les tableaux

def normalize_arabic(text):
    # Supprime les espaces multiples, les retours à la ligne, etc.
    text = re.sub(r'[\s\u200c\u200d]+', ' ', text)
    return text.strip()

def get_soffice_path():
    """Chemin de LibreOffice : variable SOFFICE_PATH, sinon emplacements usuels / PATH."""
    candidats = [
        os.environ.get("SOFFICE_PATH", ""),
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidats:
        if c and os.path.exists(c):
            return c
    return shutil.which("soffice") or shutil.which("libreoffice") or "soffice"


def extract_html_content(docx_path):
    import logging
    output_dir = os.path.dirname(docx_path)
    cmd = [
        get_soffice_path(),
        "--headless",
        "--convert-to", "html",
        "--outdir", output_dir,
        docx_path
    ]
    try:
        subprocess.run(cmd, check=True, timeout=180)
    except (FileNotFoundError, PermissionError):
        raise RuntimeError("LibreOffice est introuvable. Installez-le ou indiquez son chemin dans SOFFICE_PATH (fichier .env).")
    except subprocess.TimeoutExpired:
        raise RuntimeError("La conversion du fichier Word a pris trop de temps (LibreOffice ne répond pas).")
    except subprocess.CalledProcessError as e:
        logging.error(f"Erreur lors de la conversion LibreOffice : {e}")
        raise RuntimeError("Erreur lors de la conversion du fichier Word en HTML. Veuillez vérifier que le fichier n'est pas corrompu et que LibreOffice est bien installé.")
    base = os.path.splitext(os.path.basename(docx_path))[0]
    html_path = os.path.join(output_dir, base + ".html")
    if not os.path.exists(html_path):
        logging.error(f"Fichier HTML non généré : {html_path}")
        raise RuntimeError("La conversion du fichier Word en HTML a échoué. Aucun fichier HTML généré.")
    with open(html_path, "r", encoding="utf-8") as f:
        html = f.read()
    soup = BeautifulSoup(html, "html.parser")
    # Supprimer tous les styles inline de taille de police
    for tag in soup.find_all(['td', 'th', 'tr', 'table', 'span', 'p', 'div']):
        if tag.has_attr('style'):
            styles = tag['style'].split(';')
            # On supprime uniquement la taille de police, on garde la couleur et les autres styles
            styles = [s for s in styles if 'font-size' not in s]
            styles = [s.strip() for s in styles if s.strip()]
            if styles:
                tag['style'] = ';'.join(styles)
            else:
                del tag['style']
    # Supprimer toutes les balises <style>
    for style_tag in soup.find_all('style'):
        style_tag.decompose()
    body = soup.body
    return str(body) if body else html

# --- Extraction avancée des styles pour python-docx ---
def get_run_color(run):
    # Couleur directe
    if run.font.color and run.font.color.rgb:
        return f"#{run.font.color.rgb}"
    # Couleur héritée du style
    if run.style and hasattr(run.style, 'font') and run.style.font.color and run.style.font.color.rgb:
        return f"#{run.style.font.color.rgb}"
    return None

def is_run_underlined(run):
    # Souligné direct
    if run.underline:
        return True
    # Souligné hérité du style
    if run.style and hasattr(run.style, 'font') and run.style.font.underline:
        return True
    return False

def run_to_html(run):
    style = []
    # Couleur du texte
    color = get_run_color(run)
    if color:
        style.append(f"color: {color};")
    # Gras, italique, souligné, barré
    if run.bold:
        style.append("font-weight: bold;")
    if run.italic:
        style.append("font-style: italic;")
    if is_run_underlined(run):
        style.append("text-decoration: underline;")
    if run.font.strike:
        style.append("text-decoration: line-through;")
    # Taille de police
    if run.font.size:
        style.append(f"font-size: {run.font.size.pt}pt;")
    # Construction du span
    style_str = " ".join(style)
    text = run.text.replace('\n', '<br>')  # gestion des sauts de ligne
    return f'<span style="{style_str}">{text}</span>' if style_str else text

def run_xml_to_html(run):
    r = run._element
    rPr = r.find(qn('w:rPr'))
    style = []
    # Gras
    if rPr is not None and rPr.find(qn('w:b')) is not None:
        style.append("font-weight: bold;")
    # Italique
    if rPr is not None and rPr.find(qn('w:i')) is not None:
        style.append("font-style: italic;")
    # Souligné
    if rPr is not None and rPr.find(qn('w:u')) is not None:
        style.append("text-decoration: underline;")
    # Barré
    if rPr is not None and rPr.find(qn('w:strike')) is not None:
        style.append("text-decoration: line-through;")
    # Couleur du texte
    color = get_run_color_from_xml(run)
    if color:
        style.append(f"color: {color};")
    # Taille de police
    if rPr is not None:
        sz = rPr.find(qn('w:sz'))
        if sz is not None and sz.get(qn('w:val')):
            size_pt = int(sz.get(qn('w:val'))) / 2
            style.append(f"font-size: {size_pt}pt;")
    text = run.text.replace('\n', '<br>')
    style_str = " ".join(style)
    return f'<span style="{style_str}">{text}</span>' if style_str else text

def paragraph_to_html(paragraph):
    html = ''.join(run_xml_to_html(run) for run in paragraph.runs)
    align = 'right'
    if paragraph.alignment == 1:
        align = 'center'
    elif paragraph.alignment == 2:
        align = 'left'
    return f'<div style="text-align:{align};">{html}</div>'

def extraire_html_mammoth(fichier_path, style_map_path=None):
    if style_map_path is None:
        # Chemin absolu vers le style-map dans le dossier amendements
        style_map_path = os.path.join(os.path.dirname(__file__), "mammoth-style-map.txt")
    with open(fichier_path, "rb") as docx_file:
        with open(style_map_path, encoding="utf-8") as style_map_file:
            style_map_content = style_map_file.read()
            result = mammoth.convert_to_html(docx_file, style_map=style_map_content)
        html = result.value
    return html

def force_center_blocks(html):
    soup = BeautifulSoup(html, "html.parser")
    # Centrer tous les <div> ou <p> qui ne sont pas déjà centrés et qui ne sont pas dans un tableau
    for tag in soup.find_all(['div', 'p']):
        if not tag.has_attr('style') and not tag.find_parent('table'):
            tag['style'] = 'text-align:center;'
    return str(soup)

def add_header_table_class(html):
    soup = BeautifulSoup(html, "html.parser")
    first_table = soup.find("table")
    if first_table:
        first_table["class"] = first_table.get("class", []) + ["amendements-header-table"]
    # Ajouter dir="rtl" à tous les tableaux
    for table in soup.find_all("table"):
        table["dir"] = "rtl"
    return str(soup)

# Je retire extract_colored_html et je restaure l'extraction Mammoth avec post-traitement lors de l'upload

def creer_amendements_depuis_blocs(blocs, utilisateur, fichier_obj):
    from bs4 import BeautifulSoup
    for bloc in blocs:
        if bloc['type'] == 'tableau':
            soup = BeautifulSoup(bloc['data'], 'html.parser')
            for row in soup.find_all('tr'):
                cells = row.find_all('td')
                if len(cells) >= 10:  # Adapter ce nombre selon vos champs
                    Amendement.objects.create(
                        numero_modification=cells[0].get_text(strip=True),
                        type_modification=cells[1].get_text(strip=True),
                        texte_amendement=cells[2].get_text(strip=True),
                        numero_article=cells[3].get_text(strip=True),
                        alinea=cells[4].get_text(strip=True),
                        article=cells[5].get_text(strip=True),
                        nom_loi=cells[6].get_text(strip=True),
                        texte_original=cells[7].get_text(strip=True),
                        texte_modifie=cells[8].get_text(strip=True),
                        justification=cells[9].get_text(strip=True),
                        utilisateur=utilisateur,
                        fichier=fichier_obj
                    )

import re

IGNORED_CELL_VALUES = {"", "النص", "الثاني", "الثالث", "الرابع", "الخامس"}

def extract_fields_from_block(text):
    result = {
        "numero_modification": "",
        "numero_groupe": "",
        "type_modification": "",
        "nom_loi": "",
        "article": "",
        "alinea": "",
        "numero_article": "",
        "texte_amendement": "",
    }
    
    # Détection améliorée du numéro de modification avec plusieurs patterns
    patterns_modification = [
        r"التعديل رقم\s*(\d+)",
        r"رقم التعديل\s*(\d+)",
        r"التعديل\s*(\d+)",
        r"amendment\s*(\d+)",
        r"numéro\s*(\d+)",
    ]
    for pattern in patterns_modification:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            result["numero_modification"] = m.group(1)
            print(f"✅ Numéro de modification détecté avec pattern '{pattern}': {m.group(1)}")
            break
    
    # Detection robuste du numero_groupe
    patterns_groupe = [
        r"الفريق\s*(\d+)",
        r"groupe\s*(\d+)",
        r"group\s*(\d+)",
    ]
    for pattern in patterns_groupe:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            result["numero_groupe"] = m.group(1)
            break
    
    # Detection robuste du type_modification
    if re.search(r"إضافة\s*مادة\s*جديدة", text):
        result["type_modification"] = "إضافة مادة جديدة"
    elif re.search(r"حذف", text):
        result["type_modification"] = "حذف"
    elif re.search(r"modification", text, re.IGNORECASE):
        result["type_modification"] = "modification"
    
    # Détection du nom de la loi : toute la ligne après le mot-clé
    keywords_loi = [r"مدونة", r"ضرائب", r"اعفئات", r"جدول", r"loi", r"law"]
    for kw in keywords_loi:
        m = re.search(rf"({kw}[^\n\r]*)", text)
        if m:
            result["nom_loi"] = m.group(1).strip()
            print(f"✅ Nom de loi détecté avec mot-clé '{kw}': {result['nom_loi']}")
            break
    
    # Détection de l'article
    patterns_article = [
        r"المادة\s*(\d+)",
        r"article\s*(\d+)",
        r"art\.\s*(\d+)",
    ]
    for pattern in patterns_article:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            result["article"] = m.group(1)
            break
    
    # Détection de l'alinéa
    m = re.search(r"البند\s*([IVXLCDM]+)", text)
    if m:
        result["alinea"] = m.group(1)
    
    # Détection du numéro d'article
    patterns_numero_article = [
        r"(?:المادة|الفصل)\s*(\d+)",
        r"(?:article|section)\s*(\d+)",
    ]
    for pattern in patterns_numero_article:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            result["numero_article"] = m.group(1)
            break
    
    # Texte d'amendement = tout le texte du bloc sans les labels connus
    labels = ["التعديل رقم", "رقم التعديل", "التعديل", "الفريق", "إضافة مادة جديدة", "حذف", "ضرائب", "اعفئات", "مدونة", "المادة", "البند", "الفصل", "amendment", "group", "article", "law", "جدول", "loi"]
    texte = text
    for label in labels:
        texte = re.sub(label + r".*?(?=\s|$)", "", texte, flags=re.IGNORECASE)
    result["texte_amendement"] = texte.strip()
    
    print(f"🔍 Champs extraits du bloc:")
    for key, value in result.items():
        if value:
            print(f"  - {key}: {value}")
    
    return result

def split_blocks_by_numero_modification(text):
    # Découpe le texte en sous-blocs à chaque "التعديل رقم" ou "رقم التعديل"
    pattern = r"((?:التعديل رقم|رقم التعديل)[^التعديل رقم]*)"
    blocks = re.findall(pattern, text)
    return blocks if blocks else [text]

def extract_cell_content(cell, champ):
    sub_table = cell.find("table")
    if sub_table:
        # Extraire le texte du sous-tableau uniquement, sans dupliquer le texte principal
        rows = []
        for tr in sub_table.find_all("tr"):
            cols = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
            rows.append("\t".join(cols))
        return "\n".join(rows)
    else:
        # Extraire uniquement le texte hors sous-tableau
        # On retire les sous-tableaux du contenu avant d'extraire le texte
        for table in cell.find_all("table"):
            table.decompose()
        content = cell.get_text(strip=True)
        titres = {
            "texte_original": ["النص الأصلي", "النص كما جاء في المشروع"],
            "texte_modifie": ["نص التعديل", "النص المعدل"],
            "justification": ["التعليل"],
        }
        for titre in titres.get(champ, []):
            if content.strip() == titre:
                return ""
            if content.startswith(titre):
                content = content[len(titre):].lstrip(" :،-")
        return content.strip()

DANGEROUS_TAGS = ["script", "iframe", "object", "embed", "link", "meta", "base", "form"]


def nettoyer_html(html):
    """Retire scripts, gestionnaires d'événements (onclick...) et URLs javascript: du HTML modifié."""
    soup = BeautifulSoup(html or "", "html.parser")
    for tag in soup.find_all(DANGEROUS_TAGS):
        tag.decompose()
    for tag in soup.find_all(True):
        for attr in list(tag.attrs):
            valeur = str(tag.attrs[attr]).strip().lower()
            if attr.lower().startswith("on") or valeur.startswith(("javascript:", "data:text/html")):
                del tag.attrs[attr]
    return str(soup)


def creer_amendements_depuis_html(html, utilisateur, fichier_obj):
    """
    Découpe le document en amendements (extraction.extraire_blocs), puis détecte
    les champs de chaque en-tête avec l'IA (repli automatique sur les règles).
    Renvoie le nombre d'amendements créés.
    """
    blocs = extraire_blocs(html)
    logger.info("Fichier %s : %d bloc(s) d'amendement trouvés", fichier_obj.fichier.name, len(blocs))
    crees = 0
    for bloc in blocs:
        if not (bloc["texte_original"] or bloc["texte_modifie"] or bloc["justification"]):
            continue
        valeurs, champs_detectes = field_detector.detecter(bloc["entete"], bloc["texte_original"])
        if not valeurs["numero_modification"]:
            logger.warning("Bloc ignoré (numéro d'amendement introuvable) : %s", bloc["entete"][:120])
            continue
        Amendement.objects.create(
            numero_modification=valeurs["numero_modification"],
            numero_groupe=valeurs["numero_groupe"] or None,
            type_modification=valeurs["type_modification"] or None,
            nom_loi=valeurs["nom_loi"][:255],
            article=valeurs["article"][:50],
            alinea=valeurs["alinea"] or None,
            numero_article=valeurs["numero_article"][:50],
            texte_amendement=bloc["texte_modifie"],
            texte_entete=bloc["entete"],
            texte_original=bloc["texte_original"],
            texte_modifie=bloc["texte_modifie"],
            justification=bloc["justification"],
            utilisateur=utilisateur,
            fichier=fichier_obj,
            champs_detectes=champs_detectes,
        )
        crees += 1
    return crees


def _appliquer_detection(amendement):
    """Relance la détection IA/règles sur un amendement existant et met à jour ses champs."""
    entete = amendement.texte_entete or " ".join(filter(None, [
        f"التعديل رقم : {amendement.numero_modification}" if amendement.numero_modification else "",
        f"الفريق {amendement.numero_groupe}" if amendement.numero_groupe else "",
        amendement.type_modification or "", amendement.nom_loi or "",
    ]))
    valeurs, champs_detectes = field_detector.detecter(entete, amendement.texte_original or "")
    for champ, valeur in valeurs.items():
        if valeur:
            setattr(amendement, champ, valeur)
    amendement.champs_detectes = champs_detectes
    amendement.save()
    return champs_detectes

@login_required
def upload_amendement(request):
    if request.method == 'POST':
        if 'fichier' in request.FILES:
            form = UploadWordForm(request.POST, request.FILES)
            if form.is_valid():
                fichier = request.FILES['fichier']
                nom_fichier = fichier.name
                ext = nom_fichier.rsplit('.', 1)[-1].lower()
                if ext != 'docx':
                    messages.error(request, "Le fichier doit être au format Word (.docx). Veuillez sélectionner un fichier valide.")
                    return render(request, 'amendements/upload.html', {
                        'form': form,
                        'error': "Le fichier doit être au format Word (.docx). Veuillez sélectionner un fichier valide."
                    })
                fichier_obj = AmendementFichier.objects.create(utilisateur=request.user, fichier=fichier)
                fichier_obj.refresh_from_db()
                try:
                    html_extrait = extract_html_content(fichier_obj.fichier.path)
                except RuntimeError as e:
                    fichier_obj.fichier.delete(save=False)
                    fichier_obj.delete()
                    messages.error(request, str(e))
                    return render(request, 'amendements/upload.html', {'form': form, 'error': str(e)})
                fichier_obj.extracted_content = html_extrait
                fichier_obj.save()
                nb = creer_amendements_depuis_html(html_extrait, request.user, fichier_obj)
                fichier_obj.fichier.close()
                methode = "règles (IA indisponible)" if field_detector.derniere_erreur else field_detector.description()
                if nb:
                    messages.success(request, f"{nb} amendement(s) détecté(s) — méthode : {methode}.")
                else:
                    messages.warning(request, "Aucun amendement détecté dans ce fichier.")
                if field_detector.derniere_erreur:
                    messages.warning(request, "IA indisponible, détection par règles utilisée : " + field_detector.derniere_erreur[:150])
                return render(request, 'amendements/valider_amendements.html', {
                    'extracted_content': html_extrait,
                    'fichier_obj': fichier_obj
                })
            else:
                return render(request, 'amendements/upload.html', {
                    'form': form,
                    'error': "Le formulaire n'est pas valide. Veuillez sélectionner un fichier Word (.docx)."
                })
        else:
            messages.success(request, "Amendements validés et stockés avec succès.")
            return redirect('fichiers_amendements')
    else:
        form = UploadWordForm()
    return render(request, 'amendements/upload.html', {'form': form})

@login_required
def valider_amendements(request):
    fichier_id = request.session.get('fichier_amendement_id')
    fichier_obj = None
    if fichier_id:
        try:
            fichier_obj = AmendementFichier.objects.get(id=fichier_id, utilisateur=request.user)
        except AmendementFichier.DoesNotExist:
            fichier_obj = None
    donnees_liste = request.session.get('amendements_extraits', [])
    debug_infos = []
    debug_donnees = donnees_liste
    if request.method == 'POST':
        created = 0
        if not donnees_liste:
            # Création d'un amendement de test pour garantir l'affichage
            amendement = Amendement(
                numero_modification='TEST',
                type_modification='TEST',
                texte_amendement='TEST',
                numero_article='TEST',
                alinea='TEST',
                article='TEST',
                nom_loi='TEST',
                texte_original='TEST',
                texte_modifie='TEST',
                justification='TEST',
                utilisateur=request.user,
                fichier=fichier_obj
            )
            amendement.save()
            created = 1
        else:
            for data in donnees_liste:
                amendement = Amendement(
                    numero_modification=data.get('numero_modification', ''),
                    type_modification=data.get('type_modification', ''),
                    texte_amendement=data.get('texte_amendement', ''),
                    numero_article=data.get('numero_article', ''),
                    alinea=data.get('alinea', ''),
                    article=data.get('article', ''),
                    nom_loi=data.get('nom_loi', ''),
                    texte_original=data.get('texte_original', ''),
                    texte_modifie=data.get('texte_modifie', ''),
                    justification=data.get('justification', ''),
                    utilisateur=request.user,
                    fichier=fichier_obj
                )
                amendement.save()
                created += 1
        if created == 0:
            messages.error(request, "Aucun amendement n'a été créé. Veuillez recharger le fichier.")
            return redirect('upload_amendement')
        messages.success(request, f"{created} amendements enregistrés avec succès.")
        if 'fichier_amendement_id' in request.session:
            del request.session['fichier_amendement_id']
        if 'amendements_extraits' in request.session:
            del request.session['amendements_extraits']
        return redirect('fichiers_amendements')
    else:
        AmendementFormSet = formset_factory(AmendementForm, extra=0)
        formset = AmendementFormSet(initial=donnees_liste)
    return render(request, 'amendements/valider_amendements.html', {'formset': formset, 'fichier_obj': fichier_obj})

@login_required
def fichiers_amendements(request):
    fichiers = AmendementFichier.objects.filter(utilisateur=request.user)
    return render(request, 'amendements/fichiers.html', {'fichiers': fichiers})

@login_required
def upload_fichier_amendement(request):
    """Ancienne page d'upload sans extraction : redirige vers la page d'upload avec détection."""
    return redirect('upload_amendement')

@login_required
@require_POST
def supprimer_fichier_amendement(request, fichier_id):
    fichier = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    fichier.fichier.delete()
    fichier.delete()
    messages.success(request, "Fichier supprimé avec succès.")
    return redirect('fichiers_amendements')

@login_required
def renommer_fichier_amendement(request, fichier_id):
    fichier = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    if fichier.utilisateur != request.user:
        return HttpResponseForbidden()
    if request.method == 'POST':
        nouveau_nom = request.POST.get('nom_personnalise', '').strip()
        if not nouveau_nom:
            messages.error(request, "Le nom personnalisé ne peut pas être vide.")
        else:
            fichier.nom_personnalise = nouveau_nom
            fichier.save()
            messages.success(request, "Nom du fichier modifié avec succès.")
        return redirect('fichiers_amendements')
    return render(request, 'amendements/renommer_fichier.html', {'fichier': fichier})

@login_required
def detail_fichier_amendement(request, fichier_id):
    return redirect('fichier_mammoth', fichier_id=fichier_id)

@login_required
def afficher_fichier_mammoth(request, fichier_id):
    fichier_obj = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    html_extrait = fichier_obj.extracted_content  # On lit le HTML stocké, PAS de ré-extraction !
    return render(request, "amendements/detail_fichier.html", {
        'fichier': fichier_obj,
        'extracted_content': html_extrait,
    })

# Ajout d'une vue robuste pour la modification du contenu extrait
@login_required
def modifier_fichier_extrait(request, fichier_id):
    fichier = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    if request.method == 'POST':
        try:
            contenu = request.POST.get('extracted_content', '')
            if not contenu.strip():
                messages.error(request, "Le contenu ne peut pas être vide.")
                return render(request, 'amendements/modifier_fichier_extrait.html', {
                    'fichier': fichier,
                    'extracted_content': contenu
                })
            fichier.extracted_content = nettoyer_html(contenu)
            fichier.save()
            messages.success(request, "Le contenu a bien été enregistré.")
            return redirect('fichier_mammoth', fichier_id=fichier.id)
        except Exception as e:
            messages.error(request, f"Erreur lors de l'enregistrement : {str(e)}")
            return render(request, 'amendements/modifier_fichier_extrait.html', {
                'fichier': fichier,
                'extracted_content': request.POST.get('extracted_content', '')
            })
    else:
        return render(request, 'amendements/modifier_fichier_extrait.html', {
            'fichier': fichier,
            'extracted_content': fichier.extracted_content
        })

# Suppression de la vue export_pdf_amendement et de son code associé

@login_required
def afficher_pdf_amendement(request, fichier_id):
    fichier_obj = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    pdf_url = reverse('telecharger_pdf_libreoffice', args=[fichier_id]) + '?inline=1'
    return render(request, 'amendements/afficher_pdf.html', {
        'fichier_obj': fichier_obj,
        'pdf_url': pdf_url
    })

@login_required
def supprimer_tous_fichiers_amendement(request):
    if request.method == 'POST':
        fichiers = AmendementFichier.objects.filter(utilisateur=request.user)
        count = fichiers.count()
        for fichier in fichiers:
            fichier.fichier.delete()
            fichier.delete()
        messages.success(request, f"{count} fichier(s) supprimé(s) avec succès.")
        return redirect('fichiers_amendements')
    return render(request, 'amendements/confirm_supprimer_tous.html')

def cell_style_to_css(cell):
    css = []
    # Couleur de fond
    shading = cell._element.find('.//w:shd', namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})
    if shading is not None and shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill'):
        color = shading.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
        css.append(f'background-color: #{color};')
    # TODO: Ajouter bordures, padding, etc.
    return ' '.join(css)

def get_run_color_from_xml(run):
    r = run._element
    rPr = r.find(qn('w:rPr'))
    # Couleur directe
    if rPr is not None:
        color = rPr.find(qn('w:color'))
        if color is not None:
            val = color.get(qn('w:val'))
            if val and val.lower() != 'auto':
                val = val.upper()
                if len(val) == 6:
                    return f"#{val}"
                elif len(val) == 3:
                    return f"#{val[0]*2}{val[1]*2}{val[2]*2}"
    # Couleur héritée du style
    if run.style and hasattr(run.style, 'font') and run.style.font.color and run.style.font.color.rgb:
        return f"#{run.style.font.color.rgb}"
    return None

def convert_docx_to_pdf(docx_path, output_dir):
    cmd = [
        get_soffice_path(),
        "--headless",
        "--convert-to", "pdf",
        "--outdir", output_dir,
        docx_path
    ]
    try:
        subprocess.run(cmd, check=True, timeout=180)
    except (FileNotFoundError, PermissionError):
        raise RuntimeError("LibreOffice est introuvable. Installez-le ou indiquez son chemin dans SOFFICE_PATH (fichier .env).")
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired):
        raise RuntimeError("La conversion en PDF a échoué (LibreOffice).")
    base = os.path.splitext(os.path.basename(docx_path))[0]
    pdf_path = os.path.join(output_dir, base + ".pdf")
    return pdf_path

@login_required
def telecharger_pdf_libreoffice(request, fichier_id):
    fichier_obj = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    docx_path = fichier_obj.fichier.path
    output_dir = os.path.dirname(docx_path)
    try:
        pdf_path = convert_docx_to_pdf(docx_path, output_dir)
    except RuntimeError as e:
        messages.error(request, str(e))
        return redirect('fichiers_amendements')
    if not os.path.exists(pdf_path):
        messages.error(request, "Le PDF n'a pas pu être généré.")
        return redirect('fichiers_amendements')
    inline = request.GET.get('inline') == '1'  # affichage dans la page d'aperçu
    return FileResponse(open(pdf_path, 'rb'), as_attachment=not inline, filename=os.path.basename(pdf_path))

@login_required
def afficher_champs_detectes(request, amendement_id):
    """
    Affiche les champs détectés automatiquement par HuggingFace pour un amendement
    """
    amendement = get_object_or_404(Amendement, id=amendement_id, utilisateur=request.user)
    
    context = {
        'amendement': amendement,
        'champs_detectes': amendement.champs_detectes or {},
    }
    
    return render(request, 'amendements/champs_detectes.html', context)

@login_required
def re_detecter_champs(request, amendement_id):
    """
    Re-détecte les champs avec HuggingFace pour un amendement existant
    """
    amendement = get_object_or_404(Amendement, id=amendement_id, utilisateur=request.user)
    
    _appliquer_detection(amendement)

    messages.success(request, f"Champs re-détectés pour l'amendement {amendement.numero_modification}")
    return redirect('afficher_champs_detectes', amendement_id=amendement_id)

@login_required
def statistiques_detection(request, fichier_id):
    """
    Affiche les statistiques de détection pour un fichier
    """
    fichier = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    amendements = Amendement.objects.filter(fichier=fichier)
    
    # Statistiques
    total_amendements = amendements.count()
    amendements_avec_champs = amendements.filter(champs_detectes__isnull=False).exclude(champs_detectes={}).count()
    amendements_sans_champs = total_amendements - amendements_avec_champs
    
    # Analyse des champs détectés
    champs_stats = {}
    for amendement in amendements:
        if amendement.champs_detectes:
            for champ, donnees in amendement.champs_detectes.items():
                if champ not in champs_stats:
                    champs_stats[champ] = {
                        'count': 0,
                        'avg_confidence': 0,
                        'methods': {}
                    }
                champs_stats[champ]['count'] += 1
                champs_stats[champ]['avg_confidence'] += donnees.get('confidence', 0)
                
                method = donnees.get('method', 'unknown')
                if method not in champs_stats[champ]['methods']:
                    champs_stats[champ]['methods'][method] = 0
                champs_stats[champ]['methods'][method] += 1
    
    # Calculer les moyennes
    for champ in champs_stats:
        if champs_stats[champ]['count'] > 0:
            champs_stats[champ]['avg_confidence'] /= champs_stats[champ]['count']
    
    context = {
        'fichier': fichier,
        'total_amendements': total_amendements,
        'amendements_avec_champs': amendements_avec_champs,
        'amendements_sans_champs': amendements_sans_champs,
        'pourcentage_detection': (amendements_avec_champs / total_amendements * 100) if total_amendements > 0 else 0,
        'champs_stats': champs_stats,
        'amendements': amendements,
    }
    
    return render(request, 'amendements/statistiques_detection.html', context)
