from django.shortcuts import render, redirect, get_object_or_404
from .forms import UploadWordForm, AmendementForm
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from .models import Amendement, AmendementFichier
from .forms import AmendementFichierForm
import docx
import os
from django.forms import formset_factory
import re
from docx.table import Table as _Table
from docx.text.paragraph import Paragraph as _Paragraph
from zipfile import BadZipFile
from docx.opc.exceptions import PackageNotFoundError
import mammoth
from bs4 import BeautifulSoup
from django.views.decorators.http import require_http_methods
from django.urls import reverse
from django.http import HttpResponse
from django.utils.http import content_disposition_header
from django.template.loader import render_to_string
import weasyprint

# Nouvelle extraction : tous les tableaux

def normalize_arabic(text):
    # Supprime les espaces multiples, les retours à la ligne, etc.
    text = re.sub(r'[\s\u200c\u200d]+', ' ', text)
    return text.strip()

def extraire_blocs_word(fichier):
    doc = docx.Document(fichier)
    blocs = []
    buffer_paragraphs = []
    for element in doc.element.body:
        if element.tag.endswith('tbl'):
            if buffer_paragraphs:
                for para in buffer_paragraphs:
                    if para.text.strip():
                        blocs.append({'type': 'texte', 'data': paragraph_to_html(para)})
                buffer_paragraphs = []
            table = _Table(element, doc)
            blocs.append({'type': 'tableau', 'data': advanced_table_to_html(table)})
        elif element.tag.endswith('p'):
            para = _Paragraph(element, doc)
            buffer_paragraphs.append(para)
    if buffer_paragraphs:
        for para in buffer_paragraphs:
            if para.text.strip():
                blocs.append({'type': 'texte', 'data': paragraph_to_html(para)})
    return blocs

def advanced_table_to_html(table):
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    tbl_xml = table._element
    rows = tbl_xml.findall('.//w:tr', ns)
    grid = []
    max_cols = 0
    for r in rows:
        row = []
        cells = r.findall('./w:tc', ns)
        has_subtable = False
        for c in cells:
            cell_content = ''
            for child in c:
                if child.tag.endswith('p'):
                    para = _Paragraph(child, table._parent)
                    html_para = paragraph_to_html(para).strip()
                    if html_para and html_para != '<div></div>':
                        cell_content += html_para + '<br>'
                elif child.tag.endswith('tbl'):
                    sub_table = _Table(child, table._parent)
                    cell_content += advanced_table_to_html(sub_table)
                    has_subtable = True
            while cell_content.endswith('<br>'):
                cell_content = cell_content[:-len('<br>')]
            gridspan = c.find('.//w:gridSpan', ns)
            colspan = int(gridspan.get('{%s}val' % ns['w'])) if gridspan is not None else 1
            vmerge = c.find('.//w:vMerge', ns)
            if cell_content.strip():
                row.append({'text': cell_content, 'colspan': colspan, 'vmerge': vmerge is not None})
        total_cols = sum(cell['colspan'] for cell in row)
        if total_cols > max_cols:
            max_cols = total_cols
        if row:
            grid.append({'cells': row, 'has_subtable': has_subtable, 'total_cols': total_cols})
    max_cols = min(max_cols, 10)
    html = '<table class="am-table" dir="rtl" style="margin:10px 0; text-align:right; border:1px solid #888; background:#fff; width:100%; table-layout:fixed;">'
    for rowinfo in grid:
        row = rowinfo['cells']
        has_subtable = rowinfo['has_subtable']
        total_cols = rowinfo['total_cols']
        html += '<tr>'
        for cell in row:
            colspan_attr = f' colspan="{cell["colspan"]}"' if cell["colspan"] > 1 else ''
            html += f'<td{colspan_attr}>{cell["text"]}</td>'
        # On ne complète la ligne que si elle n'a pas de sous-tableau
        if not has_subtable and total_cols < max_cols:
            for _ in range(max_cols - total_cols):
                html += '<td></td>'
        html += '</tr>'
    html += '</table>'
    return html

# --- Extraction avancée des styles pour python-docx ---
def run_to_html(run):
    style = ''
    if run.bold:
        style += 'font-weight:bold;'
    if run.italic:
        style += 'font-style:italic;'
    if run.underline:
        style += 'text-decoration:underline;'
    if hasattr(run.font, 'strike') and run.font.strike:
        style += 'text-decoration:line-through;'
    if run.font.color and run.font.color.rgb:
        style += f'color:#{run.font.color.rgb};'
    # Ajout du surlignage (highlight)
    if hasattr(run.font, 'highlight_color') and run.font.highlight_color:
        # highlight_color peut être une valeur Enum ou None
        highlight = run.font.highlight_color
        if hasattr(highlight, 'rgb') and highlight.rgb:
            style += f'background-color:#{highlight.rgb};'
        else:
            # Couleurs Word prédéfinies (ex: YELLOW, GREEN, etc.)
            highlight_map = {
                'YELLOW': '#FFFF00',
                'GREEN': '#00FF00',
                'CYAN': '#00FFFF',
                'MAGENTA': '#FF00FF',
                'BLUE': '#0000FF',
                'RED': '#FF0000',
                'DARKBLUE': '#000080',
                'DARKCYAN': '#008080',
                'DARKGREEN': '#008000',
                'DARKMAGENTA': '#800080',
                'DARKRED': '#800000',
                'DARKYELLOW': '#808000',
                'BLACK': '#000000',
                'WHITE': '#FFFFFF',
                'NONE': '',
                'AUTO': '',
            }
            color_name = str(highlight)
            if color_name in highlight_map and highlight_map[color_name]:
                style += f'background-color:{highlight_map[color_name]};'
    text = run.text.replace('\n', '<br>')
    if style:
        return f'<span style="{style}">{text}</span>'
    else:
        return text

def paragraph_to_html(paragraph):
    align = ''
    if paragraph.alignment == 1:
        align = 'text-align:center;'
    elif paragraph.alignment == 2:
        align = 'text-align:right;'
    elif paragraph.alignment == 3:
        align = 'text-align:justify;'
    html = f'<div style="{align}">'  # ouverture du div
    for run in paragraph.runs:
        html += run_to_html(run)
    html += '</div>'  # fermeture du div
    return html

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

DANGEROUS_TAGS = ['script', 'iframe', 'object', 'embed', 'link', 'meta', 'base', 'form']


def nettoyer_html(html):
    """Remove scripts, event handlers and javascript: URLs from user-edited HTML."""
    soup = BeautifulSoup(html or '', 'html.parser')
    for tag in soup.find_all(DANGEROUS_TAGS):
        tag.decompose()
    for tag in soup.find_all(True):
        for attr in list(tag.attrs):
            value = str(tag.attrs[attr]).strip().lower()
            if attr.lower().startswith('on') or value.startswith('javascript:'):
                del tag.attrs[attr]
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

@login_required
def upload_amendement(request):
    if request.method == 'POST':
        if 'fichier' in request.FILES:
            form = UploadWordForm(request.POST, request.FILES)
            if form.is_valid():
                fichier = request.FILES['fichier']
                fichier_obj = AmendementFichier.objects.create(utilisateur=request.user, fichier=fichier)
                fichier_obj.refresh_from_db()
                # Utiliser l'extraction avancée custom (python-docx) pour garantir la détection des couleurs
                try:
                    blocs = extraire_blocs_word(fichier_obj.fichier.path)
                except (BadZipFile, PackageNotFoundError, ValueError, KeyError):
                    fichier_obj.fichier.delete()
                    fichier_obj.delete()
                    return render(request, 'amendements/upload.html', {
                        'form': UploadWordForm(),
                        'error': "Impossible de lire ce fichier. Veuillez envoyer un document Word (.docx) valide."
                    })
                html_extrait = ''
                for bloc in blocs:
                    html_extrait += bloc['data']
                fichier_obj.extracted_content = html_extrait
                fichier_obj.save()
                creer_amendements_depuis_blocs(blocs, request.user, fichier_obj)
                fichier_obj.fichier.close()
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
            # POST sans fichier = validation
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
        if donnees_liste:
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
    fichiers = AmendementFichier.objects.filter(utilisateur=request.user)
    if fichiers.exists():
        messages.error(request, "Vous avez déjà un fichier chargé. Supprimez-le ou modifiez-le avant d'en ajouter un nouveau.")
        return redirect('fichiers_amendements')
    if request.method == 'POST':
        form = AmendementFichierForm(request.POST, request.FILES)
        if form.is_valid():
            fichier_obj = form.save(commit=False)
            fichier_obj.utilisateur = request.user
            fichier_obj.save()
            messages.success(request, "Fichier uploadé avec succès.")
            return redirect('fichiers_amendements')
    else:
        form = AmendementFichierForm()
    return render(request, 'amendements/upload_fichier.html', {'form': form, 'fichiers': fichiers})

@login_required
@require_http_methods(["POST"])
def supprimer_fichier_amendement(request, fichier_id):
    fichier = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    fichier.fichier.delete()
    fichier.delete()
    messages.success(request, "Fichier supprimé avec succès.")
    return redirect('fichiers_amendements')

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

@login_required
@require_http_methods(["GET", "POST"])
def modifier_fichier_extrait(request, fichier_id):
    fichier_obj = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    if request.method == "POST":
        nouveau_html = nettoyer_html(request.POST.get("extracted_content", ""))
        fichier_obj.extracted_content = nouveau_html
        fichier_obj.save()
        messages.success(request, "Contenu extrait modifié avec succès.")
        return redirect(reverse('fichier_mammoth', args=[fichier_id]))
    return render(request, "amendements/modifier_fichier_extrait.html", {
        'fichier': fichier_obj,
        'extracted_content': fichier_obj.extracted_content,
    })

@login_required
def export_pdf_amendement(request, fichier_id):
    fichier_obj = get_object_or_404(AmendementFichier, id=fichier_id, utilisateur=request.user)
    html_content = fichier_obj.extracted_content or "<p>Aucun contenu extrait.</p>"
    html = render_to_string('amendements/pdf_template.html', {
        'fichier': fichier_obj,
        'extracted_content': html_content,
    })
    nom_fichier = fichier_obj.fichier.name
    if nom_fichier.startswith('amendements_uploads/'):
        nom_fichier = nom_fichier[len('amendements_uploads/'):]
    nom_pdf = os.path.basename(nom_fichier).rsplit('.', 1)[0] + '_extrait.pdf'
    response = HttpResponse(content_type='application/pdf')
    response['Content-Disposition'] = content_disposition_header(True, nom_pdf)
    weasyprint.HTML(string=html).write_pdf(response)
    return response
