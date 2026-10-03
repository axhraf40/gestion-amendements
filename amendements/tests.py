import io
import shutil
import tempfile

import docx
from docx.enum.text import WD_COLOR_INDEX
from django.contrib.auth.models import User
from django.utils.crypto import get_random_string
from django.core.files.uploadedfile import SimpleUploadedFile
from django.test import TestCase, override_settings
from django.urls import reverse

from .models import Amendement, AmendementFichier
from .views import nettoyer_html

TEMP_MEDIA = tempfile.mkdtemp()


def random_password():
    """Throwaway credential generated at test time (nothing stored in the repo)."""
    return get_random_string(20)

HEADERS = [
    'رقم التعديل', 'نوع التعديل', 'نص التعديل', 'رقم المادة', 'البند',
    'الفصل', 'اسم القانون', 'النص الأصلي', 'النص المعدل', 'التعليل',
]


def make_docx(rows=2):
    """Build a small synthetic amendments document (no real data)."""
    document = docx.Document()
    document.add_paragraph('الفريق التجريبي')
    table = document.add_table(rows=rows + 1, cols=len(HEADERS))
    for i, header in enumerate(HEADERS):
        table.cell(0, i).text = header
    for r in range(1, rows + 1):
        values = [str(r), 'تعديل', f'نص {r}', '3', 'I', '70', 'قانون المالية',
                  'النص القديم', 'النص الجديد', 'تعليل']
        for i, value in enumerate(values):
            table.cell(r, i).text = value
    run = document.add_paragraph().add_run('نص ملون')
    run.bold = True
    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def upload(name='test.docx', content=None):
    return SimpleUploadedFile(
        name, content if content is not None else make_docx(),
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )


@override_settings(MEDIA_ROOT=TEMP_MEDIA)
class AmendementsTests(TestCase):
    @classmethod
    def tearDownClass(cls):
        super().tearDownClass()
        shutil.rmtree(TEMP_MEDIA, ignore_errors=True)

    def setUp(self):
        self.alice_pw, self.bob_pw = random_password(), random_password()
        self.alice = User.objects.create_user('alice', password=self.alice_pw)
        self.bob = User.objects.create_user('bob', password=self.bob_pw)
        self.client.login(username='alice', password=self.alice_pw)

    def _upload(self):
        response = self.client.post(reverse('upload_amendement'), {'fichier': upload()})
        self.assertEqual(response.status_code, 200)
        return AmendementFichier.objects.get(utilisateur=self.alice)

    def test_pages_require_login(self):
        self.client.logout()
        for name in ['upload_amendement', 'fichiers_amendements', 'stats_dashboard', 'profile']:
            response = self.client.get(reverse(name))
            self.assertEqual(response.status_code, 302, name)
            self.assertIn('/accounts/login/', response['Location'])

    def test_logged_in_pages_render(self):
        for name in ['home', 'upload_amendement', 'fichiers_amendements', 'stats_dashboard',
                     'profile', 'change_password']:
            self.assertEqual(self.client.get(reverse(name)).status_code, 200, name)

    def test_upload_extracts_tables_and_creates_amendements(self):
        fichier = self._upload()
        self.assertIn('<table', fichier.extracted_content)
        self.assertIn('font-weight:bold', fichier.extracted_content)
        self.assertIn('background-color:#FFFF00', fichier.extracted_content)
        self.assertEqual(Amendement.objects.filter(fichier=fichier).count(), 3)

    def test_upload_rejects_non_docx(self):
        response = self.client.post(reverse('upload_amendement'),
                                    {'fichier': SimpleUploadedFile('x.txt', b'hello')})
        self.assertEqual(response.status_code, 200)
        self.assertEqual(AmendementFichier.objects.count(), 0)

    def test_upload_rejects_corrupt_docx(self):
        response = self.client.post(reverse('upload_amendement'),
                                    {'fichier': upload(content=b'not a zip')})
        self.assertEqual(response.status_code, 200)
        self.assertContains(response, 'valide')
        self.assertEqual(AmendementFichier.objects.count(), 0)

    def test_detail_edit_and_pdf_export(self):
        fichier = self._upload()
        self.assertEqual(self.client.get(reverse('fichier_mammoth', args=[fichier.id])).status_code, 200)
        response = self.client.post(reverse('modifier_fichier_extrait', args=[fichier.id]),
                                    {'extracted_content': '<p>Modifié</p>'})
        self.assertEqual(response.status_code, 302)
        fichier.refresh_from_db()
        self.assertEqual(fichier.extracted_content, '<p>Modifié</p>')
        response = self.client.get(reverse('export_pdf_amendement', args=[fichier.id]))
        self.assertEqual(response.status_code, 200)
        self.assertEqual(response['Content-Type'], 'application/pdf')
        self.assertTrue(response.content.startswith(b'%PDF'))

    def test_edit_strips_scripts(self):
        fichier = self._upload()
        self.client.post(reverse('modifier_fichier_extrait', args=[fichier.id]), {
            'extracted_content': '<p onclick="x()">ok</p><script>alert(1)</script><a href="javascript:x()">l</a>'})
        fichier.refresh_from_db()
        self.assertNotIn('script', fichier.extracted_content)
        self.assertNotIn('onclick', fichier.extracted_content)
        self.assertNotIn('javascript:', fichier.extracted_content)
        self.assertIn('ok', fichier.extracted_content)

    def test_other_users_cannot_access_files(self):
        fichier = self._upload()
        self.client.logout()
        self.client.login(username='bob', password=self.bob_pw)
        for name in ['fichier_mammoth', 'modifier_fichier_extrait', 'export_pdf_amendement']:
            self.assertEqual(self.client.get(reverse(name, args=[fichier.id])).status_code, 404, name)
        response = self.client.post(reverse('supprimer_fichier_amendement', args=[fichier.id]))
        self.assertEqual(response.status_code, 404)
        self.assertTrue(AmendementFichier.objects.filter(id=fichier.id).exists())

    def test_delete_requires_post(self):
        fichier = self._upload()
        url = reverse('supprimer_fichier_amendement', args=[fichier.id])
        self.assertEqual(self.client.get(url).status_code, 405)
        self.assertEqual(self.client.post(url).status_code, 302)
        self.assertFalse(AmendementFichier.objects.filter(id=fichier.id).exists())

    def test_delete_unknown_file_returns_404(self):
        response = self.client.post(reverse('supprimer_fichier_amendement', args=[99999]))
        self.assertEqual(response.status_code, 404)

    def test_stats_counts_only_own_files(self):
        self._upload()
        response = self.client.get(reverse('stats_dashboard'))
        self.assertEqual(response.context['total_fichiers'], 1)


class SanitizerTests(TestCase):
    def test_keeps_formatting(self):
        html = '<span style="font-weight:bold;">x</span><table dir="rtl"><tr><td>a</td></tr></table>'
        self.assertEqual(nettoyer_html(html), html)


class UserTests(TestCase):
    def setUp(self):
        self.old_pw, self.new_pw = random_password(), random_password()
        self.user = User.objects.create_user('carol', password=self.old_pw)
        self.client.login(username='carol', password=self.old_pw)

    def test_change_password(self):
        response = self.client.post(reverse('change_password'), {
            'old_password': self.old_pw, 'new_password1': self.new_pw, 'new_password2': self.new_pw})
        self.assertEqual(response.status_code, 302)
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password(self.new_pw))

    def test_change_password_wrong_old(self):
        self.client.post(reverse('change_password'), {
            'old_password': random_password(), 'new_password1': self.new_pw, 'new_password2': self.new_pw})
        self.user.refresh_from_db()
        self.assertTrue(self.user.check_password(self.old_pw))
