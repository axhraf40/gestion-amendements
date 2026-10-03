<div align="center">

<img src="charge_amendements/static/charge_amendements/logo.svg" alt="logo" width="72">

# PLF Amendements Manager

**Import, review and export Finance Bill amendments written in Arabic Word documents.**

[![tests](https://github.com/axhraf40/gestion-amendements/actions/workflows/tests.yml/badge.svg)](https://github.com/axhraf40/gestion-amendements/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Django](https://img.shields.io/badge/Django-5.x-092E20?logo=django&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)

[Features](#features) · [Screenshots](#screenshots) · [Getting started](#getting-started) · [Tests](#running-the-tests) · [Security](#security)

<img src="docs/screenshots/extraction.png" alt="Extracted amendments rendered right-to-left" width="820">

</div>

## About

During the preparation of a Finance Bill (*Projet de Loi de Finances*, PLF), parliamentary groups submit their
amendments as Word documents written in Arabic, full of tables, merged cells and colour-coded changes.
Re-typing them is slow and error-prone.

This Django application reads those documents, rebuilds their tables as clean **right-to-left HTML** (keeping
bold, colours, highlights and strike-through), stores each amendment in a database, and lets users review,
correct and export the result as a **PDF**.

Built during an internship.

## Features

| | |
|---|---|
| 📄 **Word (.docx) import** | Parses paragraphs, tables, nested tables and merged cells with `python-docx` |
| 🎨 **Formatting preserved** | Bold, italic, underline, strike-through, text colour and highlight |
| ↩️ **Arabic / RTL** | Tables and text rendered right-to-left |
| 🗂️ **Structured records** | Table rows saved as `Amendement` records (number, type, article, law, original text, amended text, justification…) |
| ✏️ **Review & edit** | Check the extracted content and correct it in the browser |
| 🖨️ **PDF export** | One-click export with WeasyPrint |
| 👤 **User accounts** | Login, profile, password change; each user only sees their own files |
| 📊 **Dashboard** | Number of files uploaded per user |

## Screenshots

> All screenshots use a **fictional** sample document (`scripts/generate_sample_docx.py`).

| Upload a Word file | My files |
|---|---|
| <img src="docs/screenshots/upload.png" alt="Upload page"> | <img src="docs/screenshots/files.png" alt="Files list"> |

| Edit the extracted content | PDF export |
|---|---|
| <img src="docs/screenshots/edit.png" alt="Edit page"> | <img src="docs/screenshots/pdf-export.png" alt="Exported PDF"> |

<details>
<summary>More screens (login, home, statistics)</summary>

| Login | Home | Statistics |
|---|---|---|
| <img src="docs/screenshots/login.png" alt="Login"> | <img src="docs/screenshots/home.png" alt="Home"> | <img src="docs/screenshots/stats.png" alt="Statistics"> |

</details>

## How it works

```
 .docx upload ──► python-docx parser ──► RTL HTML (tables + formatting)
                         │                        │
                         ▼                        ▼
              Amendement records          review / edit in browser
                                                  │
                                                  ▼
                                          PDF export (WeasyPrint)
```

## Tech stack

| Layer | Tools |
|---|---|
| Backend | Python, Django 5 |
| Document parsing | python-docx, lxml, BeautifulSoup, mammoth |
| PDF | WeasyPrint |
| Database | SQLite (dev), any Django-supported database in production |
| Front end | Django templates, CSS |
| CI | GitHub Actions |

## Project structure

```
gestion-amendements/
├── charge_amendements/     Project settings, URLs, base templates, CSS and logo
├── amendements/            Core app: upload, Word parsing, review/edit, PDF export, tests
├── users/                  Profile and password change
├── stats/                  Dashboard
├── scripts/                Sample-document generator and .docx inspection tools
├── docs/screenshots/       Images used in this README
├── .github/workflows/      Continuous integration (tests on every push)
├── .env.example            Template for local environment variables
└── requirements.txt
```

## Getting started

### 1. Clone and install

```bash
git clone https://github.com/axhraf40/gestion-amendements.git
cd gestion-amendements
python -m venv .venv
# Windows: .venv\Scripts\activate    macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

> WeasyPrint needs the Pango library. On Windows, follow the
> [WeasyPrint installation guide](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html).
> On Ubuntu/Debian: `sudo apt install libpango-1.0-0 libpangoft2-1.0-0`.

### 2. Configure environment variables

```bash
cp .env.example .env      # Windows: copy .env.example .env
```

For local development keep `DJANGO_DEBUG=True`. For production, set your own `DJANGO_SECRET_KEY`
(generate one with `python -c "import secrets; print(secrets.token_urlsafe(50))"`).

### 3. Create the database and a user

```bash
python manage.py migrate
python manage.py createsuperuser
```

### 4. Run

```bash
python manage.py runserver
```

Open http://127.0.0.1:8000 and log in.

### 5. Try it with a sample document

```bash
python scripts/generate_sample_docx.py exemple_amendements.docx
```

Then upload `exemple_amendements.docx` from the **Charger un amendement** page.

## Running the tests

```bash
python manage.py test
```

The tests build a synthetic Word document on the fly and cover upload and extraction, formatting, editing,
PDF export, access control between users, input validation and password changes.

## Developer tools

| Script | Purpose |
|---|---|
| `scripts/generate_sample_docx.py` | Create a fictional amendments document |
| `scripts/inspect_docx_tables.py <file.docx>` | Print the tables (and nested tables) of a document |
| `scripts/inspect_docx_xml.py <file.docx>` | Show where tables sit in the raw Word XML |
| `scripts/show_extracted_html.py <id>` | Print the HTML stored for an uploaded file |

## Security

- No secrets in the code: settings come from environment variables or a git-ignored `.env` file.
- The app refuses to start in production (`DJANGO_DEBUG=False`) without a `DJANGO_SECRET_KEY`.
- Users can only view, edit, export or delete their own files.
- Deleting requires a POST request with a CSRF token.
- Only `.docx` uploads are accepted; HTML edited by users is cleaned of scripts and event handlers.
- Uploaded documents and the database are git-ignored.

## License

[MIT](LICENSE)
