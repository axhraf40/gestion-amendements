# PLF Amendements Manager

[![tests](https://github.com/axhraf40/gestion-amendements/actions/workflows/tests.yml/badge.svg)](https://github.com/axhraf40/gestion-amendements/actions/workflows/tests.yml)
![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
![Django](https://img.shields.io/badge/Django-5.x-092E20?logo=django&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green)

A Django web application for managing **amendments to a Finance Bill (Projet de Loi de Finances, PLF)**.
Parliamentary groups submit their amendments as Word documents written in Arabic. The app reads those
documents, turns their tables into clean right-to-left HTML (keeping bold, colours and highlights),
stores each amendment in a database, and lets users review, correct and export the result as a PDF.

Built during an internship.

## Features

- **Word (.docx) import**: parses paragraphs, tables, nested tables and merged cells with `python-docx`, preserving formatting (bold, italic, underline, strike-through, text colour, highlight).
- **Arabic / RTL support**: extracted tables are rendered right-to-left.
- **Automatic amendment records**: each table row (amendment number, type, article, clause, law, original text, amended text, justification…) is saved as an `Amendement` record.
- **Review & edit**: view the extracted content and correct it in the browser.
- **PDF export**: one-click export of the extracted content with WeasyPrint.
- **User accounts**: login/logout, profile and password change. Every user only sees their own files.
- **Stats dashboard**: number of files uploaded per user.

## Tech stack

| Layer | Tools |
|---|---|
| Backend | Python, Django 5 |
| Document parsing | python-docx, lxml, mammoth, BeautifulSoup |
| PDF | WeasyPrint |
| Database | SQLite (dev), any Django-supported DB in production |
| Front end | Django templates, CSS |
| CI | GitHub Actions |

## Project structure

```
charge_amendements/   Project settings, URLs, base templates and static files
amendements/          Upload, Word parsing, review/edit, PDF export (core app)
users/                Profile and password change
stats/                Dashboard
```

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

Edit `.env` and set `DJANGO_SECRET_KEY` (generate one with
`python -c "import secrets; print(secrets.token_urlsafe(50))"`). For local development keep `DJANGO_DEBUG=True`.

### 3. Create the database and an admin user

```bash
python manage.py migrate
python manage.py createsuperuser
```

### 4. Run

```bash
python manage.py runserver
```

Open http://127.0.0.1:8000, log in, and upload a `.docx` file of amendments.

## Running the tests

```bash
python manage.py test
```

The tests build a synthetic Word document on the fly (no real data is stored in the repository) and cover
upload and extraction, editing, PDF export, access control between users, input validation, and password changes.

## Security

- No secrets in the code: the secret key and all environment-specific settings come from environment variables / `.env` (git-ignored).
- The app refuses to start in production (`DJANGO_DEBUG=False`) without a `DJANGO_SECRET_KEY`.
- Users can only view, edit, export or delete their own files.
- Deleting requires a POST request with a CSRF token.
- Only `.docx` uploads are accepted, and HTML edited by users is cleaned of scripts and event handlers.
- Uploaded documents and the database are git-ignored.

## License

[MIT](LICENSE)
