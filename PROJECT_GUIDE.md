# Charge Amendements — Project Guide

How the project works, end to end: what it does, how a Word file becomes database records, where the AI fits in, how to configure it, and how to test it.

---

## 1. What the project does

Parliamentary groups submit their proposed **amendments** (التعديلات المقترحة) to the finance bill as **Word documents (.docx)** written in Arabic. Each document contains several amendments, and each amendment has:

- a **header**: amendment number, group number, sometimes a type ("إضافة مادة جديدة", "حذف"), the law/code concerned, the article, the clause (البند), and the article of the code being modified (الفصل …);
- a **3-column table**: the original text (النص الأصلي / النص كما جاء في المشروع), the amended text (نص التعديل), and the justification (التعليل).

The app lets a logged-in user **upload** such a document. It then:

1. converts it to HTML (to display it as-is),
2. splits it into individual amendments,
3. **detects the fields** of each amendment with an **AI model** (checked by rules),
4. saves one row per amendment in the database,
5. lets the user view, edit, rename, delete, and export (PDF / CSV) the result.

---

## 2. Tech stack

| Layer | Technology |
|---|---|
| Web framework | **Django** (Python) |
| Database | **SQLite** (`db.sqlite3`, created locally, git-ignored) |
| Word → HTML / PDF | **LibreOffice** (headless, `soffice`) |
| HTML parsing | **BeautifulSoup** |
| AI | **Hugging Face `InferenceClient`** (default) or Groq / OpenRouter / Ollama |
| Config / secrets | **`.env`** file loaded with `python-dotenv` |

---

## 3. Project structure

```
stage/
├── manage.py                     # Django entry point
├── .env.example                  # template for your AI key (copy to .env)
├── requirements.txt
├── db.sqlite3                    # the database (created by migrate, git-ignored)
├── media/amendements_uploads/    # uploaded .docx + generated .html/.pdf (git-ignored)
├── scripts/generate_sample_docx.py  # creates a fictional sample document
├── test_huggingface.py           # quick script: is the AI answering?
├── HUGGINGFACE_INTEGRATION.md    # short AI setup notes
├── PROJECT_GUIDE.md              # this file
│
├── charge_amendements/           # Django project (settings, root URLs, home page)
│   ├── settings.py               # loads .env, SQLite, templates, static
│   ├── urls.py                   # /, /admin, /amendements, /users, /accounts, /stats
│   ├── templates/                # base.html (nav bar), home.html
│   └── static/                   # style.css, logo
│
├── amendements/                  # ★ main app
│   ├── models.py                 # Amendement, AmendementFichier
│   ├── views.py                  # upload, list, detail, edit, PDF, AI detection pages
│   ├── extraction.py             # ★ splits the document + rule-based detection + merge
│   ├── huggingface_utils.py      # ★ AI detector (InferenceClient / HTTP)
│   ├── admin.py                  # admin: columns, CSV export, "re-detect with AI"
│   ├── urls.py
│   ├── tests.py                  # 26 automated tests
│   ├── sample_docx.py            # fictional documents generated for the tests
│   ├── migrations/               # DB schema history (0001 → 0011)
│   └── templates/amendements/    # pages
│
├── users/                        # profile + change password
└── stats/                        # simple dashboard (number of files)
```

---

## 4. The data model

### `AmendementFichier` — one uploaded file
| Field | Meaning |
|---|---|
| `utilisateur` | owner (Django user) |
| `fichier` | the .docx, stored in `media/amendements_uploads/` |
| `date_upload` | upload date |
| `extracted_content` | the HTML version of the document (what you see on screen) |
| `nom_personnalise` | optional custom display name |

### `Amendement` — one amendment (many per file)
| Field | Meaning | Example |
|---|---|---|
| `numero_modification` | amendment number (التعديل رقم) | `4` |
| `numero_groupe` | parliamentary group (الفريق) | `1` |
| `type_modification` | type, if written | `إضافة مادة جديدة` |
| `nom_loi` | law / code / table concerned | `مدونة الجمارك والضرائب غير المباشرة` |
| `article` | article of the finance bill (1st المادة) | `3` |
| `alinea` | clause (البند) | `I` |
| `numero_article` | article of the code modified (الفصل or 2nd المادة) | `164 المكرر` |
| `texte_original` | column 1 of the table | … |
| `texte_modifie` | column 2 | … |
| `justification` | column 3 | … |
| `texte_amendement` | copy of the amended text | … |
| `texte_entete` | raw header text (input given to the AI) | `التعديل رقم : 4 الفريق 1 …` |
| `champs_detectes` | JSON: value + confidence + method for each field | see §6.4 |
| `fichier`, `utilisateur`, `date_insertion` | links and timestamp | |

---

## 5. What happens when you upload a file

```
 Browser                Django view                   Helpers
 ───────                ───────────                   ───────
 Upload .docx  ───▶  upload_amendement()
                       │ 1. checks it's a .docx
                       │ 2. saves AmendementFichier
                       │ 3. extract_html_content() ──▶ LibreOffice: .docx → .html
                       │    (stores HTML in extracted_content)
                       │ 4. creer_amendements_depuis_html()
                       │       │
                       │       ├─▶ extraire_blocs(html)          (extraction.py)
                       │       │     → list of amendments:
                       │       │       {entete, texte_original, texte_modifie, justification}
                       │       │
                       │       └─▶ for each block:
                       │             field_detector.detecter(entete)   (huggingface_utils.py)
                       │               ├─ AI model  → JSON fields
                       │               ├─ rules     → same fields
                       │               └─ merge     → final values + confidence
                       │             Amendement.objects.create(...)
                       │
                       │ 5. message: "14 amendement(s) détecté(s) — méthode : …"
                       ▼
 Validation page (shows the document as HTML)
```

### Step 3 — Word → HTML
`extract_html_content()` runs LibreOffice in headless mode. The HTML keeps the Arabic text, tables, and colors; font sizes are stripped so the page looks uniform. The LibreOffice path is found automatically (`C:\Program Files\LibreOffice\…`, or `SOFFICE_PATH` in `.env`, or the system PATH).

### Step 4a — Splitting the document (`extraction.extraire_blocs`)
The parser walks through the document **in order**, ignoring anything nested inside tables:

- every paragraph and every small table (the header table "التعديل رقم : 1 | الفريق 1 | …") is **added to a buffer**;
- when it meets a **content table** — one whose first row contains "نص التعديل" and "التعليل" — it closes the current amendment:
  - the buffer becomes the **header** (`entete`),
  - the table's columns become `texte_original`, `texte_modifie`, `justification` (nested tables, like customs tariff tables, are kept row by row, with tabs between cells),
  - the buffer is emptied for the next amendment.

So each amendment = *"everything written since the previous table"* + *its table*. This is what fixed the old bug where amendments 8–14 were saved as 1–7: the old code read the number from the wrong place (the justification text of the previous amendment).

---

## 6. How the AI works

### 6.1 The idea
The hard part is reading the **header** — a short Arabic line like:

```
التعديل رقم : 4 الفريق 1 إضافة مادة جديدة مدونة الجمارك والضرائب غير المباشرة المادة 3 البند I الفصل 164 المكرر
```

and knowing which number is the amendment, which is the article of the finance bill, which is the article of the code, where the law name starts and ends, etc. Documents vary from group to group, so a **language model (LLM)** is used: it reads the header like a human would and returns structured JSON.

> ⚠️ The project's earlier version used `dslim/bert-base-NER`, an **English** named-entity model that only finds people/places/organizations. It could not read Arabic legal text and its output was never used for the real fields. That approach (`InferenceApi`, huggingface_hub v0.13) was replaced.

### 6.2 What is sent to the model
`huggingface_utils.py` sends a **chat request**:

- a **system prompt** that explains the 7 fields, the rules to follow (e.g. *"article = the FIRST المادة after the law name"*, *"numero_article = value after الفصل, or the SECOND المادة"*, *"don't invent anything, use "" if absent"*), and **two worked examples**;
- a **user message** containing the header + the first 300 characters of the original text.

The model answers with JSON, for example:

```json
{"numero_modification":"4","numero_groupe":"1","type_modification":"إضافة مادة جديدة",
 "nom_loi":"مدونة الجمارك والضرائب غير المباشرة","article":"3","alinea":"I",
 "numero_article":"164 المكرر"}
```

The code tolerates answers wrapped in ```` ```json ```` fences or with extra text — it extracts the first `{…}` block.

Only the short header is sent (not the whole document), so each call is small, fast, and cheap. Identical headers are **cached** during a run, so they're only sent once.

### 6.3 Providers
All providers use the same "chat completion" format; you choose one in `.env`:

| `AI_PROVIDER` | How it's called | Cost | Default model |
|---|---|---|---|
| `huggingface` (default) | official **`huggingface_hub.InferenceClient(...).chat_completion(...)`**, `provider="auto"` | free monthly credits | `Qwen/Qwen2.5-72B-Instruct` |
| `groq` | HTTP to `api.groq.com/openai/v1/chat/completions` | generous free tier | `llama-3.3-70b-versatile` |
| `openrouter` | HTTP to `openrouter.ai/api/v1/chat/completions` | `:free` models | `meta-llama/llama-3.3-70b-instruct:free` |
| `ollama` | HTTP to `localhost:11434` | free, offline, runs on your PC | `qwen2.5:7b` |
| `none` | no AI, rules only | — | — |

Qwen is the default because it handles Arabic well. Any chat model can be set with `AI_MODEL`.

### 6.4 The safety net: rules + merge
The AI is never trusted alone. For every header, `extraction.detecter_par_regles()` also extracts the same 7 fields with regular expressions written for this document format (it handles "التعديل رقم : 8" with a colon, Eastern Arabic digits ١٢, diacritics like "مرّتين", "المكرر", law names without "مدونة", etc.).

Then `extraction.fusionner()` combines both, field by field:

| Situation | Value kept | Confidence | `method` |
|---|---|---|---|
| AI and rules agree | that value | 0.98 | `ia+regles` |
| They disagree | AI value (rule value kept in `regles` for review) | 0.80 | `ia` |
| Only the AI found it | AI value | 0.85 | `ia` |
| Only the rules found it | rule value | 0.70 | `regles` |

Numeric fields (number, group, article) are cleaned so that an AI answer like "رقم 7" becomes "7".

The result goes into the real columns **and** into `champs_detectes`, e.g.:

```json
{"numero_modification": {"value": "4", "confidence": 0.98, "method": "ia+regles"},
 "article": {"value": "3", "confidence": 0.98, "method": "ia+regles"}, ...}
```

### 6.5 When the AI is unavailable
No internet, quota exceeded, wrong key, invalid answer… → the error is caught, the **rules are used alone**, and after the upload the user sees:

> IA indisponible, détection par règles utilisée : HTTP 429 …

The upload never fails because of the AI.

### 6.6 Re-detecting
- Page `/amendements/amendements/<id>/champs-detectes/` shows the detected fields with confidence and method, and has a **Re-détecter** button.
- In the admin, select amendements → action **"Re-détecter les champs avec l'IA"**.

Both re-run the detection on the stored `texte_entete` and update the fields.

---

## 7. Pages and URLs

| URL | What it does |
|---|---|
| `/` | home page |
| `/accounts/login/`, `/accounts/logout/` | login / logout (Django auth) |
| `/amendements/upload/` | **upload a .docx** → extraction + AI detection |
| `/amendements/fichiers/` | list of my files (view, edit, rename, delete, download PDF) |
| `/amendements/fichier_mammoth/<id>/` | view the extracted HTML of a file |
| `/amendements/fichiers/<id>/modifier/` | edit the extracted HTML |
| `/amendements/fichiers/<id>/renommer/` | rename a file |
| `/amendements/fichiers/<id>/supprimer/` | delete a file |
| `/amendements/fichiers/supprimer_tous/` | delete all my files and their amendements |
| `/amendements/amendements/<id>/telecharger_pdf_libreoffice/` | PDF of the original Word file (LibreOffice) |
| `/amendements/amendements/<id>/champs-detectes/` | detected fields of one amendment |
| `/amendements/amendements/<id>/re-detecter-champs/` | re-run AI detection |
| `/amendements/fichiers/<id>/statistiques-detection/` | detection statistics for a file |
| `/stats/dashboard/` | number of files uploaded |
| `/users/profile/`, `/users/change-password/` | profile |
| `/admin/` | Django admin: all amendements, inline editing, **CSV export**, re-detect |

The file list links to each file's detection statistics (« Champs détectés »), from which every amendment's detected fields can be opened.

---

## 8. Setup

1. **Python packages**
   ```
   pip install -r requirements.txt
   ```
2. **LibreOffice** installed (default path is detected; otherwise add `SOFFICE_PATH=...` to `.env`).
3. **AI key**
   - Create a token at https://huggingface.co/settings/tokens → *Fine-grained* → tick only *Make calls to Inference Providers*.
   - Copy `.env.example` → `.env` and set `AI_API_KEY=hf_...`
4. **Database**
   ```
   python manage.py migrate
   ```
5. **Run**
   ```
   python manage.py runserver
   ```
   Open http://127.0.0.1:8000, log in (create a user with `python manage.py createsuperuser` if needed), and upload a file.

### `.env` options
```
DJANGO_SECRET_KEY=...          # required when DJANGO_DEBUG=False
DJANGO_DEBUG=True              # local development
DJANGO_ALLOWED_HOSTS=localhost,127.0.0.1
AI_PROVIDER=huggingface        # huggingface | groq | openrouter | ollama | none
AI_API_KEY=hf_...              # not needed for ollama
AI_MODEL=                      # optional, otherwise provider default
AI_HF_PROVIDER=auto            # optional, Hugging Face inference provider
AI_TIMEOUT=40                  # optional, seconds
SOFFICE_PATH=                  # optional, LibreOffice path
```
`.env` is in `.gitignore` — never commit it.

---

## 9. Testing

```
python manage.py test                 # 26 automated tests
python test_huggingface.py            # real AI call with your key
```

To include the real-AI test in the suite:
```
$env:AI_LIVE_TEST="1"        # PowerShell  (bash: export AI_LIVE_TEST=1)
python manage.py test
```

What the tests cover:

| Group | Checks |
|---|---|
| `ReglesTests` | rule detection on typical headers, Arabic digits, diacritics, the old numbering bug |
| `FusionTests` | AI/rules merging, confidence, cleaning |
| `DetecteurIATests` | Hugging Face goes through `InferenceClient` with your key; Groq via HTTP; fallback when the AI is unreachable or answers nonsense; no call without a key |
| `ExemplesTests` | full extraction of fictional Word documents (generated by `sample_docx.py`), field by field, and the 3 text columns compared with the Word file read by python-docx |
| `UploadTests` | real upload through the Django view (AI simulated), checks what is saved in the database; detection and statistics pages |
| `IAReelleTests` | real call to your provider (only with `AI_LIVE_TEST=1`) |

| `SecuriteTests` | users can't access each other's files, deletion is POST-only, edited HTML is sanitised, login required |

The tests never use real documents: the fictional generator covers the same difficult cases (nested tables, interleaved text and tables, diacritics, Eastern Arabic digits, law names without « مدونة »).

---

## 10. Limitations and ideas

- Documents must follow the same general layout (header, then a 3-column table with "نص التعديل" and "التعليل"). A very different layout would need the parser adjusted.
- Hugging Face free credits are limited; for heavy use, switch to Groq (free key) or Ollama (local).
- `type_modification` stays empty when the document doesn't state it.
- Possible improvements: show fields where AI and rules disagree for manual review, export amendements to Excel, background processing for large files.
