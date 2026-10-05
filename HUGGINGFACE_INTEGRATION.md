# Détection des champs par IA

À l'upload d'un fichier Word, chaque amendement est découpé puis ses champs
(`numero_modification`, `numero_groupe`, `type_modification`, `nom_loi`,
`article`, `alinea`, `numero_article`) sont détectés par **un modèle de langage
gratuit**, avec une **détection par règles** en secours. Les valeurs trouvées
sont enregistrées directement dans les colonnes de la table `Amendement`.

## Fonctionnement

1. `extraction.extraire_blocs()` parcourt le HTML (LibreOffice) dans l'ordre :
   tout le texte situé avant un tableau « النص الأصلي / نص التعديل / التعليل »
   forme l'**en-tête** de l'amendement ; les 3 colonnes du tableau donnent
   `texte_original`, `texte_modifie`, `justification`.
2. `field_detector.detecter(entete)` :
   - envoie l'en-tête au modèle IA, qui renvoie un JSON avec les 7 champs ;
   - applique en parallèle les règles (`extraction.detecter_par_regles`) ;
   - fusionne : si IA et règles sont d'accord → confiance 0.98 (`ia+regles`),
     sinon la valeur IA est gardée (la valeur des règles reste visible).
   - si l'IA est injoignable (pas d'internet, quota, mauvaise clé), les règles
     seules sont utilisées et un message l'indique après l'upload.
3. Le détail (valeur, confiance, méthode) est stocké dans `champs_detectes`
   et visible sur `/amendements/<id>/champs-detectes/`.

## Configuration (fichier `.env`)

Copier `.env.example` en `.env` à la racine du projet :

```
AI_PROVIDER=huggingface
AI_API_KEY=hf_votre_jeton
AI_MODEL=            # optionnel
```

| Fournisseur  | Gratuit ? | Clé | Modèle par défaut |
|--------------|-----------|-----|-------------------|
| `huggingface`| crédits mensuels gratuits | https://huggingface.co/settings/tokens | `Qwen/Qwen2.5-72B-Instruct` |
| `groq`       | offre gratuite généreuse | https://console.groq.com/keys | `llama-3.3-70b-versatile` |
| `openrouter` | modèles `:free` | https://openrouter.ai/keys | `meta-llama/llama-3.3-70b-instruct:free` |
| `ollama`     | 100 % gratuit, local, hors-ligne | aucune | `qwen2.5:7b` |
| `none`       | — | — | règles uniquement |

Ollama : installer https://ollama.com, puis `ollama pull qwen2.5:7b`, et mettre `AI_PROVIDER=ollama`.

## Tester

```
pip install -r requirements.txt
python manage.py migrate
python manage.py test                  # 26 tests : règles, IA simulée, documents fictifs, upload, sécurité
python test_huggingface.py             # appel réel à l'IA avec ta clé du .env
```

Pour inclure le test IA réel dans la suite : `$env:AI_LIVE_TEST="1"` (PowerShell) ou
`export AI_LIVE_TEST=1` (bash), puis `python manage.py test`.

## Sécurité

La clé n'est plus dans le code : elle est lue depuis `.env`, qui est ignoré par git
(`.gitignore`). Ne jamais committer `.env`.
