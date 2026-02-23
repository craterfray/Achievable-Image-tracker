# Medical Image Citation Tracker (Streamlit)

A non-developer-friendly local web app that:
- Reads **Complete** subchapters from your outline Google Sheet.
- Downloads each subchapter `.docx` from Google Drive.
- Extracts text, builds deterministic keyword queries, and suggests 3–5 copyright-safe images.
- Supports **Wikimedia Commons** and **OpenStax** (priority sources).
- Writes approved rows into your existing tracker sheet tab **Image citations**, replacing older rows for the same subchapter doc link.

---

## Included deliverables
- `app.py`
- `requirements.txt`
- `README.md`

---

## Default IDs (pre-filled in UI)
- Outline sheet ID: `1UscA6uH_Yh4bT83PPocGplG6zkrsK321BZnk0vq9ejQ`
- Tracker sheet ID: `1kP_7sQk64HKv_7f2N3g0KGfeKJdfFLSSb0mt0-xEx7Q`
- Tracker tab name: `Image citations`

You can override all of them in the sidebar.

---

## Features implemented

### 1) Google OAuth authorization
- Uses `google-auth-oauthlib` Installed App flow.
- Scopes:
  - Drive readonly
  - Sheets read/write

### 2) Outline loading and filtering
- Reads outline tab (`A1:ZZ`) and auto-detects columns for:
  - Status
  - Subchapter/heading/title
  - Link/doc/url
- Processes only rows where:
  - `Status == Complete`
  - doc link is not blank

### 3) DOCX download + text extraction
- Parses Drive file ID from link.
- Downloads bytes via Drive API (`files.get_media`).
- Extracts text with `python-docx`.
- On download/parse failure, displays clear error and skips.

### 4) Deterministic keyword extraction and queries
- Cleans text and tokenizes.
- Removes stopwords.
- Builds unigrams + bigrams + trigrams.
- Ranks phrases using a deterministic TF-style score.
- Generates 3–8 queries per subchapter.

### 5) Image retrieval sources

#### Wikimedia Commons
- Uses MediaWiki API.
- Pulls:
  - file page URL
  - direct image URL
  - metadata including usage/license fields
- Filters license by policy (below).

#### OpenStax
- Uses curated base book list.
- Extracts `<figure>` images and captions from pages.
- Accepts only CC BY (configured as CC BY 4.0).

### 6) License policy enforcement
Allowed:
- Public Domain
- CC0
- CC BY (any version)

Excluded:
- CC BY-SA
- Any `NC`
- Any `ND`
- All Rights Reserved
- unknown/unclear licenses

Behavior:
- If confidently verified: status `Suggested - needs upload`
- If uncertain: status `Needs license check` (and skipped during write unless explicitly changed)

### 7) Human review UI
- Select subchapters with checkboxes.
- Generate suggestions.
- Review grouped by subchapter with:
  - preview image
  - metadata
  - per-suggestion approve checkbox
  - editable alt text / caption
  - disturbing-image flag warning
  - custom query “Search again”
- 3–5 approved target warning per subchapter.

### 8) Write-back to tracker sheet
Column mapping:
- A Asset URL (source image URL placeholder)
- B Alt text
- C Image caption
- D License name
- E License link
- F Source name
- G Source Url
- H Image code formula (written as `userEnteredValue.formulaValue`)
- I License verified
- J Status
- K Textbook/quiz/exam location (`<subchapter> | <doc link>`)
- L Subchapter name

Replace behavior:
- Finds existing rows where column K contains exact doc link substring.
- Deletes matching rows first.
- Appends approved rows.

Safety controls:
- **Dry run** mode (default ON) shows would-delete/would-insert counts.
- **Backup export CSV** button before write.

---

## Setup

## 1) Create a Google Cloud project
1. Open Google Cloud Console.
2. Create/select a project.
3. Enable APIs:
   - Google Drive API
   - Google Sheets API

## 2) Configure OAuth consent screen
1. Configure OAuth consent screen as External or Internal.
2. Add yourself as test user if required.

## 3) Create OAuth client credentials
1. Go to Credentials → Create Credentials → OAuth client ID.
2. Choose **Desktop app**.
3. Download JSON.
4. Save it in this project root as `credentials.json`.

## 4) Install and run
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

---

## Typical usage
1. Start app.
2. Click **Authorize with Google**.
3. Click **Load outline**.
4. Select Complete subchapters.
5. Click **Generate suggestions**.
6. Review/edit approvals.
7. Use **Backup export CSV**.
8. Keep **Dry run** ON to verify replacement counts.
9. Turn Dry run OFF and write approved rows.

---

## Troubleshooting

### OAuth fails / browser callback issues
- Ensure `credentials.json` is Desktop OAuth credentials.
- Ensure Drive + Sheets APIs are enabled.
- Ensure your account has access to both spreadsheets and docx files.

### “Could not parse Drive file id”
- Ensure link is a valid Google Drive URL containing `/d/<id>/` or `?id=<id>`.

### No suggestions returned
- Some subchapters may be too specific.
- Try **Search again** with custom query.
- Check network access for external source APIs.

### Many results show “Needs license check”
- Metadata can be incomplete for some media.
- App intentionally skips uncertain licenses by default during writes.

### OpenStax extraction sparse
- OpenStax page structures vary.
- Wikimedia is still used in parallel and often provides broader coverage.

---

## Notes on column H formula
- The app writes column H using Sheets API `userEnteredValue.formulaValue`, row-by-row.
- This preserves formula behavior and avoids overwriting H with plain text.

