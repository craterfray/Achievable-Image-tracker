import csv
import io
import math
import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, unquote, urlparse

import requests
import streamlit as st
from docx import Document
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# -----------------------------
# Defaults and constants
# -----------------------------
DEFAULT_OUTLINE_SHEET_ID = "1UscA6uH_Yh4bT83PPocGplG6zkrsK321BZnk0vq9ejQ"
DEFAULT_TRACKER_SHEET_ID = "1kP_7sQk64HKv_7f2N3g0KGfeKJdfFLSSb0mt0-xEx7Q"
DEFAULT_TRACKER_TAB = "Image citations"

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

STOPWORDS = {
    "the",
    "and",
    "for",
    "that",
    "with",
    "this",
    "from",
    "are",
    "was",
    "were",
    "been",
    "have",
    "has",
    "had",
    "into",
    "about",
    "through",
    "between",
    "over",
    "under",
    "after",
    "before",
    "because",
    "while",
    "where",
    "when",
    "which",
    "their",
    "there",
    "also",
    "they",
    "them",
    "his",
    "her",
    "she",
    "him",
    "its",
    "you",
    "your",
    "our",
    "can",
    "may",
    "might",
    "will",
    "would",
    "should",
    "must",
    "not",
    "than",
    "such",
    "using",
    "used",
    "use",
    "within",
    "without",
    "during",
    "patient",
    "patients",
    "nurse",
    "nursing",
}

OPENSTAX_BOOKS = [
    {
        "name": "Anatomy and Physiology 2e",
        "base_url": "https://openstax.org/books/anatomy-and-physiology-2e/pages",
        "license_name": "CC BY 4.0",
        "license_link": "https://creativecommons.org/licenses/by/4.0/",
    },
    {
        "name": "Biology 2e",
        "base_url": "https://openstax.org/books/biology-2e/pages",
        "license_name": "CC BY 4.0",
        "license_link": "https://creativecommons.org/licenses/by/4.0/",
    },
    {
        "name": "Microbiology",
        "base_url": "https://openstax.org/books/microbiology/pages",
        "license_name": "CC BY 4.0",
        "license_link": "https://creativecommons.org/licenses/by/4.0/",
    },
]

WIKIMEDIA_API = "https://commons.wikimedia.org/w/api.php"


@dataclass
class Subchapter:
    title: str
    status: str
    link: str
    row_index: int


@dataclass
class ImageSuggestion:
    subchapter_title: str
    doc_link: str
    query: str
    asset_url: str
    alt_text: str
    caption: str
    license_name: str
    license_link: str
    source_name: str
    source_url: str
    status: str
    license_verified: bool
    source_priority: int
    score: float
    disturbing_flag: bool = False
    notes: str = ""
    approved: bool = True


# -----------------------------
# Auth and Google clients
# -----------------------------
def get_google_credentials(credentials_json_path: str) -> Credentials:
    if "google_credentials" not in st.session_state:
        st.session_state.google_credentials = None

    creds: Optional[Credentials] = st.session_state.google_credentials
    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        st.session_state.google_credentials = creds
        return creds

    flow = InstalledAppFlow.from_client_secrets_file(credentials_json_path, SCOPES)
    creds = flow.run_local_server(port=0)
    st.session_state.google_credentials = creds
    return creds


def get_google_clients(creds: Credentials):
    sheets = build("sheets", "v4", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return sheets, drive


# -----------------------------
# Helpers
# -----------------------------
def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


def extract_drive_file_id(url: str) -> Optional[str]:
    if not url:
        return None
    m = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if m:
        return m.group(1)

    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    if "id" in query and query["id"]:
        return query["id"][0]

    if "open?id=" in url:
        return url.split("open?id=")[-1].split("&")[0]

    return None


def tokenize(text: str) -> List[str]:
    text = text.lower()
    tokens = re.findall(r"[a-z][a-z0-9\-]{2,}", text)
    return [t for t in tokens if t not in STOPWORDS]


def ngrams(tokens: List[str], n: int) -> List[str]:
    if len(tokens) < n:
        return []
    return [" ".join(tokens[i : i + n]) for i in range(len(tokens) - n + 1)]


def tfidf_keywords(text: str, top_n: int = 15) -> List[str]:
    tokens = tokenize(text)
    if not tokens:
        return []

    token_counts = Counter(tokens)
    max_tf = max(token_counts.values())

    bigrams = Counter(ngrams(tokens, 2))
    trigrams = Counter(ngrams(tokens, 3))

    scores: Dict[str, float] = {}

    for tok, count in token_counts.items():
        tf = count / max_tf
        length_boost = 1.1 if len(tok) > 6 else 1.0
        scores[tok] = tf * length_boost

    for bg, count in bigrams.items():
        if any(word in STOPWORDS for word in bg.split()):
            continue
        scores[bg] = (count / max_tf) * 1.6

    for tg, count in trigrams.items():
        if any(word in STOPWORDS for word in tg.split()):
            continue
        scores[tg] = (count / max_tf) * 2.0

    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    return [k for k, _ in ranked[:top_n]]


def generate_queries(subchapter_title: str, text: str, min_q: int = 3, max_q: int = 8) -> List[str]:
    kw = tfidf_keywords(text, top_n=20)
    queries = []

    title_clean = normalize_space(re.sub(r"^[\d\.]+\s*", "", subchapter_title))
    if title_clean:
        queries.append(title_clean)

    for phrase in kw:
        if len(queries) >= max_q:
            break
        if phrase in queries:
            continue
        queries.append(phrase)

    if len(queries) < min_q and title_clean:
        fallback = [f"{title_clean} diagram", f"{title_clean} anatomy", f"{title_clean} pathology"]
        for q in fallback:
            if q not in queries:
                queries.append(q)
            if len(queries) >= min_q:
                break

    return queries[:max_q]


def is_disturbing(text: str) -> bool:
    text_l = text.lower()
    markers = ["blood", "wound", "trauma", "surgery", "necrosis", "ulcer", "gangrene", "injury"]
    return any(m in text_l for m in markers)


# -----------------------------
# Outline reading and doc extraction
# -----------------------------
def detect_outline_columns(header: List[str]) -> Tuple[int, int, int]:
    lower = [normalize_space(h).lower() for h in header]

    status_idx = next((i for i, h in enumerate(lower) if "status" in h), -1)
    title_idx = next(
        (
            i
            for i, h in enumerate(lower)
            if "subchapter" in h or "heading" in h or "title" in h or "section" in h
        ),
        -1,
    )
    link_idx = next((i for i, h in enumerate(lower) if "link" in h or "doc" in h or "url" in h), -1)

    if status_idx == -1 or title_idx == -1 or link_idx == -1:
        raise ValueError(
            "Could not auto-detect Status / Subchapter / Link columns in outline sheet header. "
            "Please ensure those words exist in the header row."
        )

    return status_idx, title_idx, link_idx


def read_outline_complete_subchapters(sheets_service, outline_sheet_id: str, tab_name: str) -> List[Subchapter]:
    data = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=outline_sheet_id, range=f"{tab_name}!A1:ZZ")
        .execute()
    )
    rows = data.get("values", [])
    if not rows:
        return []

    header = rows[0]
    status_idx, title_idx, link_idx = detect_outline_columns(header)

    out: List[Subchapter] = []
    for i, row in enumerate(rows[1:], start=2):
        status = normalize_space(row[status_idx] if len(row) > status_idx else "")
        title = normalize_space(row[title_idx] if len(row) > title_idx else "")
        link = normalize_space(row[link_idx] if len(row) > link_idx else "")

        if status.lower() == "complete" and link:
            out.append(Subchapter(title=title, status=status, link=link, row_index=i))

    return out


def download_docx_bytes(drive_service, doc_link: str) -> bytes:
    file_id = extract_drive_file_id(doc_link)
    if not file_id:
        raise ValueError(f"Could not parse Drive file id from link: {doc_link}")

    request = drive_service.files().get_media(fileId=file_id)
    return request.execute()


def extract_docx_text(docx_bytes: bytes) -> str:
    bio = io.BytesIO(docx_bytes)
    doc = Document(bio)
    chunks = []
    for p in doc.paragraphs:
        t = normalize_space(p.text)
        if t:
            chunks.append(t)
    return "\n".join(chunks)


# -----------------------------
# Licensing filters
# -----------------------------
def normalize_license(license_name: str, license_url: str = "") -> Tuple[Optional[str], Optional[str], str]:
    name = normalize_space(license_name)
    name_l = name.lower()
    url_l = normalize_space(license_url).lower()

    combined = f"{name_l} {url_l}"

    if "public domain" in combined:
        return "Public Domain", "https://creativecommons.org/publicdomain/mark/1.0/", "ok"
    if "cc0" in combined:
        return "CC0", "https://creativecommons.org/publicdomain/zero/1.0/", "ok"

    if "cc by" in combined and "by-sa" not in combined and "nc" not in combined and "nd" not in combined:
        # attempt to preserve version when present
        version_match = re.search(r"cc\s*by\s*([0-9]\.[0-9])", combined)
        if version_match:
            version = version_match.group(1)
            return f"CC BY {version}", f"https://creativecommons.org/licenses/by/{version}/", "ok"
        return "CC BY 4.0", "https://creativecommons.org/licenses/by/4.0/", "ok"

    if any(x in combined for x in ["by-sa", "nc", "nd", "all rights reserved", "copyright"]):
        return None, None, "disallowed"

    return None, None, "unknown"


# -----------------------------
# Wikimedia retrieval
# -----------------------------
def wikimedia_search(query: str, limit: int = 20) -> List[dict]:
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": 6,
        "gsrlimit": limit,
        "prop": "imageinfo|info",
        "iiprop": "url|extmetadata",
        "inprop": "url",
    }
    resp = requests.get(WIKIMEDIA_API, params=params, timeout=25)
    resp.raise_for_status()
    data = resp.json()
    pages = data.get("query", {}).get("pages", {})
    return list(pages.values())


def parse_wikimedia_candidate(page: dict, subchapter_title: str, doc_link: str, query: str) -> Optional[ImageSuggestion]:
    title = page.get("title", "")
    imageinfo = (page.get("imageinfo") or [{}])[0]
    ext = imageinfo.get("extmetadata", {})

    raw_license_name = ext.get("LicenseShortName", {}).get("value") or ext.get("UsageTerms", {}).get("value") or ""
    raw_license_url = ext.get("LicenseUrl", {}).get("value") or ""

    license_name, license_link, decision = normalize_license(raw_license_name, raw_license_url)
    if decision == "disallowed":
        return None

    file_url = imageinfo.get("url")
    file_page = page.get("fullurl")
    if not file_url or not file_page:
        return None

    caption = re.sub("<[^<]+?>", "", ext.get("ImageDescription", {}).get("value", "")).strip()
    if not caption:
        caption = title.replace("File:", "").replace("_", " ")

    status = "Suggested - needs upload"
    verified = decision == "ok"
    if decision != "ok":
        status = "Needs license check"

    text_blob = f"{title} {caption} {query}".lower()
    overlap = len(set(tokenize(query)) & set(tokenize(text_blob)))
    score = 1.5 + overlap

    return ImageSuggestion(
        subchapter_title=subchapter_title,
        doc_link=doc_link,
        query=query,
        asset_url=file_url,
        alt_text=caption[:300],
        caption=caption[:300],
        license_name=license_name or "",
        license_link=license_link or "",
        source_name="Wikimedia Commons",
        source_url=file_page,
        status=status,
        license_verified=verified,
        source_priority=2,
        score=score,
        disturbing_flag=is_disturbing(caption),
        notes="",
    )


# -----------------------------
# OpenStax retrieval
# -----------------------------
def fetch_openstax_figures(query: str, limit_per_book: int = 8) -> List[dict]:
    found = []
    query_tokens = set(tokenize(query))

    for book in OPENSTAX_BOOKS:
        # openstax pages are structured by numbered slugs; using the search endpoint for better match
        search_url = f"{book['base_url']}/1-introduction"
        try:
            resp = requests.get(search_url, timeout=20)
            if resp.status_code != 200:
                continue
        except requests.RequestException:
            continue

        # lightweight heuristic: gather linked chapter pages and inspect figure tags
        links = re.findall(r'href="(/books/[^"#?]+/pages/[^"#?]+)"', resp.text)
        unique_links = []
        for lk in links:
            full = f"https://openstax.org{lk}"
            if full not in unique_links:
                unique_links.append(full)
            if len(unique_links) >= 25:
                break

        for page_url in unique_links:
            try:
                page_resp = requests.get(page_url, timeout=20)
                if page_resp.status_code != 200:
                    continue
            except requests.RequestException:
                continue

            html = page_resp.text
            figures = re.findall(
                r"<figure[^>]*>(.*?)</figure>", html, flags=re.IGNORECASE | re.DOTALL
            )

            for fig in figures:
                img_match = re.search(r'<img[^>]+src="([^"]+)"[^>]*alt="([^"]*)"', fig, flags=re.IGNORECASE)
                cap_match = re.search(r"<figcaption[^>]*>(.*?)</figcaption>", fig, flags=re.IGNORECASE | re.DOTALL)
                if not img_match:
                    continue
                img_url = img_match.group(1)
                if img_url.startswith("//"):
                    img_url = "https:" + img_url
                elif img_url.startswith("/"):
                    img_url = "https://openstax.org" + img_url

                alt = re.sub("<[^<]+?>", "", img_match.group(2) or "").strip()
                cap = re.sub("<[^<]+?>", "", cap_match.group(1) if cap_match else "").strip()
                combined = f"{alt} {cap}".lower()
                overlap = len(query_tokens & set(tokenize(combined)))
                if overlap <= 0:
                    continue

                found.append(
                    {
                        "book": book,
                        "image_url": img_url,
                        "alt": alt,
                        "caption": cap or alt or "OpenStax figure",
                        "page_url": page_url,
                        "overlap": overlap,
                    }
                )

                if len(found) >= limit_per_book:
                    break

    return found


def parse_openstax_candidate(raw: dict, subchapter_title: str, doc_link: str, query: str) -> ImageSuggestion:
    book = raw["book"]
    score = 2.2 + raw.get("overlap", 0)
    caption = normalize_space(raw.get("caption") or raw.get("alt") or "OpenStax figure")

    return ImageSuggestion(
        subchapter_title=subchapter_title,
        doc_link=doc_link,
        query=query,
        asset_url=raw["image_url"],
        alt_text=(raw.get("alt") or caption)[:300],
        caption=caption[:300],
        license_name=book["license_name"],
        license_link=book["license_link"],
        source_name=f"{book['name']} (OpenStax)",
        source_url=raw["page_url"],
        status="Suggested - needs upload",
        license_verified=True,
        source_priority=3,
        score=score,
        disturbing_flag=is_disturbing(caption),
        notes="",
    )


# -----------------------------
# Suggestion engine
# -----------------------------
def dedupe_suggestions(suggestions: List[ImageSuggestion]) -> List[ImageSuggestion]:
    seen = set()
    out = []
    for s in sorted(suggestions, key=lambda x: x.score, reverse=True):
        key = (s.asset_url, s.source_url)
        if key in seen:
            continue
        seen.add(key)
        out.append(s)
    return out


def rank_and_trim(suggestions: List[ImageSuggestion], target_max: int = 5) -> List[ImageSuggestion]:
    ranked = sorted(
        suggestions,
        key=lambda s: (s.score, s.source_priority, s.license_verified),
        reverse=True,
    )
    return ranked[:target_max]


def suggest_images_for_subchapter(subchapter: Subchapter, text: str, include_owid: bool = False) -> List[ImageSuggestion]:
    _ = include_owid  # reserved for future optional source integration
    queries = generate_queries(subchapter.title, text)
    all_candidates: List[ImageSuggestion] = []

    for q in queries:
        # Wikimedia
        try:
            pages = wikimedia_search(q, limit=20)
            for page in pages:
                cand = parse_wikimedia_candidate(page, subchapter.title, subchapter.link, q)
                if cand is not None:
                    all_candidates.append(cand)
        except requests.RequestException:
            pass

        # OpenStax
        openstax_raw = fetch_openstax_figures(q, limit_per_book=6)
        for raw in openstax_raw:
            all_candidates.append(parse_openstax_candidate(raw, subchapter.title, subchapter.link, q))

    deduped = dedupe_suggestions(all_candidates)
    return rank_and_trim(deduped, target_max=5)


# -----------------------------
# Sheets write-back logic
# -----------------------------
def get_tab_gid(sheets_service, spreadsheet_id: str, tab_name: str) -> int:
    ss = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    for sh in ss.get("sheets", []):
        props = sh.get("properties", {})
        if props.get("title") == tab_name:
            return props.get("sheetId")
    raise ValueError(f"Tab not found: {tab_name}")


def read_tracker_rows(sheets_service, tracker_sheet_id: str, tracker_tab_name: str) -> List[List[str]]:
    resp = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=tracker_sheet_id, range=f"{tracker_tab_name}!A1:L")
        .execute()
    )
    return resp.get("values", [])


def build_backup_csv(rows: List[List[str]]) -> str:
    buffer = io.StringIO()
    w = csv.writer(buffer)
    for r in rows:
        w.writerow(r)
    return buffer.getvalue()


def find_rows_for_doc_link(rows: List[List[str]], doc_link: str) -> List[int]:
    matches = []
    for idx, row in enumerate(rows[1:], start=2):
        k_val = row[10] if len(row) > 10 else ""
        if doc_link and doc_link in k_val:
            matches.append(idx)
    return matches


def ensure_column_l_exists(sheets_service, tracker_sheet_id: str, tracker_tab_name: str):
    rows = read_tracker_rows(sheets_service, tracker_sheet_id, tracker_tab_name)
    if not rows:
        header = [
            "Asset URL",
            "Alt text",
            "Image caption",
            "License name",
            "License link",
            "Source name",
            "Source Url",
            "Image code",
            "License verified",
            "Status",
            "Textbook/quiz/exam location",
            "Subchapter name",
        ]
        sheets_service.spreadsheets().values().update(
            spreadsheetId=tracker_sheet_id,
            range=f"{tracker_tab_name}!A1:L1",
            valueInputOption="RAW",
            body={"values": [header]},
        ).execute()
        return

    header = rows[0]
    if len(header) < 12 or normalize_space(header[11]).lower() != "subchapter name":
        sheets_service.spreadsheets().values().update(
            spreadsheetId=tracker_sheet_id,
            range=f"{tracker_tab_name}!L1",
            valueInputOption="RAW",
            body={"values": [["Subchapter name"]]},
        ).execute()


def formula_for_row(row_num: int) -> str:
    return (
        f'=IF(A{row_num}="","",CONCAT("<Aaa.Img src=\"",A{row_num},"\" alt=\"",B{row_num},'
        f'"\" caption=\"",C{row_num},"\" licenseName=\"",D{row_num},"\" licenseLink=\"",'
        f'E{row_num},"\" sourceName=\"",F{row_num},"\" sourceUrl=\"",G{row_num},"\" />"))'
    )


def write_suggestions_for_subchapter(
    sheets_service,
    tracker_sheet_id: str,
    tracker_tab_name: str,
    doc_link: str,
    approved_suggestions: List[ImageSuggestion],
    dry_run: bool,
) -> Dict[str, int]:
    ensure_column_l_exists(sheets_service, tracker_sheet_id, tracker_tab_name)

    rows = read_tracker_rows(sheets_service, tracker_sheet_id, tracker_tab_name)
    existing_matches = find_rows_for_doc_link(rows, doc_link)

    if dry_run:
        return {"would_delete": len(existing_matches), "would_insert": len(approved_suggestions)}

    gid = get_tab_gid(sheets_service, tracker_sheet_id, tracker_tab_name)

    # Delete existing matching rows from bottom to top so indices stay valid
    requests_payload = []
    for row_num in sorted(existing_matches, reverse=True):
        requests_payload.append(
            {
                "deleteDimension": {
                    "range": {
                        "sheetId": gid,
                        "dimension": "ROWS",
                        "startIndex": row_num - 1,
                        "endIndex": row_num,
                    }
                }
            }
        )

    if requests_payload:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=tracker_sheet_id,
            body={"requests": requests_payload},
        ).execute()

    # Determine append start row
    rows_after_delete = read_tracker_rows(sheets_service, tracker_sheet_id, tracker_tab_name)
    start_row = len(rows_after_delete) + 1

    values = []
    for s in approved_suggestions:
        values.append(
            [
                s.asset_url,
                s.alt_text,
                s.caption,
                s.license_name,
                s.license_link,
                s.source_name,
                s.source_url,
                "",  # H formula written separately
                "TRUE" if s.license_verified else "",
                s.status,
                f"{s.subchapter_title} | {s.doc_link}",
                s.subchapter_title,
            ]
        )

    if values:
        sheets_service.spreadsheets().values().append(
            spreadsheetId=tracker_sheet_id,
            range=f"{tracker_tab_name}!A:L",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute()

        # apply formulas in column H using userEnteredValue.formulaValue
        formula_reqs = []
        for i, _ in enumerate(values):
            row_num = start_row + i
            formula_reqs.append(
                {
                    "updateCells": {
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {
                                            "formulaValue": formula_for_row(row_num)
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue",
                        "range": {
                            "sheetId": gid,
                            "startRowIndex": row_num - 1,
                            "endRowIndex": row_num,
                            "startColumnIndex": 7,
                            "endColumnIndex": 8,
                        },
                    }
                }
            )

        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=tracker_sheet_id,
            body={"requests": formula_reqs},
        ).execute()

    return {"deleted": len(existing_matches), "inserted": len(values)}


# -----------------------------
# Streamlit UI
# -----------------------------
def init_state():
    st.session_state.setdefault("outline_rows", [])
    st.session_state.setdefault("suggestions", {})
    st.session_state.setdefault("errors", [])


def render():
    st.set_page_config(page_title="Medical Image Citation Tracker", layout="wide")
    init_state()

    st.title("Medical Mini-Textbook Image Suggestion Tracker")
    st.caption("Non-developer workflow for collecting copyright-safe image citations.")

    with st.sidebar:
        st.header("Configuration")
        credentials_path = st.text_input("credentials.json path", value="credentials.json")
        outline_sheet_id = st.text_input("Outline sheet ID", value=DEFAULT_OUTLINE_SHEET_ID)
        outline_tab_name = st.text_input("Outline tab name", value="Sheet1")
        tracker_sheet_id = st.text_input("Tracker sheet ID", value=DEFAULT_TRACKER_SHEET_ID)
        tracker_tab_name = st.text_input("Tracker tab name", value=DEFAULT_TRACKER_TAB)
        dry_run = st.checkbox("Dry run (no writes/deletes)", value=True)
        include_owid = st.checkbox("Include Our World in Data (optional/off by default)", value=False)

    st.subheader("1) Authorize Google")
    if st.button("Authorize with Google"):
        try:
            creds = get_google_credentials(credentials_path)
            sheets, drive = get_google_clients(creds)
            st.session_state.sheets = sheets
            st.session_state.drive = drive
            st.success("Authorization successful.")
        except Exception as e:
            st.error(f"Authorization failed: {e}")

    if "sheets" not in st.session_state or "drive" not in st.session_state:
        st.info("Authorize first to continue.")
        return

    sheets = st.session_state.sheets
    drive = st.session_state.drive

    st.subheader("2) Load complete subchapters from outline")
    if st.button("Load outline"):
        try:
            rows = read_outline_complete_subchapters(sheets, outline_sheet_id, outline_tab_name)
            st.session_state.outline_rows = rows
            st.success(f"Loaded {len(rows)} complete subchapters with doc links.")
        except Exception as e:
            st.error(f"Failed to load outline: {e}")

    outline_rows: List[Subchapter] = st.session_state.outline_rows
    if not outline_rows:
        return

    selected_titles = []
    st.write("Select subchapters to process:")
    for row in outline_rows:
        if st.checkbox(f"{row.title}", value=False, key=f"sub_{row.row_index}"):
            selected_titles.append(row.title)

    selected_rows = [r for r in outline_rows if r.title in selected_titles]

    st.subheader("3) Generate suggestions")
    if st.button("Generate suggestions"):
        if not selected_rows:
            st.warning("Select at least one subchapter.")
        else:
            for sub in selected_rows:
                try:
                    bytes_data = download_docx_bytes(drive, sub.link)
                    text = extract_docx_text(bytes_data)
                    suggestions = suggest_images_for_subchapter(sub, text, include_owid=include_owid)
                    st.session_state.suggestions[sub.title] = suggestions
                except Exception as e:
                    st.session_state.errors.append(f"{sub.title}: {e}")
            st.success("Suggestion generation complete.")

    if st.session_state.errors:
        st.error("\n".join(st.session_state.errors))

    st.subheader("4) Review and approve")
    all_suggestions: Dict[str, List[ImageSuggestion]] = st.session_state.suggestions
    if not all_suggestions:
        return

    for sub in selected_rows:
        sub_suggestions = all_suggestions.get(sub.title, [])
        if not sub_suggestions:
            continue

        st.markdown(f"### {sub.title}")
        st.caption(sub.link)

        approved_count = 0
        for i, s in enumerate(sub_suggestions):
            with st.container(border=True):
                cols = st.columns([1, 2])
                with cols[0]:
                    st.image(s.asset_url, use_container_width=True)
                with cols[1]:
                    s.approved = st.checkbox(
                        f"Approve suggestion {i+1}",
                        value=s.approved,
                        key=f"approve_{sub.row_index}_{i}",
                    )
                    if s.approved:
                        approved_count += 1
                    s.alt_text = st.text_input("Alt text", value=s.alt_text, key=f"alt_{sub.row_index}_{i}")
                    s.caption = st.text_input("Caption", value=s.caption, key=f"cap_{sub.row_index}_{i}")
                    st.write(f"License: **{s.license_name or 'Unknown'}**")
                    st.write(f"Source: {s.source_name}")
                    st.write(f"Source URL: {s.source_url}")
                    st.write(f"Status: {s.status}")
                    if s.disturbing_flag:
                        st.warning("Potentially disturbing medical image. Review carefully.")

        if approved_count < 3 or approved_count > 5:
            st.warning(f"Approved images for this subchapter: {approved_count}. Target is 3–5.")

        custom_query = st.text_input(
            "Search again with custom query",
            value="",
            key=f"custom_query_{sub.row_index}",
            placeholder="Type a custom medical visual query",
        )
        if st.button("Search again", key=f"search_again_{sub.row_index}") and custom_query:
            try:
                custom_candidates = []
                for page in wikimedia_search(custom_query, limit=15):
                    cand = parse_wikimedia_candidate(page, sub.title, sub.link, custom_query)
                    if cand:
                        custom_candidates.append(cand)
                for raw in fetch_openstax_figures(custom_query, limit_per_book=6):
                    custom_candidates.append(parse_openstax_candidate(raw, sub.title, sub.link, custom_query))
                merged = dedupe_suggestions(sub_suggestions + custom_candidates)
                st.session_state.suggestions[sub.title] = rank_and_trim(merged, target_max=5)
                st.success("Custom search applied.")
            except Exception as e:
                st.error(f"Custom search failed: {e}")

        # Backup export
        tracker_rows = read_tracker_rows(sheets, tracker_sheet_id, tracker_tab_name)
        csv_text = build_backup_csv(tracker_rows)
        st.download_button(
            label=f"Backup export CSV ({sub.title})",
            data=csv_text,
            file_name=f"backup_{sub.row_index}.csv",
            mime="text/csv",
            key=f"backup_{sub.row_index}",
        )

        if st.button("Write approved (this subchapter)", key=f"write_{sub.row_index}"):
            approved = [x for x in st.session_state.suggestions[sub.title] if x.approved]
            allowed = [x for x in approved if x.status != "Needs license check"]
            if len(allowed) != len(approved):
                st.warning(
                    "Some approved rows are marked 'Needs license check' and will be skipped."
                )
            stats = write_suggestions_for_subchapter(
                sheets,
                tracker_sheet_id,
                tracker_tab_name,
                sub.link,
                allowed,
                dry_run=dry_run,
            )
            if dry_run:
                st.info(f"Dry run: would delete {stats['would_delete']} and insert {stats['would_insert']} rows.")
            else:
                st.success(f"Wrote rows: deleted {stats['deleted']}, inserted {stats['inserted']}.")

    st.subheader("5) Batch write approved")
    if st.button("Write approved (all selected)"):
        total_deleted = 0
        total_inserted = 0

        for sub in selected_rows:
            approved = [x for x in st.session_state.suggestions.get(sub.title, []) if x.approved]
            allowed = [x for x in approved if x.status != "Needs license check"]
            stats = write_suggestions_for_subchapter(
                sheets,
                tracker_sheet_id,
                tracker_tab_name,
                sub.link,
                allowed,
                dry_run=dry_run,
            )
            if dry_run:
                total_deleted += stats["would_delete"]
                total_inserted += stats["would_insert"]
            else:
                total_deleted += stats["deleted"]
                total_inserted += stats["inserted"]

        if dry_run:
            st.info(f"Dry run total: would delete {total_deleted}, would insert {total_inserted}.")
        else:
            st.success(f"Batch write complete. Deleted {total_deleted}, inserted {total_inserted}.")


if __name__ == "__main__":
    render()
