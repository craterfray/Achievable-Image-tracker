import csv
import io
import os
import re
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from urllib.parse import parse_qs, urlparse

import requests
import streamlit as st
from docx import Document
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# -----------------------------
# Local paths
# -----------------------------
APP_DIR = Path(os.environ.get("IMAGE_TRACKER_APP_DIR", Path.home() / ".image-tracker"))
DEFAULT_CREDENTIALS_PATH = str(APP_DIR / "credentials.json")
DEFAULT_TOKEN_PATH = str(APP_DIR / "token.json")

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
    "the", "and", "for", "that", "with", "this", "from", "are", "was", "were", "been", "have",
    "has", "had", "into", "about", "through", "between", "over", "under", "after", "before",
    "because", "while", "where", "when", "which", "their", "there", "also", "they", "them", "his",
    "her", "she", "him", "its", "you", "your", "our", "can", "may", "might", "will", "would",
    "should", "must", "not", "than", "such", "using", "used", "use", "within", "without", "during",
    "patient", "patients", "nurse", "nursing",
}

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


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "")).strip()


def _resolve_auth_paths(credentials_json_path: str, token_path_override: str = "") -> Tuple[Path, Path]:
    credentials_path = Path(credentials_json_path).expanduser()
    if not credentials_path.is_absolute():
        credentials_path = Path.cwd() / credentials_path

    token_path = Path(token_path_override).expanduser() if token_path_override else Path(DEFAULT_TOKEN_PATH)
    if not token_path.is_absolute():
        token_path = Path.cwd() / token_path

    token_path.parent.mkdir(parents=True, exist_ok=True)
    return credentials_path, token_path


# -----------------------------
# Auth and Google clients
# -----------------------------
def get_google_credentials(credentials_json_path: str, token_path_override: str = "") -> Credentials:
    creds: Optional[Credentials] = st.session_state.get("google_credentials")
    credentials_path, token_path = _resolve_auth_paths(credentials_json_path, token_path_override)

    if not credentials_path.exists():
        raise FileNotFoundError(f"credentials.json not found at: {credentials_path}")

    if not creds and token_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
        except Exception:
            creds = None

    if creds and creds.valid:
        st.session_state.google_credentials = creds
        return creds

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
        token_path.write_text(creds.to_json(), encoding="utf-8")
        st.session_state.google_credentials = creds
        return creds

    flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
    auth_url, _ = flow.authorization_url(
        access_type="offline",
        prompt="consent",
        include_granted_scopes="true",
    )

    st.info("Step 1: Open this URL in a new tab and authorize:")
    st.code(auth_url)

    raw = st.text_input(
        "Step 2: Paste the authorization code (or the full redirect URL) here",
        value="",
        type="password",
    )

    if not raw:
        raise RuntimeError("Authorization code not provided yet.")

    auth_code = raw.strip()
    if auth_code.lower().startswith("http"):
        parsed = urlparse(auth_code)
        qs = parse_qs(parsed.query)
        auth_code = qs.get("code", [""])[0]

    if not auth_code:
        raise RuntimeError("Could not extract authorization code. Paste the code or full redirect URL.")

    flow.fetch_token(code=auth_code)
    creds = flow.credentials
    token_path.write_text(creds.to_json(), encoding="utf-8")

    st.success(f"token.json saved to: {token_path}")
    st.session_state.google_credentials = creds
    return creds


def get_google_clients(creds: Credentials):
    sheets = build("sheets", "v4", credentials=creds)
    drive = build("drive", "v3", credentials=creds)
    return sheets, drive


def list_sheet_tabs(sheets_service, spreadsheet_id: str) -> List[str]:
    ss = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    return [s.get("properties", {}).get("title", "") for s in ss.get("sheets", [])]


# -----------------------------
# Outline reading and doc extraction
# -----------------------------
def detect_outline_columns(header: List[str]) -> Tuple[int, int, int]:
    lower = [normalize_space(h).lower() for h in header]
    status_idx = next((i for i, h in enumerate(lower) if "status" in h), -1)
    title_idx = next((i for i, h in enumerate(lower) if any(k in h for k in ["subchapter", "heading", "title", "section"])), -1)
    link_idx = next((i for i, h in enumerate(lower) if any(k in h for k in ["link", "doc", "url"])), -1)

    if status_idx == -1 or title_idx == -1 or link_idx == -1:
        raise ValueError("Could not auto-detect Status / Subchapter / Link columns in outline sheet header.")
    return status_idx, title_idx, link_idx


def read_outline_complete_subchapters(sheets_service, outline_sheet_id: str, tab_name: str) -> List[Subchapter]:
    available_tabs = list_sheet_tabs(sheets_service, outline_sheet_id)
    if tab_name not in available_tabs:
        raise ValueError(f"Tab '{tab_name}' was not found. Available tabs: {available_tabs}")

    data = sheets_service.spreadsheets().values().get(
        spreadsheetId=outline_sheet_id,
        range=f"'{tab_name}'!A1:Z2000",
    ).execute()
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


def tokenize(text: str) -> List[str]:
    tokens = re.findall(r"[a-z][a-z0-9\-]{2,}", text.lower())
    return [t for t in tokens if t not in STOPWORDS]


def ngrams(tokens: List[str], n: int) -> List[str]:
    return [" ".join(tokens[i : i + n]) for i in range(max(0, len(tokens) - n + 1))] if len(tokens) >= n else []


def tfidf_keywords(text: str, top_n: int = 15) -> List[str]:
    tokens = tokenize(text)
    if not tokens:
        return []
    token_counts = Counter(tokens)
    max_tf = max(token_counts.values())
    scores: Dict[str, float] = {}

    for tok, count in token_counts.items():
        scores[tok] = (count / max_tf) * (1.1 if len(tok) > 6 else 1.0)
    for bg, count in Counter(ngrams(tokens, 2)).items():
        if not any(word in STOPWORDS for word in bg.split()):
            scores[bg] = (count / max_tf) * 1.6
    for tg, count in Counter(ngrams(tokens, 3)).items():
        if not any(word in STOPWORDS for word in tg.split()):
            scores[tg] = (count / max_tf) * 2.0

    return [k for k, _ in sorted(scores.items(), key=lambda x: x[1], reverse=True)[:top_n]]


def render():
    st.set_page_config(page_title="Medical Image Citation Tracker", layout="wide")
    st.title("Medical Mini-Textbook Image Suggestion Tracker")

    with st.sidebar:
        st.header("Configuration")
        credentials_path = st.text_input("credentials.json path", value=DEFAULT_CREDENTIALS_PATH)
        token_path = st.text_input("token.json path", value=DEFAULT_TOKEN_PATH)
        outline_sheet_id = st.text_input("Outline sheet ID", value=DEFAULT_OUTLINE_SHEET_ID)
        outline_tab_name = st.text_input("Outline tab name", value="Textbook/quizzes/exams")

    st.subheader("1) Authorize Google")
    if st.button("Authorize with Google"):
        try:
            creds = get_google_credentials(credentials_path, token_path)
            sheets, drive = get_google_clients(creds)
            st.session_state.sheets = sheets
            st.session_state.drive = drive
            st.success("Authorization successful.")
        except Exception as e:
            st.error(f"Authorization failed: {e}")

    if "sheets" not in st.session_state:
        st.info("Authorize first to continue.")
        return

    st.subheader("2) Load complete subchapters from outline")
    if st.button("Load outline"):
        try:
            rows = read_outline_complete_subchapters(st.session_state.sheets, outline_sheet_id, outline_tab_name)
            st.session_state.outline_rows = rows
            st.success(f"Loaded {len(rows)} complete subchapters with doc links.")
        except Exception as e:
            st.error(f"Failed to load outline: {e}")


if __name__ == "__main__":
    render()
