"""
EWB — Agent Auto Stat Submission + Admin Dashboard

AGENT SIDE  (normal link)
  Agents log in, fill in the Auto Stat form. Each submission becomes ONE row in
  the worksheet EWB_Submissions, in the EWB upload column order:

    CH CODE | ACCOUNT NUMBER | STATUS CODE | REMARKS | REMARKS BY |
    REMARKS DATE (MM/DD/YYYY HH:MM:SS) | PTP DATE (MM/DD/YYYY) | PTP KEPT (amount)

  - PTP statuses  -> PTP DATE = promised date,  PTP KEPT = promised amount
  - KEPT statuses -> PTP DATE = date paid,      PTP KEPT = amount paid

  Three bookkeeping columns are added after those 8 (ID, EXPORTED, EXPORTED AT).
  They are never included in the downloaded Excel.

ADMIN SIDE  (open the app link with  ?admin=1  on the end)
  Review every submission, filter / search, see per-agent totals, download the
  EWB-format Excel (new rows get marked as exported), clear test data.
"""

import hashlib
import hmac
import math
import re
import time
import uuid
from datetime import date, datetime
from io import BytesIO

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
CAMPAIGN = "EWB"

EWB_AGENTS = [
    "SFBALPONGO",
    "JATADEO",
    "CPAGAUITAN",
    "GARUFIL",
    "STERNOLA",
]

STATUS_CODES = [
    "CALL - KEPT PAYOFF",
    "CALL - KEPT_REPO 3RD PARTY",
    "CALL - KEPT_REPO CLIENT",
    "CARAVAN - KEPT PAYOFF",
    "CARAVAN - KEPT_REPO 3RD PARTY",
    "CARAVAN - KEPT_REPO CLIENT",
    "EMAIL - KEPT PAYOFF",
    "EMAIL - KEPT_ REPO CLIENT",
    "EMAIL - KEPT_REPO 3RD PARTY",
    "FIELD - KEPT PAYOFF",
    "FIELD - KEPT_REPO 3RD PARTY",
    "FIELD - KEPT_REPO CLIENT",
    "REPO AI - KEPT PAYOFF",
    "REPO AI - KEPT_REPO 3RD PARTY",
    "REPO AI - KEPT_REPO CLIENT",
    "SKIP - KEPT PAYOFF",
    "SKIP - KEPT_REPO 3RD PARTY",
    "SKIP - KEPT_REPO CLIENT",
    "SKIP - PTP PAYOFF_CON",
    "SKIP - PTP_FULL UPDATE_CON",
    "SKIP - PTP_PARTIAL_CON",
    "SKIP - PTP_REPO 3RD PARTY_CON",
    "SKIP - PTP_REPO CLIENT_CON",
    "SKIP -PTP_PUSH BACK",
    "SMEDIA - KEPT PAYOFF",
    "SMEDIA - KEPT_REPO 3RD PARTY",
    "SMEDIA - KEPT_REPO CLIENT",
    "SMS - KEPT PAYOFF",
    "SMS - KEPT_REPO 3RD PARTY",
    "SMS - KEPT_REPO CLIENT",
    "VIBER - KEPT PAYOFF",
    "VIBER - KEPT_REPO 3RD PARTY",
    "VIBER - KEPT_REPO CLIENT",
    "CALL - FOLLOW UP COMPLIANT",
    "CALL - FOLLOW UP KOR",
    "CALL - FOLLOW UP LMTRC",
    "CALL - FOLLOW UP UNCONTACTABLE",
    "CARAVAN - FOLLOW UP COMPLIANT",
    "EMAIL - FOLLOW UP COMPLIANT",
    "EMAIL - FOLLOW UP MESSAGE",
    "FIELD - FOLLOW UP COMPLIANT",
    "SMEDIA - FOLLOW UP COMPLIANT",
    "SMEDIA - FOLLOW UP MESSAGE",
    "SMS - FOLLOW UP COMPLIANT",
    "SMS - FOLLOW UP MESSAGE",
    "VIBER - FOLLOW UP COMPLIANT",
    "VIBER - FOLLOW UP MESSAGE",
    "CALL - PTP PAYOFF",
    "CALL - PTP REPO",
    "CARAVAN - PTP PAYOFF",
    "CARAVAN - PTP REPO",
    "EMAIL - PTP PAYOFF",
    "EMAIL - PTP REPO",
    "FIELD - PTP REPO",
    "FIELD - PTP_PAYOFF",
    "REPO AI - PTP FULL UPDATE",
    "REPO AI - PTP PAY OFF",
    "REPO AI - PTP REPO",
    "SKIP - PTP PAYOFF",
    "SMEDIA - PTP PAYOFF",
    "SMEDIA - PTP REPO",
    "SMS - PTP PAYOFF",
    "SMS - PTP REPO",
    "VIBER - PTP PARTIAL",
    "VIBER - PTP PAYOFF",
    "VIBER - PTP REPO",
    "CALL - CLAIMING PAID",
    "CALL - INSURANCE CLAIM",
    "CALL - UNDERNEGO",
    "CALL - UNIT DAMAGE OR WRECK",
    "CALL - UNIT UNDER HPG",
    "CALL - UNIT_ASSUMED",
    "CALL - UNIT_CARNAPPED",
    "CALL - UNIT_IMPOUNDED",
    "CARAVAN - CLAIMING PAID",
    "CARAVAN - INSURANCE CLAIM",
    "CARAVAN - NO INTENTION TO PAY",
    "CARAVAN - UNIT ASSUMED",
    "CARAVAN - UNIT CARNAPPED",
    "CARAVAN - UNIT DAMAGE OR WRECK",
    "CARAVAN - UNIT IMPOUNDED",
    "CARAVAN - UNIT UNDER HPG",
    "EMAIL - CLAIMING PAID",
    "EMAIL - INSURANCE CLAIM",
    "EMAIL - RESPONSIVE",
    "EMAIL - UNDERNEGO",
    "EMAIL - UNIT ASSUMED",
    "EMAIL - UNIT CARNAPPED",
    "EMAIL - UNIT DAMAGE OR WRECK",
    "EMAIL - UNIT UNDER HPG",
    "EMAIL - UNIT_IMPOUNDED",
    "FIELD - CLAIMING PAID",
    "FIELD - INSURANCE CLAIM",
    "FIELD - NO INTENTION TO PAY",
    "FIELD - UNDERNEGO",
    "FIELD - UNIT ASSUMED",
    "FIELD - UNIT CARNAPPED",
    "FIELD - UNIT DAMAGE OR WRECK",
    "FIELD - UNIT IMPOUNDED",
    "FIELD - UNIT UNDER HPG",
    "REPO AI - UNDERNEGO",
    "SMEDIA - RESPONSIVE",
    "SMEDIA - UNDERNEGO",
    "SMS - CLAIMING PAID",
    "SMS - INSURANCE CLAIM",
    "SMS - RESPONSIVE",
    "SMS - UNDER HPG",
    "SMS - UNDERNEGO",
    "SMS - UNIT ASSUMED",
    "SMS - UNIT CARNAPPED",
    "SMS - UNIT DAMAGE OR WRECK",
    "SMS - UNIT IMPOUNDED",
    "VIBER - CLAIMING PAID",
    "VIBER - INSURANCE CLAIM",
    "VIBER - NO INTENTION TO PAY",
    "VIBER - RESPONSIVE",
    "VIBER - UNDERNEGO",
    "VIBER - UNIT ASSUMED",
    "VIBER - UNIT CARNAPPED",
    "VIBER - UNIT DAMAGE OR WRECK",
    "VIBER - UNIT IMPOUNDED",
    "VIBER - UNIT UNDER HPG",
    "CALL - POS_LEAVE MSG TO 3RD PARTY",
    "CARAVAN - LEAVE MESSAGE TO 3RD PARTY",
    "FIELD - LEAVE_MESSAGE TO 3RD PARTY",
    "EMAIL - NO EMAIL",
    "VIBER - NO VIBER",
    "CALL - DECEASED",
    "CALL - NEG_BUSY",
    "CALL - NEG_CBR",
    "CALL - NEG_DROPPED",
    "CALL - NEG_EMPLOYER NLC",
    "CALL - NEG_KOR",
    "CALL - NEG_LEAVE MSG TO 3RD PARTY",
    "CALL - NEG_NOT IN SERVICE",
    "CALL - NEG_UNATTENDED",
    "CALL - NEG_WRONG NUMBER",
    "CALL - NO INTENTION TO PAY",
    "CALL - POS_BUSY",
    "CALL - POS_CBR",
    "CALL - POS_DROPPED",
    "CALL - POS_KOR",
    "CALL - POS_UNATTENDED",
    "CARAVAN - CLIENT OUT OF AREA",
    "CARAVAN - CLIENT UNKNOWN",
    "CARAVAN - DECEASED",
    "CARAVAN - HOUSED CLOSED UNVERIFIED",
    "CARAVAN - HOUSED CLOSED VERIFIED",
    "CARAVAN - LOT ONLY",
    "CARAVAN - MOVED OUT",
    "CARAVAN - NOT ALLOWED TO ENTER",
    "CARAVAN - RESULT",
    "CARAVAN - UNLOCATED",
    "CEASE - CLAIMING PAID",
    "CEASE - INSURANCE CLAIM",
    "CEASE - PENDING COMPLAINT",
    "CEASE - POSSIBLE COMPLAINT",
    "CEASE - REPOSSESSED BY OTHER ECA",
    "CEASE - REQUESTED BY BANK",
    "CEASE - VALID COMPLAINT",
    "EMAIL - CEASE COLLECTION",
    "EMAIL - DECEASED",
    "EMAIL - GOT NEW CONTACT",
    "EMAIL - NEG_SENT MESSAGE",
    "EMAIL - NO INTENTION TO PAY",
    "EMAIL - POS_SENT MESSAGE",
    "FIELD - CLIENT_OUT OF AREA",
    "FIELD - CLIENT_UNKNOWN",
    "FIELD - DECEASED",
    "FIELD - HOUSED CLOSED VERIFIED",
    "FIELD - HOUSED_CLOSED UNVERIFIED",
    "FIELD - LOT_ONLY",
    "FIELD - MOVED_OUT",
    "FIELD - NOT_ALLOWED TO ENTER",
    "FIELD - RESULT",
    "FIELD - UNLOCATED",
    "FIELD REQUEST - BP_NC",
    "FIELD REQUEST - FOR REVISIT",
    "FIELD REQUEST - NEW_ADDRESS",
    "FIELD REQUEST - OTS SURE REPO",
    "SKIP - CONTACT NUMBER",
    "SKIP - NEGATIVE",
    "SKIP - NEW ADDRESS",
    "SKIP - POSSIBLE LEADS",
    "SKIP - SMEDIA ACCOUNT",
    "SKIP - UNIT ASSUMED",
    "SKIP - UNIT CARNAPPED",
    "SKIP - UNIT DAMAGE OR WRECK",
    "SKIP - UNIT IMPOUNDED",
    "SKIP - UNIT UNDER HPG",
    "SMEDIA - CLAIMING PAID",
    "SMEDIA - INSURANCE CLAIM",
    "SMEDIA - NEG_SENT A MESSAGE",
    "SMEDIA - NO INTENTION TO PAY",
    "SMEDIA - POS_SENT A MESSAGE",
    "SMEDIA - UNIT ASSUMED",
    "SMEDIA - UNIT CARNAPPED",
    "SMEDIA - UNIT DAMAGE OR WRECK",
    "SMEDIA - UNIT IMPOUNDED",
    "SMEDIA - UNIT UNDER HPG",
    "SMS - DECEASED",
    "SMS - GOT NEW CONTACT NUM",
    "SMS - NEG_SENT MESSAGE",
    "SMS - NO INTENTION TO PAY",
    "SMS - POS_SENT MESSAGE",
    "SMS - WRONG NUMBER",
    "VIBER - BOUNCED",
    "VIBER - DELIVERED",
    "VIBER - NEG_SENT A MESSAGE",
    "VIBER - PENDING",
    "VIBER - POS_SENT A MESSAGE",
    "VIBER - READ",
]

REQUIRE_PASSWORD   = True
AGENT_PASSWORD     = "EWB2026"
SESSION_HOURS      = 12
LOGIN_SIGNING_KEY  = "ewb-change-this-long-random-string"

ADMIN_PASSWORD     = "changeme123"      # <-- change this

MIN_ACCOUNT_DIGITS = 10   # EWB accounts look like 200057765304 (12 digits)

SUBMISSIONS_SHEET = "EWB_Submissions"
# The 8 columns of the EWB upload format (this is what gets downloaded)
EWB_COLUMNS = [
    "CH CODE", "ACCOUNT NUMBER", "STATUS CODE", "REMARKS", "REMARKS BY",
    "REMARKS DATE", "PTP DATE", "PTP KEPT",
]
# Sheet columns = EWB columns + bookkeeping
SUBMISSIONS_HEADERS = EWB_COLUMNS + ["ID", "EXPORTED", "EXPORTED AT"]

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ----------------------------------------------------------------------
# GOOGLE SHEETS
# ----------------------------------------------------------------------

@st.cache_resource
def _get_gspread_client():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=GOOGLE_SCOPES
    )
    return gspread.authorize(creds)


@st.cache_resource
def _get_spreadsheet():
    gc = _get_gspread_client()
    sheets_cfg = st.secrets["sheets"]
    url = sheets_cfg.get("ewb_spreadsheet_url") or sheets_cfg["spreadsheet_url"]
    return gc.open_by_url(url)


@st.cache_resource
def _get_or_create_ws(name, headers_tuple):
    headers = list(headers_tuple)
    sh = _get_spreadsheet()
    try:
        ws = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=2000, cols=len(headers) + 2)
        ws.append_row(headers)
        return ws
    if not ws.row_values(1):
        ws.append_row(headers)
    return ws


def _ws():
    return _get_or_create_ws(SUBMISSIONS_SHEET, tuple(SUBMISSIONS_HEADERS))


def ensure_schema(ws=None):
    """Upgrade an older 8-column EWB_Submissions tab: add the ID / EXPORTED /
    EXPORTED AT headers and give every existing row an ID. Safe to run often."""
    ws = ws or _ws()
    n = len(SUBMISSIONS_HEADERS)
    header = ws.row_values(1)
    if header[:n] != SUBMISSIONS_HEADERS:
        if ws.col_count < n:
            ws.add_cols(n - ws.col_count)
        ws.batch_update(
            [{"range": f"A1:{gspread.utils.rowcol_to_a1(1, n)}", "values": [SUBMISSIONS_HEADERS]}],
            value_input_option="RAW",
        )
    id_idx = SUBMISSIONS_HEADERS.index("ID")
    updates = []
    for row_num, row in enumerate(ws.get_all_values()[1:], start=2):
        has_data = any(str(c).strip() for c in row[: len(EWB_COLUMNS)])
        has_id = len(row) > id_idx and str(row[id_idx]).strip()
        if has_data and not has_id:
            updates.append(
                {
                    "range": gspread.utils.rowcol_to_a1(row_num, id_idx + 1),
                    "values": [[uuid.uuid4().hex[:10]]],
                }
            )
    if updates:
        ws.batch_update(updates, value_input_option="RAW")


def init_db():
    ensure_schema(_ws())


def normalize_account(raw) -> str:
    """Digits only, no dashes, no padding (EWB format)."""
    if raw is None:
        return ""
    if isinstance(raw, float):
        if math.isnan(raw) or math.isinf(raw):
            return ""
        if raw.is_integer():
            raw = int(raw)
    return re.sub(r"\D", "", str(raw))


def insert_submission(ch_code, account_number, status_code, remarks,
                      agent, remark_dt, date_val, amount, ws=None):
    """Append one row in the EWB column order. Everything is written as text
    (RAW) so Sheets never reformats dates, amounts or account numbers."""
    ws = ws or _ws()
    clean_remarks = re.sub(r"\s*[\r\n]+\s*", " ", remarks or "").strip()
    ws.append_row(
        [
            ch_code.strip().upper(),
            normalize_account(account_number),
            status_code.strip(),
            clean_remarks,
            agent.strip().upper(),
            remark_dt.strftime("%m/%d/%Y %H:%M:%S"),
            date_val.strftime("%m/%d/%Y") if date_val else "",
            f"{amount:,.2f}" if amount else "",
            uuid.uuid4().hex[:10],   # ID
            0,                        # EXPORTED
            "",                       # EXPORTED AT
        ],
        value_input_option="RAW",
    )


def _bucket(status_code: str) -> str:
    s = (status_code or "").upper()
    if "PTP" in s:
        return "PTP"
    if "KEPT" in s:
        return "KEPT"
    return "OTHER"


def load_submissions(ws=None) -> pd.DataFrame:
    ws = ws or _ws()
    try:
        # keep every cell as text so account numbers keep leading zeros
        records = ws.get_all_records(numericise_ignore=["all"])
    except TypeError:
        records = ws.get_all_records()
    df = pd.DataFrame(records) if records else pd.DataFrame(columns=SUBMISSIONS_HEADERS)
    for col in SUBMISSIONS_HEADERS:
        if col not in df.columns:
            df[col] = ""
    df = df[SUBMISSIONS_HEADERS].copy()
    for col in SUBMISSIONS_HEADERS:
        if col != "EXPORTED":
            df[col] = df[col].astype(str)
    df["EXPORTED"] = pd.to_numeric(df["EXPORTED"], errors="coerce").fillna(0).astype(int)
    df = df[df["ACCOUNT NUMBER"].str.strip() != ""]
    df["_bucket"] = df["STATUS CODE"].apply(_bucket)
    df["_dt"] = pd.to_datetime(df["REMARKS DATE"], format="%m/%d/%Y %H:%M:%S", errors="coerce")
    return df.iloc[::-1].reset_index(drop=True)   # newest first


def mark_exported(ids, ws=None):
    if not ids:
        return
    ids = set(str(i) for i in ids)
    ws = ws or _ws()
    all_values = ws.get_all_values()
    if not all_values:
        return
    header = all_values[0]
    id_col = header.index("ID")
    exp_col = header.index("EXPORTED")
    exp_at_col = header.index("EXPORTED AT")
    now = datetime.now().isoformat(timespec="seconds")
    updates = []
    for row_num, row in enumerate(all_values[1:], start=2):
        if len(row) > id_col and row[id_col] in ids:
            updates.append({"range": gspread.utils.rowcol_to_a1(row_num, exp_col + 1), "values": [[1]]})
            updates.append({"range": gspread.utils.rowcol_to_a1(row_num, exp_at_col + 1), "values": [[now]]})
    if updates:
        ws.batch_update(updates, value_input_option="RAW")


def clear_submissions(ws=None):
    ws = ws or _ws()
    ws.clear()
    ws.append_row(SUBMISSIONS_HEADERS)


def build_export_workbook(df: pd.DataFrame) -> bytes:
    """Excel with exactly the 8 EWB columns, oldest row first, all as text."""
    out_df = df[EWB_COLUMNS].iloc[::-1].reset_index(drop=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Sheet1", index=False)
    return output.getvalue()


# ----------------------------------------------------------------------
# STYLES
# ----------------------------------------------------------------------

def inject_styles(wide: bool = False):
    max_w = "1250px" if wide else "680px"
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

        *, *::before, *::after { box-sizing: border-box; }
        html, body, [class*="css"] { font-family: 'Inter', sans-serif !important; }

        #MainMenu, footer, header { visibility: hidden; }
        .stDeployButton { display: none !important; }
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="stToolbar"] { display: none !important; }

        .stApp {
            background: linear-gradient(160deg, #e8f4fd 0%, #dbeeff 40%, #c8e4f8 100%);
            min-height: 100vh;
        }
        [data-testid="stAppViewContainer"] { background: transparent !important; }
        .block-container {
            padding-top: 0 !important;
            padding-bottom: 2rem !important;
            max-width: __MAXW__ !important;
        }

        /* text that must stay readable on the light page */
        .stApp h1, .stApp h2, .stApp h3, .stApp h4 { color: #1e3a5f !important; }
        [data-testid="stCaptionContainer"], [data-testid="stCaptionContainer"] * { color: #64748b !important; }
        [data-testid="stMetric"] {
            background: #fff; border: 1px solid #bfdbfe; border-radius: 12px; padding: 0.8rem 1rem;
            box-shadow: 0 2px 10px rgba(37,99,235,0.06);
        }
        [data-testid="stMetricLabel"], [data-testid="stMetricLabel"] * { color: #64748b !important; }
        [data-testid="stMetricValue"], [data-testid="stMetricValue"] * { color: #1e3a5f !important; }
        [data-testid="stExpander"] summary, [data-testid="stExpander"] summary * { color: #1e3a5f !important; }

        /* NAV BAR */
        .ewb-nav {
            display: flex; align-items: center; justify-content: space-between;
            padding: 0.75rem 1.5rem; background: #1d4ed8; margin-bottom: 0;
        }
        .ewb-nav-logo { display: flex; align-items: center; gap: 0.65rem; }
        .ewb-badge {
            background: #fff; color: #1d4ed8; font-size: 0.72rem; font-weight: 900;
            letter-spacing: 0.20em; padding: 0.30rem 0.75rem; border-radius: 5px;
        }
        .ewb-nav-title {
            color: rgba(255,255,255,0.75); font-size: 0.74rem; font-weight: 500;
            letter-spacing: 0.08em; text-transform: uppercase;
        }
        .ewb-agent-label { font-size: 0.72rem; color: rgba(255,255,255,0.65); font-weight: 400; }
        .ewb-agent-name  { font-size: 0.78rem; color: #fff; font-weight: 700; margin-left: 0.25rem; }

        /* HERO */
        .ewb-hero { padding: 1.8rem 0 1rem; }
        .ewb-hero.center { text-align: center; padding: 3rem 0 1.5rem; }
        .ewb-eyebrow {
            font-size: 0.65rem; font-weight: 700; letter-spacing: 0.20em;
            color: #1d4ed8; text-transform: uppercase; margin-bottom: 0.35rem;
        }
        .ewb-h1 {
            font-size: 2rem; font-weight: 900; letter-spacing: -0.02em;
            color: #1e3a5f; line-height: 1.1; margin: 0 0 0.25rem;
        }
        .ewb-h1 span {
            background: linear-gradient(90deg, #2563eb, #1d4ed8);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .ewb-sub { font-size: 0.82rem; color: #64748b; font-weight: 400; }
        .ewb-logo-badge {
            display: inline-block; background: linear-gradient(135deg, #2563eb, #1e40af);
            color: #fff; font-size: 1.05rem; font-weight: 900; letter-spacing: 0.22em;
            padding: 0.55rem 1.4rem; border-radius: 9px;
            box-shadow: 0 4px 20px rgba(29,78,216,0.35);
        }

        /* CARDS: st.container(border=True) */
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: #ffffff;
            border: 1px solid #bfdbfe !important;
            border-radius: 14px !important;
            box-shadow: 0 2px 12px rgba(37,99,235,0.07);
        }
        .ewb-sec-label {
            font-size: 0.64rem; font-weight: 700; letter-spacing: 0.16em;
            color: #3b82f6; text-transform: uppercase;
            margin: 0.2rem 0 0.6rem; display: flex; align-items: center; gap: 0.5rem;
        }
        .ewb-sec-label::after {
            content: ''; flex: 1; height: 1px;
            background: linear-gradient(90deg, #bfdbfe, transparent);
        }
        .ewb-box-tag {
            font-size: 0.66rem; font-weight: 700; letter-spacing: 0.12em;
            text-transform: uppercase; margin: 0.6rem 0 0.2rem;
        }
        .ewb-box-tag.ptp  { color: #1d4ed8; }
        .ewb-box-tag.kept { color: #059669; }

        /* INPUTS */
        [data-testid="stTextInput"] input,
        [data-testid="stDateInput"] input,
        [data-testid="stNumberInput"] input,
        [data-testid="stTextArea"] textarea {
            background: #f0f7ff !important;
            border: 1.5px solid #bfdbfe !important;
            border-radius: 9px !important;
            color: #1e293b !important;
            font-family: 'Inter', sans-serif !important;
            font-size: 0.88rem !important;
        }
        [data-testid="stTextInput"] input::placeholder,
        [data-testid="stTextArea"] textarea::placeholder { color: #94a3b8 !important; }
        [data-testid="stTextInput"] input:focus,
        [data-testid="stTextArea"] textarea:focus {
            border-color: #2563eb !important;
            box-shadow: 0 0 0 3px rgba(37,99,235,0.12) !important;
            background: #fff !important;
        }
        [data-testid="stTextInput"] input:disabled {
            color: #64748b !important; background: #e2e8f0 !important; border-color: #cbd5e1 !important;
            -webkit-text-fill-color: #64748b !important;
        }
        [data-testid="stSelectbox"] > div > div,
        [data-testid="stMultiSelect"] > div > div {
            background: #f0f7ff !important;
            border: 1.5px solid #bfdbfe !important;
            border-radius: 9px !important;
            color: #1e293b !important;
        }
        [data-testid="stNumberInput"] button {
            background: #e0f0ff !important; border-color: #bfdbfe !important; color: #1d4ed8 !important;
        }
        [data-testid="stTimeInput"] input {
            background: #f0f7ff !important; border: 1.5px solid #bfdbfe !important;
            color: #1e293b !important; border-radius: 9px !important;
        }

        /* LABELS */
        label, .stTextInput label, .stSelectbox label, .stTextArea label,
        .stDateInput label, .stNumberInput label, .stTimeInput label {
            color: #475569 !important;
            font-size: 0.70rem !important; font-weight: 700 !important;
            letter-spacing: 0.08em !important; text-transform: uppercase !important;
        }

        /* SECONDARY BUTTONS (blue outline) */
        .stButton > button, [data-testid="stDownloadButton"] > button {
            background: #ffffff !important;
            color: #1d4ed8 !important;
            border: 1.5px solid #93c5fd !important;
            border-radius: 9px !important;
            font-size: 0.78rem !important; font-weight: 700 !important;
            transition: all 0.18s !important;
        }
        .stButton > button:hover, [data-testid="stDownloadButton"] > button:hover {
            background: #eff6ff !important; border-color: #2563eb !important;
        }

        /* PRIMARY BUTTONS (solid blue): submit, log in, download */
        button[kind="primary"], [data-testid="stBaseButton-primary"],
        [data-testid="stBaseButton-primary"] {
            background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
            color: #fff !important; border: none !important;
            border-radius: 10px !important;
            font-size: 0.85rem !important; font-weight: 800 !important;
            letter-spacing: 0.10em !important; text-transform: uppercase !important;
            padding: 0.7rem 1.5rem !important;
            box-shadow: 0 4px 20px rgba(29,78,216,0.35) !important;
        }
        button[kind="primary"] *, [data-testid="stBaseButton-primary"] * { color: #fff !important; }
        button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {
            box-shadow: 0 6px 28px rgba(29,78,216,0.50) !important;
            transform: translateY(-1px) !important;
        }

        /* FLASH / ALERTS */
        .ewb-flash {
            background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 10px;
            padding: 0.85rem 1.1rem; color: #15803d; font-size: 0.83rem; font-weight: 600;
            display: flex; align-items: center; gap: 0.5rem; margin-bottom: 0.85rem;
        }
        [data-testid="stAlert"] { border-radius: 10px !important; font-size: 0.82rem !important; }
        hr { border-color: #bfdbfe !important; }
        </style>
        """.replace("__MAXW__", max_w),
        unsafe_allow_html=True,
    )


# ----------------------------------------------------------------------
# AGENT AUTH
# ----------------------------------------------------------------------

def _sign_login(agent: str, exp: int) -> str:
    msg = f"{agent}|{exp}|{AGENT_PASSWORD}".encode()
    return hmac.new(LOGIN_SIGNING_KEY.encode(), msg, hashlib.sha256).hexdigest()


def _make_login_token(agent: str) -> str:
    exp = int(time.time()) + SESSION_HOURS * 3600
    return f"{agent}|{exp}|{_sign_login(agent, exp)}"


def _verify_login_token(token: str):
    try:
        agent, exp, sig = str(token).split("|")
        exp = int(exp)
    except ValueError:
        return None
    if agent not in EWB_AGENTS or exp < time.time():
        return None
    if not hmac.compare_digest(sig, _sign_login(agent, exp)):
        return None
    return agent


def restore_agent_login():
    if st.session_state.get("agent_user"):
        return
    token = st.query_params.get("t")
    if token:
        agent = _verify_login_token(token)
        if agent:
            st.session_state["agent_user"] = agent


def agent_logout():
    st.session_state["agent_user"] = None
    if "t" in st.query_params:
        del st.query_params["t"]


# ----------------------------------------------------------------------
# AGENT PAGES
# ----------------------------------------------------------------------

def agent_login_page():
    st.markdown(
        f"""
        <div class="ewb-hero center">
            <div class="ewb-logo-badge">{CAMPAIGN}</div>
            <div class="ewb-h1" style="font-size:1.5rem; margin-top:0.8rem;">Agent Portal</div>
            <div class="ewb-sub" style="text-transform:uppercase; letter-spacing:0.08em;">Auto Stat Submission</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    agent = st.selectbox(
        "AGENT", EWB_AGENTS, index=None,
        placeholder="Select your name...", key="login_agent",
    )
    pwd = ""
    if REQUIRE_PASSWORD:
        pwd = st.text_input("PASSWORD", type="password",
                            placeholder="Enter shared password", key="login_pwd")

    if st.button("→  LOG IN", type="primary", use_container_width=True, key="login_btn"):
        pwd_ok = (not REQUIRE_PASSWORD) or hmac.compare_digest(str(pwd), AGENT_PASSWORD)
        if agent and pwd_ok:
            st.session_state["agent_user"] = agent
            st.query_params["t"] = _make_login_token(agent)
            st.rerun()
        else:
            st.error("Wrong agent or password — please try again.")


def autostat_form_page():
    if "as_v" not in st.session_state:
        st.session_state["as_v"] = 0
    v = st.session_state["as_v"]

    collector_raw = st.session_state.get("agent_user") or ""
    collector = collector_raw.strip().upper()

    st.markdown(
        f"""
        <div class="ewb-hero">
            <div class="ewb-eyebrow">{CAMPAIGN} · Collection Activity</div>
            <div class="ewb-h1">Auto <span>Stat</span></div>
            <div class="ewb-sub">Fill in the details below and submit your status update.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.session_state.pop("as_flash", False):
        st.markdown(
            '<div class="ewb-flash">✅ &nbsp;Submission saved — you may submit another one.</div>',
            unsafe_allow_html=True,
        )

    # ── CARD 1: Account Info ──
    with st.container(border=True):
        st.markdown('<div class="ewb-sec-label">Account Info</div>', unsafe_allow_html=True)

        st.text_input("REMARKS BY (AGENT)", value=collector, disabled=True, key=f"as_col_{v}")
        ch_code = st.text_input("CH CODE", placeholder="e.g. EWAL02F2608-1079", key=f"as_ch_{v}")
        account_number = st.text_input("ACCOUNT NUMBER", placeholder="e.g. 200057765304", key=f"as_acct_{v}")
        status_code = st.selectbox(
            "STATUS CODE", STATUS_CODES, index=None,
            placeholder="Type to search status...", key=f"as_sc_{v}",
        )

        sc       = status_code or ""
        has_ptp  = "PTP"  in sc.upper()
        has_kept = "KEPT" in sc.upper() and not has_ptp

        # PTP and KEPT statuses fill the same two EWB columns: PTP DATE + PTP KEPT
        date_val = None
        amount   = None

        if has_ptp:
            st.markdown('<div class="ewb-box-tag ptp">📅 PTP — Promise to Pay</div>', unsafe_allow_html=True)
            p1, p2 = st.columns(2)
            with p1:
                date_val = st.date_input("PTP DATE", value=None, format="MM/DD/YYYY", key=f"as_pd_{v}")
            with p2:
                amount = st.number_input(
                    "PTP AMOUNT (₱)", min_value=0.0, step=100.0, format="%.2f", key=f"as_pa_{v}"
                )
        elif has_kept:
            st.markdown('<div class="ewb-box-tag kept">💳 Payment Kept</div>', unsafe_allow_html=True)
            k1, k2 = st.columns(2)
            with k1:
                date_val = st.date_input("DATE PAID", value=None, format="MM/DD/YYYY", key=f"as_kd_{v}")
            with k2:
                amount = st.number_input(
                    "AMOUNT PAID (₱)", min_value=0.0, step=100.0, format="%.2f", key=f"as_ka_{v}"
                )

    # ── CARD 2: Remarks & Timestamp ──
    with st.container(border=True):
        st.markdown('<div class="ewb-sec-label">Remarks & Timestamp</div>', unsafe_allow_html=True)

        remarks = st.text_area(
            "REMARKS", placeholder="Enter full remarks / spiel notes here...",
            height=120, key=f"as_rem_{v}",
        )
        rd1, rd2 = st.columns(2)
        with rd1:
            remark_date = st.date_input(
                "REMARK DATE", value=date.today(), format="MM/DD/YYYY", key=f"as_rd_{v}"
            )
        with rd2:
            remark_time = st.time_input(
                "REMARK TIME",
                value=datetime.now().time().replace(second=0, microsecond=0),
                step=60, key=f"as_rt_{v}",
            )

    # ── SUBMIT ──
    if st.button("✦  SUBMIT STATUS", type="primary", use_container_width=True, key=f"as_sub_{v}"):
        errors = []
        clean_ch = (ch_code or "").strip().upper()
        clean_account = normalize_account(account_number)

        if not clean_ch:
            errors.append("CH Code is required.")
        if not clean_account or len(clean_account) < MIN_ACCOUNT_DIGITS:
            errors.append(f"Account Number looks invalid — enter at least {MIN_ACCOUNT_DIGITS} digits.")
        if not status_code:
            errors.append("Status Code is required.")
        if not collector:
            errors.append("Agent is required — please log in again.")
        elif collector not in [a.upper() for a in EWB_AGENTS]:
            errors.append(f"'{collector_raw.strip()}' isn't a recognized agent.")
        if has_ptp and not date_val:
            errors.append("PTP Date is required for this status.")
        if has_ptp and (amount in (None, 0, 0.0)):
            errors.append("PTP Amount must be greater than 0 for this status.")
        if has_kept and not date_val:
            errors.append("Date Paid is required for this status.")
        if has_kept and (amount in (None, 0, 0.0)):
            errors.append("Amount Paid must be greater than 0 for this status.")

        if errors:
            for e in errors:
                st.error(e)
        else:
            remark_dt = datetime.combine(remark_date, remark_time)
            try:
                insert_submission(
                    clean_ch, clean_account, status_code, remarks,
                    collector, remark_dt, date_val, amount,
                )
            except Exception as e:
                st.error(f"Could not save to Google Sheets: {e}")
                return
            st.session_state["as_flash"] = True
            st.session_state["as_v"] += 1
            st.rerun()


# ----------------------------------------------------------------------
# ADMIN
# ----------------------------------------------------------------------

def admin_login():
    st.markdown(
        f"""
        <div class="ewb-hero center">
            <div class="ewb-logo-badge">{CAMPAIGN}</div>
            <div class="ewb-h1" style="font-size:1.5rem; margin-top:0.8rem;">Admin Access</div>
            <div class="ewb-sub" style="text-transform:uppercase; letter-spacing:0.08em;">Restricted — Authorized Personnel Only</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    pwd = st.text_input("PASSWORD", type="password", placeholder="Enter admin password", key="admin_pwd")
    if st.button("Unlock Dashboard", type="primary", use_container_width=True, key="admin_unlock"):
        if hmac.compare_digest(str(pwd), ADMIN_PASSWORD):
            st.session_state["is_admin"] = True
            st.rerun()
        else:
            st.error("Wrong password.")


def _on_export_download(ids):
    """Runs BEFORE the rerun, so the table redraws without the exported rows."""
    try:
        mark_exported(ids)
        st.session_state["export_msg"] = ("ok", f"Marked {len(ids)} row(s) as exported.")
    except Exception as e:
        st.session_state["export_msg"] = (
            "err",
            f"Download worked, but marking rows as exported FAILED: {e}. "
            "Rows were NOT marked - they will still show as new.",
        )


def admin_dashboard():
    st.markdown(
        f"""
        <div class="ewb-hero">
            <div class="ewb-eyebrow">{CAMPAIGN} · Control Panel</div>
            <div class="ewb-h1">Admin <span>Dashboard</span></div>
            <div class="ewb-sub">Review Auto Stat submissions and download the EWB upload file.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    b1, b2, _ = st.columns([1, 1, 4])
    with b1:
        if st.button("🔄 Refresh", key="admin_refresh", use_container_width=True):
            st.rerun()
    with b2:
        if st.button("🔒 Log out", key="admin_logout", use_container_width=True):
            st.session_state["is_admin"] = False
            st.rerun()

    try:
        df = load_submissions()
    except Exception as e:
        st.error(f"Could not read the Google Sheet: {e}")
        return

    # ---------- reset ----------
    with st.expander("🗑️ Clear all EWB data (test data reset)", expanded=False):
        st.warning("This **permanently deletes** every row in the EWB_Submissions tab. Use only to clear test data.")
        if st.button("I understand — clear everything", key="clear_btn"):
            st.session_state["confirm_clear"] = True
        if st.session_state.get("confirm_clear"):
            if st.button("✅ YES, delete everything", key="clear_confirm"):
                clear_submissions()
                st.session_state["confirm_clear"] = False
                st.success("EWB data cleared.")
                st.rerun()

    if df.empty:
        st.info("No submissions yet.")
        return

    # ---------- metrics ----------
    today = date.today()
    today_count = int((df["_dt"].dt.date == today).sum())
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("📋 Total", len(df))
    m2.metric("🆕 Not exported", int((df["EXPORTED"] == 0).sum()))
    m3.metric("📅 Today", today_count)
    m4.metric("🤝 PTP", int((df["_bucket"] == "PTP").sum()))
    m5.metric("💳 KEPT", int((df["_bucket"] == "KEPT").sum()))

    # ---------- per-agent summary ----------
    with st.expander("👥 Totals per agent", expanded=False):
        pivot = pd.pivot_table(
            df, index="REMARKS BY", columns="_bucket", values="ID",
            aggfunc="count", fill_value=0,
        )
        for col in ["PTP", "KEPT", "OTHER"]:
            if col not in pivot.columns:
                pivot[col] = 0
        pivot = pivot[["PTP", "KEPT", "OTHER"]]
        pivot["TOTAL"] = pivot.sum(axis=1)
        pivot = pivot.sort_values("TOTAL", ascending=False).reset_index().rename(columns={"REMARKS BY": "AGENT"})
        st.dataframe(pivot, use_container_width=True, hide_index=True)

    st.divider()

    # ---------- review ----------
    st.subheader("① Review Submissions")

    f1, f2, f3 = st.columns(3)
    with f1:
        agent_opts = sorted(df["REMARKS BY"].unique().tolist())
        agent_filter = st.multiselect("Agent", agent_opts, default=agent_opts)
    with f2:
        type_filter = st.multiselect("Type", ["PTP", "KEPT", "OTHER"], default=["PTP", "KEPT", "OTHER"])
    with f3:
        search = st.text_input("Search account / CH code", placeholder="type to search...")

    hide_exported = st.checkbox("Hide exported rows", value=True)

    filtered = df[df["REMARKS BY"].isin(agent_filter) & df["_bucket"].isin(type_filter)]
    if search.strip():
        q = search.strip().upper()
        filtered = filtered[
            filtered["ACCOUNT NUMBER"].str.contains(q, na=False)
            | filtered["CH CODE"].str.upper().str.contains(q, na=False)
        ]
    if hide_exported:
        filtered = filtered[filtered["EXPORTED"] == 0]

    st.dataframe(
        filtered[EWB_COLUMNS + ["EXPORTED", "ID"]],
        use_container_width=True,
        hide_index=True,
    )
    st.caption(f"Showing {len(filtered)} of {len(df)} total submissions.")

    st.divider()

    # ---------- download ----------
    st.subheader("② Download — EWB Upload File")
    st.caption(
        "The Excel has exactly the 8 EWB columns (CH CODE → PTP KEPT), oldest row first. "
        "The first button downloads only rows NOT yet exported and marks them as exported."
    )

    flash = st.session_state.pop("export_msg", None)
    if flash:
        (st.success if flash[0] == "ok" else st.error)(flash[1])

    to_export = filtered[filtered["EXPORTED"] == 0]

    d1, d2 = st.columns(2)
    with d1:
        if to_export.empty:
            st.info("Nothing new to export with the current filters.")
        else:
            st.download_button(
                label=f"⬇ Download {len(to_export)} new row(s) & mark exported",
                data=build_export_workbook(to_export),
                file_name=f"EWB_export_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_new",
                type="primary",
                on_click=_on_export_download,
                args=(to_export["ID"].tolist(),),
                use_container_width=True,
            )
    with d2:
        if filtered.empty:
            st.info("No rows in the current view.")
        else:
            st.download_button(
                label=f"⬇ Download current view ({len(filtered)}) — does NOT mark",
                data=build_export_workbook(filtered),
                file_name=f"EWB_view_{date.today().isoformat()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_view",
                use_container_width=True,
            )
    st.caption("To re-download rows that were already exported, untick “Hide exported rows” and use the second button.")


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------

def main():
    st.set_page_config(page_title="EWB Auto Stat", page_icon="🔵", layout="centered")

    is_admin_url = st.query_params.get("admin") == "1"
    if "is_admin" not in st.session_state:
        st.session_state["is_admin"] = False
    if "agent_user" not in st.session_state:
        st.session_state["agent_user"] = None

    inject_styles(wide=is_admin_url or st.session_state["is_admin"])

    if "db_initialized" not in st.session_state:
        init_db()
        st.session_state["db_initialized"] = True

    # ---- ADMIN SIDE ----
    if st.session_state["is_admin"]:
        admin_dashboard()
        return
    if is_admin_url:
        admin_login()
        return

    # ---- AGENT SIDE ----
    restore_agent_login()
    if not st.session_state["agent_user"]:
        agent_login_page()
        return

    agent_name = st.session_state["agent_user"]
    st.markdown(
        f"""
        <div class="ewb-nav">
            <div class="ewb-nav-logo">
                <span class="ewb-badge">EWB</span>
                <span class="ewb-nav-title">Auto Stat Portal</span>
            </div>
            <div>
                <span class="ewb-agent-label">Logged in as</span>
                <span class="ewb-agent-name">{agent_name}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    _, btn_col = st.columns([5, 1])
    with btn_col:
        if st.button("Log out", key="agent_logout_btn", use_container_width=True):
            agent_logout()
            st.rerun()

    autostat_form_page()


if __name__ == "__main__":
    main()
