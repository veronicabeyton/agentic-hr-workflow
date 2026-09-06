"""

agent_workflow.py

Agentic HR Workflow: reads new "Advance / Expense Report" form submissions

from a Google Sheet, has Gemini run an AI audit against your HR policy

(`hr_policy.txt`), generates a clean PDF report from a Google Docs template

(template chosen by whether a previous advance was received), and emails HR

with the AI audit verdict + financial reconciliation in the body and the

clean PDF attached.

ARCHITECTURE (September 2026)

------------------------------

Google Form -> Google Sheet -> Python polling script -> financial

reconciliation -> Gemini policy audit -> Google Docs template fill ->

PDF export -> email via Gmail SMTP -> mark Sheet row PROCESSED.

TWO SEPARATE GOOGLE IDENTITIES ARE USED, ON PURPOSE:

1. A SERVICE ACCOUNT, used ONLY for Google Sheets (read submissions, write

   "PROCESSED" back). Service accounts have no personal Drive storage, so

   they must never be the identity that creates/copies/owns Drive files.

2. Your own human Google account, authorized once via OAuth 2.0 (a normal

   browser consent screen), used for Google Docs + Drive: copying the

   template, filling placeholders, and exporting the PDF. Because a real

   Drive copy still has to be created somewhere in order for the Docs API

   to fill placeholders and export a PDF, that copy is created in YOUR

   Drive (where you have normal storage), not the service account's. The

   temporary copy is trashed immediately after the PDF is exported.

This is the fix for the previous `403 storageQuotaExceeded` error: the

service account was trying to own a Drive file copy and service accounts

functionally have 0 bytes of personal Drive quota outside a Shared Drive.

No Shared Drive is required with this design.

This is a Python port of an original Google Apps Script `onFormSubmit`

trigger. A local Python script has no equivalent event trigger, so this

script runs CONTINUOUSLY: every POLL_INTERVAL_SECONDS (default 30s) it

scans the sheet for rows whose `Processed` column is blank, processes each

one, and writes "PROCESSED" into that column. The sheet itself is the

tracking state -- no separate local state database is used. A row is only

marked PROCESSED after reconciliation, the Gemini audit (or its safe

fallback), PDF generation, AND email sending have all succeeded. Rows that

fail are retried later with an in-memory exponential backoff so a

persistently broken row does not get hammered every 30 seconds forever.

Setup

-----

1. pip install -r requirements.txt

2. Enable these APIs in your Google Cloud project: Sheets API, Docs API,

   Drive API.

3. SERVICE ACCOUNT (Sheets only):

   - Create a Service Account, download its JSON key as `credentials.json`

     (same folder as this script, or point GOOGLE_SERVICE_ACCOUNT_FILE at

     it).

   - Share the Google Sheet with the service account's `client_email`

     (found inside credentials.json) as Editor. The service account does

     NOT need access to the Docs templates -- it never touches Docs/Drive.

4. OAUTH CLIENT (Docs + Drive, run as you):

   - In Google Cloud Console -> APIs & Services -> Credentials, create an

     OAuth 2.0 Client ID of type "Desktop app".

   - Download its JSON and save it as `oauth_client_secret.json` next to

     this script (or point GOOGLE_OAUTH_CLIENT_SECRET_FILE at it). This

     file identifies your app to Google; it is not a login credential by

     itself and still requires the interactive consent below.

   - Make sure the Google account you intend to authorize with already

     owns (or has Editor access to) both Google Docs templates.

   - The FIRST time you run this script, a browser window will open

     asking you to log in as that account and approve access. After you

     approve, a token is cached locally at `token.json` (path configurable

     via GOOGLE_OAUTH_TOKEN_FILE) and silently refreshed on every future

     run -- you will not be asked again unless you delete that file or

     revoke access.

   - Treat `oauth_client_secret.json` and `token.json` like secrets: do

     not commit them to source control.

5. Write `hr_policy.txt` with your company's expense/advance policy in

   plain language -- it's sent to Gemini as the audit rulebook.

6. Create a `.env` file (copy `.env.example`) and fill in the values.

7. Fill in the CONFIG section below with your template IDs and HR

   recipient address.

8. Run it: `python agent_workflow.py`. On first run it will pop a browser

   window for the one-time OAuth consent, then start polling.

"""

from __future__ import annotations

import json

import logging

import os

import re

import smtplib

import time

from dataclasses import dataclass, field

from datetime import datetime

from email.mime.application import MIMEApplication

from email.mime.multipart import MIMEMultipart

from email.mime.text import MIMEText

from pathlib import Path

from typing import Any

from dotenv import load_dotenv

from google.auth.exceptions import RefreshError

from google.auth.transport.requests import Request as GoogleAuthRequest

from google.oauth2 import service_account

from google.oauth2.credentials import Credentials as OAuthCredentials

from google_auth_oauthlib.flow import InstalledAppFlow

from googleapiclient.discovery import build

from pydantic import BaseModel

from google import genai

from google.genai import types as genai_types

from google.genai.errors import APIError as GenAIAPIError

# --------------------------------------------------------------------------

# CONFIGURATION

# --------------------------------------------------------------------------

load_dotenv()

_SCRIPT_DIR = Path(__file__).parent

def _require_env(name: str) -> str:

    """Read a required environment variable, failing with a clear, actionable

    message instead of a raw KeyError traceback if it's missing/blank."""

    value = os.environ.get(name)

    if not value or not value.strip():

        raise SystemExit(

            f"Missing required environment variable: {name}\n"

            f"Set it in your .env file (see the setup steps in this script's "

            f"module docstring for the full list of required variables)."

        )

    return value

GEMINI_API_KEY = _require_env("GEMINI_API_KEY")

GOOGLE_SHEET_ID = _require_env("GOOGLE_SHEET_ID")

EMAIL_SENDER_ADDRESS = _require_env("EMAIL_SENDER_ADDRESS")

EMAIL_SENDER_APP_PASSWORD = _require_env("EMAIL_SENDER_APP_PASSWORD")

def _resolve_path(env_name: str, default_filename: str) -> str:

    raw = os.environ.get(env_name, default_filename)

    if os.path.isabs(raw) or os.path.dirname(raw):

        return raw

    return str(_SCRIPT_DIR / raw)

SERVICE_ACCOUNT_FILE = _resolve_path(

    "GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json"

)

OAUTH_CLIENT_SECRET_FILE = _resolve_path(

    "GOOGLE_OAUTH_CLIENT_SECRET_FILE", "oauth_client_secret.json"

)

OAUTH_TOKEN_FILE = _resolve_path(

    "GOOGLE_OAUTH_TOKEN_FILE", "token.json"

)

# Gemini model.

#

# gemini-1.5-flash and gemini-2.5-flash are both being retired by Google in

# 2026. gemini-3.7-flash is the current generally-available stable Flash

# model as of September 2026. Override via GEMINI_MODEL_NAME in .env if

# Google ships a newer stable Flash model later.

GEMINI_MODEL_NAME = os.environ.get(

    "GEMINI_MODEL_NAME",

    "gemini-3.7-flash"

)

# 1. Google Docs template IDs.

ADVANCE_YES_TEMPLATE_ID = "1PswZ0tSW_tL8YAoMTXZUfFTNMhX8uvhtZx1QWG0g1zw"

ADVANCE_NO_TEMPLATE_ID = "1m5ee1ncDlq5p5Ce_nWggqFu2VEqdT_SeY-s2BsEkUGk"

# 2. Who in HR receives the audit email, and where the PDF is exported locally

# before being attached.

HR_EMAIL_ADDRESS = "veronicabeyton@gmail.com"

LOCAL_PDF_PATH = Path(

    os.environ.get(

        "LOCAL_PDF_PATH",

        str(_SCRIPT_DIR / "temp_report.pdf")

    )

)

# 3. Sheet tab holding form responses, timestamp header, and the policy rulebook.

SHEET_RESPONSE_TAB = "Form responses 1"

HEADER_TIMESTAMP = "Timestamp"

HR_POLICY_FILE = Path(__file__).parent / "hr_policy.txt"

# 4. Continuous polling + failure backoff.

PROCESSED_COLUMN_HEADER = "Processed"

POLL_INTERVAL_SECONDS = int(

    os.environ.get("POLL_INTERVAL_SECONDS", "30")

)

# Cap on how long a repeatedly-failing row is left alone between retries.

MAX_RETRY_BACKOFF_SECONDS = int(

    os.environ.get("MAX_RETRY_BACKOFF_SECONDS", str(30 * 60))

)

# Google Sheet header (left) -> template placeholder key (right).

#

# NOTE:

# "Total expenses" and "Closing balance" may no longer exist in the Google

# Form/Sheet. They remain in this mapping because the corresponding

# placeholders still exist in the Google Doc template. Their values are

# injected programmatically by compute_financial_reconciliation().

HEADER_MAPPING: dict[str, str] = {

    "Employee's position": "Employee_Position",

    "Employee's name and surname": "Employee_NameSurn",

    "Manager's position": "Manager_Position",

    "Manager's name and surname": "Manager_NameSurn",

    "Did you receive an advance?": "AdvanceYesNo",

    "Total expenses": "Total_Expenses",

    "Closing balance": "Closing_Balance",

    "Advance amount": "Advance_Amount",

    "Advance date": "Advance_Date",

    "Receipt No.1 date": "Receipt1_Date",

    "Receipt No.1 company": "Receipt1_Company",

    "Receipt No.1 number": "Receipt1_Num",

    "Receipt No.1 amount": "Receipt1_Amount",

    "Receipt No.2 date": "Receipt2_Date",

    "Receipt No.2 company": "Receipt2_Company",

    "Receipt No.2 number": "Receipt2_Num",

    "Receipt No.2 amount": "Receipt2_Amount",

    "Receipt No.3 date": "Receipt3_Date",

    "Receipt No.3 company": "Receipt3_Company",

    "Receipt No.3 number": "Receipt3_Num",

    "Receipt No.3 amount": "Receipt3_Amount",

}

DATE_KEYS = {k for k in HEADER_MAPPING.values() if "Date" in k}

AMOUNT_KEYS = {

    k

    for k in HEADER_MAPPING.values()

    if any(tok in k for tok in ("Amount", "Balance", "Expenses"))

}

RECEIPT_NUMBERS = (1, 2, 3)

# Service account: Sheets ONLY. It must never hold Docs/Drive scopes -- see

# the module docstring for why.

SHEETS_SCOPES = [

    "https://www.googleapis.com/auth/spreadsheets",

]

# OAuth (human account): Docs + Drive, since this identity is the one that

# actually owns/creates the temporary Doc copy and therefore has real quota

# for it.

OAUTH_SCOPES = [

    "https://www.googleapis.com/auth/documents",

    "https://www.googleapis.com/auth/drive",

]

logging.basicConfig(

    level=logging.INFO,

    format="%(asctime)s [%(levelname)s] %(message)s"

)

log = logging.getLogger("agent_workflow")

class Verdict(str):

    APPROVED = "APPROVED"

    REJECTED = "REJECTED"

    NEEDS_HUMAN_REVIEW = "NEEDS HUMAN REVIEW"

class AuditVerdictSchema(BaseModel):

    """Structured-output schema handed to the Gemini API directly, so the

    model is constrained to return exactly this shape instead of relying on

    regex-stripping markdown fences from free-form text."""

    verdict: str

    reasoning: str

    compliance_notes: str

@dataclass

class AuditResult:

    verdict: str = Verdict.NEEDS_HUMAN_REVIEW

    reasoning: str = ""

    compliance_notes: str = ""

    raw: dict[str, Any] = field(default_factory=dict)

@dataclass

class Reconciliation:

    receipts_total: float

    reported_total_expenses: float | None

    advance_amount: float | None

    reported_closing_balance: float | None

    expected_closing_balance: float | None

    discrepancies: list[str] = field(default_factory=list)

    def as_text(self) -> str:

        lines = [

            f"Sum of submitted receipts: {self.receipts_total:.2f}"

        ]

        if self.advance_amount is not None:

            lines.append(

                f"Advance received:          {self.advance_amount:.2f}"

            )

        if self.expected_closing_balance is not None:

            lines.append(

                f"Expected closing balance:  {self.expected_closing_balance:.2f}"

            )

        if self.discrepancies:

            lines.append("Discrepancies found:")

            lines.extend(

                f"  - {d}" for d in self.discrepancies

            )

        else:

            lines.append("No numerical discrepancies found.")

        return "\n".join(lines)

# --------------------------------------------------------------------------

# RETRY / COOLDOWN TRACKING (in-memory only -- the Sheet stays the source of

# truth for what's actually PROCESSED; this just prevents hammering a row

# that is failing every single poll).

# --------------------------------------------------------------------------

class RowRetryTracker:

    def __init__(self, base_seconds: int, max_seconds: int) -> None:

        self._base_seconds = base_seconds

        self._max_seconds = max_seconds

        self._fail_counts: dict[int, int] = {}

        self._next_attempt_at: dict[int, float] = {}

    def should_skip(self, row_number: int) -> bool:

        next_attempt = self._next_attempt_at.get(row_number)

        return next_attempt is not None and time.time() < next_attempt

    def seconds_until_retry(self, row_number: int) -> float:

        next_attempt = self._next_attempt_at.get(row_number, 0.0)

        return max(0.0, next_attempt - time.time())

    def record_failure(self, row_number: int) -> None:

        fail_count = self._fail_counts.get(row_number, 0) + 1

        self._fail_counts[row_number] = fail_count

        backoff = min(

            self._base_seconds * (2 ** (fail_count - 1)),

            self._max_seconds

        )

        self._next_attempt_at[row_number] = time.time() + backoff

        log.info(

            "Row %s has now failed %d time(s); next retry in %ds.",

            row_number,

            fail_count,

            backoff

        )

    def record_success(self, row_number: int) -> None:

        self._fail_counts.pop(row_number, None)

        self._next_attempt_at.pop(row_number, None)

# --------------------------------------------------------------------------

# AUTH: SERVICE ACCOUNT (Sheets only)

# --------------------------------------------------------------------------

_REQUIRED_SERVICE_ACCOUNT_FIELDS = (

    "type",

    "project_id",

    "private_key_id",

    "private_key",

    "client_email",

    "client_id",

    "token_uri",

)

def _validate_service_account_json(

    service_account_file: str

) -> dict[str, Any]:

    """Structurally validate the credentials file without exposing secrets."""

    path = Path(service_account_file)

    if not path.exists():

        raise SystemExit(

            f"Service account credentials file not found: {path}\n"

            f"Set GOOGLE_SERVICE_ACCOUNT_FILE in .env, or place credentials.json "

            f"next to this script."

        )

    try:

        raw = path.read_text(encoding="utf-8")

    except OSError as exc:

        raise SystemExit(

            f"Could not read credentials file {path}: {exc}"

        ) from exc

    try:

        data = json.loads(raw)

    except json.JSONDecodeError as exc:

        raise SystemExit(

            f"The credentials file at {path} is not valid JSON ({exc}).\n"

            f"It may have been partially downloaded, edited, or corrupted."

        ) from exc

    if not isinstance(data, dict):

        raise SystemExit(

            f"The credentials file at {path} is not a JSON object."

        )

    missing = [

        f for f in _REQUIRED_SERVICE_ACCOUNT_FIELDS

        if not data.get(f)

    ]

    if missing:

        raise SystemExit(

            "The credentials file does not appear to be a valid Google "

            f'service-account key. Missing/empty field(s): {", ".join(missing)}.\n'

            f"File: {path}"

        )

    if data.get("type") != "service_account":

        raise SystemExit(

            f'The credentials file at {path} has type={data.get("type")!r}, '

            f'expected type="service_account". This looks like the wrong kind '

            f"of Google credentials file."

        )

    private_key = data.get("private_key", "")

    if (

        "BEGIN PRIVATE KEY" not in private_key

        or "END PRIVATE KEY" not in private_key

    ):

        raise SystemExit(

            f"The private_key field in {path} does not look like a well-formed "

            f"PEM key. Re-download a fresh key JSON from Google Cloud Console."

        )

    return data

def _preflight_google_auth(

    creds: service_account.Credentials,

    client_email: str

) -> None:

    """Force a real token exchange before the polling loop starts."""

    log.info("[Sheets] Authenticating service account with Google...")

    try:

        creds.refresh(GoogleAuthRequest())

    except RefreshError as exc:

        reason = exc.args[0] if exc.args else str(exc)

        raise SystemExit(

            "\n"

            "ERROR: [Sheets] Google service-account authentication failed.\n"

            f"Reason: {reason}\n"

            "\n"

            "Check:\n"

            "1. credentials.json is the correct, CURRENT service-account key\n"

            "   (not a stale one left over from a regenerated key).\n"

            "2. That key is ACTIVE in Google Cloud Console.\n"

            f"3. The system clock is correct (service account: {client_email}).\n"

            "4. GOOGLE_SERVICE_ACCOUNT_FILE points to the intended file.\n"

            "\n"

            "This is a credentials/configuration problem that this script\n"

            "cannot self-repair. Application stopped."

        ) from exc

    log.info("[Sheets] Service account authentication successful.")

def get_sheets_service(

    service_account_file: str = SERVICE_ACCOUNT_FILE

) -> Any:

    """Validate + authenticate as the service account and build the Sheets

    API client. This identity touches Sheets ONLY -- see module docstring."""

    log.info("[Sheets] Loading service-account credentials...")

    data = _validate_service_account_json(service_account_file)

    client_email = data.get("client_email", "(unknown)")

    log.info(

        "[Sheets] Using service account credentials: %s",

        Path(service_account_file).resolve()

    )

    log.info("[Sheets] Service account: %s", client_email)

    creds = service_account.Credentials.from_service_account_file(

        service_account_file,

        scopes=SHEETS_SCOPES

    )

    _preflight_google_auth(creds, client_email)

    return build("sheets", "v4", credentials=creds)

# --------------------------------------------------------------------------

# AUTH: OAUTH (human account) FOR DOCS + DRIVE

# --------------------------------------------------------------------------

def _load_cached_oauth_credentials() -> OAuthCredentials | None:

    token_path = Path(OAUTH_TOKEN_FILE)

    if not token_path.exists():

        return None

    try:

        return OAuthCredentials.from_authorized_user_file(

            str(token_path), OAUTH_SCOPES

        )

    except (ValueError, json.JSONDecodeError) as exc:

        log.warning(

            "[Docs/Drive] Could not parse cached OAuth token at %s (%s); "

            "a fresh browser authorization will be requested.",

            token_path,

            exc

        )

        return None

def _save_oauth_credentials(creds: OAuthCredentials) -> None:

    Path(OAUTH_TOKEN_FILE).write_text(creds.to_json(), encoding="utf-8")

def get_oauth_credentials() -> OAuthCredentials:

    """Return valid OAuth user credentials for Docs + Drive, refreshing or

    running the one-time interactive browser consent flow as needed. The

    resulting token (including its refresh token) is cached locally so this

    interactive step normally only happens once per machine."""

    creds = _load_cached_oauth_credentials()

    if creds and creds.valid:

        return creds

    if creds and creds.expired and creds.refresh_token:

        log.info("[Docs/Drive] Refreshing cached OAuth token...")

        try:

            creds.refresh(GoogleAuthRequest())

            _save_oauth_credentials(creds)

            log.info("[Docs/Drive] OAuth token refreshed.")

            return creds

        except RefreshError as exc:

            log.warning(

                "[Docs/Drive] Cached OAuth token could not be refreshed "

                "(%s); a fresh browser authorization will be requested.",

                exc

            )

    if not Path(OAUTH_CLIENT_SECRET_FILE).exists():

        raise SystemExit(

            "\n"

            "ERROR: [Docs/Drive] OAuth client secret file not found: "

            f"{OAUTH_CLIENT_SECRET_FILE}\n"

            "Create a Desktop-app OAuth Client ID in Google Cloud Console, "

            "download its JSON, and save it there (or point "

            "GOOGLE_OAUTH_CLIENT_SECRET_FILE at it). See the setup steps in "

            "this script's module docstring."

        )

    log.info(

        "[Docs/Drive] No valid cached token found -- opening a browser "

        "window for one-time authorization. Log in as the Google account "

        "that owns the Docs templates."

    )

    flow = InstalledAppFlow.from_client_secrets_file(

        OAUTH_CLIENT_SECRET_FILE, OAUTH_SCOPES

    )

    creds = flow.run_local_server(port=0)

    _save_oauth_credentials(creds)

    log.info(

        "[Docs/Drive] OAuth authorization complete; token cached at %s.",

        OAUTH_TOKEN_FILE

    )

    return creds

def get_docs_drive_services() -> dict[str, Any]:

    creds = get_oauth_credentials()

    return {

        "docs": build("docs", "v1", credentials=creds),

        "drive": build("drive", "v3", credentials=creds),

    }

def get_google_services() -> dict[str, Any]:

    """Build every Google API client the workflow needs, using the right

    identity for each: service account for Sheets, OAuth human account for

    Docs/Drive."""

    services = {"sheets": get_sheets_service()}

    services.update(get_docs_drive_services())

    return services

# --------------------------------------------------------------------------

# GEMINI

# --------------------------------------------------------------------------

def configure_gemini() -> genai.Client:

    return genai.Client(api_key=GEMINI_API_KEY)

def load_hr_policy(path: Path = HR_POLICY_FILE) -> str:

    if not path.exists():

        log.warning(

            "%s not found; auditing with no policy text (not recommended).",

            path

        )

        return "(No written policy provided.)"

    return path.read_text(encoding="utf-8").strip()

# --------------------------------------------------------------------------

# STEP 1: READ THE SHEET AND FIND UNPROCESSED ROWS

# --------------------------------------------------------------------------

def _clean(s: str) -> str:

    return re.sub(r"\s+", " ", str(s)).strip()

def get_column_letter(col_index: int) -> str:

    """1-indexed column number -> A1-notation column letters."""

    letters = ""

    while col_index > 0:

        col_index, remainder = divmod(col_index - 1, 26)

        letters = chr(65 + remainder) + letters

    return letters

def read_all_rows(

    sheets_service,

    spreadsheet_id: str,

    tab_name: str

) -> tuple[list[str], list[list[str]]]:

    """Return (headers, data_rows)."""

    try:

        result = (

            sheets_service.spreadsheets()

            .values()

            .get(

                spreadsheetId=spreadsheet_id,

                range=f"'{tab_name}'"

            )

            .execute()

        )

    except Exception as exc:  # noqa: BLE001

        raise RuntimeError(f"[Sheets] Failed to read sheet: {exc}") from exc

    rows = result.get("values", [])

    if not rows:

        return [], []

    return rows[0], rows[1:]

def ensure_processed_column(

    sheets_service,

    spreadsheet_id: str,

    tab_name: str,

    headers: list[str]

) -> int:

    """Return the 1-indexed Processed column, creating it if necessary."""

    clean_headers = [_clean(h) for h in headers]

    if PROCESSED_COLUMN_HEADER in clean_headers:

        return clean_headers.index(PROCESSED_COLUMN_HEADER) + 1

    col_index = len(headers) + 1

    col_letter = get_column_letter(col_index)

    try:

        (

            sheets_service.spreadsheets()

            .values()

            .update(

                spreadsheetId=spreadsheet_id,

                range=f"'{tab_name}'!{col_letter}1",

                valueInputOption="RAW",

                body={"values": [[PROCESSED_COLUMN_HEADER]]},

            )

            .execute()

        )

    except Exception as exc:  # noqa: BLE001

        raise RuntimeError(

            f"[Sheets] Failed to add '{PROCESSED_COLUMN_HEADER}' header: {exc}"

        ) from exc

    log.info(

        "[Sheets] Added missing '%s' header column at %s1",

        PROCESSED_COLUMN_HEADER,

        col_letter

    )

    return col_index

def find_unprocessed_rows(

    data_rows: list[list[str]],

    processed_col_index: int

) -> list[tuple[int, list[str]]]:

    """Return every row whose Processed cell is blank."""

    unprocessed = []

    for i, row in enumerate(data_rows):

        row_number = i + 2

        cell_value = (

            row[processed_col_index - 1].strip()

            if len(row) >= processed_col_index

            else ""

        )

        if not cell_value:

            unprocessed.append((row_number, row))

    return unprocessed

def mark_row_processed(

    sheets_service,

    spreadsheet_id: str,

    tab_name: str,

    row_number: int,

    processed_col_index: int

) -> None:

    col_letter = get_column_letter(processed_col_index)

    try:

        (

            sheets_service.spreadsheets()

            .values()

            .update(

                spreadsheetId=spreadsheet_id,

                range=f"'{tab_name}'!{col_letter}{row_number}",

                valueInputOption="RAW",

                body={"values": [["PROCESSED"]]},

            )

            .execute()

        )

    except Exception as exc:  # noqa: BLE001

        raise RuntimeError(

            f"[Sheets] Failed to mark row {row_number} PROCESSED: {exc}"

        ) from exc

# --------------------------------------------------------------------------

# STEP 2: MAP RAW SHEET VALUES -> TEMPLATE PLACEHOLDER DATA

# --------------------------------------------------------------------------

def map_headers_to_report_data(

    headers: list[str],

    row: list[str]

) -> dict[str, str]:

    """Map Google Sheet headers to Google Doc template placeholders.

    Total_Expenses and Closing_Balance are deliberately allowed to remain

    blank at this stage because those Form questions have been removed.

    They are calculated later by compute_financial_reconciliation().

    """

    if len(row) < len(headers):

        row = row + [""] * (len(headers) - len(row))

    clean_headers = [_clean(h) for h in headers]

    report_data: dict[str, str] = {}

    for long_header, short_key in HEADER_MAPPING.items():

        target = _clean(long_header)

        value = " "

        if target in clean_headers:

            raw_value = row[clean_headers.index(target)]

            if raw_value not in (None, "", " "):

                if short_key in DATE_KEYS:

                    value = _format_date(raw_value)

                elif short_key in AMOUNT_KEYS:

                    value = _format_amount(raw_value)

                else:

                    value = str(raw_value)

        elif short_key == "Employee_NameSurn":

            raise ValueError(

                f"Critical column header not found: {long_header}"

            )

        report_data[short_key] = value

    if HEADER_TIMESTAMP in clean_headers:

        raw_ts = row[clean_headers.index(HEADER_TIMESTAMP)]

        report_data["Report_Date"] = (

            _format_date(raw_ts) if raw_ts else "N/A"

        )

    else:

        report_data["Report_Date"] = "N/A"

    return report_data

def _format_date(raw_value: str) -> str:

    """Normalize common Google Sheets date/time formats."""

    raw_value = str(raw_value).strip()

    for fmt in (

        "%m/%d/%Y %H:%M:%S",

        "%m/%d/%Y",

        "%Y-%m-%d",

        "%d.%m.%Y"

    ):

        try:

            return datetime.strptime(

                raw_value,

                fmt

            ).strftime("%d.%m.%Y")

        except ValueError:

            continue

    return raw_value

def _parse_amount(value: str | int | float | None) -> float | None:
    """Parse money values from dot or comma decimal formats."""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    text = re.sub(r"[^0-9,.-]", "", text)
    if not text or text in {"-", ".", ","}:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        if text.count(",") == 1 and len(text.rsplit(",", 1)[1]) in (1, 2):
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")
    elif text.count(".") > 1:
        parts = text.split(".")
        text = "".join(parts[:-1]) + "." + parts[-1]

    try:
        return float(text)
    except ValueError:
        return None


def _format_amount(raw_value: str) -> str:
    parsed = _parse_amount(raw_value)
    return f"{parsed:.2f}" if parsed is not None else str(raw_value).strip()


def _to_float(value: str | None) -> float | None:
    return _parse_amount(value)


# STEP 3: FINANCIAL RECONCILIATION

#

# IMPORTANT:

# The Google Form no longer asks for "Total expenses" or "Closing balance".

# Therefore those values MUST NOT be expected from the Sheet.

#

# Python calculates:

#   Total_Expenses  = sum of submitted receipt amounts

#   Closing_Balance = Advance_Amount - Total_Expenses

#

# These calculated values are written directly into report_data so they exist

# BEFORE either Gemini or generate_pdf_report() is called.

# --------------------------------------------------------------------------

def get_present_receipts(
    report_data: dict[str, str]
) -> list[dict[str, str]]:
    """Return only receipt slots that contain actual receipt information.

    Receipt #1/#2/#3 are optional. An entirely unused slot is ignored. A slot
    is considered a real receipt only when company, receipt number, or amount
    is filled. The date alone does not activate a receipt slot because a form
    or default value may populate it even when no receipt was submitted.
    Once a slot is active, missing fields are reported as incomplete.
    """
    present = []
    for n in RECEIPT_NUMBERS:
        amount = str(report_data.get(f"Receipt{n}_Amount", "")).strip()
        date = str(report_data.get(f"Receipt{n}_Date", "")).strip()
        company = str(report_data.get(f"Receipt{n}_Company", "")).strip()
        number = str(report_data.get(f"Receipt{n}_Num", "")).strip()

        receipt_is_present = any((company, number, amount))

        if receipt_is_present:
            present.append({
                "receipt_number": n,
                "date": date or "(missing)",
                "company": company or "(missing)",
                "receipt_no": number or "(missing)",
                "amount": amount or "(missing)",
            })
    return present


def compute_financial_reconciliation(

    report_data: dict[str, str],

    present_receipts: list[dict[str, str]]

) -> Reconciliation:

    """Calculate and inject the financial totals into report_data.

    Total expenses is the deterministic sum of all valid submitted receipt

    amounts.

    Closing balance is:

        advance amount - calculated total expenses

    The calculated values are formatted to exactly two decimal places and

    stored in report_data under the exact Google Docs placeholder keys:

        report_data["Total_Expenses"]

        report_data["Closing_Balance"]

    This function therefore MUST run before Gemini and before

    generate_pdf_report(). Gemini never performs this arithmetic itself --

    it is only told the result and asked to treat it as ground truth.

    """

    # ------------------------------------------------------------------

    # 1. Calculate the sum of all submitted receipt amounts.

    # ------------------------------------------------------------------

    receipts_total = 0.0

    for receipt in present_receipts:

        amount = _to_float(receipt.get("amount"))

        if amount is not None:

            receipts_total += amount

    # Round at the reconciliation boundary to avoid floating-point artifacts

    # such as 149.9999999997 appearing in downstream output.

    receipts_total = round(receipts_total, 2)

    # ------------------------------------------------------------------

    # 2. Read the advance amount.

    # ------------------------------------------------------------------

    advance_received = _to_float(

        report_data.get("Advance_Amount")

    )

    # ------------------------------------------------------------------

    # 3. Calculate the closing balance.

    #

    # If no valid advance amount was submitted, there is no meaningful

    # numerical closing balance to calculate. In that case the placeholder

    # receives a blank space rather than the string "None".

    # ------------------------------------------------------------------

    if advance_received is not None:

        expected_closing_balance = round(

            advance_received - receipts_total,

            2

        )

    else:

        expected_closing_balance = None

    # ------------------------------------------------------------------

    # 4. INJECT THE CALCULATED VALUES INTO report_data.

    #

    # generate_pdf_report() and the Gemini audit both receive this SAME

    # dictionary later, so these values are available to replace

    # {{Total_Expenses}} and {{Closing_Balance}} in the Google Docs template.

    # ------------------------------------------------------------------

    report_data["Total_Expenses"] = f"{receipts_total:.2f}"

    if expected_closing_balance is not None:

        report_data["Closing_Balance"] = (

            f"{expected_closing_balance:.2f}"

        )

    else:

        report_data["Closing_Balance"] = " "

    discrepancies: list[str] = []

    # ------------------------------------------------------------------

    # 5. Validate receipt completeness.

    # ------------------------------------------------------------------

    for receipt in present_receipts:

        missing = [

            key

            for key in ("date", "company", "receipt_no", "amount")

            if receipt[key] == "(missing)"

        ]

        if missing:

            discrepancies.append(

                f"Receipt #{receipt['receipt_number']} is missing: "

                f"{', '.join(missing)}."

            )

    return Reconciliation(

        receipts_total=receipts_total,

        # There is no longer a user-reported total. The calculated value is

        # now the authoritative total used by the workflow.

        reported_total_expenses=receipts_total,

        advance_amount=advance_received,

        # There is no longer a user-reported closing balance. The calculated

        # value is authoritative.

        reported_closing_balance=expected_closing_balance,

        expected_closing_balance=expected_closing_balance,

        discrepancies=discrepancies,

    )

# --------------------------------------------------------------------------

# STEP 4: GEMINI AI AUDIT LAYER

# --------------------------------------------------------------------------

AUDIT_PROMPT_TEMPLATE = """\\

You are an HR expense-policy auditor. Evaluate this advance/expense report

submission strictly against the company policy below, using the pre-computed

financial reconciliation as ground truth for the math (do not redo the

arithmetic yourself -- trust the numbers given).

COMPANY HR POLICY:

{policy_text}

SUBMISSION FIELDS:

{data_json}

RECEIPTS SUBMITTED (only receipts with actual data are listed -- #2 and #3

are optional and should not be flagged as missing if absent):

{receipts_json}

PRE-COMPUTED FINANCIAL RECONCILIATION:

{reconciliation_text}

Evaluate receipt validity (date, vendor, receipt number, amount present and

sensible) and whether the submission complies with policy. Then choose one

verdict:

- "APPROVED": no policy violations or discrepancies found.

- "REJECTED": clear policy violation or unresolvable discrepancy.

- "NEEDS HUMAN REVIEW": ambiguous, borderline, or missing information that a

  human should judge.

"""

def evaluate_hr_policy_compliance(

    client: genai.Client,

    report_data: dict[str, str],

    present_receipts: list[dict[str, str]],

    policy_text: str,

    reconciliation: Reconciliation

) -> AuditResult:

    """Call Gemini to audit the submission.

    The calculated financial values are already present in report_data when

    this function runs. Uses the current google-genai client with a

    structured-output schema so the model is constrained to return valid

    JSON in the required shape, rather than relying on regex-stripping

    markdown fences from free-form text.

    If Gemini is unavailable or returns something unusable for any reason,

    this NEVER raises -- it safely falls back to NEEDS HUMAN REVIEW with an

    explanation, so a Gemini outage can never crash the workflow or block a

    row from otherwise being processed.

    """

    prompt = AUDIT_PROMPT_TEMPLATE.format(

        policy_text=policy_text,

        data_json=json.dumps(

            report_data,

            indent=2

        ),

        receipts_json=json.dumps(

            present_receipts,

            indent=2

        ),

        reconciliation_text=reconciliation.as_text(),

    )

    try:

        response = client.models.generate_content(

            model=GEMINI_MODEL_NAME,

            contents=prompt,

            config=genai_types.GenerateContentConfig(

                response_mime_type="application/json",

                response_schema=AuditVerdictSchema,

            ),

        )

        parsed_model = response.parsed

        if parsed_model is not None:

            parsed = parsed_model.model_dump()

        else:

            # Fallback: the SDK didn't auto-parse (e.g. truncated output).

            # response.text should still be raw JSON since response_mime_type

            # was set, but strip markdown fences defensively just in case.

            text = re.sub(

                r"^```(?:json)?|```$",

                "",

                (response.text or "").strip(),

                flags=re.MULTILINE

            ).strip()

            parsed = json.loads(text)

        verdict = parsed.get("verdict", Verdict.NEEDS_HUMAN_REVIEW)

        if verdict not in (

            Verdict.APPROVED,

            Verdict.REJECTED,

            Verdict.NEEDS_HUMAN_REVIEW,

        ):

            verdict = Verdict.NEEDS_HUMAN_REVIEW

        return AuditResult(

            verdict=verdict,

            reasoning=parsed.get("reasoning", ""),

            compliance_notes=parsed.get("compliance_notes", ""),

            raw=parsed,

        )

    except GenAIAPIError as exc:
        # Temporary Gemini outages are retried by the existing row-level
        # exponential backoff instead of becoming a final human-review result.
        status_code = getattr(exc, "code", None)
        if status_code is None:
            match = re.search(r"\b(429|500|502|503|504)\b", str(exc))
            status_code = int(match.group(1)) if match else None

        if status_code in {429, 500, 502, 503, 504}:
            log.warning(
                "[Gemini] Temporary API failure (%s): %s -- row will be retried.",
                status_code,
                exc,
            )
            raise RuntimeError(
                f"[Gemini] Temporary API failure ({status_code}): {exc}"
            ) from exc

        log.warning(
            "[Gemini] Audit failed (%s): %s -- falling back to %s.",
            type(exc).__name__,
            exc,
            Verdict.NEEDS_HUMAN_REVIEW
        )
        return AuditResult(
            verdict=Verdict.NEEDS_HUMAN_REVIEW,
            reasoning=(
                f"Automated audit failed ({type(exc).__name__}: {exc}); "
                f"please review manually."
            ),
            compliance_notes="",
        )

    except (json.JSONDecodeError, ValueError, KeyError) as exc:
        log.warning(
            "[Gemini] Audit returned unusable data (%s): %s -- falling back to %s.",
            type(exc).__name__,
            exc,
            Verdict.NEEDS_HUMAN_REVIEW
        )
        return AuditResult(
            verdict=Verdict.NEEDS_HUMAN_REVIEW,
            reasoning=(
                f"Automated audit returned unusable data ({type(exc).__name__}: {exc}); "
                f"please review manually."
            ),
            compliance_notes="",
        )

    except Exception as exc:  # noqa: BLE001

        # Broad catch is deliberate and final: no matter what goes wrong

        # inside the Gemini call, the workflow must degrade gracefully

        # instead of crashing the whole row.

        log.warning(

            "[Gemini] Unexpected audit failure (%s): %s -- falling back to %s.",

            type(exc).__name__,

            exc,

            Verdict.NEEDS_HUMAN_REVIEW

        )

        return AuditResult(

            verdict=Verdict.NEEDS_HUMAN_REVIEW,

            reasoning=(

                f"Automated audit failed unexpectedly ({type(exc).__name__}); "

                f"please review manually."

            ),

            compliance_notes="",

        )

# --------------------------------------------------------------------------

# STEP 5: TEMPLATE SELECTION

# --------------------------------------------------------------------------

def select_template(

    report_data: dict[str, str]

) -> str:

    advance_value = str(

        report_data.get("AdvanceYesNo", "")

    ).strip().lower()

    if "yes" in advance_value:

        return ADVANCE_YES_TEMPLATE_ID

    if "no" in advance_value:

        return ADVANCE_NO_TEMPLATE_ID

    raise ValueError(

        f"Advance status not recognized: "

        f"{report_data.get('AdvanceYesNo')!r}"

    )

# --------------------------------------------------------------------------

# STEP 6: FILL THE TEMPLATE AND EXPORT PDF (Docs + Drive, via OAuth identity)

# --------------------------------------------------------------------------

def replace_placeholders(

    docs_service,

    document_id: str,

    data: dict[str, str]

) -> None:

    """Replace every {{Key}} placeholder in the Google Doc."""

    requests = [

        {

            "replaceAllText": {

                "containsText": {

                    "text": f"{{{{{key}}}}}",

                    "matchCase": True

                },

                "replaceText": (

                    value

                    if value not in (None, "")

                    else " "

                ),

            }

        }

        for key, value in data.items()

    ]

    if requests:

        try:

            (

                docs_service.documents()

                .batchUpdate(

                    documentId=document_id,

                    body={"requests": requests}

                )

                .execute()

            )

        except Exception as exc:  # noqa: BLE001

            raise RuntimeError(

                f"[Google Docs] Failed to replace placeholders: {exc}"

            ) from exc

    _blank_out_unmapped_placeholders(

        docs_service,

        document_id

    )

def _blank_out_unmapped_placeholders(

    docs_service,

    document_id: str

) -> None:

    """Blank any leftover {{...}} placeholders."""

    try:

        doc = (

            docs_service.documents()

            .get(documentId=document_id)

            .execute()

        )

    except Exception as exc:  # noqa: BLE001

        raise RuntimeError(

            f"[Google Docs] Failed to read document back: {exc}"

        ) from exc

    leftover_text = _extract_document_text(doc)

    leftover_keys = set(

        re.findall(

            r"\{\{([^}]+)\}\}",

            leftover_text

        )

    )

    if not leftover_keys:

        return

    requests = [

        {

            "replaceAllText": {

                "containsText": {

                    "text": f"{{{{{key}}}}}",

                    "matchCase": True

                },

                "replaceText": " ",

            }

        }

        for key in leftover_keys

    ]

    try:

        (

            docs_service.documents()

            .batchUpdate(

                documentId=document_id,

                body={"requests": requests}

            )

            .execute()

        )

    except Exception as exc:  # noqa: BLE001

        raise RuntimeError(

            f"[Google Docs] Failed to blank leftover placeholders: {exc}"

        ) from exc

def _extract_document_text(

    doc: dict

) -> str:

    """Walk Docs API structural elements including table cells."""

    chunks: list[str] = []

    def walk(elements: list[dict]) -> None:

        for el in elements:

            if "paragraph" in el:

                for pe in el["paragraph"].get(

                    "elements",

                    []

                ):

                    text_run = pe.get("textRun")

                    if text_run:

                        chunks.append(

                            text_run.get(

                                "content",

                                ""

                            )

                        )

            elif "table" in el:

                for row in el["table"].get(

                    "tableRows",

                    []

                ):

                    for cell in row.get(

                        "tableCells",

                        []

                    ):

                        walk(

                            cell.get(

                                "content",

                                []

                            )

                        )

            elif "tableOfContents" in el:

                walk(

                    el["tableOfContents"].get(

                        "content",

                        []

                    )

                )

    walk(

        doc.get("body", {}).get(

            "content",

            []

        )

    )

    return "".join(chunks)

def generate_pdf_report(

    services: dict[str, Any],

    template_id: str,

    report_data: dict[str, str],

    employee_name: str

) -> tuple[Path, bytes, str]:

    """Copy template, fill placeholders, export PDF, then trash temp Doc.

    The copy/export/trash all run under the OAuth human identity in

    `services["drive"]` / `services["docs"]`, so the temporary Doc copy is

    created in a Drive that actually has storage quota -- this is the fix

    for the service-account `403 storageQuotaExceeded` error.

    """

    drive = services["drive"]

    docs = services["docs"]

    date_stamp = datetime.now().strftime("%d-%m-%Y")

    new_doc_name = (

        f"Advance Report - "

        f"{employee_name} - "

        f"{date_stamp}"

    )

    try:

        copy = (

            drive.files()

            .copy(

                fileId=template_id,

                body={"name": new_doc_name},

                fields="id"

            )

            .execute()

        )

    except Exception as exc:  # noqa: BLE001

        raise RuntimeError(

            f"[Google Drive] Failed to copy template {template_id}: {exc}"

        ) from exc

    temp_doc_id = copy["id"]

    try:

        # IMPORTANT:

        # report_data already contains the calculated Total_Expenses and

        # Closing_Balance values because compute_financial_reconciliation()

        # runs before this function.

        replace_placeholders(

            docs,

            temp_doc_id,

            report_data

        )

        try:

            pdf_bytes = (

                drive.files()

                .export(

                    fileId=temp_doc_id,

                    mimeType="application/pdf"

                )

                .execute()

            )

        except Exception as exc:  # noqa: BLE001

            raise RuntimeError(

                f"[Google Drive] Failed to export PDF: {exc}"

            ) from exc

        pdf_filename = f"{new_doc_name}.pdf"

        try:

            LOCAL_PDF_PATH.write_bytes(pdf_bytes)

        except OSError as exc:

            raise RuntimeError(

                f"[PDF generation] Failed to write local PDF "

                f"{LOCAL_PDF_PATH}: {exc}"

            ) from exc

        log.info(

            "[PDF generation] Exported clean PDF locally: %s",

            LOCAL_PDF_PATH

        )

        return (

            LOCAL_PDF_PATH,

            pdf_bytes,

            pdf_filename

        )

    finally:

        try:

            drive.files().update(

                fileId=temp_doc_id,

                body={"trashed": True}

            ).execute()

        except Exception as exc:  # noqa: BLE001

            # Don't let cleanup failure mask the real error, but do surface

            # it -- an un-trashed copy is exactly the kind of thing that

            # quietly eats into Drive storage over time.

            log.warning(

                "[Google Drive] Failed to trash temporary copy %s: %s",

                temp_doc_id,

                exc

            )

# --------------------------------------------------------------------------

# STEP 7: HR EMAIL

# --------------------------------------------------------------------------

def build_hr_email_body(

    employee_name: str,

    audit: AuditResult,

    reconciliation: Reconciliation

) -> str:

    lines = [

        (

            f"A new Advance Report has been submitted and processed "

            f"for {employee_name}."

        ),

        "",

        f"AI AUDIT VERDICT: {audit.verdict}",

        f"Reasoning: {audit.reasoning}",

    ]

    if audit.compliance_notes:

        lines += [

            "",

            f"Compliance notes: {audit.compliance_notes}"

        ]

    lines += [

        "",

        "FINANCIAL RECONCILIATION",

        reconciliation.as_text(),

        "",

        "The clean PDF report is attached to this email.",

        "",

        "Thank you.",

    ]

    return "\n".join(lines)

def send_hr_email(

    employee_name: str,

    pdf_bytes: bytes,

    pdf_filename: str,

    audit: AuditResult,

    reconciliation: Reconciliation

) -> None:

    message = MIMEMultipart()

    message["From"] = EMAIL_SENDER_ADDRESS

    message["To"] = HR_EMAIL_ADDRESS

    message["Subject"] = (

        f"[ADVANCE REPORT] "

        f"{audit.verdict} — "

        f"{employee_name}"

    )

    body = build_hr_email_body(

        employee_name,

        audit,

        reconciliation

    )

    message.attach(

        MIMEText(

            body,

            "plain"

        )

    )

    attachment = MIMEApplication(

        pdf_bytes,

        _subtype="pdf"

    )

    attachment.add_header(

        "Content-Disposition",

        "attachment",

        filename=pdf_filename

    )

    message.attach(attachment)

    try:

        with smtplib.SMTP_SSL(

            "smtp.gmail.com",

            465

        ) as smtp:

            smtp.login(

                EMAIL_SENDER_ADDRESS,

                EMAIL_SENDER_APP_PASSWORD

            )

            smtp.send_message(message)

        log.info(

            "[SMTP] HR audit email sent to %s",

            HR_EMAIL_ADDRESS

        )

    except smtplib.SMTPException as exc:

        # Re-raised (not swallowed): a failed send must prevent the row from

        # being marked PROCESSED.

        raise RuntimeError(f"[SMTP] Failed to send HR email: {exc}") from exc

# --------------------------------------------------------------------------

# ORCHESTRATION

# --------------------------------------------------------------------------

def process_submission(

    services: dict[str, Any],

    gemini_client: genai.Client,

    policy_text: str,

    headers: list[str],

    row: list[str],

    row_number: int,

    processed_col_index: int

) -> None:

    """Run the complete audit -> PDF -> email pipeline for one row.

    The row is marked PROCESSED only if every step through send_hr_email()

    completes without raising. Any exception here propagates to the caller

    (run_once), which records the failure against the retry tracker instead

    of marking the row processed.

    """

    report_data = map_headers_to_report_data(

        headers,

        row

    )

    employee_name = report_data[

        "Employee_NameSurn"

    ]

    log.info(

        "Processing submission for %s (row %s)",

        employee_name,

        row_number

    )

    present_receipts = get_present_receipts(

        report_data

    )

    # CRITICAL ORDER:

    #

    # compute_financial_reconciliation() calculates:

    #

    #   report_data["Total_Expenses"]

    #   report_data["Closing_Balance"]

    #

    # BEFORE either Gemini or generate_pdf_report() receives report_data.

    reconciliation = compute_financial_reconciliation(

        report_data,

        present_receipts

    )

    log.info(

        "Row %s calculated financials: "

        "Total_Expenses=%s, Closing_Balance=%s",

        row_number,

        report_data["Total_Expenses"],

        report_data["Closing_Balance"]

    )

    # Gemini audit failure never raises -- it safely falls back to

    # NEEDS HUMAN REVIEW inside evaluate_hr_policy_compliance().

    audit = evaluate_hr_policy_compliance(

        gemini_client,

        report_data,

        present_receipts,

        policy_text,

        reconciliation

    )

    log.info(

        "Row %s audit verdict: %s",

        row_number,

        audit.verdict

    )

    template_id = select_template(

        report_data

    )

    # The same report_data dictionary containing the calculated financial

    # values is passed directly to generate_pdf_report(). A failure here

    # raises and propagates up -- the row must NOT be marked PROCESSED.

    pdf_path, pdf_bytes, pdf_filename = generate_pdf_report(

        services,

        template_id,

        report_data,

        employee_name

    )

    try:

        # A failure here also raises and propagates up -- the row must NOT

        # be marked PROCESSED if the email never went out.

        send_hr_email(

            employee_name,

            pdf_bytes,

            pdf_filename,

            audit,

            reconciliation

        )

    finally:

        # Local PDF is scratch space only -- always clean it up, whether or

        # not the email send succeeded.

        try:

            pdf_path.unlink(

                missing_ok=True

            )

        except OSError as exc:

            log.warning(

                "Could not delete local temp PDF %s: %s",

                pdf_path,

                exc

            )

    mark_row_processed(

        services["sheets"],

        GOOGLE_SHEET_ID,

        SHEET_RESPONSE_TAB,

        row_number,

        processed_col_index

    )

    log.info(

        "Row %s done.",

        row_number

    )

def run_once(

    services: dict[str, Any],

    gemini_client: genai.Client,

    policy_text: str,

    retry_tracker: RowRetryTracker

) -> None:

    """One full pass over all unprocessed rows."""

    headers, data_rows = read_all_rows(

        services["sheets"],

        GOOGLE_SHEET_ID,

        SHEET_RESPONSE_TAB

    )

    if not headers:

        log.info(

            "Sheet is empty; nothing to do."

        )

        return

    processed_col_index = ensure_processed_column(

        services["sheets"],

        GOOGLE_SHEET_ID,

        SHEET_RESPONSE_TAB,

        headers

    )

    if len(headers) < processed_col_index:

        headers = headers + [

            PROCESSED_COLUMN_HEADER

        ]

    unprocessed = find_unprocessed_rows(

        data_rows,

        processed_col_index

    )

    if not unprocessed:

        log.info(

            "No unprocessed rows."

        )

        return

    ready_rows = []

    for row_number, row in unprocessed:

        if retry_tracker.should_skip(row_number):

            log.info(

                "Skipping row %s -- in cooldown for another %.0fs after a "

                "previous failure.",

                row_number,

                retry_tracker.seconds_until_retry(row_number)

            )

            continue

        ready_rows.append((row_number, row))

    if not ready_rows:

        log.info(

            "Found %d unprocessed row(s), all currently in retry cooldown.",

            len(unprocessed)

        )

        return

    log.info(

        "Found %d unprocessed row(s), %d ready to process now.",

        len(unprocessed),

        len(ready_rows)

    )

    for row_number, row in ready_rows:

        try:

            process_submission(

                services,

                gemini_client,

                policy_text,

                headers,

                row,

                row_number,

                processed_col_index

            )

            retry_tracker.record_success(row_number)

        except Exception as exc:  # noqa: BLE001

            log.error(

                "Failed to process row %s (%s): %s",

                row_number,

                type(exc).__name__,

                exc

            )

            retry_tracker.record_failure(row_number)

def run_forever() -> None:

    """Authenticate once, then continuously poll the sheet."""

    services = get_google_services()

    gemini_client = configure_gemini()

    policy_text = load_hr_policy()

    retry_tracker = RowRetryTracker(

        base_seconds=POLL_INTERVAL_SECONDS,

        max_seconds=MAX_RETRY_BACKOFF_SECONDS

    )

    log.info(

        "Starting continuous polling loop "

        "(every %ss). Press Ctrl+C to stop.",

        POLL_INTERVAL_SECONDS

    )

    while True:

        try:

            run_once(

                services,

                gemini_client,

                policy_text,

                retry_tracker

            )

        except Exception as exc:  # noqa: BLE001

            log.error(

                "Unexpected error during polling pass (%s): %s",

                type(exc).__name__,

                exc

            )

        try:

            time.sleep(

                POLL_INTERVAL_SECONDS

            )

        except KeyboardInterrupt:

            log.info(

                "Stopped by user."

            )

            return

if __name__ == "__main__":

    run_forever()