"""
agent_workflow.py

Agentic HR Workflow: reads new "Advance / Expense Report" form submissions
from a Google Sheet, has Gemini Flash run an AI audit against your HR policy
(`hr_policy.txt`), generates a clean PDF report from a Google Docs template
(template chosen by whether a previous advance was received), and emails HR
with the AI audit verdict + financial reconciliation in the body and the
clean PDF attached.

This is a Python port of an original Google Apps Script `onFormSubmit`
trigger. Apps Script runs automatically inside the bound Sheet; a local
Python script has no equivalent event trigger, so this script instead POLLS
for the latest row and uses a small local state file (`.last_row_state.json`)
to avoid reprocessing the same submission twice. Run it on a schedule (cron,
Cloud Scheduler + Cloud Run, GitHub Actions, etc.) to approximate the
"on form submit" behavior.

Setup
-----
1. pip install -r requirements.txt
2. Enable these APIs in your Google Cloud project: Sheets, Docs, Drive.
3. Create a Service Account, download its JSON key as `credentials.json`
   (same folder as this script), and SHARE these with the service account's
   `client_email` (found inside credentials.json) as Editor:
     - the Google Sheet (Viewer is enough)
     - both Google Docs templates
     - the Drive output folder
4. Write `hr_policy.txt` with your company's expense/advance policy in plain
   language — it's sent to Gemini as the audit rulebook.
5. Create a `.env` file (see `.env.example`) with:
       GEMINI_API_KEY=...
       GOOGLE_SHEET_ID=...
       EMAIL_SENDER_ADDRESS=...        # Gmail address the report is sent FROM
       EMAIL_SENDER_APP_PASSWORD=...   # Gmail app password (not your login password)
   Note: HR's inbox here (HR_EMAIL_ADDRESS below) is a personal Gmail
   address, not a Workspace mailbox, so sending via the Gmail API would
   require domain-wide delegation that a personal account can't grant. SMTP
   with an app password is the correct fit and needs no extra Cloud Console
   setup beyond enabling 2FA + generating the app password on the sender
   account.
6. Fill in the CONFIG section below with your template/folder IDs.
"""

from __future__ import annotations

import json
import logging
import os
import re
import smtplib
from dataclasses import dataclass, field
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaInMemoryUpload

import google.generativeai as genai

# --------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------

load_dotenv()

GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]
GOOGLE_SHEET_ID = os.environ["GOOGLE_SHEET_ID"]
EMAIL_SENDER_ADDRESS = os.environ["EMAIL_SENDER_ADDRESS"]
EMAIL_SENDER_APP_PASSWORD = os.environ["EMAIL_SENDER_APP_PASSWORD"]

SERVICE_ACCOUNT_FILE = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json")

# Stable, cost-effective Flash model as of mid-2026. Bump to "gemini-3.7-flash"
# for the most capable Flash variant if you want deeper audit reasoning.
GEMINI_MODEL_NAME = os.environ.get("GEMINI_MODEL_NAME", "gemini-3.5-flash")

# 1. Google Docs template IDs.
ADVANCE_YES_TEMPLATE_ID = "1PswZ0tSW_tL8YAoMTXZUfFTNMhX8uvhtZx1QWG0g1zw"
ADVANCE_NO_TEMPLATE_ID = "1m5ee1ncDlq5p5Ce_nWggqFu2VEqdT_SeY-s2BsEkUGk"

# 2. Drive folder where clean PDFs are stored, and who in HR receives the audit email.
OUTPUT_FOLDER_ID = "1IyRVV7mInrIFL6KLE8DJ6g4tkHVPRS0u"
HR_EMAIL_ADDRESS = "veronicabeyton@gmail.com"

# 3. Sheet tab holding form responses, timestamp header, and the policy rulebook.
SHEET_RESPONSE_TAB = "Form responses 1"
HEADER_TIMESTAMP = "Timestamp"
HR_POLICY_FILE = Path(__file__).parent / "hr_policy.txt"

# Google Sheet header (left) -> template placeholder key (right). These are
# the ONLY keys ever written into the Doc/PDF — the AI audit never adds keys
# here, so the exported PDF stays exactly as originally formatted.
HEADER_MAPPING: dict[str, str] = {
    "Employee's position": "Employee_Position",
    "Employee's name and surname": "Employee_NameSurn",
    "Manager's position": "Manager_Position",
    "Manager's name and surname": "Manager_NameSurn",
    "Did you receive an advance?": "AdvanceYesNo",  # "Previous Advance" Yes/No
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
    k for k in HEADER_MAPPING.values()
    if any(tok in k for tok in ("Amount", "Balance", "Expenses"))
}
RECEIPT_NUMBERS = (1, 2, 3)  # Receipts #2 and #3 are optional.

# Service accounts don't need the Gmail scope — HR email goes out over SMTP.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive",
]

STATE_FILE = Path(__file__).parent / ".last_row_state.json"

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
log = logging.getLogger("agent_workflow")


@dataclass
class AuditResult:
    verdict: str = "NEEDS HUMAN REVIEW"  # APPROVED | REJECTED | NEEDS HUMAN REVIEW
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
        lines = [f"Sum of submitted receipts: {self.receipts_total:.2f}"]
        if self.reported_total_expenses is not None:
            lines.append(f"Reported total expenses:   {self.reported_total_expenses:.2f}")
        if self.advance_amount is not None:
            lines.append(f"Advance received:          {self.advance_amount:.2f}")
        if self.expected_closing_balance is not None:
            lines.append(f"Expected closing balance:  {self.expected_closing_balance:.2f}")
        if self.reported_closing_balance is not None:
            lines.append(f"Reported closing balance:  {self.reported_closing_balance:.2f}")
        if self.discrepancies:
            lines.append("Discrepancies found:")
            lines.extend(f"  - {d}" for d in self.discrepancies)
        else:
            lines.append("No numerical discrepancies found.")
        return "\n".join(lines)


# --------------------------------------------------------------------------
# AUTH
# --------------------------------------------------------------------------

def get_google_services(service_account_file: str = SERVICE_ACCOUNT_FILE) -> dict[str, Any]:
    """Authenticate as the service account and build Sheets/Docs/Drive clients.
    Remember to share the Sheet, both Doc templates, and the output Drive
    folder with the service account's client_email (Editor access)."""
    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES
    )
    return {
        "sheets": build("sheets", "v4", credentials=creds),
        "docs": build("docs", "v1", credentials=creds),
        "drive": build("drive", "v3", credentials=creds),
    }


def configure_gemini() -> "genai.GenerativeModel":
    genai.configure(api_key=GEMINI_API_KEY)
    return genai.GenerativeModel(GEMINI_MODEL_NAME)


def load_hr_policy(path: Path = HR_POLICY_FILE) -> str:
    if not path.exists():
        log.warning("%s not found; auditing with no policy text (not recommended).", path)
        return "(No written policy provided.)"
    return path.read_text(encoding="utf-8").strip()


# --------------------------------------------------------------------------
# STEP 1: READ THE LATEST FORM SUBMISSION
# --------------------------------------------------------------------------

def _clean(s: str) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def read_latest_submission(sheets_service, spreadsheet_id: str, tab_name: str
                            ) -> tuple[list[str], list[str], int] | None:
    """Return (headers, last_row_values, row_number) for the sheet's last row,
    or None if the sheet has no data rows yet."""
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"'{tab_name}'")
        .execute()
    )
    rows = result.get("values", [])
    if len(rows) < 2:
        return None

    headers = rows[0]
    last_row_number = len(rows)  # 1-indexed, matches Sheets row numbering
    last_row = rows[-1]
    if len(last_row) < len(headers):
        last_row = last_row + [""] * (len(headers) - len(last_row))
    return headers, last_row, last_row_number


def has_new_submission(row_number: int) -> bool:
    if not STATE_FILE.exists():
        return True
    try:
        state = json.loads(STATE_FILE.read_text())
        return row_number > state.get("last_processed_row", 0)
    except (json.JSONDecodeError, OSError):
        return True


def mark_submission_processed(row_number: int) -> None:
    STATE_FILE.write_text(json.dumps({"last_processed_row": row_number}))


# --------------------------------------------------------------------------
# STEP 2: MAP RAW SHEET VALUES -> TEMPLATE PLACEHOLDER DATA
# --------------------------------------------------------------------------

def map_headers_to_report_data(headers: list[str], row: list[str]) -> dict[str, str]:
    """Mirrors the Apps Script header-lookup + date/number formatting logic.
    Every key produced here is a legitimate template placeholder — this dict
    (or a subset of it) is exactly what gets written into the PDF."""
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
            raise ValueError(f"Critical column header not found: {long_header}")
        report_data[short_key] = value

    if HEADER_TIMESTAMP in clean_headers:
        raw_ts = row[clean_headers.index(HEADER_TIMESTAMP)]
        report_data["Report_Date"] = _format_date(raw_ts) if raw_ts else "N/A"
    else:
        report_data["Report_Date"] = "N/A"

    return report_data


def _format_date(raw_value: str) -> str:
    """Sheets API returns dates as plain strings (no Date objects like Apps
    Script's bound SpreadsheetApp), so this normalizes common formats."""
    raw_value = str(raw_value).strip()
    for fmt in ("%m/%d/%Y %H:%M:%S", "%m/%d/%Y", "%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(raw_value, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue
    return raw_value


def _format_amount(raw_value: str) -> str:
    try:
        return f"{float(raw_value):.2f}"
    except (TypeError, ValueError):
        return str(raw_value)


def _to_float(value: str | None) -> float | None:
    if value is None:
        return None
    value = str(value).strip()
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# STEP 3: FINANCIAL RECONCILIATION (computed deterministically, not by the LLM)
# --------------------------------------------------------------------------

def get_present_receipts(report_data: dict[str, str]) -> list[dict[str, str]]:
    """Only include a receipt if it actually has data — receipts #2 and #3
    are optional and should never be audited (or flagged as missing) when blank."""
    present = []
    for n in RECEIPT_NUMBERS:
        amount = report_data.get(f"Receipt{n}_Amount", " ").strip()
        date = report_data.get(f"Receipt{n}_Date", " ").strip()
        company = report_data.get(f"Receipt{n}_Company", " ").strip()
        number = report_data.get(f"Receipt{n}_Num", " ").strip()
        if any(field_val not in ("", " ") for field_val in (amount, date, company, number)):
            present.append({
                "receipt_number": n,
                "date": date or "(missing)",
                "company": company or "(missing)",
                "receipt_no": number or "(missing)",
                "amount": amount or "(missing)",
            })
    return present


def compute_financial_reconciliation(report_data: dict[str, str],
                                      present_receipts: list[dict[str, str]]) -> Reconciliation:
    receipts_total = sum(
        _to_float(r["amount"]) or 0.0
        for r in present_receipts
        if _to_float(r["amount"]) is not None
    )
    reported_total_expenses = _to_float(report_data.get("Total_Expenses"))
    advance_received = _to_float(report_data.get("Advance_Amount"))
    reported_closing_balance = _to_float(report_data.get("Closing_Balance"))

    expected_closing_balance = None
    discrepancies: list[str] = []

    if reported_total_expenses is not None:
        if abs(receipts_total - reported_total_expenses) > 0.01:
            discrepancies.append(
                f"Submitted receipts sum to {receipts_total:.2f}, but reported total "
                f"expenses is {reported_total_expenses:.2f}."
            )

    if advance_received is not None and reported_total_expenses is not None:
        expected_closing_balance = advance_received - reported_total_expenses
        if reported_closing_balance is not None and abs(
            expected_closing_balance - reported_closing_balance
        ) > 0.01:
            discrepancies.append(
                f"Expected closing balance is {expected_closing_balance:.2f} "
                f"(advance {advance_received:.2f} - expenses {reported_total_expenses:.2f}), "
                f"but reported closing balance is {reported_closing_balance:.2f}."
            )

    for r in present_receipts:
        missing = [k for k in ("date", "company", "receipt_no", "amount") if r[k] == "(missing)"]
        if missing:
            discrepancies.append(f"Receipt #{r['receipt_number']} is missing: {', '.join(missing)}.")

    return Reconciliation(
        receipts_total=receipts_total,
        reported_total_expenses=reported_total_expenses,
        advance_amount=advance_received,
        reported_closing_balance=reported_closing_balance,
        expected_closing_balance=expected_closing_balance,
        discrepancies=discrepancies,
    )


# --------------------------------------------------------------------------
# STEP 4: GEMINI AI AUDIT LAYER  (runs BEFORE the report/email is finalized)
# --------------------------------------------------------------------------

AUDIT_PROMPT_TEMPLATE = """\
You are an HR expense-policy auditor. Evaluate this advance/expense report
submission strictly against the company policy below, using the pre-computed
financial reconciliation as ground truth for the math (do not redo the
arithmetic yourself — trust the numbers given).

COMPANY HR POLICY:
{policy_text}

SUBMISSION FIELDS:
{data_json}

RECEIPTS SUBMITTED (only receipts with actual data are listed — #2 and #3
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

Respond with ONLY a JSON object, no markdown fences, no extra text, in
exactly this shape:
{{
  "verdict": "APPROVED" | "REJECTED" | "NEEDS HUMAN REVIEW",
  "reasoning": "2-4 sentences explaining the verdict, referencing specific receipts or numbers",
  "compliance_notes": "any policy-specific notes HR should see (or empty string if none)"
}}
"""


def evaluate_hr_policy_compliance(model: "genai.GenerativeModel",
                                   report_data: dict[str, str],
                                   present_receipts: list[dict[str, str]],
                                   policy_text: str,
                                   reconciliation: Reconciliation) -> AuditResult:
    """Calls Gemini Flash to audit the submission. This result is used ONLY for
    the HR email body — it is never written into the Doc/PDF template."""
    prompt = AUDIT_PROMPT_TEMPLATE.format(
        policy_text=policy_text,
        data_json=json.dumps(report_data, indent=2),
        receipts_json=json.dumps(present_receipts, indent=2),
        reconciliation_text=reconciliation.as_text(),
    )
    try:
        response = model.generate_content(prompt)
        text = re.sub(r"^```(?:json)?|```$", "", response.text.strip(), flags=re.MULTILINE).strip()
        parsed = json.loads(text)
        verdict = parsed.get("verdict", "NEEDS HUMAN REVIEW")
        if verdict not in ("APPROVED", "REJECTED", "NEEDS HUMAN REVIEW"):
            verdict = "NEEDS HUMAN REVIEW"
        return AuditResult(
            verdict=verdict,
            reasoning=parsed.get("reasoning", ""),
            compliance_notes=parsed.get("compliance_notes", ""),
            raw=parsed,
        )
    except Exception as exc:  # noqa: BLE001 - degrade to human review, never silently drop
        log.warning("Gemini audit failed: %s", exc)
        return AuditResult(
            verdict="NEEDS HUMAN REVIEW",
            reasoning=f"Automated audit failed ({exc}); please review manually.",
            compliance_notes="",
        )


# --------------------------------------------------------------------------
# STEP 5: TEMPLATE SELECTION (unchanged logic)
# --------------------------------------------------------------------------

def select_template(report_data: dict[str, str]) -> str:
    advance_value = str(report_data.get("AdvanceYesNo", "")).strip().lower()
    if "yes" in advance_value:
        return ADVANCE_YES_TEMPLATE_ID
    if "no" in advance_value:
        return ADVANCE_NO_TEMPLATE_ID
    raise ValueError(f"Advance status not recognized: {report_data.get('AdvanceYesNo')!r}")


# --------------------------------------------------------------------------
# STEP 6: FILL THE TEMPLATE (Docs API) AND EXPORT A CLEAN PDF (Drive API)
# --------------------------------------------------------------------------

def replace_placeholders(docs_service, document_id: str, data: dict[str, str]) -> None:
    """Replace every {{Key}} placeholder — including inside tables, which
    replaceAllText already searches — with its mapped value. `data` here is
    ONLY the standard report fields; no AI/audit keys are ever passed in, so
    the exported PDF keeps its original formatting exactly."""
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": f"{{{{{key}}}}}", "matchCase": True},
                "replaceText": (value if value not in (None, "") else " "),
            }
        }
        for key, value in data.items()
    ]
    if requests:
        docs_service.documents().batchUpdate(
            documentId=document_id, body={"requests": requests}
        ).execute()
    _blank_out_unmapped_placeholders(docs_service, document_id)


def _blank_out_unmapped_placeholders(docs_service, document_id: str) -> None:
    """Second pass: blank out any leftover {{...}} with no mapped value
    (mirrors the Apps Script fallback), keeping the PDF free of stray tags."""
    doc = docs_service.documents().get(documentId=document_id).execute()
    leftover_text = _extract_document_text(doc)
    leftover_keys = set(re.findall(r"\{\{([^}]+)\}\}", leftover_text))
    if not leftover_keys:
        return
    requests = [
        {
            "replaceAllText": {
                "containsText": {"text": f"{{{{{key}}}}}", "matchCase": True},
                "replaceText": " ",
            }
        }
        for key in leftover_keys
    ]
    docs_service.documents().batchUpdate(
        documentId=document_id, body={"requests": requests}
    ).execute()


def _extract_document_text(doc: dict) -> str:
    """Walks Docs API structural elements (including table cells) so leftover
    placeholders can be found by regex."""
    chunks: list[str] = []

    def walk(elements: list[dict]) -> None:
        for el in elements:
            if "paragraph" in el:
                for pe in el["paragraph"].get("elements", []):
                    text_run = pe.get("textRun")
                    if text_run:
                        chunks.append(text_run.get("content", ""))
            elif "table" in el:
                for row in el["table"].get("tableRows", []):
                    for cell in row.get("tableCells", []):
                        walk(cell.get("content", []))
            elif "tableOfContents" in el:
                walk(el["tableOfContents"].get("content", []))

    walk(doc.get("body", {}).get("content", []))
    return "".join(chunks)


def generate_pdf_report(services: dict[str, Any], template_id: str,
                         report_data: dict[str, str], employee_name: str
                         ) -> tuple[str, bytes, str]:
    """Copies the template, fills ONLY the standard placeholders, exports a
    clean PDF into OUTPUT_FOLDER_ID, deletes the intermediate Doc copy, and
    returns (webViewLink, pdf_bytes, pdf_filename)."""
    drive = services["drive"]
    docs = services["docs"]

    date_stamp = datetime.now().strftime("%d-%m-%Y")
    new_doc_name = f"Advance Report - {employee_name} - {date_stamp}"

    copy = drive.files().copy(fileId=template_id, body={"name": new_doc_name}).execute()
    temp_doc_id = copy["id"]

    try:
        replace_placeholders(docs, temp_doc_id, report_data)

        pdf_bytes = drive.files().export(fileId=temp_doc_id, mimeType="application/pdf").execute()
        pdf_filename = f"{new_doc_name}.pdf"

        media = MediaInMemoryUpload(pdf_bytes, mimetype="application/pdf")
        pdf_file = drive.files().create(
            body={"name": pdf_filename, "parents": [OUTPUT_FOLDER_ID]},
            media_body=media,
            fields="id, webViewLink",
        ).execute()

        log.info("Generated clean PDF: %s", pdf_filename)
        return pdf_file["webViewLink"], pdf_bytes, pdf_filename
    finally:
        drive.files().update(fileId=temp_doc_id, body={"trashed": True}).execute()


# --------------------------------------------------------------------------
# STEP 7: HR EMAIL — AUDIT VERDICT + RECONCILIATION IN BODY, CLEAN PDF ATTACHED
# --------------------------------------------------------------------------

def build_hr_email_body(employee_name: str, doc_url: str, audit: AuditResult,
                         reconciliation: Reconciliation) -> str:
    lines = [
        f"A new Advance Report has been submitted and processed for {employee_name}.",
        "",
        f"AI AUDIT VERDICT: {audit.verdict}",
        f"Reasoning: {audit.reasoning}",
    ]
    if audit.compliance_notes:
        lines += ["", f"Compliance notes: {audit.compliance_notes}"]
    lines += [
        "",
        "FINANCIAL RECONCILIATION",
        reconciliation.as_text(),
        "",
        f"Drive copy of the report: {doc_url}",
        "The clean PDF report is attached to this email.",
        "",
        "Thank you.",
    ]
    return "\n".join(lines)


def send_hr_email(employee_name: str, doc_url: str, pdf_bytes: bytes, pdf_filename: str,
                   audit: AuditResult, reconciliation: Reconciliation) -> None:
    message = MIMEMultipart()
    message["From"] = EMAIL_SENDER_ADDRESS
    message["To"] = HR_EMAIL_ADDRESS
    message["Subject"] = f"[ADVANCE REPORT] {audit.verdict} — {employee_name}"

    body = build_hr_email_body(employee_name, doc_url, audit, reconciliation)
    message.attach(MIMEText(body, "plain"))

    attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
    attachment.add_header("Content-Disposition", "attachment", filename=pdf_filename)
    message.attach(attachment)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_SENDER_ADDRESS, EMAIL_SENDER_APP_PASSWORD)
            smtp.send_message(message)
        log.info("HR audit email sent to %s", HR_EMAIL_ADDRESS)
    except smtplib.SMTPException as exc:
        log.error("Error sending HR email: %s", exc)


# --------------------------------------------------------------------------
# ORCHESTRATION
# --------------------------------------------------------------------------

def run_once() -> None:
    services = get_google_services()
    gemini_model = configure_gemini()
    policy_text = load_hr_policy()

    submission = read_latest_submission(services["sheets"], GOOGLE_SHEET_ID, SHEET_RESPONSE_TAB)
    if submission is None:
        log.info("No form responses found yet.")
        return

    headers, row, row_number = submission
    if not has_new_submission(row_number):
        log.info("Row %s already processed; nothing new.", row_number)
        return

    report_data = map_headers_to_report_data(headers, row)
    employee_name = report_data["Employee_NameSurn"]
    log.info("Processing submission for %s (row %s)", employee_name, row_number)

    present_receipts = get_present_receipts(report_data)
    reconciliation = compute_financial_reconciliation(report_data, present_receipts)

    audit = evaluate_hr_policy_compliance(
        gemini_model, report_data, present_receipts, policy_text, reconciliation
    )
    log.info("Audit verdict: %s", audit.verdict)

    # Only standard fields go into the template — the PDF stays 100% clean.
    template_id = select_template(report_data)
    doc_url, pdf_bytes, pdf_filename = generate_pdf_report(
        services, template_id, report_data, employee_name
    )

    send_hr_email(employee_name, doc_url, pdf_bytes, pdf_filename, audit, reconciliation)

    mark_submission_processed(row_number)
    log.info("Done. Report: %s", doc_url)


if __name__ == "__main__":
    run_once()