import os
import re
import json
import time
import random
import base64
from datetime import datetime, timezone
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime

import requests
from openpyxl import load_workbook
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from archive_storage import (
    archive_email_package,
    connect_to_archive,
    verify_existing_archive,
)


# ============================================================
# SRNR EMAIL AGENT - EMERGENCY CLEANUP V3
# ============================================================
#
# PURPOSE
# -------
# Archive each listed Gmail message to a configured path, verify
# the archive, then permanently delete the Gmail copy.
#
# Resume from an existing Emergency Cleanup Excel audit.
#
# IMPORTANT RESUME RULES
# ----------------------
#
# 1. If Excel says:
#       Delete Status = PERMANENTLY DELETED
#    -> SKIP COMPLETELY.
#       No Gmail lookup.
#       No archive lookup.
#       No re-archive.
#
# 2. If Excel says:
#       Archive Status = VERIFIED
#       Delete Status != PERMANENTLY DELETED
#    -> Verify the EXISTING archive.
#    -> If the archive is still valid, DELETE FROM GMAIL ONLY.
#    -> Do NOT download/re-archive the email.
#
# 3. If the email is not already archived:
#    -> Download the exact Gmail Message ID from the Excel file.
#    -> Parse body and attachments locally from the RAW email.
#    -> Classify locally with Qwen.
#    -> Archive to the configured path (local folder by default).
#    -> Verify archive.
#    -> Only then permanently delete from Gmail.
#
# 4. If archive verification fails:
#    -> Gmail message is NOT deleted.
#
# 5. This script NEVER rescans the Gmail mailbox.
#    It uses only Message IDs already present in the Excel report.
#
# ============================================================


# ============================================================
# CONFIGURATION
# ============================================================

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(PROJECT_DIR, "reports")

CREDENTIALS_FILE = os.path.join(
    PROJECT_DIR,
    "credentials",
    "credentials.json"
)

# Token for this project's permanent-delete Gmail scope.
DELETE_TOKEN_FILE = os.path.join(
    PROJECT_DIR,
    "credentials",
    "token.json"
)

# Permanent Gmail deletion requires full Gmail access.
SCOPES = [
    "https://mail.google.com/"
]

# ------------------------------------------------------------
# INPUT REPORT
# ------------------------------------------------------------
#
# SAFEST OPTION:
# Set this to the exact latest V2/V3 result workbook that you
# want to continue from.
#
# Example:
# INPUT_REPORT = os.path.join(
#     REPORT_DIR,
#     "Emergency_Cleanup_V3_Result_2026-09-08_09-26-38.xlsx"
# )
#
# If None, V3 chooses the newest eligible Emergency Cleanup
# workbook in the reports folder.
#
INPUT_REPORT = None

# ------------------------------------------------------------
# BATCH SIZE
# ------------------------------------------------------------
#
# The successful 5-email test has already demonstrated that the
# archive -> verify -> permanent delete sequence works.
#
# 50 is a sensible next batch.
# You can later change this to 200 if desired.
#
PROCESS_LIMIT = 100

# Delay between emails to reduce Gmail quota bursts.
DELAY_BETWEEN_EMAILS_SECONDS = 3

# Gmail automatic retry/backoff.
MAX_API_RETRIES = 7
INITIAL_RETRY_SECONDS = 10
MAX_RETRY_SECONDS = 120

# Local Ollama classification.
OLLAMA_URL = "http://127.0.0.1:11434/api/generate"
OLLAMA_MODEL = "qwen2.5:3b"

# Exact confirmation phrase.
CONFIRMATION_PREFIX = "ARCHIVE AND PERMANENTLY DELETE"


# ============================================================
# BASIC HELPERS
# ============================================================

def format_size(value):
    try:
        size = float(value or 0)
    except Exception:
        size = 0.0

    units = ["B", "KB", "MB", "GB", "TB"]

    for unit in units:
        if size < 1024:
            return f"{size:.2f} {unit}"
        size /= 1024

    return f"{size:.2f} PB"


def normalize_status(value):
    return str(value or "").strip().upper()


def is_permanently_deleted_status(value):
    status = normalize_status(value)

    return (
        "PERMANENTLY DELETED" in status
        or "ALREADY ABSENT FROM GMAIL" in status
    )


def is_verified_archive_status(value):
    return normalize_status(value) == "VERIFIED"


# ============================================================
# FIND INPUT WORKBOOK
# ============================================================

def find_input_report():
    if INPUT_REPORT:
        if not os.path.exists(INPUT_REPORT):
            raise FileNotFoundError(
                "Configured INPUT_REPORT does not exist:\n"
                f"{INPUT_REPORT}"
            )
        return INPUT_REPORT

    if not os.path.isdir(REPORT_DIR):
        raise FileNotFoundError(
            f"Reports folder not found:\n{REPORT_DIR}"
        )

    candidates = []

    for name in os.listdir(REPORT_DIR):
        lower = name.lower()

        if not lower.endswith(".xlsx"):
            continue

        # Accept original emergency reports and V2/V3 result reports.
        if (
            lower.startswith("emergency_cleanup_")
            or lower.startswith("emergency_cleanup_v2_result_")
            or lower.startswith("emergency_cleanup_v3_result_")
        ):
            full_path = os.path.join(REPORT_DIR, name)
            candidates.append(full_path)

    if not candidates:
        raise FileNotFoundError(
            "No Emergency Cleanup Excel workbook was found in:\n"
            f"{REPORT_DIR}"
        )

    candidates.sort(
        key=lambda path: os.path.getmtime(path),
        reverse=True
    )

    return candidates[0]


# ============================================================
# GMAIL AUTHENTICATION
# ============================================================

def authenticate_gmail():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            "Google OAuth credentials file not found:\n"
            f"{CREDENTIALS_FILE}"
        )

    creds = None

    if os.path.exists(DELETE_TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(
                DELETE_TOKEN_FILE,
                SCOPES
            )
        except Exception:
            creds = None

    # If an old token exists without the required scope,
    # preserve it as a backup and force fresh authorization.
    if creds and not creds.has_scopes(SCOPES):
        backup_path = (
            DELETE_TOKEN_FILE
            + ".backup_"
            + datetime.now().strftime("%Y%m%d_%H%M%S")
        )

        try:
            os.replace(
                DELETE_TOKEN_FILE,
                backup_path
            )

            print(
                "\nOld emergency token did not contain the "
                "required scope."
            )
            print(
                "It was backed up to:\n"
                f"{backup_path}"
            )
        except Exception:
            pass

        creds = None

    if not creds or not creds.valid:
        if (
            creds
            and creds.expired
            and creds.refresh_token
        ):
            try:
                creds.refresh(Request())
            except Exception:
                creds = None

        if not creds:
            print("\nGoogle authorization is required.")
            print("A browser window will open.")
            print(
                "Authorize the Gmail account used by "
                "SRNR Email Agent.\n"
            )

            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE,
                SCOPES
            )

            creds = flow.run_local_server(
                port=0,
                prompt="consent",
                access_type="offline"
            )

        with open(
            DELETE_TOKEN_FILE,
            "w",
            encoding="utf-8"
        ) as token_file:
            token_file.write(
                creds.to_json()
            )

    if not creds.has_scopes(SCOPES):
        raise RuntimeError(
            "OAuth token does not contain the required scope:\n"
            "https://mail.google.com/"
        )

    return build(
        "gmail",
        "v1",
        credentials=creds,
        cache_discovery=False
    )


# ============================================================
# GMAIL RETRY / RATE LIMIT HANDLING
# ============================================================

def is_retryable_http_error(error):
    status = getattr(
        error.resp,
        "status",
        None
    )

    text = str(error).lower()

    if status in (
        429,
        500,
        502,
        503,
        504
    ):
        return True

    if status == 403 and (
        "ratelimitexceeded" in text
        or "userratelimitexceeded" in text
        or "quota exceeded" in text
        or "backenderror" in text
    ):
        return True

    return False


def execute_with_retry(
    request_factory,
    description
):
    delay = INITIAL_RETRY_SECONDS

    for attempt in range(
        1,
        MAX_API_RETRIES + 1
    ):
        try:
            return request_factory().execute()

        except HttpError as error:
            if not is_retryable_http_error(error):
                raise

            if attempt >= MAX_API_RETRIES:
                raise

            wait_seconds = (
                delay
                + random.uniform(0, 2.0)
            )

            print(
                f"  Gmail temporary limit during "
                f"{description}."
            )
            print(
                f"  Waiting {wait_seconds:.1f} seconds "
                f"before retry "
                f"{attempt}/{MAX_API_RETRIES - 1}..."
            )

            time.sleep(wait_seconds)

            delay = min(
                delay * 2,
                MAX_RETRY_SECONDS
            )


# ============================================================
# RAW GMAIL DOWNLOAD
# ============================================================

def download_raw_message(
    service,
    message_id
):
    response = execute_with_retry(
        lambda: (
            service
            .users()
            .messages()
            .get(
                userId="me",
                id=message_id,
                format="raw"
            )
        ),
        f"RAW download for {message_id}"
    )

    encoded = response.get(
        "raw",
        ""
    )

    if not encoded:
        raise RuntimeError(
            "Gmail returned no RAW email data."
        )

    encoded += "=" * (
        -len(encoded) % 4
    )

    raw_bytes = (
        base64.urlsafe_b64decode(
            encoded.encode("ascii")
        )
    )

    return response, raw_bytes


# ============================================================
# PARSE RAW EMAIL LOCALLY
# ============================================================

def decode_part_text(part):
    try:
        content = part.get_content()

        if isinstance(
            content,
            str
        ):
            return content
    except Exception:
        pass

    try:
        payload = (
            part.get_payload(
                decode=True
            )
        )

        if payload:
            charset = (
                part.get_content_charset()
                or "utf-8"
            )

            return payload.decode(
                charset,
                errors="replace"
            )

    except Exception:
        pass

    return ""


def strip_html(html):
    if not html:
        return ""

    text = re.sub(
        r"<script.*?>.*?</script>",
        " ",
        html,
        flags=(
            re.IGNORECASE
            | re.DOTALL
        )
    )

    text = re.sub(
        r"<style.*?>.*?</style>",
        " ",
        text,
        flags=(
            re.IGNORECASE
            | re.DOTALL
        )
    )

    text = re.sub(
        r"<br\s*/?>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"</p>",
        "\n",
        text,
        flags=re.IGNORECASE
    )

    text = re.sub(
        r"<[^>]+>",
        " ",
        text
    )

    text = re.sub(
        r"\s+",
        " ",
        text
    )

    return text.strip()


def parse_raw_email(raw_bytes):
    message = BytesParser(
        policy=policy.default
    ).parsebytes(
        raw_bytes
    )

    sender = str(
        message.get(
            "From",
            ""
        )
    )

    recipient = str(
        message.get(
            "To",
            ""
        )
    )

    subject = str(
        message.get(
            "Subject",
            ""
        )
    )

    date_header = str(
        message.get(
            "Date",
            ""
        )
    )

    email_date = None

    if date_header:
        try:
            email_date = (
                parsedate_to_datetime(
                    date_header
                )
            )

            if email_date.tzinfo is None:
                email_date = (
                    email_date.replace(
                        tzinfo=timezone.utc
                    )
                )

        except Exception:
            email_date = None

    plain_parts = []
    html_parts = []
    attachments = []

    parts = (
        message.walk()
        if message.is_multipart()
        else [message]
    )

    for part in parts:
        if part.is_multipart():
            continue

        filename = (
            part.get_filename()
        )

        content_type = (
            part.get_content_type()
        )

        disposition = (
            part.get_content_disposition()
        )

        if filename:
            data = (
                part.get_payload(
                    decode=True
                )
                or b""
            )

            attachments.append({
                "filename":
                    str(filename),

                "mime_type":
                    content_type,

                "data":
                    data,

                "content_disposition":
                    disposition or ""
            })

            continue

        if content_type == "text/plain":
            text = decode_part_text(
                part
            )

            if text:
                plain_parts.append(
                    text
                )

        elif content_type == "text/html":
            text = decode_part_text(
                part
            )

            if text:
                html_parts.append(
                    text
                )

    if plain_parts:
        body = "\n".join(
            plain_parts
        )

    elif html_parts:
        body = strip_html(
            "\n".join(
                html_parts
            )
        )

    else:
        body = ""

    body = re.sub(
        r"\s+",
        " ",
        body
    ).strip()

    return {
        "sender":
            sender,

        "recipient":
            recipient,

        "subject":
            subject,

        "date":
            email_date,

        "body":
            body,

        "attachments":
            attachments
    }


# ============================================================
# LOCAL QWEN CATEGORY
# ============================================================

ALLOWED_CATEGORIES = [
    "Business",
    "Important",
    "Reports",
    "Advertisement",
    "OTP",
    "Spam",
    "Notifications",
    "Newsletter",
    "Finance",
    "Banking",
    "Investment",
    "Personal",
    "Shareholder / RTA",
    "Government",
    "Bills",
    "Orders",
    "Travel",
    "Job",
    "Education",
    "Other"
]


def classify_category(parsed):
    attachment_names = "; ".join(
        item.get(
            "filename",
            ""
        )
        for item
        in parsed[
            "attachments"
        ]
    )

    prompt = f"""
Classify this office email into exactly ONE category.

Allowed categories:
{", ".join(ALLOWED_CATEGORIES)}

Return ONLY JSON:
{{"category":"Business"}}

Prefer "Shareholder / RTA" for emails about:
shareholder claims, share certificates, folio numbers, unpaid dividend,
registrar correspondence, duplicate securities, transmission/transposition,
or registrar/shareholder work.

Sender:
{parsed["sender"]}

Recipient:
{parsed["recipient"]}

Subject:
{parsed["subject"]}

Attachments:
{attachment_names}

Body:
{parsed["body"][:4000]}
"""

    try:
        response = requests.post(
            OLLAMA_URL,
            json={
                "model":
                    OLLAMA_MODEL,

                "prompt":
                    prompt,

                "stream":
                    False,

                "format":
                    "json"
            },
            timeout=120
        )

        response.raise_for_status()

        result = json.loads(
            response.json().get(
                "response",
                "{}"
            )
        )

        category = str(
            result.get(
                "category",
                "Business"
            )
        ).strip()

        if category in ALLOWED_CATEGORIES:
            return category

    except Exception as error:
        print(
            "  Qwen category unavailable; "
            f"using Business. ({error})"
        )

    return "Business"


# Archive verification lives in archive_storage.py
# (local folder by default; optional NAS backend).


# ============================================================
# GMAIL PERMANENT DELETE
# ============================================================

def permanently_delete_message(
    service,
    message_id
):
    return execute_with_retry(
        lambda: (
            service
            .users()
            .messages()
            .delete(
                userId="me",
                id=message_id
            )
        ),
        f"permanent delete for {message_id}"
    )


def is_http_404(error):
    return (
        getattr(
            error.resp,
            "status",
            None
        )
        == 404
    )


# ============================================================
# WORKBOOK HELPERS
# ============================================================

REQUIRED_COLUMNS = [
    "Date",
    "Sender",
    "Recipient",
    "Subject",
    "Message ID",
    "Message Size Bytes",
    "Message Size",
    "Attachment Count",
    "Attachment Names",
    "Archive Status",
    "Archive Path",
    "Delete Status",
    "Error"
]


EXTRA_COLUMNS = [
    "V3 Category",
    "V3 Resume Action",
    "Archive Recheck",
    "V3 Processed At"
]


def load_cleanup_workbook(path):
    workbook = load_workbook(
        path
    )

    if "Cleanup Audit" not in (
        workbook.sheetnames
    ):
        raise RuntimeError(
            'Workbook does not contain a '
            '"Cleanup Audit" sheet.'
        )

    sheet = workbook[
        "Cleanup Audit"
    ]

    header_map = {}

    for cell in sheet[1]:
        if cell.value is not None:
            header_map[
                str(
                    cell.value
                ).strip()
            ] = cell.column

    missing = [
        name
        for name in REQUIRED_COLUMNS
        if name not in header_map
    ]

    if missing:
        raise RuntimeError(
            "Input report is missing columns:\n"
            + ", ".join(
                missing
            )
        )

    for extra in EXTRA_COLUMNS:
        if extra not in header_map:
            new_col = (
                sheet.max_column
                + 1
            )

            sheet.cell(
                row=1,
                column=new_col
            ).value = extra

            header_map[
                extra
            ] = new_col

    return (
        workbook,
        sheet,
        header_map
    )


def cell_value(
    sheet,
    row,
    header_map,
    name
):
    return sheet.cell(
        row=row,
        column=header_map[name]
    ).value


def set_cell(
    sheet,
    row,
    header_map,
    name,
    value
):
    sheet.cell(
        row=row,
        column=header_map[name]
    ).value = value


# ============================================================
# BUILD QUEUE
# ============================================================

def build_processing_queue(
    sheet,
    header_map
):
    """
    IMPORTANT:
    Rows already permanently deleted are excluded here, before
    Gmail or archive operations begin.
    """

    queue = []
    already_deleted_count = 0
    already_deleted_bytes = 0

    for row in range(
        2,
        sheet.max_row + 1
    ):
        message_id = cell_value(
            sheet,
            row,
            header_map,
            "Message ID"
        )

        if not message_id:
            continue

        delete_status = cell_value(
            sheet,
            row,
            header_map,
            "Delete Status"
        )

        size_bytes = int(
            cell_value(
                sheet,
                row,
                header_map,
                "Message Size Bytes"
            )
            or 0
        )

        if is_permanently_deleted_status(
            delete_status
        ):
            already_deleted_count += 1
            already_deleted_bytes += (
                size_bytes
            )

            # SKIP COMPLETELY.
            # No Gmail lookup.
            # No archive lookup.
            continue

        archive_status = (
            cell_value(
                sheet,
                row,
                header_map,
                "Archive Status"
            )
        )

        archive_path = str(
            cell_value(
                sheet,
                row,
                header_map,
                "Archive Path"
            )
            or ""
        ).strip()

        if (
            is_verified_archive_status(
                archive_status
            )
            and archive_path
        ):
            resume_action = (
                "VERIFY EXISTING ARCHIVE -> DELETE ONLY"
            )
        else:
            resume_action = (
                "ARCHIVE -> VERIFY -> DELETE"
            )

        queue.append({
            "row":
                row,

            "message_id":
                str(
                    message_id
                ).strip(),

            "resume_action":
                resume_action,

            "size_bytes":
                size_bytes
        })

    if PROCESS_LIMIT:
        queue = queue[
            :PROCESS_LIMIT
        ]

    return (
        queue,
        already_deleted_count,
        already_deleted_bytes
    )


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 78)
    print(
        "SRNR EMAIL AGENT - EMERGENCY CLEANUP V3"
    )
    print(
        "RESUME-SAFE ARCHIVE + PERMANENT GMAIL DELETE"
    )
    print("=" * 78)

    try:
        input_report = find_input_report()
    except FileNotFoundError as error:
        print(f"\n{error}")
        print(
            "\nPut a Cleanup Audit workbook in the reports folder "
            "and run again. Gmail was not contacted."
        )
        return

    print(
        f"\nInput Excel:\n"
        f"{input_report}"
    )

    workbook, sheet, header_map = (
        load_cleanup_workbook(
            input_report
        )
    )

    (
        queue,
        already_deleted_count,
        already_deleted_bytes
    ) = build_processing_queue(
        sheet,
        header_map
    )

    if not queue:
        print(
            "\nNo undeleted Message IDs remain "
            "in this report."
        )

        print(
            f"Rows already completed/skipped: "
            f"{already_deleted_count}"
        )

        return

    selected_bytes = sum(
        item[
            "size_bytes"
        ]
        for item in queue
    )

    existing_archive_queue = sum(
        1
        for item in queue
        if item[
            "resume_action"
        ].startswith(
            "VERIFY EXISTING"
        )
    )

    new_archive_queue = (
        len(queue)
        - existing_archive_queue
    )

    print(
        f"\nAlready permanently deleted "
        f"(SKIPPED COMPLETELY): "
        f"{already_deleted_count}"
    )

    print(
        f"Previously deleted estimated size: "
        f"{format_size(already_deleted_bytes)}"
    )

    print(
        f"\nMessages queued now: "
        f"{len(queue)}"
    )

    print(
        f"Already archived; delete-only candidates: "
        f"{existing_archive_queue}"
    )

    print(
        f"Need new archive first: "
        f"{new_archive_queue}"
    )

    print(
        f"Estimated queued Gmail size: "
        f"{format_size(selected_bytes)}"
    )

    print("\nV3 SAFETY RULES:")
    print(
        "  1. Already permanently deleted rows are "
        "ignored completely."
    )
    print(
        "  2. Already VERIFIED archives are rechecked "
        "on disk and are NOT re-archived."
    )
    print(
        "  3. New/unverified rows are archived and "
        "verified before Gmail deletion."
    )
    print(
        "  4. If archive verification fails, Gmail is "
        "NOT deleted."
    )
    print(
        "  5. This script does NOT rescan the mailbox."
    )

    # --------------------------------------------------------
    # CREATE ONE PROGRESS/RESULT FILE
    # --------------------------------------------------------

    run_timestamp = (
        datetime.now().strftime(
            "%Y-%m-%d_%H-%M-%S"
        )
    )

    result_report = os.path.join(
        REPORT_DIR,
        (
            "Emergency_Cleanup_V3_Result_"
            f"{run_timestamp}.xlsx"
        )
    )

    # Save immediately, and then reuse the SAME result file
    # after every row. This avoids creating dozens of reports.
    workbook.save(
        result_report
    )

    print(
        f"\nV3 result/progress workbook:\n"
        f"{result_report}"
    )

    # --------------------------------------------------------
    # CONFIRMATION
    # --------------------------------------------------------

    expected_confirmation = (
        f"{CONFIRMATION_PREFIX} "
        f"{len(queue)}"
    )

    print(
        "\nPermanent deletion cannot be undone."
    )
    print(
        "To continue, type exactly:"
    )
    print(
        expected_confirmation
    )

    confirmation = input(
        "\nConfirmation: "
    ).strip()

    if confirmation != expected_confirmation:
        print(
            "\nConfirmation did not match."
        )
        print(
            "Nothing was changed in Gmail."
        )
        return

    # --------------------------------------------------------
    # AUTHENTICATE ONLY AFTER CONFIRMATION
    # --------------------------------------------------------

    print(
        "\nAuthenticating Gmail..."
    )

    service = (
        authenticate_gmail()
    )

    print(
        "Gmail authorization: OK"
    )

    # Connect the archive backend once up front.
    print(
        "\nConnecting to archive storage..."
    )

    connect_to_archive()

    print(
        "Archive connection: OK"
    )

    attempted = 0
    archive_verified_count = 0
    existing_archive_reused = 0
    newly_archived = 0
    permanently_deleted = 0
    already_absent = 0
    failed = 0
    deleted_bytes = 0

    # ========================================================
    # PROCESS QUEUE
    # ========================================================

    for position, item in enumerate(
        queue,
        start=1
    ):
        row = item["row"]
        message_id = item[
            "message_id"
        ]

        size_bytes = item[
            "size_bytes"
        ]

        subject = str(
            cell_value(
                sheet,
                row,
                header_map,
                "Subject"
            )
            or ""
        )

        archive_status = (
            cell_value(
                sheet,
                row,
                header_map,
                "Archive Status"
            )
        )

        archive_path = str(
            cell_value(
                sheet,
                row,
                header_map,
                "Archive Path"
            )
            or ""
        ).strip()

        attachment_count = int(
            cell_value(
                sheet,
                row,
                header_map,
                "Attachment Count"
            )
            or 0
        )

        set_cell(
            sheet,
            row,
            header_map,
            "V3 Resume Action",
            item[
                "resume_action"
            ]
        )

        print(
            "\n"
            + "-"
            * 78
        )

        print(
            f"{position}/{len(queue)} | "
            f"{message_id} | "
            f"{format_size(size_bytes)}"
        )

        print(
            f"Subject: "
            f"{subject[:120]}"
        )

        print(
            "Resume action: "
            + item[
                "resume_action"
            ]
        )

        gmail_delete_allowed = False

        try:
            # =================================================
            # CASE A:
            # ALREADY ARCHIVED + VERIFIED
            # =================================================

            if (
                is_verified_archive_status(
                    archive_status
                )
                and archive_path
            ):
                print(
                    "  Rechecking existing archive..."
                )

                (
                    nas_ok,
                    nas_reason
                ) = (
                    verify_existing_archive(
                        archive_path,
                        attachment_count
                    )
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Archive Recheck",
                    (
                        "VERIFIED: "
                        + nas_reason
                        if nas_ok
                        else
                        "FAILED: "
                        + nas_reason
                    )
                )

                if not nas_ok:
                    raise RuntimeError(
                        "Existing archive recheck FAILED. "
                        "Gmail will not be deleted. "
                        + nas_reason
                    )

                archive_verified_count += 1
                existing_archive_reused += 1
                gmail_delete_allowed = True

                print(
                    "  Existing archive: VERIFIED"
                )
                print(
                    "  Re-archive: SKIPPED"
                )

            # =================================================
            # CASE B:
            # NOT YET VERIFIED -> DOWNLOAD + ARCHIVE
            # =================================================

            else:
                print(
                    "  Downloading original Gmail message..."
                )

                try:
                    (
                        gmail_record,
                        raw_bytes
                    ) = download_raw_message(
                        service,
                        message_id
                    )

                except HttpError as error:
                    if is_http_404(
                        error
                    ):
                        raise RuntimeError(
                            "Gmail message is already absent, "
                            "but this row does not have a "
                            "verified archive. "
                            "No safe archive can now be created."
                        )

                    raise

                parsed = parse_raw_email(
                    raw_bytes
                )

                if not parsed[
                    "subject"
                ]:
                    parsed[
                        "subject"
                    ] = subject

                print(
                    "  Parsed attachments: "
                    f"{len(parsed['attachments'])}"
                )

                category = (
                    classify_category(
                        parsed
                    )
                )

                print(
                    f"  Category: "
                    f"{category}"
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "V3 Category",
                    category
                )

                metadata = {
                    "message_id":
                        message_id,

                    "thread_id":
                        gmail_record.get(
                            "threadId",
                            ""
                        ),

                    "gmail_labels":
                        gmail_record.get(
                            "labelIds",
                            []
                        ),

                    "gmail_size_estimate":
                        gmail_record.get(
                            "sizeEstimate",
                            size_bytes
                        ),

                    "source_excel":
                        input_report,

                    "source_excel_row":
                        row,

                    "sender":
                        parsed[
                            "sender"
                        ],

                    "recipient":
                        parsed[
                            "recipient"
                        ],

                    "subject":
                        parsed[
                            "subject"
                        ],

                    "date":
                        (
                            parsed[
                                "date"
                            ].isoformat()
                            if parsed[
                                "date"
                            ]
                            else ""
                        ),

                    "category":
                        category,

                    "emergency_cleanup_version":
                        "V3"
                }

                print(
                    "  Archiving..."
                )

                archive_result = (
                    archive_email_package(
                        category=
                            category,

                        email_date=
                            parsed[
                                "date"
                            ],

                        sender=
                            parsed[
                                "sender"
                            ],

                        subject=
                            parsed[
                                "subject"
                            ],

                        message_id=
                            message_id,

                        raw_email_bytes=
                            raw_bytes,

                        body_text=
                            parsed[
                                "body"
                            ],

                        metadata=
                            metadata,

                        attachments=
                            parsed[
                                "attachments"
                            ]
                    )
                )

                if not archive_result.get(
                    "verified"
                ):
                    raise RuntimeError(
                        "Archive did not return VERIFIED."
                    )

                archive_path = (
                    archive_result.get(
                        "archive_path",
                        ""
                    )
                )

                # Recheck it independently after archive_email_package
                # reports success.
                (
                    nas_ok,
                    nas_reason
                ) = (
                    verify_existing_archive(
                        archive_path,
                        len(
                            parsed[
                                "attachments"
                            ]
                        )
                    )
                )

                if not nas_ok:
                    raise RuntimeError(
                        "Post-archive recheck FAILED. "
                        + nas_reason
                    )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Archive Status",
                    "VERIFIED"
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Archive Path",
                    archive_path
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Archive Recheck",
                    (
                        "VERIFIED: "
                        + nas_reason
                    )
                )

                archive_verified_count += 1
                newly_archived += 1
                gmail_delete_allowed = True

                print(
                    "  New archive: VERIFIED"
                )

            # =================================================
            # DELETE ONLY AFTER VERIFIED ARCHIVE
            # =================================================

            if not gmail_delete_allowed:
                raise RuntimeError(
                    "Internal safety gate blocked Gmail deletion."
                )

            print(
                "  Permanently deleting from Gmail..."
            )

            try:
                permanently_delete_message(
                    service,
                    message_id
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Delete Status",
                    "PERMANENTLY DELETED"
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Error",
                    ""
                )

                permanently_deleted += 1
                deleted_bytes += (
                    size_bytes
                )

                print(
                    "  Gmail: PERMANENTLY DELETED"
                )

            except HttpError as error:
                if is_http_404(
                    error
                ):
                    # At this point the archive is VERIFIED.
                    # A 404 therefore means Gmail no longer has this
                    # message. Treat it as complete, without retrying.
                    set_cell(
                        sheet,
                        row,
                        header_map,
                        "Delete Status",
                        "ALREADY ABSENT FROM GMAIL"
                    )

                    set_cell(
                        sheet,
                        row,
                        header_map,
                        "Error",
                        ""
                    )

                    already_absent += 1

                    print(
                        "  Gmail: ALREADY ABSENT"
                    )
                    print(
                        "  Existing archive is VERIFIED, "
                        "so this row is considered complete."
                    )

                else:
                    raise

        except HttpError as error:
            failed += 1

            error_text = str(
                error
            )

            if not is_permanently_deleted_status(
                cell_value(
                    sheet,
                    row,
                    header_map,
                    "Delete Status"
                )
            ):
                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Delete Status",
                    "NOT DELETED"
                )

            set_cell(
                sheet,
                row,
                header_map,
                "Error",
                error_text[
                    :32000
                ]
            )

            print(
                f"  FAILED: "
                f"{error}"
            )

            # Stop immediately on permissions errors.
            if (
                getattr(
                    error.resp,
                    "status",
                    None
                )
                == 403
                and "insufficient" in (
                    error_text.lower()
                )
            ):
                print(
                    "\nSTOPPING: Gmail token lacks "
                    "permanent-delete permission."
                )
                print(
                    "No further messages will be attempted."
                )

                set_cell(
                    sheet,
                    row,
                    header_map,
                    "V3 Processed At",
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                )

                workbook.save(
                    result_report
                )

                attempted += 1
                break

        except Exception as error:
            failed += 1

            error_text = str(
                error
            )

            # Never overwrite a completed deletion marker.
            if not is_permanently_deleted_status(
                cell_value(
                    sheet,
                    row,
                    header_map,
                    "Delete Status"
                )
            ):
                set_cell(
                    sheet,
                    row,
                    header_map,
                    "Delete Status",
                    "NOT DELETED"
                )

            set_cell(
                sheet,
                row,
                header_map,
                "Error",
                error_text[
                    :32000
                ]
            )

            print(
                f"  FAILED / NOT DELETED: "
                f"{error}"
            )

        finally:
            set_cell(
                sheet,
                row,
                header_map,
                "V3 Processed At",
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            )

            # Save the SAME progress workbook after every row.
            workbook.save(
                result_report
            )

            attempted += 1

        if position < len(queue):
            print(
                f"  Waiting "
                f"{DELAY_BETWEEN_EMAILS_SECONDS} "
                f"seconds..."
            )

            time.sleep(
                DELAY_BETWEEN_EMAILS_SECONDS
            )

    # ========================================================
    # FINAL SAVE + SUMMARY
    # ========================================================

    workbook.save(
        result_report
    )

    print(
        "\n"
        + "="
        * 78
    )

    print(
        "EMERGENCY CLEANUP V3 COMPLETE"
    )

    print(
        "="
        * 78
    )

    print(
        f"Already deleted before this run: "
        f"{already_deleted_count}"
    )

    print(
        f"Rows attempted this run:         "
        f"{attempted}"
    )

    print(
        f"Existing archives reused:        "
        f"{existing_archive_reused}"
    )

    print(
        f"New archives created:            "
        f"{newly_archived}"
    )

    print(
        f"Archives verified this run:      "
        f"{archive_verified_count}"
    )

    print(
        f"Permanently deleted this run:    "
        f"{permanently_deleted}"
    )

    print(
        f"Already absent from Gmail:       "
        f"{already_absent}"
    )

    print(
        f"Failed / not deleted:            "
        f"{failed}"
    )

    print(
        "Estimated Gmail size permanently "
        "deleted this run: "
        f"{format_size(deleted_bytes)}"
    )

    print(
        f"\nAudit/progress report:\n"
        f"{result_report}"
    )

    print(
        "\nRESUME NOTE:"
    )

    print(
        "If you run V3 again, set INPUT_REPORT "
        "to this V3 result workbook."
    )

    print(
        "Rows marked PERMANENTLY DELETED or "
        "ALREADY ABSENT FROM GMAIL will then "
        "be skipped completely."
    )


if __name__ == "__main__":
    main()
