import os
import json
import re
from datetime import datetime


# ============================================================
# ARCHIVE STORAGE
# ============================================================
#
# Default: write verified email packages to a local folder.
# Optional: set ARCHIVE_MODE = "nas" to write over SMB instead.
#
# Never commit real hostnames, share paths, or passwords.
# ============================================================


PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

# ------------------------------------------------------------
# BACKEND
# ------------------------------------------------------------
#
# "local" = <this project>/archive
# "nas"   = SMB share (placeholders below; fill in locally)
#
ARCHIVE_MODE = "local"

LOCAL_ARCHIVE_ROOT = os.path.join(PROJECT_DIR, "archive")

# Placeholders only. Used when ARCHIVE_MODE = "nas".
NAS_SERVER = "YOUR_NAS_HOST"
NAS_ARCHIVE_ROOT = r"\\YOUR_NAS_HOST\share\Archive_Emails"
NAS_CREDENTIALS_FILE = os.path.join(
    PROJECT_DIR,
    "config",
    "archive_credentials.json"
)


def get_archive_root():
    if ARCHIVE_MODE == "nas":
        return NAS_ARCHIVE_ROOT
    return LOCAL_ARCHIVE_ROOT


def is_nas_mode():
    return ARCHIVE_MODE == "nas"


def archive_join(*parts):
    cleaned = [str(part).rstrip("\\/") for part in parts if part is not None]

    if is_nas_mode():
        return "\\".join(cleaned)

    return os.path.join(*cleaned)


# ============================================================
# CREDENTIALS (NAS MODE ONLY)
# ============================================================

def load_nas_credentials():
    if not os.path.exists(NAS_CREDENTIALS_FILE):
        raise FileNotFoundError(
            "NAS credentials file not found:\n"
            f"{NAS_CREDENTIALS_FILE}\n\n"
            "Copy config/archive_credentials.json.example to "
            "config/archive_credentials.json and fill in the "
            "placeholders. NAS mode is optional."
        )

    with open(NAS_CREDENTIALS_FILE, "r", encoding="utf-8") as file:
        data = json.load(file)

    username = str(data.get("username", "")).strip()
    password = data.get("password", "")

    if (
        not username
        or not password
        or username.startswith("YOUR_")
        or NAS_SERVER.startswith("YOUR_")
    ):
        raise ValueError(
            "NAS placeholders are still present. "
            "Set NAS_SERVER, NAS_ARCHIVE_ROOT, and real "
            "credentials before using ARCHIVE_MODE = \"nas\"."
        )

    return username, password


def _smbclient():
    try:
        import smbclient
    except ImportError as error:
        raise ImportError(
            "NAS mode requires the smbprotocol package.\n"
            "Install it with: pip install smbprotocol"
        ) from error

    return smbclient


# ============================================================
# CONNECTION
# ============================================================

def connect_to_archive():
    if is_nas_mode():
        smbclient = _smbclient()
        username, password = load_nas_credentials()

        smbclient.register_session(
            NAS_SERVER,
            username=username,
            password=password
        )

        root = get_archive_root()

        if not smbclient.path.exists(root):
            raise FileNotFoundError(
                "NAS archive path is not accessible:\n"
                f"{root}"
            )

        return True

    os.makedirs(get_archive_root(), exist_ok=True)
    return True


# Keep the old name as an alias so callers can stay explicit.
initialize_archive = connect_to_archive


# ============================================================
# STORAGE PRIMITIVES
# ============================================================

def storage_exists(path):
    if is_nas_mode():
        return _smbclient().path.exists(path)
    return os.path.exists(path)


def storage_getsize(path):
    if is_nas_mode():
        return _smbclient().path.getsize(path)
    return os.path.getsize(path)


def storage_makedirs(path, exist_ok=True):
    if is_nas_mode():
        _smbclient().makedirs(path, exist_ok=exist_ok)
        return

    os.makedirs(path, exist_ok=exist_ok)


def storage_listdir(path):
    if is_nas_mode():
        return _smbclient().listdir(path)
    return os.listdir(path)


def storage_write_text(path, text):
    if is_nas_mode():
        with _smbclient().open_file(
            path,
            mode="w",
            encoding="utf-8"
        ) as file:
            file.write(text or "")
        return

    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)

    with open(path, "w", encoding="utf-8") as file:
        file.write(text or "")


def storage_write_binary(path, data):
    if is_nas_mode():
        with _smbclient().open_file(path, mode="wb") as file:
            file.write(data)
        return

    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)

    with open(path, "wb") as file:
        file.write(data)


def storage_read_json(path):
    if is_nas_mode():
        with _smbclient().open_file(
            path,
            mode="r",
            encoding="utf-8"
        ) as file:
            return json.load(file)

    with open(path, "r", encoding="utf-8") as file:
        return json.load(file)


def storage_is_dir(path):
    if is_nas_mode():
        return _smbclient().path.isdir(path)
    return os.path.isdir(path)


def remove_archive_directory(path):
    """
    Remove a test archive folder. Used by test_archive.py.
    """

    if not storage_exists(path):
        return

    if not is_nas_mode():
        import shutil
        shutil.rmtree(path)
        return

    smbclient = _smbclient()

    def remove_tree(target):
        if not smbclient.path.exists(target):
            return

        if smbclient.path.isdir(target):
            for name in smbclient.listdir(target):
                if name in (".", ".."):
                    continue
                remove_tree(archive_join(target, name))
            smbclient.rmdir(target)
        else:
            smbclient.remove(target)

    remove_tree(path)


# ============================================================
# SAFE FILE/FOLDER NAME
# ============================================================

def safe_filename(value, max_length=100):
    if not value:
        return "Unknown"

    value = str(value)
    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", " ", value).strip()
    value = value.rstrip(". ")

    if not value:
        value = "Unknown"

    return value[:max_length]


# ============================================================
# CATEGORY NORMALIZATION
# ============================================================

def normalize_category(category):
    mapping = {
        "Business": "Business",
        "Important": "Important",
        "Reports": "Reports",
        "Advertisement": "Advertisement",
        "Marketing": "Advertisement",
        "OTP": "OTP",
        "Spam": "Spam",
        "Notification": "Notifications",
        "Notifications": "Notifications",
        "Newsletter": "Newsletter",
        "Finance": "Finance",
        "Banking": "Banking",
        "Investment": "Investment",
        "Personal": "Personal",
        "Shareholder / RTA": "Shareholder_RTA",
        "Shareholder_RTA": "Shareholder_RTA",
        "Government": "Government",
        "Bills": "Bills",
        "Order": "Orders",
        "Orders": "Orders",
        "Travel": "Travel",
        "Job": "Job",
        "Education": "Education",
        "Other": "Other",
    }

    return mapping.get(category, "Other")


# ============================================================
# BUILD EMAIL ARCHIVE DIRECTORY
# ============================================================

def build_archive_directory(
    category,
    email_date,
    sender,
    subject,
    message_id
):
    category_folder = normalize_category(category)

    if email_date:
        year = email_date.strftime("%Y")
        month = email_date.strftime("%m")
        date_text = email_date.strftime("%Y-%m-%d")
    else:
        now = datetime.now()
        year = now.strftime("%Y")
        month = now.strftime("%m")
        date_text = now.strftime("%Y-%m-%d")

    sender_safe = safe_filename(sender, 40)
    subject_safe = safe_filename(subject, 70)
    short_id = safe_filename(message_id[-10:], 10)

    folder_name = (
        f"{date_text}_"
        f"{sender_safe}_"
        f"{subject_safe}_"
        f"{short_id}"
    )

    return archive_join(
        get_archive_root(),
        category_folder,
        year,
        month,
        folder_name
    )


# ============================================================
# VERIFY FILE
# ============================================================

def verify_file(path, expected_size=None):
    if not storage_exists(path):
        return False

    if expected_size is not None:
        actual_size = storage_getsize(path)
        if actual_size != expected_size:
            return False

    return True


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


# ============================================================
# VERIFY AN EXISTING ARCHIVE PACKAGE
# ============================================================

def verify_existing_archive(archive_path, expected_attachment_count=0):
    """
    Re-check an already written archive before allowing Gmail deletion.

    Returns:
        (True, "reason") on success
        (False, "reason") on failure
    """

    if not archive_path:
        return False, "Archive path is blank."

    try:
        connect_to_archive()

        if not storage_exists(archive_path):
            return False, "Existing archive folder does not exist."

        eml_path = archive_join(archive_path, "email.eml")
        metadata_path = archive_join(archive_path, "metadata.json")
        body_path = archive_join(archive_path, "body.txt")

        for required_file in (eml_path, metadata_path, body_path):
            if not storage_exists(required_file):
                return (
                    False,
                    "Required archive file missing: " + required_file
                )

        eml_size = storage_getsize(eml_path)

        if eml_size <= 0:
            return False, "email.eml exists but is empty."

        try:
            metadata = storage_read_json(metadata_path)
        except Exception as error:
            return (
                False,
                "metadata.json could not be read: " + str(error)
            )

        archived_attachments = (
            metadata.get("archived_attachments", []) or []
        )

        attachments_directory = archive_join(
            archive_path,
            "attachments"
        )

        if archived_attachments:
            if not storage_exists(attachments_directory):
                return False, "Attachments folder is missing."

            for attachment in archived_attachments:
                archived_filename = attachment.get(
                    "archived_filename",
                    ""
                )
                expected_size = attachment.get("size_bytes")

                if not archived_filename:
                    return (
                        False,
                        "metadata.json contains an attachment "
                        "without archived_filename."
                    )

                attachment_path = archive_join(
                    attachments_directory,
                    archived_filename
                )

                if not storage_exists(attachment_path):
                    return (
                        False,
                        "Archived attachment missing: "
                        + archived_filename
                    )

                if expected_size is not None:
                    actual_size = storage_getsize(attachment_path)
                    if int(actual_size) != int(expected_size):
                        return (
                            False,
                            "Archived attachment size mismatch: "
                            + archived_filename
                        )

        elif int(expected_attachment_count or 0) > 0:
            if not storage_exists(attachments_directory):
                return (
                    False,
                    "Excel expects attachments but the "
                    "attachments folder is missing."
                )

            try:
                entries = storage_listdir(attachments_directory)
                file_count = len(
                    [
                        name
                        for name in entries
                        if name not in (".", "..")
                    ]
                )

                if file_count < int(expected_attachment_count):
                    return (
                        False,
                        "Archive attachment count is lower than "
                        "the Excel attachment count."
                    )

            except Exception as error:
                return (
                    False,
                    "Could not verify the existing "
                    "attachments folder: " + str(error)
                )

        return (
            True,
            (
                f"Existing archive verified. "
                f"email.eml={format_size(eml_size)}, "
                f"attachments={len(archived_attachments)}"
            )
        )

    except Exception as error:
        return False, "Archive verification error: " + str(error)


# ============================================================
# ARCHIVE COMPLETE EMAIL PACKAGE
# ============================================================

def archive_email_package(
    category,
    email_date,
    sender,
    subject,
    message_id,
    raw_email_bytes,
    body_text,
    metadata,
    attachments
):
    connect_to_archive()

    archive_directory = build_archive_directory(
        category=category,
        email_date=email_date,
        sender=sender,
        subject=subject,
        message_id=message_id
    )

    attachments_directory = archive_join(
        archive_directory,
        "attachments"
    )

    storage_makedirs(attachments_directory, exist_ok=True)

    eml_path = archive_join(archive_directory, "email.eml")
    storage_write_binary(eml_path, raw_email_bytes)

    body_path = archive_join(archive_directory, "body.txt")
    storage_write_text(body_path, body_text)

    archived_attachments = []

    for index, attachment in enumerate(attachments, start=1):
        filename = safe_filename(
            attachment.get("filename", f"attachment_{index}"),
            150
        )

        destination_name = f"{index:02d}_{filename}"
        destination_path = archive_join(
            attachments_directory,
            destination_name
        )

        data = attachment.get("data", b"")
        storage_write_binary(destination_path, data)

        if not verify_file(destination_path, len(data)):
            raise RuntimeError(
                f"Attachment verification failed: {destination_name}"
            )

        archived_attachments.append({
            "original_filename": attachment.get("filename", ""),
            "archived_filename": destination_name,
            "mime_type": attachment.get("mime_type", ""),
            "size_bytes": len(data),
        })

    metadata_copy = dict(metadata)
    metadata_copy["archive_timestamp"] = datetime.now().isoformat()
    metadata_copy["archive_path"] = archive_directory
    metadata_copy["archive_mode"] = ARCHIVE_MODE
    metadata_copy["archived_attachments"] = archived_attachments

    metadata_path = archive_join(archive_directory, "metadata.json")
    metadata_json = json.dumps(
        metadata_copy,
        indent=2,
        ensure_ascii=False,
        default=str
    )
    storage_write_text(metadata_path, metadata_json)

    for required_file in (eml_path, body_path, metadata_path):
        if not verify_file(required_file):
            raise RuntimeError(
                "Archive verification failed:\n"
                f"{required_file}"
            )

    if not verify_file(eml_path, len(raw_email_bytes)):
        raise RuntimeError("EML size verification failed.")

    return {
        "status": "VERIFIED",
        "verified": True,
        "archive_path": archive_directory,
        "archive_timestamp": metadata_copy["archive_timestamp"],
        "attachment_count": len(archived_attachments),
    }
