# Emergency Cleanup

Archive a Gmail message to a configured path, **verify** the archive, then permanently delete the Gmail copy.

This project does not rescan the mailbox. It only processes Message IDs that already appear in an Excel audit workbook.

## Safety rules

1. Rows already marked `PERMANENTLY DELETED` (or `ALREADY ABSENT FROM GMAIL`) are skipped completely.
2. Rows with `Archive Status = VERIFIED` are rechecked on disk. If the archive is still valid, Gmail is deleted and the message is **not** re-archived.
3. All other rows are downloaded, archived, and verified before Gmail deletion.
4. If archive verification fails, Gmail is **not** deleted.
5. You must type an exact confirmation phrase before anything is deleted.

Permanent Gmail deletion requires the full Gmail scope: `https://mail.google.com/`

## What you need

- Python 3.10+
- [Ollama](https://ollama.com/) running locally with `qwen2.5:3b` (or change `OLLAMA_MODEL` in `emergency_cleanup.py`)
- A Google Cloud OAuth **Desktop** client authorized for full Gmail access
- An Excel workbook in `reports/` with a sheet named `Cleanup Audit`

Optional: Ollama is used only when a new archive must be created (category classification). Delete-only rows do not call the model.

## Setup

```text
pip install -r requirements.txt
```

1. Copy `credentials/credentials.json.example` to `credentials/credentials.json` and fill in your Google OAuth Desktop client.
2. Put an audit workbook in `reports/`. Leave `INPUT_REPORT = None` to use the newest matching file, or set it to a specific workbook path inside `reports/`.

Required columns on `Cleanup Audit`:

- Date
- Sender
- Recipient
- Subject
- Message ID
- Message Size Bytes
- Message Size
- Attachment Count
- Attachment Names
- Archive Status
- Archive Path
- Delete Status
- Error

Workbook names that are auto-detected start with `emergency_cleanup_`, `emergency_cleanup_v2_result_`, or `emergency_cleanup_v3_result_`.

## Archive location

Default: a local folder named `archive/` inside this project.

To use a NAS/SMB share instead, edit `archive_storage.py`:

```python
ARCHIVE_MODE = "nas"
NAS_SERVER = "YOUR_NAS_HOST"
NAS_ARCHIVE_ROOT = r"\\YOUR_NAS_HOST\share\Archive_Emails"
```

Then copy `config/archive_credentials.json.example` to `config/archive_credentials.json` and fill in the placeholders. Do not commit real credentials.

## Test the archive backend

This does not touch Gmail. It writes a dummy package, verifies it, and removes it:

```text
python test_archive.py
```

## Run cleanup

```text
python emergency_cleanup.py
```

You will be asked to type:

```text
ARCHIVE AND PERMANENTLY DELETE <count>
```

where `<count>` is the number of queued messages. If the phrase does not match, nothing is changed.

Progress is saved after every row to `reports/Emergency_Cleanup_V3_Result_<timestamp>.xlsx`. To resume, set `INPUT_REPORT` to that result workbook.

## Paths

All default paths are relative to this project folder. There is no hardcoded machine path, NAS IP, or organization-specific share.
