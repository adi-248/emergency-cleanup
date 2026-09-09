from datetime import datetime, timezone

from archive_storage import (
    ARCHIVE_MODE,
    archive_email_package,
    connect_to_archive,
    get_archive_root,
    remove_archive_directory,
    verify_existing_archive,
)


def main():
    print("=" * 60)
    print("EMERGENCY CLEANUP - ARCHIVE STORAGE TEST")
    print("=" * 60)

    print(f"\nArchive mode : {ARCHIVE_MODE}")
    print(f"Archive root : {get_archive_root()}")
    print("\nThis test does not connect to Gmail.")
    print("It writes a dummy email package, verifies it, then removes it.")

    try:
        connect_to_archive()
        print("\nArchive connection: OK")

        raw_email = (
            b"From: archive-test@example.com\r\n"
            b"To: you@example.com\r\n"
            b"Subject: Archive storage test\r\n"
            b"Date: Tue, 08 Sep 2026 09:00:00 +0000\r\n"
            b"\r\n"
            b"This is a dummy message used only by test_archive.py.\r\n"
        )

        result = archive_email_package(
            category="Other",
            email_date=datetime.now(timezone.utc),
            sender="archive-test@example.com",
            subject="Archive storage test",
            message_id="TESTARCHIVE0001",
            raw_email_bytes=raw_email,
            body_text="This is a dummy message used only by test_archive.py.",
            metadata={
                "message_id": "TESTARCHIVE0001",
                "test": True,
            },
            attachments=[
                {
                    "filename": "test.txt",
                    "mime_type": "text/plain",
                    "data": b"hello archive test\n",
                }
            ],
        )

        archive_path = result["archive_path"]
        print(f"\nWrote test package:\n{archive_path}")

        ok, reason = verify_existing_archive(
            archive_path,
            expected_attachment_count=1
        )

        if not ok:
            raise RuntimeError(reason)

        print(f"Verification: OK ({reason})")

        remove_archive_directory(archive_path)
        print("Test package removed.")

        print("\n" + "=" * 60)
        print("ARCHIVE TEST SUCCESSFUL")
        print("=" * 60)

        print(
            "\nThe cleanup tool can write, verify, and "
            "remove an archive package using the configured backend."
        )

    except Exception as error:
        print("\n" + "=" * 60)
        print("ARCHIVE TEST FAILED")
        print("=" * 60)
        print(f"\n{error}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
