"""Yahoo Inbox attachment downloader used by the Tok Tenging GUI."""

from __future__ import annotations

import imaplib
import re
from datetime import date, timedelta
from email import policy
from email.header import decode_header, make_header
from email.parser import BytesParser
from pathlib import Path
from typing import Callable


IMAP_HOST = "export.imap.mail.yahoo.com"
MAX_FILENAME_LENGTH = 140
SKIP_EXTENSIONS = {".xml", ".jpg", ".jpeg", ".png"}


class DownloadCancelled(RuntimeError):
    """Raised internally when the user cancels an active download."""


def parse_date_range(start_value: str, end_value: str) -> tuple[date, date]:
    """Parse an inclusive ISO date range."""
    try:
        start_date = date.fromisoformat((start_value or "").strip())
        end_date = date.fromisoformat((end_value or "").strip())
    except ValueError as exc:
        raise ValueError("Dagsetningar verða að vera á forminu ÁÁÁÁ-MM-DD.") from exc
    if end_date < start_date:
        raise ValueError("Lokadagsetning má ekki vera á undan upphafsdagsetningu.")
    if end_date == date.max:
        raise ValueError("Lokadagsetning er utan studds dagsetningabils.")
    return start_date, end_date


def imap_date_bounds(start_date: date, end_date: date) -> tuple[str, str]:
    """Return inclusive SINCE and exclusive BEFORE values for IMAP search."""
    return (
        start_date.strftime("%d-%b-%Y"),
        (end_date + timedelta(days=1)).strftime("%d-%b-%Y"),
    )


def readable_header(value: str | None) -> str:
    if not value:
        return ""
    try:
        return str(make_header(decode_header(value)))
    except (LookupError, UnicodeError):
        return value


def safe_filename(filename: str) -> str:
    filename = Path(filename.replace("\\", "/")).name
    filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename).strip(" .")
    filename = filename or "unnamed_attachment"
    path = Path(filename)
    suffix = path.suffix[:20]
    stem_limit = max(1, MAX_FILENAME_LENGTH - len(suffix))
    return f"{path.stem[:stem_limit]}{suffix}"


def unused_path(directory: Path, filename: str) -> Path:
    safe_name = safe_filename(filename)
    original = Path(safe_name)
    candidate = directory / safe_name
    counter = 2
    while candidate.exists():
        counter_text = f"_{counter}"
        stem_limit = max(1, MAX_FILENAME_LENGTH - len(original.suffix) - len(counter_text))
        candidate = directory / f"{original.stem[:stem_limit]}{counter_text}{original.suffix}"
        counter += 1
    return candidate


def new_output_directory(output_root: str | Path, start_date: date, end_date: date) -> Path:
    """Create a unique date-labelled run folder below ``email downloads``."""
    root = Path(output_root)
    root.mkdir(parents=True, exist_ok=True)
    folder_name = f"download_{start_date.isoformat()}_to_{end_date.isoformat()}"
    candidate = root / folder_name
    counter = 2
    while candidate.exists():
        candidate = root / f"{folder_name}_{counter}"
        counter += 1
    candidate.mkdir()
    return candidate


def _received_date_key(mailbox: imaplib.IMAP4_SSL, uid: bytes) -> tuple:
    """Sort by Yahoo's received timestamp; use UID if metadata is unavailable."""
    status, response = mailbox.uid("fetch", uid, "(INTERNALDATE)")
    if status == "OK":
        for item in response:
            metadata = item[0] if isinstance(item, tuple) else item
            if isinstance(metadata, bytes):
                parsed = imaplib.Internaldate2tuple(metadata)
                if parsed is not None:
                    return (0, *parsed[:6], int(uid))
    return (1, int(uid))


def download_inbox_attachments(
    *,
    email_address: str,
    app_password: str,
    start_value: str,
    end_value: str,
    output_root: str | Path,
    emit: Callable[[str], None] | None = None,
    is_cancelled: Callable[[], bool] | None = None,
) -> dict:
    """Download allowed Inbox attachments and stream progress through ``emit``."""
    emit = emit or (lambda _message: None)
    is_cancelled = is_cancelled or (lambda: False)
    start_date, end_date = parse_date_range(start_value, end_value)
    imap_start, imap_end_exclusive = imap_date_bounds(start_date, end_date)
    email_address = (email_address or "").strip()
    app_password = (app_password or "").replace(" ", "")
    if not email_address or not app_password:
        raise ValueError("Yahoo-netfang og app-lykilorð vantar í stillingar.")

    output_dir = new_output_directory(output_root, start_date, end_date)
    skipped_path = output_dir / "skipped_emails.txt"
    attachment_count = 0
    failed_messages = 0
    skipped_emails: list[str] = []
    message_count = 0
    processed_count = 0
    cancelled = False

    emit("Yahoo Inbox attachment downloader")
    emit(f"Date range: {start_date.isoformat()} through {end_date.isoformat()} (inclusive)")
    emit(f"Output folder: {output_dir}")

    try:
        with imaplib.IMAP4_SSL(IMAP_HOST, 993) as mailbox:
            try:
                mailbox.login(email_address, app_password)
            except imaplib.IMAP4.error as exc:
                raise RuntimeError(
                    "Yahoo hafnaði innskráningu. Notaðu Yahoo app-lykilorð í stillingum."
                ) from exc

            status, _ = mailbox.select("INBOX", readonly=True)
            if status != "OK":
                raise RuntimeError("Ekki tókst að opna Yahoo Inbox.")

            status, data = mailbox.uid(
                "search", None, "SINCE", imap_start, "BEFORE", imap_end_exclusive
            )
            if status != "OK":
                raise RuntimeError("Yahoo gat ekki leitað á völdu dagsetningabili.")

            message_uids = data[0].split() if data and data[0] else []
            message_count = len(message_uids)
            emit(f"Found {message_count} matching Inbox message(s).")
            if message_uids:
                emit("Sorting messages by received date (oldest first, newest last)...")
                dated_uids = []
                for sort_number, uid in enumerate(message_uids, start=1):
                    if is_cancelled():
                        raise DownloadCancelled()
                    dated_uids.append((_received_date_key(mailbox, uid), uid))
                    if sort_number % 10 == 0 or sort_number == message_count:
                        emit(f"Reading received dates: {sort_number}/{message_count}")
                message_uids = [uid for _key, uid in sorted(dated_uids, key=lambda item: item[0])]

            for number, uid in enumerate(message_uids, start=1):
                if is_cancelled():
                    raise DownloadCancelled()
                uid_text = uid.decode("ascii", errors="replace")
                status, message_data = mailbox.uid("fetch", uid, "(BODY.PEEK[])")
                if status != "OK":
                    failed_messages += 1
                    emit(f"[{number}/{message_count}] Could not download message UID {uid_text}.")
                    continue

                raw_message = next(
                    (item[1] for item in message_data if isinstance(item, tuple)), None
                )
                if raw_message is None:
                    failed_messages += 1
                    emit(f"[{number}/{message_count}] Message UID {uid_text} had no readable body.")
                    continue

                message = BytesParser(policy=policy.default).parsebytes(raw_message)
                subject = readable_header(message.get("Subject")) or "(no subject)"
                sender = readable_header(message.get("From"))
                message_date = readable_header(message.get("Date"))
                saved_for_message = 0

                for part in message.walk():
                    if is_cancelled():
                        raise DownloadCancelled()
                    filename = part.get_filename()
                    if not filename:
                        continue
                    decoded_filename = readable_header(filename)
                    if Path(decoded_filename).suffix.lower() in SKIP_EXTENSIONS:
                        continue
                    payload = part.get_payload(decode=True)
                    if payload is None:
                        continue
                    output_path = unused_path(output_dir, decoded_filename)
                    output_path.write_bytes(payload)
                    attachment_count += 1
                    saved_for_message += 1

                if saved_for_message == 0:
                    skipped_emails.append(
                        f"Message {number}/{message_count} | UID {uid_text}\n"
                        f"Date: {message_date}\n"
                        f"From: {sender}\n"
                        f"Subject: {subject}\n"
                    )
                processed_count += 1
                emit(f"[{number}/{message_count}] Saved {saved_for_message} attachment(s): {subject}")
    except DownloadCancelled:
        cancelled = True
        emit("Download cancelled; no additional messages or attachments will be saved.")
    except Exception:
        # Keep the unique run folder so the user can inspect any files that were
        # successfully downloaded before a connection or message error.
        raise

    report_lines = [
        "Emails with zero downloaded attachments",
        "=======================================",
        "",
    ]
    if skipped_emails:
        report_lines.extend("\n".join(skipped_emails).splitlines())
    else:
        report_lines.append("None")
    # The report is intentionally created after all attachments, so it sorts as
    # the newest-created file in Windows Explorer.
    skipped_path.write_text("\n".join(report_lines) + "\n", encoding="utf-8")

    emit("")
    emit("Summary")
    emit(f"Emails matched: {message_count}")
    emit(f"Emails processed: {processed_count}")
    emit(f"Attachments saved: {attachment_count}")
    emit(f"Emails with zero saved attachments: {len(skipped_emails)}")
    emit(f"Messages that failed: {failed_messages}")
    emit(f"Cancelled: {'yes' if cancelled else 'no'}")
    emit(f"Attachment folder: {output_dir}")
    emit(f"Skipped-email report: {skipped_path}")
    return {
        "output_dir": str(output_dir),
        "skipped_path": str(skipped_path),
        "matched": message_count,
        "messages": processed_count,
        "attachments": attachment_count,
        "skipped": len(skipped_emails),
        "failed": failed_messages,
        "cancelled": cancelled,
    }
