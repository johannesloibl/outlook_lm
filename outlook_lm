# /// script
# requires-python = ">=3.12"
# dependencies = [
#     "extract-msg>=0.54.1",
#     "markdown-pdf>=1.7",
#     "pywin32>=310",
# ]
# ///
"""Script to recursively download and extract Outlook emails, saving them as Markdown and PDF.
The PDF may be used as a knowledge base for a language model or Google's NotebookLM.

- Downloads all messages from a specified Outlook folder (and subfolders) as .msg files into a temporary directory.
- Extracts message content and metadata, writing all emails as Markdown, reflecting the folder structure.
- Outputs a Markdown file and a PDF file with the extracted emails.
- Supports filtering by minimum received date.
- Designed for CLI and programmatic use.

Usage:
    uv run download_and_extract.py --account <ACCOUNT> --folder <FOLDER_PATH> [--min-date YYYY-MM-DD] [--output OUTPUT_PATH]

"""

import hashlib
import json
import logging
import re
import tempfile
from datetime import UTC, datetime
from pathlib import Path

import extract_msg

# Use the correct constant for .msg format
MSG_FORMAT = 3  # olMsg (not always present in constants)


def process_and_extract_outlook_folder(
    outlook_folder,
    min_date: datetime | None = None,
    output_path: Path | None = None,
) -> None:
    """Recursively downloads all messages from an Outlook folder (and subfolders) as .msg files into a temp directory,
    extracts their content, and writes all emails as Markdown, reflecting the folder structure and including all available fields.
    The temp directory is cleaned up automatically.

    :param outlook_folder: The Outlook folder COM object.
    :param min_date: Optional minimum received date for filtering messages.
    :param output_path: Optional output file path for the Markdown file. PDF will use the same base name.
    """
    logger = logging.getLogger(__name__)
    md_lines = ["# Outlook Email Export\n"]
    try:
        with tempfile.TemporaryDirectory(
            prefix="outlooklm_",
            ignore_cleanup_errors=True,
        ) as temp_dir:
            temp_dir_path = Path(temp_dir)

            def process_folder(folder, parent_parts=None) -> None:
                if parent_parts is None:
                    parent_parts = []
                folder_name = folder.Name
                current_parts = [*parent_parts, folder_name]
                logger.info(f"Processing folder: {'/'.join(current_parts)}")
                messages = folder.Items
                # Sort and restrict if min_date is set
                messages.Sort("[ReceivedTime]", True)
                if min_date is not None:
                    filter_str = (
                        f"[ReceivedTime] >= '{min_date.strftime('%m/%d/%Y %H:%M')}'"
                    )
                    logger.info(f"Applying date filter: {filter_str}")
                    messages = messages.Restrict(filter_str)
                message = messages.GetFirst()
                i = 1
                while message is not None:
                    try:
                        msg_date = message.ReceivedTime
                        logger.debug(
                            f"Processing message {i} in folder {'/'.join(current_parts)}: Received {msg_date}",
                        )
                        if min_date is None or (
                            hasattr(msg_date, "utctimetuple")
                            and datetime.fromtimestamp(
                                msg_date.timestamp(),
                                tz=UTC,
                            )
                            >= min_date
                        ):
                            # Save to temp file
                            safe_subject = "".join(
                                c
                                for c in (message.Subject or f"message_{i}")
                                if c.isalnum() or c in (" ", "_", "-")
                            ).rstrip()
                            # Use a hash of the message entry ID for the filename to ensure uniqueness and avoid issues with special characters
                            entry_id = (
                                message.EntryID
                                if hasattr(message, "EntryID")
                                else f"{safe_subject}_{i}"
                            )
                            filename_hash = hashlib.sha1(
                                str(entry_id).encode(),
                            ).hexdigest()
                            filename = f"{filename_hash}.msg"
                            filepath = temp_dir_path / filename
                            logger.debug(f"Saving message to {filepath}")
                            message.SaveAs(str(filepath), MSG_FORMAT)
                            # Extract and add to markdown
                            msg = extract_msg.openMsg(filepath)
                            if isinstance(msg, extract_msg.Message):
                                msg_as_dict = json.loads(msg.getJson())
                                # --- Markdown heading logic ---
                                folder_parts = current_parts
                                subject = msg_as_dict.get("subject", "(No Subject)")
                                hash_short = hashlib.sha1(
                                    str(filepath).encode(),
                                ).hexdigest()[:8]
                                # Level 2 heading for each mail
                                md_lines.append(f"## {subject} [{hash_short}]")
                                logger.info(
                                    f"Added message: {subject} [{hash_short}] from {'/'.join(folder_parts)}",
                                )
                                # --- Markdown fields ---
                                # Add folder as a field, including parent parts
                                full_folder_path = [*parent_parts, folder_name]
                                md_lines.append(
                                    f"- **folder**: {' - '.join(full_folder_path)}",
                                )
                                for k, v in msg_as_dict.items():
                                    if k.lower() == "body":
                                        body = (
                                            str(v)
                                            .replace("\r\n", "\n")
                                            .replace("\r", "\n")
                                        )
                                        body = re.sub(
                                            r"(\n\s*){2,}",
                                            "\n\n",
                                            body.strip(),
                                        )
                                        indented_body = "\n".join(
                                            f"    {line}" for line in body.splitlines()
                                        )
                                        md_lines.append(f"- **{k}**:\n\n```")
                                        md_lines.append(indented_body)
                                        md_lines.append("```")
                                    else:
                                        md_lines.append(f"- **{k}**: {v}")
                                md_lines.append("")

                    except Exception as e:
                        if "private key" in str(e):
                            pass
                        else:
                            logger.exception(
                                f"Error processing message {i} ({safe_subject}) in folder {'/'.join(current_parts)}: {e}",
                                exc_info=False,
                            )
                    message = messages.GetNext()
                    i += 1
                # Recurse into subfolders
                for subfolder in folder.Folders:
                    logger.info(
                        f"Descending into subfolder: {subfolder.Name} of {'/'.join(current_parts)}",
                    )
                    process_folder(subfolder, current_parts)

            process_folder(outlook_folder)
    finally:
        if output_path is None:
            md_file = Path("outlook_emails.md").absolute()
        else:
            md_file = Path(output_path).absolute()
        md_content = "\n".join(md_lines)
        md_file.parent.mkdir(parents=True, exist_ok=True)
        md_file.write_text(md_content, encoding="utf-8")

        from markdown_pdf import MarkdownPdf, Section

        pdf = MarkdownPdf()
        pdf.meta["title"] = "Title"
        pdf.add_section(Section(md_content, toc=False))
        pdf.save(md_file.with_suffix(".pdf").absolute())


def get_outlook_folder(account_name: str, folder_path: str):
    """Returns the Outlook folder object for a given account and __-separated folder path.
    :param account_name: The Outlook account name/email.
    :param folder_path: The __-separated folder path (e.g., 'Inbox__Subfolder1__Subfolder2').
    :return: The Outlook folder COM object.
    """
    from win32com.client import Dispatch

    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    account = outlook.Folders.Item(account_name)
    folder = account
    for part in folder_path.split("__"):
        folder = folder.Folders.Item(part)
    return folder


def cli() -> None:
    """CLI entry point: parses arguments and runs the download/extract process.

    Adds --output option to specify the output Markdown/PDF file path (PDF will use the same base name).
    """
    import argparse

    from win32com.client import Dispatch

    logging.basicConfig(level=logging.INFO)
    parser = argparse.ArgumentParser(
        description="Download and extract Outlook emails to Markdown.",
    )
    parser.add_argument("--account", required=True, help="Outlook account name/email")
    parser.add_argument(
        "--folder",
        required=True,
        help="Folder path, e.g. 'Inbox/Subfolder1/Subfolder2'",
    )
    parser.add_argument(
        "--min-date",
        type=str,
        default=None,
        help="Minimum date (YYYY-MM-DD) for messages",
    )
    parser.add_argument(
        "--output",
        type=str,
        default=None,
        help="Output Markdown file path (PDF will use the same base name)",
    )
    args = parser.parse_args()
    outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
    account = outlook.Folders.Item(args.account)
    # Support /-separated folder path
    folder = account
    for part in args.folder.split("/"):
        folder = folder.Folders.Item(part)
    min_date = None
    if args.min_date:
        min_date = datetime.strptime(args.min_date, "%Y-%m-%d").replace(
            tzinfo=UTC,
        )
    output_path = Path(args.output) if args.output else None
    process_and_extract_outlook_folder(
        folder, min_date=min_date, output_path=output_path
    )


def example_debug_run() -> None:
    """Example function for debugging: runs the extraction for a hardcoded account and folder.
    Adjust the account, folder, and min_date as needed for your environment.
    """
    logging.basicConfig(level=logging.INFO)
    account_name = "foo.bar@abc.com"
    folder_path = "Inbox"
    min_date = datetime(2025, 1, 1, tzinfo=UTC)
    folder = get_outlook_folder(account_name, folder_path)
    process_and_extract_outlook_folder(folder, min_date=min_date)


if __name__ == "__main__":
    cli()
