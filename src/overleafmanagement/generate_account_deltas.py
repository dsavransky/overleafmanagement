"""Compute account deltas between two Overleaf bundle sheets.

Compares the ``Email`` column of two worksheets and writes an
``account_deltas.xlsx`` workbook listing which accounts need to be deleted
(present before, gone now) and which need to be added (new since before).
The two worksheets come from either of two sources:

* the ``All Accounts`` worksheet of two separate Google Sheets workbooks (a
  previous snapshot and a current one), or
* two named worksheets within a single Google Sheets workbook.

Requires the ``openpyxl`` package as pandas's .xlsx write engine.
"""

import argparse
import re
from typing import Dict, List, Tuple

import gspread
import pandas as pd

ALL_ACCOUNTS_TITLE = "All Accounts"
DEFAULT_CURR_SHEET_NAME = "Revised Overleaf Bundle"
OUTPUT_FILENAME = "account_deltas.xlsx"
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(email: str) -> bool:
    """Check whether email looks like a syntactically valid address.

    Args:
        email (str):
            The candidate string to validate, as sourced from a worksheet
            cell.

    Returns:
        bool:
            True if email matches the simple ``local@domain.tld`` pattern,
            False otherwise.
    """
    return bool(EMAIL_PATTERN.match(email.strip()))


def get_worksheet_emails(
    spreadsheet: gspread.Spreadsheet, worksheet_title: str = ALL_ACCOUNTS_TITLE
) -> List[str]:
    """Extract valid email addresses from a worksheet's Email column.

    Args:
        spreadsheet (gspread.Spreadsheet):
            The workbook to read from.
        worksheet_title (str):
            Title of the worksheet to read, defaulting to "All Accounts".

    Returns:
        List[str]:
            Valid email addresses in original casing/order, with blank or
            malformed entries discarded.

    Raises:
        gspread.exceptions.WorksheetNotFound:
            If no worksheet titled worksheet_title exists in spreadsheet.
    """
    worksheet = spreadsheet.worksheet(worksheet_title)
    records = worksheet.get_all_records()
    emails = [str(record.get("Email", "")) for record in records]
    return [email for email in emails if is_valid_email(email)]


def build_case_insensitive_index(emails: List[str]) -> Dict[str, str]:
    """Map lowercased emails to an original-cased representative.

    Args:
        emails (List[str]):
            Email addresses, in original casing.

    Returns:
        Dict[str, str]:
            Mapping of lowercase email to the first original-cased
            occurrence.
    """
    index: Dict[str, str] = {}
    for email in emails:
        key = email.strip().lower()
        index.setdefault(key, email.strip())
    return index


def compute_deltas(
    prev_emails: List[str], curr_emails: List[str]
) -> Tuple[List[str], List[str]]:
    """Compute case-insensitive email set differences between two sources.

    Args:
        prev_emails (List[str]):
            Emails from the previous sheet's All Accounts.
        curr_emails (List[str]):
            Emails from the current sheet's All Accounts.

    Returns:
        tuple:
            to_delete (List[str]):
                Prev-cased emails present in prev but absent from curr
                (case-insensitively).
            to_add (List[str]):
                Curr-cased emails present in curr but absent from prev
                (case-insensitively).
    """
    prev_index = build_case_insensitive_index(prev_emails)
    curr_index = build_case_insensitive_index(curr_emails)

    delete_keys = prev_index.keys() - curr_index.keys()
    add_keys = curr_index.keys() - prev_index.keys()

    to_delete = sorted(prev_index[key] for key in delete_keys)
    to_add = sorted(curr_index[key] for key in add_keys)
    return to_delete, to_add


def write_deltas_xlsx(
    to_delete: List[str], to_add: List[str], output_path: str = OUTPUT_FILENAME
) -> None:
    """Write delete/add email lists to a two-sheet .xlsx workbook.

    Args:
        to_delete (List[str]):
            Emails to write to the "Delete" sheet.
        to_add (List[str]):
            Emails to write to the "Add" sheet.
        output_path (str):
            Destination file path, defaulting to OUTPUT_FILENAME in the
            current working directory.

    Raises:
        ImportError:
            If no Excel writer engine (e.g. openpyxl) is installed.
    """
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame({"Email": to_delete}).to_excel(
            writer, sheet_name="Delete", index=False
        )
        pd.DataFrame({"Email": to_add}).to_excel(
            writer, sheet_name="Add", index=False
        )


def main() -> None:
    """Compute and write account deltas between two Overleaf bundle sheets.

    Two modes are selected from the command line:

    * Two workbooks (default): the previous workbook name is a positional
      argument; the current workbook name is set with ``-c/--curr_sheet_name``
      and defaults to DEFAULT_CURR_SHEET_NAME. The ``All Accounts`` worksheet
      of each is compared.
    * One workbook: ``-w/--workbook_name`` names a single workbook, and
      ``-p/--prev_tab_name`` and ``-t/--curr_tab_name`` name the two worksheets
      within it to compare.

    Raises:
        gspread.exceptions.SpreadsheetNotFound:
            If a workbook name cannot be opened.
        gspread.exceptions.WorksheetNotFound:
            If a required worksheet does not exist in its workbook.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Compute Overleaf account deltas between two sheets, either in two "
            "separate workbooks (prev_sheet_name [-c]) or in two tabs of one "
            "workbook (-w -p -t)."
        )
    )
    parser.add_argument(
        "prev_sheet_name",
        type=str,
        nargs="?",
        help=(
            "Name of the previous Google Sheets workbook to compare from "
            f"(compares the {ALL_ACCOUNTS_TITLE!r} tab of each workbook)."
        ),
    )
    parser.add_argument(
        "-c",
        "--curr_sheet_name",
        type=str,
        default=None,
        help=(
            "Name of the current Google Sheets workbook to compare against "
            f"(default: {DEFAULT_CURR_SHEET_NAME!r}). Not valid with -w."
        ),
    )
    parser.add_argument(
        "-w",
        "--workbook_name",
        type=str,
        default=None,
        help=(
            "Name of a single Google Sheets workbook whose two tabs "
            "(-p and -t) are compared."
        ),
    )
    parser.add_argument(
        "-p",
        "--prev_tab_name",
        type=str,
        default=None,
        help="Name of the previous tab to compare from. Requires -w.",
    )
    parser.add_argument(
        "-t",
        "--curr_tab_name",
        type=str,
        default=None,
        help="Name of the current tab to compare against. Requires -w.",
    )
    args = parser.parse_args()

    if args.workbook_name is None:
        if args.prev_sheet_name is None:
            parser.error("provide either prev_sheet_name or -w/--workbook_name")
        if args.prev_tab_name is not None or args.curr_tab_name is not None:
            parser.error("-p/--prev_tab_name and -t/--curr_tab_name require -w")
    else:
        if args.prev_sheet_name is not None or args.curr_sheet_name is not None:
            parser.error(
                "-w/--workbook_name cannot be combined with prev_sheet_name or "
                "-c/--curr_sheet_name"
            )
        if args.prev_tab_name is None or args.curr_tab_name is None:
            parser.error("-w/--workbook_name requires both -p and -t")

    gc = gspread.oauth()
    if args.workbook_name is None:
        prev_spreadsheet = gc.open(args.prev_sheet_name)
        curr_spreadsheet = gc.open(args.curr_sheet_name or DEFAULT_CURR_SHEET_NAME)
        prev_emails = get_worksheet_emails(prev_spreadsheet)
        curr_emails = get_worksheet_emails(curr_spreadsheet)
    else:
        spreadsheet = gc.open(args.workbook_name)
        prev_emails = get_worksheet_emails(spreadsheet, args.prev_tab_name)
        curr_emails = get_worksheet_emails(spreadsheet, args.curr_tab_name)

    to_delete, to_add = compute_deltas(prev_emails, curr_emails)
    write_deltas_xlsx(to_delete, to_add)

    print(
        f"Wrote {len(to_delete)} deletion(s) and {len(to_add)} addition(s) "
        f"to {OUTPUT_FILENAME}"
    )


if __name__ == "__main__":
    main()
