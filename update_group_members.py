"""Apply account deltas to the Overleaf group members management page.

Reads the ``account_deltas.xlsx`` file produced by
``generate_account_deltas.py`` (sheets ``"Delete"`` and ``"Add"``, each a
single ``Email`` column) and drives an authenticated Selenium session
against the Overleaf group members management page to remove the deleted
accounts (page by page, with user confirmation before each removal) and
invite the added accounts. Reuses the login/URL plumbing already defined in
``update_current_accounts.py``.
"""

import argparse
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from generate_account_deltas import OUTPUT_FILENAME
from update_current_accounts import BASE_URL, wait_for_login

DELETE_SHEET_NAME = "Delete"
ADD_SHEET_NAME = "Add"
EMAIL_COLUMN = "Email"
WAIT_TIMEOUT_SECONDS = 15

TABLE_ROW_SELECTOR = "table.managed-entities-table tbody tr.managed-entity-row"
CHECKBOX_SELECTOR = 'td.cell-checkbox input[data-testid="select-single-checkbox"]'
EMAIL_CELL_SELECTOR = "td.cell-email"
INVITE_INPUT_ID = "add-members-emails"
REMOVE_BUTTON_TEXT = "Remove from group"
INVITE_BUTTON_TEXT = "Invite"


def read_email_lists(input_path: str) -> Tuple[List[str], List[str]]:
    """Read the Delete/Add email lists from an account deltas workbook.

    Args:
        input_path (str):
            Path to the ``.xlsx`` file produced by
            ``generate_account_deltas.write_deltas_xlsx``, containing
            "Delete" and "Add" sheets with a single "Email" column each.

    Returns:
        tuple:
            delete_emails (List[str]):
                Emails to remove from the group, as read from the
                "Delete" sheet.
            add_emails (List[str]):
                Emails to invite to the group, as read from the "Add"
                sheet.

    Raises:
        FileNotFoundError:
            If input_path does not exist.
        ValueError:
            If either sheet is missing from the workbook.
    """
    delete_emails = (
        pd.read_excel(input_path, sheet_name=DELETE_SHEET_NAME)[EMAIL_COLUMN]
        .dropna()
        .astype(str)
        .tolist()
    )
    add_emails = (
        pd.read_excel(input_path, sheet_name=ADD_SHEET_NAME)[EMAIL_COLUMN]
        .dropna()
        .astype(str)
        .tolist()
    )
    return delete_emails, add_emails


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


def get_row_email(driver: webdriver.Firefox, row: WebElement) -> str:
    """Extract the plain email address from a members table row.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, used to run a small script that
            reads only the row's raw email text node (the cell also
            contains a "Pending invite" badge for unaccepted invites,
            which must not be included in the result).
        row (WebElement):
            A ``tr.managed-entity-row`` element from the members table.

    Returns:
        str:
            The row's email address, with any "Pending invite" badge
            text excluded.
    """
    cell = row.find_element(By.CSS_SELECTOR, EMAIL_CELL_SELECTOR)
    email = driver.execute_script(
        "const cell = arguments[0];"
        "for (const node of cell.childNodes) {"
        "  if (node.nodeType === 3) {"
        "    const text = node.textContent.trim();"
        "    if (text) { return text; }"
        "  }"
        "}"
        "return null;",
        cell,
    )
    if not email:
        email = cell.text.split()[0]
        print(
            "Warning: falling back to whitespace-split email parsing for a "
            f"row; got {cell.text!r}"
        )
    return email


def find_button_by_text(
    driver: webdriver.Firefox, css_selector: str, text: str
) -> Optional[WebElement]:
    """Find a button matching css_selector whose visible text equals text.

    Neither the "Remove from group" nor the "Invite" button exposes a
    stable ``id``/``data-testid``, so both are located by exact text match.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver to search within.
        css_selector (str):
            CSS selector narrowing the candidate buttons (e.g. by class).
        text (str):
            Exact, stripped visible text to match against each candidate.

    Returns:
        Optional[WebElement]:
            The first matching button, or None if no button matches.
    """
    for element in driver.find_elements(By.CSS_SELECTOR, css_selector):
        if element.text.strip() == text:
            return element
    return None


def prompt_yes_no(message: str) -> bool:
    """Ask the user a yes/no question on the console.

    Args:
        message (str):
            The question to display, without a trailing prompt suffix.

    Returns:
        bool:
            True if the user answered "y"/"yes" (case-insensitive), False
            for any other input, including a blank response.
    """
    response = input(f"{message} [y/N]: ").strip().lower()
    return response in {"y", "yes"}


def find_matches_on_page(
    driver: webdriver.Firefox,
    remaining_lower: Set[str],
    lower_to_original: Dict[str, str],
) -> List[Tuple[WebElement, str]]:
    """Find rows on the current page matching a set of target emails.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, positioned on a members page.
        remaining_lower (Set[str]):
            Lowercased target emails not yet accounted for.
        lower_to_original (Dict[str, str]):
            Mapping of lowercase email to original casing, as returned by
            build_case_insensitive_index.

    Returns:
        List[Tuple[WebElement, str]]:
            One (row, lowercased email) pair per row on this page whose
            email is in remaining_lower.
    """
    WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, TABLE_ROW_SELECTOR))
    )
    matches = []
    for row in driver.find_elements(By.CSS_SELECTOR, TABLE_ROW_SELECTOR):
        email_lower = get_row_email(driver, row).strip().lower()
        if email_lower in remaining_lower:
            matches.append((row, email_lower))
    return matches


def check_row_checkboxes(
    driver: webdriver.Firefox, matches: List[Tuple[WebElement, str]]
) -> None:
    """Check the selection checkbox for each matched row.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver.
        matches (List[Tuple[WebElement, str]]):
            Rows to select, as returned by find_matches_on_page.
    """
    for row, _ in matches:
        checkbox = row.find_element(By.CSS_SELECTOR, CHECKBOX_SELECTOR)
        WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
            EC.element_to_be_clickable(checkbox)
        )
        if not checkbox.is_selected():
            checkbox.click()


def click_remove_from_group_button(driver: webdriver.Firefox) -> None:
    """Click the "Remove from group" button and wait for it to disappear.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, with at least one row checkbox
            already checked so the button is present in the DOM.

    Raises:
        RuntimeError:
            If the button does not appear within WAIT_TIMEOUT_SECONDS,
            indicating the checkbox selection was not registered.
    """
    try:
        button = WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
            lambda d: find_button_by_text(
                d, "button.btn.btn-danger", REMOVE_BUTTON_TEXT
            )
        )
    except TimeoutException as exc:
        raise RuntimeError(
            "Remove from group button did not appear after selecting "
            "checkboxes; the page may not have registered the selection."
        ) from exc
    button.click()
    WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(EC.staleness_of(button))


def get_next_page_button(driver: webdriver.Firefox) -> Optional[WebElement]:
    """Find the pagination button that advances to the next page.

    Note:
        The fallback "»" control's exact next-vs-last semantics were not
        verified against the live page; this assumes it advances one page
        at a time, matching typical pagination widgets. Verify this live
        before relying on it for a Delete list long enough to need it.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, positioned on a members page.

    Returns:
        Optional[WebElement]:
            The button that navigates to the next page, or None if the
            current page is the last one.

    Raises:
        RuntimeError:
            If the current page number cannot be determined from the
            pagination controls.
    """
    buttons = driver.find_elements(By.CSS_SELECTOR, "button[aria-label]")
    current_page = None
    for button in buttons:
        aria_label = button.get_attribute("aria-label") or ""
        if aria_label.endswith(", Current Page"):
            try:
                current_page = int(aria_label.split(",")[0].replace("Page", "").strip())
            except ValueError:
                continue
            break
    if current_page is None:
        raise RuntimeError("Could not determine the current pagination page.")

    next_label = f"Go to page {current_page + 1}"
    for button in buttons:
        if button.get_attribute("aria-label") == next_label:
            return button

    return find_button_by_text(driver, "button", "»")


def go_to_next_page(driver: webdriver.Firefox, next_button: WebElement) -> None:
    """Click a pagination button and wait for the table to refresh.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver.
        next_button (WebElement):
            The pagination button to click, as returned by
            get_next_page_button.
    """
    rows = driver.find_elements(By.CSS_SELECTOR, TABLE_ROW_SELECTOR)
    anchor = rows[0] if rows else None
    WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
        EC.element_to_be_clickable(next_button)
    )
    next_button.click()
    if anchor is not None:
        WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(EC.staleness_of(anchor))
    WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, TABLE_ROW_SELECTOR))
    )


def process_delete_flow(
    driver: webdriver.Firefox, delete_emails: List[str]
) -> Tuple[List[str], List[str], List[str]]:
    """Remove the given emails from the group, page by page, with prompts.

    Walks the members table one page at a time. On each page, any rows
    matching delete_emails are checked and the user is asked to confirm
    before "Remove from group" is clicked; a decline is recorded and never
    retried, since each account appears in exactly one row across the
    whole table. Stops once every email has been accounted for or once
    pagination is exhausted.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, already on the members page.
        delete_emails (List[str]):
            Emails to remove from the group.

    Returns:
        tuple:
            removed (List[str]):
                Emails the user confirmed and were removed, sorted.
            declined (List[str]):
                Emails found on a page but the user declined to remove,
                sorted.
            not_found (List[str]):
                Emails never encountered on any page, sorted.
    """
    lower_to_original = build_case_insensitive_index(delete_emails)
    remaining: Set[str] = set(lower_to_original)
    removed: List[str] = []
    declined: List[str] = []

    while True:
        matches = find_matches_on_page(driver, remaining, lower_to_original)
        if matches:
            matched_originals = sorted(
                lower_to_original[email_lower] for _, email_lower in matches
            )
            print("Found the following account(s) to remove on this page:")
            for email in matched_originals:
                print(f"  {email}")
            if prompt_yes_no(
                f"Remove these {len(matched_originals)} account(s) from the group?"
            ):
                check_row_checkboxes(driver, matches)
                click_remove_from_group_button(driver)
                for _, email_lower in matches:
                    remaining.discard(email_lower)
                    removed.append(lower_to_original[email_lower])
                print(f"Removed {len(matched_originals)} account(s).")
            else:
                for _, email_lower in matches:
                    remaining.discard(email_lower)
                    declined.append(lower_to_original[email_lower])
                print("Skipped removal for this page's matches.")

        if not remaining:
            break
        next_button = get_next_page_button(driver)
        if next_button is None:
            break
        go_to_next_page(driver, next_button)

    not_found = sorted(lower_to_original[email_lower] for email_lower in remaining)
    return sorted(removed), sorted(declined), not_found


def process_add_flow(driver: webdriver.Firefox, add_emails: List[str]) -> None:
    """Invite the given emails to the group via the invite textbox.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, already on the members page.
        add_emails (List[str]):
            Emails to invite to the group. If empty, no action is taken.

    Raises:
        RuntimeError:
            If the Invite button cannot be located within
            WAIT_TIMEOUT_SECONDS.
    """
    if not add_emails:
        print("No accounts to invite.")
        return

    joined = ",".join(add_emails)
    input_box = WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
        EC.presence_of_element_located((By.ID, INVITE_INPUT_ID))
    )
    input_box.clear()
    input_box.send_keys(joined)

    try:
        invite_button = WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
            lambda d: find_button_by_text(
                d, "button.btn.btn-primary", INVITE_BUTTON_TEXT
            )
        )
    except TimeoutException as exc:
        raise RuntimeError("Could not locate the Invite button.") from exc
    invite_button.click()
    print(f"Invited {len(add_emails)} account(s): {joined}")


def main(input_path: str) -> None:
    """Apply the Delete/Add lists from input_path to the Overleaf group.

    Args:
        input_path (str):
            Path to the account deltas ``.xlsx`` file, as produced by
            generate_account_deltas.write_deltas_xlsx.
    """
    delete_emails, add_emails = read_email_lists(input_path)
    print(
        f"Loaded {len(delete_emails)} account(s) to delete and "
        f"{len(add_emails)} account(s) to add."
    )

    driver = webdriver.Firefox()
    try:
        wait_for_login(driver, BASE_URL)
        removed, declined, not_found = process_delete_flow(driver, delete_emails)
        process_add_flow(driver, add_emails)
    finally:
        driver.quit()

    print("\nSummary")
    print(
        f"  Removed ({len(removed)}): "
        f"{', '.join(removed) if removed else 'none'}"
    )
    print(
        f"  Declined ({len(declined)}): "
        f"{', '.join(declined) if declined else 'none'}"
    )
    print(
        f"  Not found ({len(not_found)}): "
        f"{', '.join(not_found) if not_found else 'none'}"
    )


def _parse_args() -> argparse.Namespace:
    """Parse command-line arguments for this script.

    Returns:
        argparse.Namespace:
            Parsed arguments with a single ``input`` path.
    """
    parser = argparse.ArgumentParser(
        description="Apply account deltas to the Overleaf group members page."
    )
    parser.add_argument(
        "-i",
        "--input",
        type=str,
        default=OUTPUT_FILENAME,
        help=(
            "Path to the account deltas .xlsx file produced by "
            f"generate_account_deltas.py (default: {OUTPUT_FILENAME!r})."
        ),
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    main(args.input)
