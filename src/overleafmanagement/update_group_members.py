"""Apply account deltas to the Overleaf group members management page.

Reads the ``account_deltas.xlsx`` file produced by
``generate_account_deltas.py`` (sheets ``"Delete"`` and ``"Add"``, each a
single ``Email`` column) and drives an authenticated Selenium session
against the Overleaf group members management page to remove the deleted
accounts (one at a time, via the members-search box, with user confirmation
before each removal) and invite the added accounts. Reuses the login/URL
plumbing already defined in ``update_current_accounts.py``.
"""

import argparse
from typing import Dict, List, Optional, Tuple

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from overleafmanagement.generate_account_deltas import OUTPUT_FILENAME
from overleafmanagement.update_current_accounts import BASE_URL, wait_for_login

DELETE_SHEET_NAME = "Delete"
ADD_SHEET_NAME = "Add"
EMAIL_COLUMN = "Email"
WAIT_TIMEOUT_SECONDS = 15
MAX_REMOVE_ATTEMPTS = 2

TABLE_ROW_SELECTOR = "table.managed-entities-table tbody tr.managed-entity-row"
CHECKBOX_SELECTOR = 'td.cell-checkbox input[data-testid="select-single-checkbox"]'
EMAIL_CELL_SELECTOR = "td.cell-email"
SEARCH_INPUT_SELECTOR = 'input[data-testid="search-members-input"]'
NO_MEMBERS_XPATH = "//*[normalize-space(text())='No members']"
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


def find_member_row(
    driver: webdriver.Firefox, email_lower: str
) -> Optional[WebElement]:
    """Find the currently displayed row for a lowercased email, if any.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver.
        email_lower (str):
            The target email, already lowercased.

    Returns:
        Optional[WebElement]:
            The matching row, or None if no currently displayed row's
            email matches.
    """
    for row in driver.find_elements(By.CSS_SELECTOR, TABLE_ROW_SELECTOR):
        if get_row_email(driver, row).strip().lower() == email_lower:
            return row
    return None


def search_for_email(
    driver: webdriver.Firefox, email: str
) -> Optional[WebElement]:
    """Search the members list for email and return its row, if found.

    Types email into the members-search box and waits for the filtered
    results to settle, since the search box filters the table in place
    without a page reload.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, already on the members page.
        email (str):
            The email address to search for.

    Returns:
        Optional[WebElement]:
            The row for email once the search results settle, or None if
            the search reports no matching members.
    """
    search_box = WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, SEARCH_INPUT_SELECTOR))
    )
    search_box.clear()
    search_box.send_keys(email)

    email_lower = email.strip().lower()

    def _results_settled(d: webdriver.Firefox):
        row = find_member_row(d, email_lower)
        if row is not None:
            return row
        return bool(d.find_elements(By.XPATH, NO_MEMBERS_XPATH))

    result = WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(_results_settled)
    return result if isinstance(result, WebElement) else None


def check_checkbox(driver: webdriver.Firefox, row: WebElement) -> None:
    """Check a single row's selection checkbox if not already checked.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver.
        row (WebElement):
            The row whose checkbox should be checked.
    """
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
            If the button does not appear within WAIT_TIMEOUT_SECONDS
            (the checkbox selection was not registered), or if it does not
            go stale after being clicked (the click did not register; this
            has been observed live as a transient stuck state in
            Overleaf's page, recoverable by reloading and retrying).
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
    try:
        WebDriverWait(driver, WAIT_TIMEOUT_SECONDS).until(EC.staleness_of(button))
    except TimeoutException as exc:
        raise RuntimeError(
            "Remove from group button did not go stale after being "
            "clicked; the removal likely did not register."
        ) from exc


def remove_email_with_retry(
    driver: webdriver.Firefox, email: str, row: WebElement
) -> None:
    """Check row's checkbox and remove it, retrying once via reload on failure.

    If checking the box and clicking "Remove from group" fails (see
    click_remove_from_group_button), the page is reloaded and email is
    searched for again before a second, final attempt. This mirrors the
    only recovery observed live for the stuck-page failure mode: a full
    page reload.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, already on the members page.
        email (str):
            The original-cased email being removed, used to re-search
            after a retry reload.
        row (WebElement):
            The row located by the initial search_for_email call.

    Raises:
        RuntimeError:
            If removal still fails after MAX_REMOVE_ATTEMPTS attempts, or
            if email can no longer be found after a retry reload.
    """
    for attempt in range(1, MAX_REMOVE_ATTEMPTS + 1):
        try:
            check_checkbox(driver, row)
            click_remove_from_group_button(driver)
            return
        except RuntimeError as exc:
            if attempt == MAX_REMOVE_ATTEMPTS:
                raise RuntimeError(
                    f"Failed to remove {email} after {MAX_REMOVE_ATTEMPTS} "
                    "attempts."
                ) from exc
            print(
                f"Warning: removal attempt {attempt} for {email} failed "
                f"({exc}); reloading the page and retrying."
            )
            driver.get(BASE_URL)
            retried_row = search_for_email(driver, email)
            if retried_row is None:
                raise RuntimeError(
                    f"Could not find {email} again after reloading to "
                    "retry removal."
                ) from exc
            row = retried_row


def process_delete_flow(
    driver: webdriver.Firefox, delete_emails: List[str]
) -> Tuple[List[str], List[str], List[str]]:
    """Remove the given emails from the group, one at a time, with prompts.

    For each email, searches the members list via the search box; if a
    matching row is found, checks it, asks the user to confirm, and clicks
    "Remove from group" on a yes (see remove_email_with_retry for the
    one-reload retry on failure). A no is recorded as declined rather than
    retried.

    Args:
        driver (webdriver.Firefox):
            The active Selenium driver, already on the members page.
        delete_emails (List[str]):
            Emails to remove from the group.

    Raises:
        RuntimeError:
            If removing a confirmed email fails twice in a row (see
            remove_email_with_retry).

    Returns:
        tuple:
            removed (List[str]):
                Emails the user confirmed and were removed, sorted.
            declined (List[str]):
                Emails found but the user declined to remove, sorted.
            not_found (List[str]):
                Emails the search reported no match for, sorted.
    """
    lower_to_original = build_case_insensitive_index(delete_emails)
    removed: List[str] = []
    declined: List[str] = []
    not_found: List[str] = []

    for _, original_email in sorted(
        lower_to_original.items(), key=lambda item: item[1]
    ):
        row = search_for_email(driver, original_email)
        if row is None:
            print(f"Not found: {original_email}")
            not_found.append(original_email)
            continue

        print(f"Found account to remove: {original_email}")
        if prompt_yes_no(f"Remove {original_email} from the group?"):
            remove_email_with_retry(driver, original_email, row)
            removed.append(original_email)
            print(f"Removed {original_email}.")
        else:
            declined.append(original_email)
            print(f"Skipped {original_email}.")

    return sorted(removed), sorted(declined), sorted(not_found)


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


def main() -> None:
    """Apply the Delete/Add lists from an account deltas file to the Overleaf group.

    The path to the account deltas ``.xlsx`` file (as produced by
    generate_account_deltas.write_deltas_xlsx) is read from the command line via
    ``-i/--input``.
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
    args = parser.parse_args()

    delete_emails, add_emails = read_email_lists(args.input)
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


if __name__ == "__main__":
    main()
