"""Download the current Overleaf group members export into Google Sheets.

Overleaf's group members export requires an authenticated browser session,
so this module uses Selenium only to let the user log in interactively in a
real Firefox window. Once logged in, the session cookies are handed off to
``requests`` to fetch the export CSV directly, which is downloaded to a
temporary file and loaded into a pandas DataFrame. The result is written to the
``Current Accounts`` worksheet of the default bundle sheet and, optionally, of a
second user-supplied Google Sheets workbook.
"""

import argparse
import os
import tempfile

import dotenv
import gspread
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.firefox.options import Options

dotenv.load_dotenv()
GROUP_ID = os.getenv("GROUP_ID", "5ba29530a9a3c57d4039f59d")
SHEET_NAME = os.getenv("SHEET_NAME", "Revised Overleaf Bundle")
CURRENT_ACCOUNTS_TITLE = os.getenv("CURRENT_ACCOUNTS_TITLE", "Current Accounts")
BASE_URL = f"https://www.overleaf.com/manage/groups/{GROUP_ID}/members"
EXPORT_URL = f"{BASE_URL}/export"

def wait_for_login(driver: webdriver.Firefox, url: str) -> None:
    """Navigate to url and block until the user confirms they are logged in.

    Args:
        driver (webdriver.Firefox):
            The Selenium Firefox driver to navigate.
        url (str):
            The URL to open, which Overleaf will redirect to its login page
            for if the browser session is not already authenticated.

    Example:
        >>> driver = webdriver.Firefox()
        >>> wait_for_login(driver, BASE_URL)
    """
    driver.get(url)
    input(
        "Log in to Overleaf in the browser window, then press Enter here to continue..."
    )


def build_authenticated_session(driver: webdriver.Firefox) -> requests.Session:
    """Build a requests session carrying the driver's authenticated cookies.

    Args:
        driver (webdriver.Firefox):
            A Selenium Firefox driver holding an authenticated Overleaf
            browser session.

    Returns:
        requests.Session:
            A session with the browser's cookies and user agent, usable to
            make authenticated HTTP requests without the browser.
    """
    session = requests.Session()
    for cookie in driver.get_cookies():
        session.cookies.set(cookie["name"], cookie["value"], domain=cookie["domain"])
    session.headers.update(
        {"User-Agent": driver.execute_script("return navigator.userAgent")}
    )
    return session


def download_export(session: requests.Session, url: str, dest_dir: str) -> str:
    """Download the export at url into dest_dir using an authenticated session.

    Args:
        session (requests.Session):
            An authenticated session, as returned by
            :func:`build_authenticated_session`.
        url (str):
            The export URL to download.
        dest_dir (str):
            The directory to write the downloaded file into.

    Returns:
        str:
            The full path to the downloaded file.

    Raises:
        RuntimeError:
            If the response looks like an HTML page (e.g. a login page)
            rather than an export file, which indicates the session was not
            actually authenticated.
    """
    response = session.get(url)
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "")
    if "html" in content_type.lower():
        raise RuntimeError(
            "Expected a file download but received an HTML page; the "
            "session may not be authenticated."
        )

    filename = "members_export.csv"
    disposition = response.headers.get("Content-Disposition", "")
    if "filename=" in disposition:
        filename = disposition.split("filename=")[-1].strip('"; ')

    dest_path = os.path.join(dest_dir, filename)
    with open(dest_path, "wb") as f:
        f.write(response.content)
    return dest_path


def main() -> None:
    """Authenticate, download the group members export, and upload it.

    The export's email and last login columns are written to the "Current
    Accounts" worksheet of the default bundle sheet (SHEET_NAME) and, if
    ``-s/--second_sheet_name`` is given on the command line, of that workbook
    as well. All target worksheets are located before logging in to Overleaf so
    that a bad sheet name fails immediately.

    Raises:
        gspread.exceptions.SpreadsheetNotFound:
            If a target workbook name cannot be opened.
        gspread.exceptions.WorksheetNotFound:
            If a target workbook lacks a "Current Accounts" worksheet.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Update the Current Accounts tab of the bundle sheet from the "
            "Overleaf group members export."
        )
    )
    parser.add_argument(
        "-s",
        "--second_sheet_name",
        type=str,
        default=None,
        help=(
            "Name of an additional Google Sheets workbook whose "
            f"{CURRENT_ACCOUNTS_TITLE!r} tab is also updated "
            f"(in addition to {SHEET_NAME!r})."
        ),
    )
    args = parser.parse_args()

    sheet_names = [SHEET_NAME]
    if args.second_sheet_name is not None and args.second_sheet_name != SHEET_NAME:
        sheet_names.append(args.second_sheet_name)

    gc = gspread.oauth()
    worksheets = [
        gc.open(sheet_name).worksheet(CURRENT_ACCOUNTS_TITLE)
        for sheet_name in sheet_names
    ]

    driver = webdriver.Firefox()
    try:
        wait_for_login(driver, BASE_URL)
        session = build_authenticated_session(driver)
    finally:
        driver.quit()

    with tempfile.TemporaryDirectory() as tmp_dir:
        csv_path = download_export(session, EXPORT_URL, tmp_dir)
        members = pd.read_csv(csv_path)
    members = members[["email", "last_logged_in_at"]]
    # Rename the "email" column to "Email" to match the header in Google Sheets
    members.rename(columns={"email": "Email"}, inplace=True)
    payload = [members.columns.values.tolist()] + members.fillna("").values.tolist()
    for sheet_name, worksheet in zip(sheet_names, worksheets):
        worksheet.update(payload)
        # Check for length mismatch, remove excess if necessary
        data_rows = worksheet.get_all_values()
        curr_member_len = len(data_rows) - 1 # subtract for header row
        new_member_len = members.shape[0]  # doesn't include header row
        if new_member_len < curr_member_len:
            worksheet.batch_clear([f"A{new_member_len + 1}:{curr_member_len + 1}"])
        print(f"Sheet row count: {curr_member_len}, New members count: {new_member_len}")
        print(f"Updated {CURRENT_ACCOUNTS_TITLE!r} in {sheet_name!r}")

if __name__ == "__main__":
    main()
