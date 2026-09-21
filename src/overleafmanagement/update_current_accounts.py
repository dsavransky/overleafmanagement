"""Download the current Overleaf group members export into a DataFrame.

Overleaf's group members export requires an authenticated browser session,
so this module uses Selenium only to let the user log in interactively in a
real Firefox window. Once logged in, the session cookies are handed off to
``requests`` to fetch the export CSV directly, which is downloaded to a
temporary file and loaded into a pandas DataFrame.
"""

import os
import tempfile

import gspread
import pandas as pd
import requests
from selenium import webdriver

GROUP_ID = "5ba29530a9a3c57d4039f59d"
BASE_URL = f"https://www.overleaf.com/manage/groups/{GROUP_ID}/members"
EXPORT_URL = f"{BASE_URL}/export"
SHEET_NAME = "Revised Overleaf Bundle"


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
    Accounts" worksheet of the bundle sheet.
    """
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

    gc = gspread.oauth()
    bundle = gc.open(SHEET_NAME)
    sheets = bundle.worksheets()

    # find Current Accounts tab
    for sheet in sheets:
        if sheet.title == "Current Accounts":
            break
    assert sheet.title == "Current Accounts", "Could not find Current Accounts in sheet"

    sheet.update([members.columns.values.tolist()] + members.fillna("").values.tolist())


if __name__ == "__main__":
    main()
