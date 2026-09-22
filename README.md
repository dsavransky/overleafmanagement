# overleafmanagement

Tools for keeping an Overleaf group's membership in sync with a Google Sheets bundle.

## Installation

If you want to use a venv:
```
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Then, install the package:

```
pip install .
```

Use `pip install -e .` for a development install. Google access uses
`gspread.oauth()`, so [OAuth credentials must be configured for gspread](https://docs.gspread.org/en/v6.0.0/oauth2.html), and the
Overleaf steps require Firefox.
On MacOS, you'll have to grant full disk access to the calling application (terminal, VSCode, etc) to allow selenium to automate Firefox.

Customize .env.example and copy to .env if you want to set your own default group or other settings.

## Commands

- `overleaf-update-accounts [-s SECOND_SHEET_NAME]`: download the current Overleaf
  group members export and write it to the `Current Accounts` tab of the bundle sheet,
  and also of a second sheet if `-s` is given.
  Will truncate the list if there are less members than before. 
- `overleaf-account-deltas PREV_SHEET_NAME [-c CURR_SHEET_NAME]`: compare the
  `All Accounts` tabs of two sheets and write `account_deltas.xlsx` listing accounts
  to delete and add. Useful for comparing membership across years.
- `overleaf-account-deltas -w WORKBOOK_NAME -p PREV_TAB_NAME -t CURR_TAB_NAME`:
  same, but compare the `Email` columns of two tabs within a single sheet.
  Useful for comparing differences between Current and All.
  `overleaf-account-deltas -a OLD_WORKBOOK_NAME -b NEW_WORKBOOK_NAME -p PREV_TAB_NAME -t CURR_TAB_NAME`:
  as above, but allows specifying arbitrary tabs in different workbooks.
  Useful for comparing differences for a given group across years.

(The default user export from Overleaf is lower case - overleaf-update-accounts changes it to match the existing header in the bundle, since python dictionaries are case sensitive.)

- `overleaf-update-members [-i INPUT]`: apply the Delete/Add lists in
  `account_deltas.xlsx` to the Overleaf group members page.

Each is also runnable as `python -m overleafmanagement.<module>`.

## License

MIT



----

 

