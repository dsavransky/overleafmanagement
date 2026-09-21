# overleafmanagement

Tools for keeping an Overleaf group's membership in sync with a Google Sheets bundle.

## Installation

```
pip install .
```

Use `pip install -e .` for a development install. Google access uses
`gspread.oauth()`, so OAuth credentials must be configured for gspread, and the
Overleaf steps require Firefox.

## Commands

- `overleaf-update-accounts [-s SECOND_SHEET_NAME]`: download the current Overleaf
  group members export and write it to the `Current Accounts` tab of the bundle sheet,
  and also of a second sheet if `-s` is given.
- `overleaf-account-deltas PREV_SHEET_NAME [-c CURR_SHEET_NAME]`: compare the
  `All Accounts` tabs of two sheets and write `account_deltas.xlsx` listing accounts
  to delete and add.
- `overleaf-account-deltas -w WORKBOOK_NAME -p PREV_TAB_NAME -t CURR_TAB_NAME`:
  same, but compare the `Email` columns of two tabs within a single sheet.
- `overleaf-update-members [-i INPUT]`: apply the Delete/Add lists in
  `account_deltas.xlsx` to the Overleaf group members page.

Each is also runnable as `python -m overleafmanagement.<module>`.

## License

MIT
