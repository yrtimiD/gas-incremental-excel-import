# GAS Excel importer

Minimal Google Apps Script that reads Excel files from a configured Drive folder
and appends their rows into a target Google Sheet.

Setup
- Open the Apps Script project and paste `Code.gs`.
- Enable the Drive advanced service: Resources → Advanced Google services → Drive API (turn on).
- If your Apps Script project is connected to a Google Cloud project, make sure the Drive API is also enabled there for that project.
- When you run the script for the first time, approve the authorization prompts so Apps Script can access your Google Drive files and Google Sheets.
- Configure constants at the top of `Code.gs`: `FOLDER_ID`, `TARGET_SHEET_ID`, `TARGET_SHEET_NAME`.
- (Optional) Set `PROCESSED_FOLDER_ID` to move processed files after import.
- If you use a time-based trigger, grant trigger permission when prompted.
- The source folder, target spreadsheet, and processed folder should all be accessible with the same Google account that runs the script.

Run
- Run `main()` manually or add a time-based trigger.

Filename/date rules
- Filenames must include a date in `YYYY-MM-DD` or `YYYYMMDD` form; the script uses
  the first match it finds. The sheet is expected to have a date in the first column
  of the last row; the script only imports files whose filename date is newer.
