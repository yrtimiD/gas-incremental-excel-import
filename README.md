# GAS Excel importer

Minimal Google Apps Script that reads Excel files from a configured Drive folder
and appends their rows into a target Google Sheet.

## Container-bound mode
> script is attached to specific google sheet

### Setup
- Open google sheet
- Extensions > Apps Script
- Paste `Code.gs` and `Config.gs`
- Add `Drive API` service
- Return to sheet and refresh
### Run
- New "Quick Scripts" menu should appear, click it and select "Run..."
- (on the first time run permissions prompt will be shown)


## Stand-alone mode

### Setup
- Open the Apps Script project and paste `Code.gs` and `Config.gs`.
- Enable the Drive advanced service: Resources → Advanced Google services → Drive API (turn on).
- If your Apps Script project is connected to a Google Cloud project, make sure the Drive API is also enabled there for that project.
- When you run the script for the first time, approve the authorization prompts so Apps Script can access your Google Drive files and Google Sheets.
- Configure constants at the top of `Code.gs`: `FOLDER_ID`, `TARGET_SHEET_ID`, `TARGET_SHEET_NAME`.
- (Optional) Set `PROCESSED_FOLDER_ID` to move processed files after import.
- If you use a time-based trigger, grant trigger permission when prompted.
- The source folder, target spreadsheet, and processed folder should all be accessible with the same Google account that runs the script.

### Run
- Run `runImportFromDrive()` manually or add a time-based trigger.

File rules
- Filenames must be string-sortable (e.g. named with dates `YYYY-MM-DD` or padded number `001`, `002`...);
- The script starts import from the row which is exact match of existing last row. If not found - whole file is skipped.
