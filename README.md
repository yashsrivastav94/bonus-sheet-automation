# Bonus Sheet Automation (Google Apps Script)

This script automates the creation of monthly bonus tracking sheets in Google Sheets.
It supports dynamic generation, formula replication, dropdown preservation, and personalized bonus descriptions.

## Features
- Monthly sheet creation
- Data import from multiple sources
- Formula & formatting replication
- Bonus-specific comment generation

## Setup

1. Copy the script to your Google Apps Script project.
2. Replace the placeholders in `config`:
   - `contractorSpreadsheetId`: Your Contractors sheet ID
   - Update sheet names if different.

## How to Use

1. Open the master Google Sheet.
2. Go to Extensions > Apps Script.
3. Paste the script into your project.
4. Run `automateNextBonusSheetImproved()`.

## License
[MIT](LICENSE)
