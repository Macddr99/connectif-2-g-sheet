# connectif-2-g-sheets

This CLI (Command Line Interface) allows generating and downloading Data Explorer reports from Connectif, and then uploading them to Google Sheets. It is ideal for automating the data extraction and spreadsheet update process.

## Features

- Generates reports from Connectif using the API.
- Downloads and extracts ZIP files containing the reports in CSV format.
- Uploads CSV data to Google Sheets, allowing both the creation of new sheets and updating existing ones.

## Requirements

- Node.js (v20 or higher)
- A Google account with access to Google Sheets.
- Google Sheets API Enabled
- Google Service Account Credentials API
- A Connectif API key with the scopes "Read" and "Write" of "Exports"

## Installation

1. Install the [NodeJs](https://nodejs.org/en) runtime.

2. Now, from your favourite shell, install the CLI by typing the following command:
```
$ npm install -i connectif-2-g-sheets -g
```

## Usage

The CLI offers several configurable options through the command line:

connectif-2-g-sheets -k <apikey> -c <credentials> -s <spreadsheetid> -r <reportid> -f <filename> -o <fromdate> -t <todate> [options]

Options
- -k, --apikey <apikey>: Connectif API key (required).
- -c, --credentials <credentials>: Path to the Google JSON credentials file (required).
- -s, --spreadsheetid <spreadsheetid>: Destination Google Sheets spreadsheet ID (required).
- -r, --reportid <reportid>: Connectif report ID (required).
- -f, --filename <filename>: Renamed CSV file name (required).
- -o, --fromdate <fromdate>: From date to send report in YYYY-MM-DD format (required).
- -t, --todate <todate>: To date to send report in YYYY-MM-DD format (required).
- -d, --destination <destination>: Destination path to download and extract the ZIP file (required).
- -n, --newspreadsheet: Add data to a new Google Sheet (optional, default false).
- -a, --append: Append data to an existing spreadsheet without overwriting previous data (optional).
- -e, --existingsheet <existingsheet>: Name of the existing sheet to which data will be appended (optional, requires -a).

Usage Example:
```
connectif-2-g-sheets -k YOUR_CONNECTIF_API_KEY -c ./path/to/credentials.json -s YOUR_SPREADSHEET_ID -r YOUR_REPORT_ID -f report.csv -o 2023-01-01 -t 2023-01-31 -n -d 'C:\Users\YOUR_USER'
```

This command will generate a report for January 2023, download and extract it, and then upload the data to a new sheet in Google Sheets.

```
connectif-2-g-sheets -k YOUR_CONNECTIF_API_KEY -c ./path/to/credentials.json -s YOUR_SPREADSHEET_ID -r YOUR_REPORT_ID -f report.csv -o 2023-01-01 -t 2023-01-31 -a -e SHEET_NAME -d 'C:\Users\YOUR_USER'
```
This command will generate a report for January 2023, download it, extract it and then upload it to an existing sheet called SHEET_NAME adding the report data.



## Notes
- The range between fromDate and toDate cannot be greater than 31 days.
- Ensure that the Google credentials file has the necessary permissions to access and modify Google Sheets.
- Ensure that the Google Sheets file is shared with the email address of the service account associated to the Google Sheets API Key.
- The script automatically handles downloading and extracting the ZIP file from Connectif.
