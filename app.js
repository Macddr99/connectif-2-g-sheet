#!/usr/bin/env node

const axios = require('axios');
const AdmZip = require('adm-zip');
const fs = require('fs');
const csv = require('csv-parser');
const { google } = require('googleapis');
const cliProgress = require('cli-progress');
const { Command } = require('commander');
const path = require('path');

const program = new Command();
program
    .version('1.0.0')
    .description('CLI to generated and download reports from Connectif and upload them to Google Sheets')
    .requiredOption('-k, --apikey <apikey>', 'Connectif API key')
    .requiredOption('-c, --credentials <credentials>', 'Path to the Google JSON credentials file')
    .requiredOption('-s, --spreadsheetid <spreadsheetid>', 'Destination Google Sheets spreadsheet ID')
    .requiredOption('-r, --reportid <reportid>', 'Connectif report ID')
    .requiredOption('-f, --filename <filename>', 'Renamed CSV file name')
    .requiredOption('-o, --fromdate <fromdate>', 'From date to send report YYYY-MM-DD')
    .requiredOption('-t, --todate <todate>', 'To date to send report YYYY-MM-DD')
    .requiredOption('-d, --destination <destination>', 'Destination path to download and extract the ZIP file')
    .option('-n, --newspreadsheet', 'Adding data to a new Google Sheet', false)
    .option('-a, --append', 'Append data to an existing spreadsheet.')
    .option('-e, --existingsheet <existingsheet>', 'Name of the existing spreadsheet to which the data will be appended')

    .parse(process.argv);

const options = program.opts();

// ApiKey & Credentials 
const apiKey = options.apikey;
const credentials = require(path.resolve(options.credentials));

// Options
const spreadsheetId = options.spreadsheetid;
const reportCnId = options.reportid;
const reportCnFileName = options.filename;
const addToNewSpreadsheet = options.newspreadsheet;
const destination = options.destination;
const appendToExistingSheet = options.append;
const existingSheetName = options.existingsheet;

async function generateReport(fromDateReport,toDateReport) {
    try {
        const response = await axios.post(
            'https://api.connectif.cloud/exports/type/data-explorer',
            {
                delimiter: ";",
                filters: {
                    fromDate: fromDateReport,
                    reportId: reportCnId,
                    toDate: toDateReport
                }
            },
            {
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': apiKey
                }
            }
        );

        const reportId = response.data.id;
        return reportId;
    } catch (error) {
        console.error('Error generating report:', error);
        return null;
    }
}

async function getReportFileUrl(reportId) {
    const progressBar = new cliProgress.SingleBar({}, cliProgress.Presets.shades_classic);
    progressBar.start(100, 0);

    let progress = 0;
    const maxWaitTime = 40 * 1000;
    const startTime = Date.now();

    while (true) {
        try {
            const response = await axios.get(
                `https://api.connectif.cloud/exports/${reportId}`,
                {
                    headers: {
                        'Content-Type': 'application/json',
                        'Authorization': apiKey
                    }
                }
            );

            const status = response.data.status;
            const fileUrl = response.data.fileUrl;

            if (status === 'queued') {
                progress = (progress + 1) % 100;
                progressBar.update(progress);
                console.log(`The report status is 'queued'. Waiting 5 second before trying again.`);

                if (Date.now() - startTime > maxWaitTime) {
                    console.error('Timed out while getting report file URL.');
                    progressBar.stop();
                    return null;
                }

                await new Promise(resolve => setTimeout(resolve, 5000));
                continue;
            }

            progressBar.stop();
            return fileUrl;
        } catch (error) {
            console.error('Error getting report file URL:', error.response.data);
            progressBar.stop();
            return null;
        }
    }
}

async function downloadAndExtract(url, destination) {
    try {
        const response = await axios({
            method: 'GET',
            url: url,
            responseType: 'arraybuffer'
        });

        const zipData = response.data;
        fs.writeFileSync(`${destination}/temp.zip`, zipData);

        const zip = new AdmZip(`${destination}/temp.zip`);
        zip.extractAllTo(destination, true);

        const extractedFiles = zip.getEntries();
        const firstEntry = extractedFiles[0];
        const originalFileName = firstEntry.entryName;

        fs.renameSync(`${destination}/${originalFileName}`, `${destination}/${reportCnFileName}.csv`);

        console.log('ZIP file downloaded and unzipped successfully.');
    } catch (error) {
        console.error('Error downloading or unzipping ZIP file:', error);
    }
}

async function importCsvToGoogleSheets(csvFilePath, addToNewSpreadsheet, appendToExistingSheet, existingSheetName) {
    try {
        const auth = new google.auth.GoogleAuth({
            credentials: credentials,
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });
        const sheets = google.sheets({ version: 'v4', auth });

        const rows = [];
        let isFirstRow = true;
        let startRow = 1; // By default, we start from the first row

        // If we are adding data to an existing sheet
        if (appendToExistingSheet) {
            // Get the current range
            const response = await sheets.spreadsheets.values.get({
                spreadsheetId: spreadsheetId,
                range: existingSheetName
            });

            const numRows = response.data.values ? response.data.values.length : 0;
            startRow = numRows + 1; // Start from the next row after the last row with data
        }


        fs.createReadStream(csvFilePath)
            .pipe(csv({ separator: ';' }))
            .on('data', (data) => {

                if (isFirstRow) {
                    const headers = Object.keys(data);
                    rows.push(headers);
                    if (appendToExistingSheet || addToNewSpreadsheet) {
                        isFirstRow = false; 
                    }; 
                }

                const values = Object.values(data);
                rows.push(values);
            })

            .on('end', async () => {
                if (addToNewSpreadsheet) {
                    rows.push([]);
                    try {
                    const addSheetResponse = await sheets.spreadsheets.batchUpdate({
                        spreadsheetId: spreadsheetId,
                        resource: {
                            requests: [
                                {
                                    addSheet: {
                                        properties: {
                                            title: `Datos ${new Date().toLocaleDateString()}`
                                        }
                                    }
                                }
                            ]
                        }
                    });

                    const newSheetTitle = addSheetResponse.data.replies[0].addSheet.properties.title;

                    await sheets.spreadsheets.values.update({
                        spreadsheetId: spreadsheetId,
                        range: `${newSheetTitle}!A1`,
                        valueInputOption: 'USER_ENTERED',
                        resource: {
                            values: rows
                        }
                    });

                    console.log('Data imported to a new sheet in Google Sheets successfully.');
                } catch (error) {
                    console.error('Error adding new sheet in Google Sheets:',error.response.data);
                }
                } else {
                    const range = `${existingSheetName}!A${startRow}`;
                    const dataWithoutHeader = rows.slice(1);

                    await sheets.spreadsheets.values.update({
                        spreadsheetId: spreadsheetId,
                        range: range,
                        valueInputOption: 'USER_ENTERED',
                        resource: {
                            values: dataWithoutHeader
                        }
                    });

                    console.log('Data imported to existing sheet in Google Sheets successfully.');
                }
            });

    } catch (error) {
        console.error('Error importing data to Google Sheets:', error.response.data);
    }
}

async function main() {

    try{const reportId = await generateReport(options.fromdate,options.todate);
        if (!reportId) {
            console.error('The report could not be generated.');
            return;
        }
    
        const fileUrl = await getReportFileUrl(reportId);
        if (!fileUrl) {
            console.error('Could not get report file URL.');
            return;
        }
    
        await downloadAndExtract(fileUrl, destination); 
    
        const csvFilePath = `${destination}/${reportCnFileName}.csv`;
       
        await importCsvToGoogleSheets(csvFilePath, addToNewSpreadsheet,appendToExistingSheet,existingSheetName);

    } catch (error){
        console.error('Error during execution', error);
        process.exit(1);

    }
}
    
main();
