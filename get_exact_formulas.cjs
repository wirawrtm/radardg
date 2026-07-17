const { google } = require('googleapis');

const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const privateKey = process.env.GOOGLE_PRIVATE_KEY
  ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  : undefined;

const auth = new google.auth.JWT({
  email: clientEmail,
  key: privateKey,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

async function run() {
  try {
    const res = await sheets.spreadsheets.get({
      spreadsheetId,
      ranges: ['channel!G700:G715'],
      includeGridData: true,
    });
    const rowData = res.data.sheets[0].data[0].rowData || [];
    rowData.forEach((row, idx) => {
      const cell = row.values[0] || {};
      console.log(`Cell G${700 + idx}:`, {
        userEnteredValue: cell.userEnteredValue,
        effectiveValue: cell.effectiveValue,
        formattedValue: cell.formattedValue,
      });
    });
  } catch (err) {
    console.error('Error:', err);
  }
}
run();
