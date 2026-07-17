const { google } = require('googleapis');

const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const privateKey = process.env.GOOGLE_PRIVATE_KEY
  ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n')
  : undefined;

if (!spreadsheetId || !clientEmail || !privateKey) {
  console.log('Credentials not configured.');
  process.exit(1);
}

const auth = new google.auth.JWT({
  email: clientEmail,
  key: privateKey,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });

async function run() {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId });
    const sheetNames = meta.data.sheets.map(s => s.properties.title);
    console.log('Sheets:', sheetNames);

    for (const name of sheetNames) {
      const res = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `${name}!A1:Z2000`,
      });
      const rows = res.data.values || [];
      rows.forEach((row, idx) => {
        const rowStr = JSON.stringify(row);
        if (rowStr.toLowerCase().includes('iing') || rowStr.toLowerCase().includes('qqqq')) {
          console.log(`[${name}] Row ${idx + 1}:`, row);
        }
      });
    }
  } catch (err) {
    console.error('Error:', err);
  }
}
run();
