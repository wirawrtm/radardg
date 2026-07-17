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
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'employee!A1:Z500',
    });
    const rows = res.data.values || [];
    console.log('Total employee rows:', rows.length);
    console.log('Headers:', rows[0]);
    // Log unique combinations of Division/Group, Province/Area, and Employee Name
    const map = {};
    rows.slice(1).forEach(row => {
      const name = row[0] || '';
      const division = row[3] || ''; // Let's check which columns are what
      const province = row[4] || '';
      const area = row[5] || '';
      const group = row[9] || '';
      const key = `${province} | ${group}`;
      if (!map[key]) map[key] = [];
      map[key].push(name);
    });
    console.log('\nProvince & Group Mapping in employee sheet:');
    for (const [key, names] of Object.entries(map)) {
      console.log(`- ${key} => ${[...new Set(names)].join(', ')}`);
    }
  } catch (err) {
    console.error('Error:', err);
  }
}
run();
