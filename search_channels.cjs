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
      range: 'channel!A1:G1500',
    });
    const rows = res.data.values || [];
    const targets = [
      'TRISHAKTI',
      'TANI SUBUR AGRO',
      'Lemsi Triguna',
      'BINGEI AGUNG',
      'Sentra Agronusa',
      'KARISMA INDOAGRO',
      'APOTIK TANI',
      'AGRO KIMIA ASIA',
      'BINTANG TANI GROUP',
      'INAGRI JAYA',
      'CAHAYA KARUNIA',
      'HIJAU BUMI',
      'SEMI',
      'PANCA AGRO',
      'SAPROTAN',
    ];
    rows.forEach((row, idx) => {
      const name = row[0] || '';
      const group = row[1] || '';
      const cat = row[2] || '';
      const prov = row[5] || '';
      const pic = row[6] || '';
      
      const matched = targets.some(t => name.toLowerCase().includes(t.toLowerCase()));
      if (matched) {
        console.log(`Row ${idx + 1}: Name='${name}', Group='${group}', Cat='${cat}', Prov='${prov}', PIC='${pic}'`);
      }
    });
  } catch (err) {
    console.error('Error:', err);
  }
}
run();
