const { google } = require('googleapis');
const fs = require('fs');

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

function cleanForMatch(val) {
  if (!val) return "";
  return String(val).replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
}

async function run() {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'channel!A1:G1500',
    });
    const data = res.data.values || [];
    const headers = data[0];
    const idx = {
      pic: headers.findIndex((h) =>
        /pic|user|nama|analyst|solution/i.test(String(h).trim()),
      ),
      channel: headers.findIndex((h) =>
        /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()),
      ),
    };

    console.log('Headers:', headers);
    console.log('idx:', idx);

    // Let's find "Qqqq" rows
    data.forEach((row, rIdx) => {
      if (row[0] && row[0].toLowerCase() === 'qqqq') {
        console.log(`Found Qqqq at Row ${rIdx + 1}:`, row);
      }
    });

    const body = {
      id: 1337, // or other index
      name: "Qqqq",
      pic: "Iing Mubarok"
    };

    let rowIndex = -1;
    const rowNum = Number(body.id);

    if (!isNaN(rowNum) && rowNum > 1 && rowNum <= data.length) {
      const potentialRow = data[rowNum - 1];
      if (potentialRow && idx.channel !== -1 && body.name) {
        const currentClean = cleanForMatch(potentialRow[idx.channel]);
        const cleanName = cleanForMatch(body.name);
        console.log(`Checking bounds matching: Row=${rowNum}, CurrentClean='${currentClean}', TargetClean='${cleanName}'`);
        if (currentClean === cleanName) {
          rowIndex = rowNum - 1;
        }
      }
    }

    if (rowIndex === -1 && body.name && idx.channel !== -1) {
      const targetClean = cleanForMatch(body.name);
      rowIndex = data.findIndex(
        (row, idxVal) =>
          idxVal > 0 &&
          cleanForMatch(row[idx.channel]) === targetClean,
      );
      console.log(`Fallback search index:`, rowIndex);
    }

    console.log(`Final rowIndex: ${rowIndex}`);
    if (rowIndex > 0 && rowIndex < data.length) {
      console.log(`Row index is valid: Row ${rowIndex + 1}`);
    } else {
      console.log(`Row index is INVALID! rowIndex: ${rowIndex}, data.length: ${data.length}`);
    }

  } catch (err) {
    console.error('Error:', err);
  }
}
run();
