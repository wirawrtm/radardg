import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import { google } from "googleapis";
import dotenv from "dotenv";
import fs from "fs";

dotenv.config();

const app = express();
const PORT = 3000;

// Enable CORS for all requests to support external clients (like Vercel deployments)
app.use((req, res, next) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, X-Requested-With");
  if (req.method === "OPTIONS") {
    return res.sendStatus(200);
  }
  next();
});

const LOCAL_DB_PATH = path.join(process.cwd(), "local_sheets_db.json");

function initLocalDb() {
  if (fs.existsSync(LOCAL_DB_PATH)) {
    return;
  }
  const defaultDb: Record<string, any[][]> = {
    employee: [
      ["Nama", "Email", "User", "Position", "Province", "Area", "Upline", "Password", "Level", "Group"],
      ["Aditya Wiratama", "ADITYAHEADOFFICE", "adityaheadoffice", "Business Analyst", "Head Office", "Head Office", "-", "123", "5", "All"],
      ["Suryanto Budi Santoso", "SURYANTOHEAD", "suryantohead", "Vegetables Sales Manager", "Head Office", "Head Office", "Yash Pal Rathore", "123", "2", "Vegetables"],
      ["Dani Adi Prasetya", "DANIHEADOFFICE", "daniheadoffice", "Commercial Lead", "Head Office", "Head Office", "Yash Pal Rathore", "123", "4", "Field Corn"],
      ["Yash Pal Rathore", "YASHHEADOFFICE", "yashheadoffice", "Country Head", "Head Office", "Head Office", "-", "123", "5", "All"]
    ],
    channel: [
      ["Name", "PIC", "Category", "Province", "Area"],
      ["Kiosk Maju", "Listianto", "RTL", "Jawa Timur", "East Java"],
      ["Kiosk Jaya", "Listianto", "RTL", "Jawa Timur", "East Java"],
      ["Kiosk Makmur", "Listianto", "RTL", "Jawa Timur", "East Java"],
      ["Kiosk Tani", "Agus Herdianto", "RTL", "Jawa Timur", "East Java"]
    ],
    working: [
      ["Timestamp", "Channel", "Name Checker", "Lot", "Quantity (kg)", "Aging (month)", "Exp Date", "Crops", "Condition", "Shipping Date", "POG", "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Ags", "Sep", "Okt", "Nov", "Des", "upd_jan", "upd_feb", "upd_mar", "upd_apr", "upd_mei", "upd_jun", "upd_jul", "upd_ags", "upd_sep", "upd_okt", "upd_nov", "upd_des"],
      ["01/01/2026 10:00:00", "Kiosk Maju", "Listianto", "LOT001", "120", "2", "01/Dec/2026", "Field Corn", "tetap", "01/Jan/2026", "100", "100", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "Listianto", "", "", "", "", "", "", "", "", "", "", ""]
    ],
    dr: [
      ["Lot No", "Date", "Qty", "Hybrid", "Crops"],
      ["LOT001", "01/Jan/2026", "200", "ADV808", "Field Corn"],
      ["LOT002", "15/Jan/2026", "150", "ADV808", "Field Corn"]
    ],
    hybrid: [
      ["Material", "Hybrid", "Crops"],
      ["ADV808", "ADV808", "Field Corn"]
    ],
    access: [
      ["position", "home", "partner", "stock", "pog", "overview", "temp", "access"],
      ["Business Analyst", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE"],
      ["Sales Manager", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE"],
      ["Area Sales Manager", "TRUE", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE"],
      ["Sales Agronomist", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE", "FALSE"],
      ["Business Solution", "TRUE", "TRUE", "TRUE", "TRUE", "FALSE", "FALSE", "FALSE"]
    ]
  };
  fs.writeFileSync(LOCAL_DB_PATH, JSON.stringify(defaultDb, null, 2), "utf8");
}

// Helpers for Jakarta (WIB) Timezone (UTC+7)
function getJakartaParts(): {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
} {
  const now = new Date();
  const options: Intl.DateTimeFormatOptions = {
    timeZone: "Asia/Jakarta",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false
  };
  const formatter = new Intl.DateTimeFormat("en-US", options);
  const parts = formatter.formatToParts(now);
  const partMap = parts.reduce((acc, part) => {
    acc[part.type] = part.value;
    return acc;
  }, {} as Record<string, string>);

  return {
    year: parseInt(partMap.year, 10),
    month: parseInt(partMap.month, 10),
    day: parseInt(partMap.day, 10),
    hour: parseInt(partMap.hour, 10),
    minute: parseInt(partMap.minute, 10),
    second: parseInt(partMap.second, 10)
  };
}

function getJakartaDate(): Date {
  const parts = getJakartaParts();
  return new Date(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
}

app.use(express.json({ limit: "50mb" }));
app.use(express.text({ type: "text/plain", limit: "50mb" }));
app.use((req, res, next) => {
  if (typeof req.body === "string" && req.body.trim().startsWith("{")) {
    try {
      req.body = JSON.parse(req.body);
    } catch (e) {}
  }
  next();
});

// Google Sheets API Helpers
const spreadsheetId = process.env.GOOGLE_SPREADSHEET_ID;
const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
const privateKey = process.env.GOOGLE_PRIVATE_KEY
  ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n")
  : undefined;

const isDirectConfigured = false;

console.log(
  `[Google Sheets API] Direct connection configured: ${isDirectConfigured}`,
);

let sheetsClient: any = null;

function getSheetsClient() {
  if (!isDirectConfigured) return null;
  if (!sheetsClient) {
    try {
      const auth = new google.auth.JWT({
        email: clientEmail,
        key: privateKey,
        scopes: ["https://www.googleapis.com/auth/spreadsheets"],
      });
      sheetsClient = google.sheets({ version: "v4", auth });
    } catch (e) {
      console.error("Error creating Google Sheets client:", e);
    }
  }
  return sheetsClient;
}

// Helpers mirroring code.gs
function cleanForMatch(val: any): string {
  return String(val || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function normalizePosition(pos: string | undefined): string {
  if (!pos) return "Business Solution";
  const clean = String(pos).toLowerCase().replace(/\s+/g, "");
  if (clean === "businessanalyst" || clean === "analyst")
    return "Business Analyst";
  if (clean === "salesmanager" || clean === "sm") return "Vegetables Sales Manager";
  if (clean === "areasalesmanager" || clean === "asm")
    return "Area Sales Manager";
  if (clean === "salesagronomist" || clean === "sa") return "Sales Agronomist";
  if (clean === "businesssolution" || clean === "bs")
    return "Business Solution";
  return String(pos).trim();
}

function formatMyDate(dateObj: any): string {
  if (!dateObj || dateObj === "") return "N/A";
  const date = new Date(dateObj);
  if (isNaN(date.getTime())) return "N/A";
  const day = String(date.getDate()).padStart(2, "0");
  const months = [
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "OCT",
    "NOV",
    "DEC",
  ];
  const year = String(date.getFullYear()).substring(2);
  return `${day}/${months[date.getMonth()]}/${year}`;
}

function parseGasDate(val: any): Date {
  if (!val) return new Date(0);
  if (val instanceof Date) return val;
  const str = String(val).trim();
  let d = new Date(str);
  if (!isNaN(d.getTime())) return d;
  if (str.includes("/")) {
    const parts = str.split(/[\s/:]+/);
    if (parts.length >= 3) {
      const dPart = parseInt(parts[0], 10);
      let mPart = parts[1];
      const yPart = parseInt(parts[2], 10);
      let m = parseInt(mPart, 10) - 1;
      if (isNaN(m)) {
        const months = [
          "jan",
          "feb",
          "mar",
          "apr",
          "may",
          "jun",
          "jul",
          "aug",
          "sep",
          "oct",
          "nov",
          "dec",
        ];
        const lowerM = mPart.toLowerCase();
        m = months.findIndex((name) => lowerM.startsWith(name));
        if (m === -1) m = 0;
      }
      return new Date(yPart, m, dPart);
    }
  }
  if (str.includes("-")) {
    const parts = str.split(/[\s\-:]+/);
    if (parts.length >= 3) {
      const dPart = parseInt(parts[0], 10);
      let mPart = parts[1];
      let yPart = parseInt(parts[2], 10);
      if (yPart < 100) yPart += 2000;
      let m = parseInt(mPart, 10) - 1;
      if (isNaN(m)) {
        const months = [
          "jan",
          "feb",
          "mar",
          "apr",
          "may",
          "jun",
          "jul",
          "aug",
          "sep",
          "oct",
          "nov",
          "dec",
        ];
        const lowerM = mPart.toLowerCase();
        m = months.findIndex((name) => lowerM.startsWith(name));
        if (m === -1) m = 0;
      }
      return new Date(yPart, m, dPart);
    }
  }
  return new Date(0);
}

function getMonthIndices(headers: string[]): number[] {
  const synonyms = [
    ["jan", "januari", "january"],
    ["feb", "februari", "february"],
    ["mar", "maret", "march"],
    ["apr", "april"],
    ["mei", "may"],
    ["jun", "juni", "june"],
    ["jul", "juli", "july"],
    ["ags", "agu", "agst", "agustus", "aug", "august"],
    ["sep", "sept", "september"],
    ["okt", "oct", "oktober", "october"],
    ["nov", "november"],
    ["des", "dec", "desember", "december"],
  ];
  const matchedIndices = Array(12).fill(-1);
  for (let m = 0; m < 12; m++) {
    const list = synonyms[m];
    const idx = headers.findIndex((h) => {
      const hStr = String(h || "")
        .trim()
        .toLowerCase();
      return list.some((syn) => hStr === syn);
    });
    matchedIndices[m] = idx;
  }
  return matchedIndices;
}

function getUpdMonthIndices(headers: string[]): number[] {
  const synonyms = [
    ["jan", "januari", "january"],
    ["feb", "februari", "february"],
    ["mar", "maret", "march"],
    ["apr", "april"],
    ["mei", "may"],
    ["jun", "juni", "june"],
    ["jul", "juli", "july"],
    ["ags", "agu", "agst", "agustus", "aug", "august"],
    ["sep", "sept", "september"],
    ["okt", "oct", "oktober", "october"],
    ["nov", "november"],
    ["des", "dec", "desember", "december"],
  ];
  const matchedIndices = Array(12).fill(-1);
  for (let m = 0; m < 12; m++) {
    const list = synonyms[m];
    const idx = headers.findIndex((h) => {
      if (h === undefined || h === null) return false;
      const hStr = String(h)
        .trim()
        .toLowerCase()
        .replace(/[\s_\-\/]/g, "");
      if (!hStr.startsWith("upd")) return false;
      const remains = hStr.substring(3);
      return list.some((syn) => remains === syn);
    });
    matchedIndices[m] = idx;
  }
  return matchedIndices;
}

function getMonthIndexFromDateString(dateStr: any): number {
  if (!dateStr) return new Date().getMonth();
  if (dateStr instanceof Date) return dateStr.getMonth();
  const str = String(dateStr).trim();
  if (str.includes("/")) {
    const parts = str.split(/[\s/:]+/);
    if (parts.length >= 2) {
      const mVal = parseInt(parts[1], 10);
      if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) return mVal - 1;
      const months = [
        "jan",
        "feb",
        "mar",
        "apr",
        "may",
        "jun",
        "jul",
        "aug",
        "sep",
        "oct",
        "nov",
        "dec",
      ];
      const lowerM = parts[1].toLowerCase();
      const m = months.findIndex((name) => lowerM.startsWith(name));
      if (m !== -1) return m;
    }
  }
  if (str.includes("-")) {
    const parts = str.split(/[\s\-:]+/);
    if (parts.length >= 2) {
      const mVal = parseInt(parts[1], 10);
      if (!isNaN(mVal) && mVal >= 1 && mVal <= 12) return mVal - 1;
      const months = [
        "jan",
        "feb",
        "mar",
        "apr",
        "may",
        "jun",
        "jul",
        "aug",
        "sep",
        "oct",
        "nov",
        "des",
      ];
      const lowerM = parts[1].toLowerCase();
      const m = months.findIndex((name) => lowerM.startsWith(name));
      if (m !== -1) return m;
    }
  }
  const d = new Date(str);
  return !isNaN(d.getTime()) ? d.getMonth() : new Date().getMonth();
}

// In-memory cache map for Google Sheets data to avoid repetitive API requests and make operations extremely fast
const activeRequests: Record<string, Promise<any[][] | null>> = {};
const cacheMap: Record<string, { data: any[][]; timestamp: number }> = {};
const CACHE_DURATION_MS = 15000; // Cache for 15 seconds to handle simultaneous or rapid read spikes beautifully

function invalidateCache(sheetName: string) {
  delete cacheMap[sheetName];
  delete activeRequests[sheetName];
}

// Fetch helper from a specified sheet
async function getSheetValues(sheetName: string): Promise<any[][] | null> {
  if (!isDirectConfigured) {
    initLocalDb();
    if (fs.existsSync(LOCAL_DB_PATH)) {
      try {
        const db = JSON.parse(fs.readFileSync(LOCAL_DB_PATH, "utf8"));
        return db[sheetName] || [];
      } catch (e) {
        console.error(`Error reading local db for ${sheetName}:`, e);
      }
    }
    return [];
  }

  const sheets = getSheetsClient();
  if (!sheets) return null;

  const now = Date.now();
  // Return valid cached response if available
  if (
    cacheMap[sheetName] &&
    now - cacheMap[sheetName].timestamp < CACHE_DURATION_MS
  ) {
    return cacheMap[sheetName].data;
  }

  // Coalesce overlapping requests to avoid redundant simultaneous calls to Google API
  if (activeRequests[sheetName]) {
    return activeRequests[sheetName];
  }

  const fetchPromise = (async () => {
    try {
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId,
        range: sheetName,
      });
      const data = response.data.values || [];
      cacheMap[sheetName] = { data, timestamp: Date.now() };
      return data;
    } catch (e) {
      console.error(`Error reading sheet ${sheetName} from API:`, e);
      throw new Error("Failed to read from Google Sheets API");
    } finally {
      delete activeRequests[sheetName];
    }
  })();

  activeRequests[sheetName] = fetchPromise;
  return fetchPromise;
}

// Write helper to overwrite a sheet's content
async function updateSheetValues(
  sheetName: string,
  values: any[][],
): Promise<boolean> {
  const sanitizedValues = (values || []).map(row => {
    if (!Array.isArray(row)) return [];
    const len = row.length;
    const sanitizedRow = [];
    for (let i = 0; i < len; i++) {
      const val = row[i];
      sanitizedRow.push(val === undefined || val === null ? "" : val);
    }
    return sanitizedRow;
  });

  if (!isDirectConfigured) {
    initLocalDb();
    try {
      let db: Record<string, any[][]> = {};
      if (fs.existsSync(LOCAL_DB_PATH)) {
        db = JSON.parse(fs.readFileSync(LOCAL_DB_PATH, "utf8"));
      }
      db[sheetName] = sanitizedValues;
      fs.writeFileSync(LOCAL_DB_PATH, JSON.stringify(db, null, 2), "utf8");
      invalidateCache(sheetName);
      return true;
    } catch (e) {
      console.error(`Error updating local db for ${sheetName}:`, e);
      return false;
    }
  }

  const sheets = getSheetsClient();
  if (!sheets) return false;
  invalidateCache(sheetName);
  try {
    // Clear first to avoid leftover values
    await sheets.spreadsheets.values.clear({
      spreadsheetId,
      range: sheetName,
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: sanitizedValues },
    });
    return true;
  } catch (e) {
    console.error(`Error updating sheet ${sheetName}:`, e);
    throw new Error(`Failed to update sheet ${sheetName}`);
  }
}

// Append helper to add a row
async function appendSheetRow(
  sheetName: string,
  rowValues: any[],
): Promise<boolean> {
  const sanitizedRow = (rowValues || []).map(cell => (cell === undefined || cell === null ? "" : cell));

  if (!isDirectConfigured) {
    initLocalDb();
    try {
      let db: Record<string, any[][]> = {};
      if (fs.existsSync(LOCAL_DB_PATH)) {
        db = JSON.parse(fs.readFileSync(LOCAL_DB_PATH, "utf8"));
      }
      if (!db[sheetName]) {
        db[sheetName] = [];
      }
      db[sheetName].push(sanitizedRow);
      fs.writeFileSync(LOCAL_DB_PATH, JSON.stringify(db, null, 2), "utf8");
      invalidateCache(sheetName);
      return true;
    } catch (e) {
      console.error(`Error appending to local db for ${sheetName}:`, e);
      return false;
    }
  }

  const sheets = getSheetsClient();
  if (!sheets) return false;
  try {
    const currentValues = await getSheetValues(sheetName);
    if (!currentValues) return false;
    currentValues.push(sanitizedRow);
    return await updateSheetValues(sheetName, currentValues);
  } catch (e) {
    console.error(`Error appending to sheet ${sheetName}:`, e);
    throw new Error(`Failed to append to sheet ${sheetName}`);
  }
}

// User mappings cached during operations
function findEmployeeRow(user: string, empData: any[][]): any | null {
  if (!user || !empData || empData.length <= 1) return null;
  const headers = empData[0];
  const emailIdx = headers.findIndex((h: any) => /email/i.test(String(h).trim()));
  const userIdx = headers.findIndex((h: any) => /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()));
  const nameIdx = headers.findIndex((h: any) => /nama|name|pic/i.test(String(h).trim()));

  const targetClean = cleanForMatch(user);

  for (let i = 1; i < empData.length; i++) {
    const row = empData[i];
    const rowName = nameIdx !== -1 ? cleanForMatch(row[nameIdx]) : "";
    const rowEmail = emailIdx !== -1 ? cleanForMatch(row[emailIdx]) : "";
    const rowUser = userIdx !== -1 ? cleanForMatch(row[userIdx]) : "";

    if (rowUser === targetClean || rowName === targetClean || rowEmail === targetClean) {
      return row;
    }
  }
  return null;
}

function getUserGroup(userName: string, empData: any[][]): string {
  if (!userName || !empData || empData.length <= 1) return "";
  const headers = empData[0];
  const groupIdx = headers.findIndex((h: any) =>
    /group|tim|divisi|division/i.test(String(h).trim()),
  );
  if (groupIdx === -1) return "";
  const row = findEmployeeRow(userName, empData);
  if (row) {
    return String(row[groupIdx] || "").trim();
  }
  return "";
}

function getUserProvince(userName: string, empData: any[][]): string {
  if (!userName || !empData || empData.length <= 1) return "";
  const headers = empData[0];
  const provIdx = headers.findIndex((h: any) =>
    /province|provinsi/i.test(String(h).trim()),
  );
  const areaIdx = headers.findIndex((h: any) => /area/i.test(String(h).trim()));
  const row = findEmployeeRow(userName, empData);
  if (row) {
    if (
      provIdx !== -1 &&
      row[provIdx] !== "" &&
      row[provIdx] !== undefined
    ) {
      return String(row[provIdx]).trim();
    }
    if (
      areaIdx !== -1 &&
      row[areaIdx] !== "" &&
      row[areaIdx] !== undefined
    ) {
      return String(row[areaIdx]).trim();
    }
  }
  return "";
}

function findEmployeeDetails(picName: string, empData: any[][]) {
  const result = { upline: "", area: "" };
  if (!picName || !empData || empData.length <= 1) return result;
  const headers = empData[0];
  const idx = {
    upline: headers.findIndex((h: any) =>
      /upline|spv|supervisor|atasan|manager/i.test(String(h).trim()),
    ),
    area: headers.findIndex((h: any) => /area/i.test(String(h).trim())),
    prov: headers.findIndex((h: any) =>
      /province|provinsi/i.test(String(h).trim()),
    ),
  };
  const row = findEmployeeRow(picName, empData);
  if (row) {
    if (idx.upline !== -1)
      result.upline = String(row[idx.upline] || "").trim();
    if (
      idx.area !== -1 &&
      row[idx.area] !== "" &&
      row[idx.area] !== undefined
    ) {
      result.area = String(row[idx.area]).trim();
    } else if (
      idx.prov !== -1 &&
      row[idx.prov] !== "" &&
      row[idx.prov] !== undefined
    ) {
      result.area = String(row[idx.prov]).trim();
    }
  }
  return result;
}

// Endpoints re-implementing Apps Script handlers
async function handleGetWorkingData(user: string) {
  const data = await getSheetValues("working");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  const headers = data[0];
  const getIdx = (patterns: RegExp) =>
    headers.findIndex((h) => patterns.test(String(h).trim()));
  const idx = {
    lot: getIdx(/^lot package$|^lot$/i),
    hybrid: getIdx(/^hybrid$|^material$/i),
    stock: getIdx(/^quantity \(kg\)|^qty$|^stock$|^kg$/i),
    aging: getIdx(/^aging \(month\)/i),
    exp: getIdx(/^exp date$|^expired$/i),
    kiosk: getIdx(/^channel$|^kiosk$/i),
    crops: getIdx(/^crops$/i),
    time: getIdx(/^tgl$|^waktu$|^date$|^timestamp$|^tanggal$/i),
    cond: getIdx(/^condition$|^kondisi$/i),
    dr: getIdx(/^shipping date$|^dr date$/i),
    user: getIdx(/^name checker$|^nama checker$|^user$|^pic$|^checker$/i),
    pog: getIdx(/^pog$|^selisih$/i),
    area: getIdx(/^area$|^region$/i),
  };
  const monthIndices = getMonthIndices(headers);
  const updMonthIndices = getUpdMonthIndices(headers);

  const result = data
    .slice(1)
    .filter((row) => row[0] !== "" && row[0] !== undefined)
    .map((row) => {
      const rowItem: any = {
        lot: idx.lot !== -1 ? row[idx.lot] : "",
        hybrid: idx.hybrid !== -1 ? row[idx.hybrid] : "",
        crops:
          idx.crops !== -1 &&
          row[idx.crops] !== "" &&
          row[idx.crops] !== undefined
            ? row[idx.crops]
            : "Uncategorized Crops",
        stock:
          idx.stock !== -1 &&
          row[idx.stock] !== "" &&
          row[idx.stock] !== undefined
            ? Number(row[idx.stock]) || 0
            : 0,
        aging:
          idx.aging !== -1 &&
          row[idx.aging] !== "" &&
          row[idx.aging] !== undefined
            ? row[idx.aging]
            : "-",
        expired:
          idx.exp !== -1 && row[idx.exp] ? formatMyDate(row[idx.exp]) : "N/A",
        drDate:
          idx.dr !== -1 && row[idx.dr] ? formatMyDate(row[idx.dr]) : "N/A",
        kiosk: idx.kiosk !== -1 ? row[idx.kiosk] : "",
        timestamp: idx.time !== -1 && row[idx.time] ? row[idx.time] : "",
        condition: idx.cond !== -1 && row[idx.cond] ? row[idx.cond] : "tetap",
        user:
          idx.user !== -1 && row[idx.user] ? String(row[idx.user]).trim() : "",
        pog:
          idx.pog !== -1 && row[idx.pog] !== "" && row[idx.pog] !== undefined
            ? Number(row[idx.pog]) || 0
            : 0,
        area:
          idx.area !== -1 && row[idx.area] !== "" && row[idx.area] !== undefined
            ? String(row[idx.area]).trim()
            : "",
      };

      const months = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "Mei",
        "Jun",
        "Jul",
        "Ags",
        "Sep",
        "Okt",
        "Nov",
        "Des",
      ];
      months.forEach((m, mIdx) => {
        const colIdx = monthIndices[mIdx];
        if (colIdx !== -1 && colIdx < row.length) {
          rowItem[m.toLowerCase()] =
            row[colIdx] !== "" && row[colIdx] !== undefined
              ? Number(row[colIdx]) || 0
              : 0;
        } else {
          rowItem[m.toLowerCase()] = 0;
        }

        const updColIdx = updMonthIndices[mIdx];
        if (updColIdx !== -1 && updColIdx < row.length) {
          rowItem["upd_" + m.toLowerCase()] =
            row[updColIdx] !== "" && row[updColIdx] !== undefined
              ? String(row[updColIdx]).trim()
              : "";
        } else {
          rowItem["upd_" + m.toLowerCase()] = "";
        }
      });
      return rowItem;
    });
  return { status: "success", data: result };
}

async function handleGetChannels(user: string) {
  const data = await getSheetValues("channel");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  const lowerUser = String(user).trim().toLowerCase();

  const empData = await getSheetValues("employee");
  let authorizedPICs = [lowerUser];
  let isBusinessAnalyst =
    lowerUser === "adityawiratama" ||
    lowerUser.includes("adityawiratama") ||
    lowerUser === "analyst" ||
    lowerUser === "businessanalyst";

  if (empData && empData.length > 0) {
    const empHeaders = empData[0];
    const idxE = {
      name: empHeaders.findIndex((h: any) =>
        /nama|name|pic/i.test(String(h).trim()),
      ),
      email: empHeaders.findIndex((h: any) =>
        /email|user/i.test(String(h).trim()),
      ),
      upline: empHeaders.findIndex((h: any) =>
        /upline|spv|supervisor|atasan|manager/i.test(String(h).trim()),
      ),
      pos: empHeaders.findIndex((h: any) =>
        /position|jabatan/i.test(String(h).trim()),
      ),
    };

    const matchedRow = findEmployeeRow(user, empData);
    isBusinessAnalyst =
      lowerUser === "adityawiratama" ||
      lowerUser.includes("adityawiratama") ||
      lowerUser === "analyst" ||
      lowerUser === "businessanalyst";
    const userAliases = new Set([lowerUser]);

    if (matchedRow) {
      const emailIdx = empHeaders.findIndex((h: any) => /email/i.test(String(h).trim()));
      const userIdx = empHeaders.findIndex((h: any) => /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()));
      const rowName = idxE.name !== -1 ? String(matchedRow[idxE.name] || "").trim().toLowerCase() : "";
      const rowEmail = emailIdx !== -1 ? String(matchedRow[emailIdx] || "").trim().toLowerCase() : "";
      const rowUser = userIdx !== -1 ? String(matchedRow[userIdx] || "").trim().toLowerCase() : "";
      const rowPos = idxE.pos !== -1 ? String(matchedRow[idxE.pos] || "").trim().toLowerCase() : "";

      if (rowName !== "") userAliases.add(rowName);
      if (rowEmail !== "") userAliases.add(rowEmail);
      if (rowUser !== "") userAliases.add(rowUser);

      const cleanRowPos = rowPos.replace(/\s+/g, "");
      const levelIdx = empHeaders.findIndex((h: any) => /level|grade/i.test(String(h).trim()));
      const rowLevel = levelIdx !== -1 ? String(matchedRow[levelIdx] || "").trim().toLowerCase() : "";
      if (
        cleanRowPos === "businessanalyst" ||
        cleanRowPos === "analyst" ||
        rowLevel === "admin"
      ) {
        isBusinessAnalyst = true;
      }
    }

    if (isBusinessAnalyst) {
      const emailIdx = empHeaders.findIndex((h: any) => /email/i.test(String(h).trim()));
      const userIdx = empHeaders.findIndex((h: any) => /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()));
      const groupIdx = empHeaders.findIndex((h: any) => /group|tim|divisi|division/i.test(String(h).trim()));
      const levelIdx = empHeaders.findIndex((h: any) => /level|grade/i.test(String(h).trim()));
      
      let adminGroup = "";
      if (matchedRow && levelIdx !== -1 && groupIdx !== -1) {
        const rowLevel = String(matchedRow[levelIdx] || "").trim().toLowerCase();
        if (rowLevel === "admin") {
          adminGroup = String(matchedRow[groupIdx] || "").trim().toLowerCase();
        }
      }

      empData.slice(1).forEach((row) => {
        const rowGroup = groupIdx !== -1 ? String(row[groupIdx] || "").trim().toLowerCase() : "";
        const rowName =
          idxE.name !== -1
            ? String(row[idxE.name] || "")
                .trim()
                .toLowerCase()
            : "";
        const rowEmail =
          emailIdx !== -1
            ? String(row[emailIdx] || "")
                .trim()
                .toLowerCase()
            : "";
        const rowUser =
          userIdx !== -1
            ? String(row[userIdx] || "")
                .trim()
                .toLowerCase()
            : "";

        // If admin has a specific group, only include employees of that group (or if they are the admin themselves)
        if (adminGroup && adminGroup !== "all" && adminGroup !== "") {
          if (rowGroup !== adminGroup && rowName !== lowerUser && rowEmail !== lowerUser && rowUser !== lowerUser) {
            return;
          }
        }

        if (rowName !== "") userAliases.add(rowName);
        if (rowEmail !== "") userAliases.add(rowEmail);
        if (rowUser !== "") userAliases.add(rowUser);
      });
    }

    const queue = Array.from(userAliases);
    const visited = new Set(queue);
    userAliases.forEach((alias) => {
      if (!authorizedPICs.includes(alias)) authorizedPICs.push(alias);
    });

    while (queue.length > 0) {
      const currentUpline = queue.shift();
      empData.slice(1).forEach((row) => {
        const empNameRaw =
          idxE.name !== -1 ? String(row[idxE.name] || "").trim() : "";
        const empNameLower = empNameRaw.toLowerCase();
        const empEmailRaw =
          idxE.email !== -1 ? String(row[idxE.email] || "").trim() : "";
        const empEmailLower = empEmailRaw.toLowerCase();
        const empUplineRaw =
          idxE.upline !== -1 ? String(row[idxE.upline] || "").trim() : "";
        const empUplineLower = empUplineRaw.toLowerCase();

        if (empUplineLower !== "") {
          const isMatch =
            empUplineLower === currentUpline ||
            empUplineLower.includes(currentUpline!) ||
            currentUpline!.includes(empUplineLower);
          if (isMatch) {
            let addedAny = false;
            if (empNameLower !== "" && !visited.has(empNameLower)) {
              visited.add(empNameLower);
              queue.push(empNameLower);
              if (!authorizedPICs.includes(empNameLower))
                authorizedPICs.push(empNameLower);
              addedAny = true;
            }
            if (empEmailLower !== "" && !visited.has(empEmailLower)) {
              visited.add(empEmailLower);
              queue.push(empEmailLower);
              if (!authorizedPICs.includes(empEmailLower))
                authorizedPICs.push(empEmailLower);
              addedAny = true;
            }
          }
        }
      });
    }
  }

  const headers = data[0];
  const idx = {
    pic: headers.findIndex((h: any) =>
      /pic|user|nama|analyst|solution/i.test(String(h).trim()),
    ),
    channel: headers.findIndex((h: any) =>
      /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()),
    ),
    cat: headers.findIndex((h: any) =>
      /kategori|category|klasifikasi|^cat$/i.test(String(h).trim()),
    ),
    upline: headers.findIndex((h: any) =>
      /upline|spv|supervisor/i.test(String(h).trim()),
    ),
    area: headers.findIndex((h: any) =>
      /area|provinsi|province|wilayah/i.test(String(h).trim()),
    ),
    group: headers.findIndex((h: any) =>
      /group|tim|divisi|division/i.test(String(h).trim()),
    ),
  };

  const picToAreaMap: Record<string, string> = {};
  if (empData && empData.length > 1) {
    const empHeadersVal = empData[0];
    const nameCol = empHeadersVal.findIndex((h: any) =>
      /nama|name|pic/i.test(String(h).trim()),
    );
    const emailCol = empHeadersVal.findIndex((h: any) =>
      /email|user/i.test(String(h).trim()),
    );
    const areaCol = empHeadersVal.findIndex((h: any) =>
      /area/i.test(String(h).trim()),
    );
    const provCol = empHeadersVal.findIndex((h: any) =>
      /province|provinsi/i.test(String(h).trim()),
    );

    empData.slice(1).forEach((empRow) => {
      const empName =
        nameCol !== -1
          ? String(empRow[nameCol] || "")
              .trim()
              .toLowerCase()
          : "";
      const empEmail =
        emailCol !== -1
          ? String(empRow[emailCol] || "")
              .trim()
              .toLowerCase()
          : "";
      const empArea =
        areaCol !== -1 &&
        empRow[areaCol] !== "" &&
        empRow[areaCol] !== undefined
          ? String(empRow[areaCol]).trim()
          : provCol !== -1 &&
              empRow[provCol] !== "" &&
              empRow[provCol] !== undefined
            ? String(empRow[provCol]).trim()
            : "";
      if (empArea) {
        if (empName) picToAreaMap[empName] = empArea;
        if (empEmail) picToAreaMap[empEmail] = empArea;
      }
    });
  }

  const channels = data
    .slice(1)
    .map((row, i) => {
      if (row[0] === "" || row[0] === undefined) return null;
      const catValue =
        idx.cat !== -1 && row[idx.cat] !== "" && row[idx.cat] !== undefined
          ? String(row[idx.cat]).trim()
          : "Uncategorized";
      const rowGroup =
        idx.group !== -1 &&
        row[idx.group] !== "" &&
        row[idx.group] !== undefined
          ? String(row[idx.group]).trim()
          : "";
      if (idx.pic !== -1 && idx.channel !== -1) {
        const picLower = String(row[idx.pic] || "")
          .trim()
          .toLowerCase();
        const uplineLower =
          idx.upline !== -1
            ? String(row[idx.upline] || "")
                .trim()
                .toLowerCase()
            : "";

        const isAuth =
          isBusinessAnalyst ||
          picLower === "" ||
          picLower === "tanpa pic" ||
          picLower === "tidak ada" ||
          picLower === lowerUser ||
          (lowerUser !== "" && picLower.includes(lowerUser)) ||
          uplineLower === lowerUser ||
          (lowerUser !== "" && uplineLower.includes(lowerUser)) ||
          authorizedPICs.some(
            (auth) =>
              picLower === auth || (auth !== "" && picLower.includes(auth)),
          );
        if (isAuth) {
          const sheetArea =
            idx.area !== -1 &&
            row[idx.area] !== "" &&
            row[idx.area] !== undefined
              ? String(row[idx.area]).trim()
              : "";
          const resolvedArea = sheetArea || picToAreaMap[picLower] || "-";
          return {
            id: i + 2,
            name: row[idx.channel],
            category: catValue,
            pic: String(row[idx.pic] || "").trim(),
            upline:
              idx.upline !== -1 ? String(row[idx.upline] || "").trim() : "",
            area: resolvedArea,
            group: rowGroup,
          };
        }
      } else if (idx.channel !== -1) {
        const sheetArea =
          idx.area !== -1 && row[idx.area] !== "" && row[idx.area] !== undefined
            ? String(row[idx.area]).trim()
            : "";
        return {
          id: i + 2,
          name: row[idx.channel],
          category: catValue,
          pic: "",
          upline: "",
          area: sheetArea || "-",
          group: rowGroup,
        };
      }
      return null;
    })
    .filter(Boolean);

  return { status: "success", data: channels };
}

async function handleGetDrSalesData(user: string) {
  const hybridMap: Record<string, { hybrid: string; crops: string }> = {};
  const hData = await getSheetValues("hybrid");
  if (hData && hData.length > 1) {
    const hHeaders = hData[0];
    const idxH = {
      desc: hHeaders.findIndex((h: any) =>
        /material.*desc|description/i.test(String(h).trim()),
      ),
      hybrid: hHeaders.findIndex((h: any) =>
        /^hybrid$/i.test(String(h).trim()),
      ),
      crops: hHeaders.findIndex((h: any) => /^crops$/i.test(String(h).trim())),
    };
    if (idxH.desc !== -1 && idxH.hybrid !== -1) {
      hData.slice(1).forEach((row) => {
        const mDesc = String(row[idxH.desc] || "")
          .trim()
          .toLowerCase();
        if (mDesc) {
          hybridMap[mDesc] = {
            hybrid: String(row[idxH.hybrid] || "").trim(),
            crops:
              idxH.crops !== -1 && row[idxH.crops] !== undefined
                ? String(row[idxH.crops]).trim()
                : "",
          };
        }
      });
    }
  }

  const data = await getSheetValues("dr");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  const headers = data[0];

  const hIdx = {
    qty: headers.findIndex((h: any) => /qty|quantity/i.test(String(h).trim())),
    type: headers.findIndex((h: any) => /order type/i.test(String(h).trim())),
    channel: headers.findIndex((h: any) =>
      /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()),
    ),
    lot: headers.findIndex((h: any) => /lot/i.test(String(h).trim())),
    desc: headers.findIndex((h: any) =>
      /material.*desc|description/i.test(String(h).trim()),
    ),
    dr: headers.findIndex((h: any) =>
      /dr date|shipping date/i.test(String(h).trim()),
    ),
    exp: headers.findIndex((h: any) =>
      /exp date|expired/i.test(String(h).trim()),
    ),
  };

  const result = data
    .slice(1)
    .filter((row) => {
      return (
        hIdx.type !== -1 &&
        String(row[hIdx.type] || "")
          .trim()
          .toLowerCase() === "sales"
      );
    })
    .map((row) => {
      const rawDesc =
        hIdx.desc !== -1 ? String(row[hIdx.desc] || "").trim() : "";
      const mapInfo = hybridMap[rawDesc.toLowerCase()] || {
        hybrid: rawDesc,
        crops: "",
      };
      const drValue =
        hIdx.dr !== -1 && row[hIdx.dr] ? formatMyDate(row[hIdx.dr]) : "N/A";
      const expValue =
        hIdx.exp !== -1 && row[hIdx.exp] ? formatMyDate(row[hIdx.exp]) : "N/A";
      return {
        lot:
          hIdx.lot !== -1
            ? String(row[hIdx.lot] || "")
                .trim()
                .toUpperCase()
            : "",
        hybrid: mapInfo.hybrid,
        crops: mapInfo.crops,
        channel:
          hIdx.channel !== -1 ? String(row[hIdx.channel] || "").trim() : "",
        qty: hIdx.qty !== -1 ? Number(row[hIdx.qty]) || 0 : 0,
        drDate: drValue,
        expired: expValue,
      };
    });
  return { status: "success", data: result };
}

async function handleGetLotInfo(lotNo: string) {
  if (!lotNo) return { status: "error", message: "Lot number is required" };
  const hybridMap: Record<string, { hybrid: string; crops: string }> = {};
  const hData = await getSheetValues("hybrid");
  if (hData && hData.length > 1) {
    const hHeaders = hData[0];
    const idxH = {
      desc: hHeaders.findIndex((h: any) =>
        /material.*desc|description/i.test(String(h).trim()),
      ),
      hybrid: hHeaders.findIndex((h: any) =>
        /^hybrid$/i.test(String(h).trim()),
      ),
      crops: hHeaders.findIndex((h: any) => /^crops$/i.test(String(h).trim())),
    };
    if (idxH.desc !== -1 && idxH.hybrid !== -1) {
      hData.slice(1).forEach((row) => {
        const mDesc = String(row[idxH.desc] || "")
          .trim()
          .toLowerCase();
        if (mDesc) {
          hybridMap[mDesc] = {
            hybrid: String(row[idxH.hybrid] || "").trim(),
            crops:
              idxH.crops !== -1 && row[idxH.crops] !== undefined
                ? String(row[idxH.crops]).trim()
                : "",
          };
        }
      });
    }
  }
  const data = await getSheetValues("dr");
  if (!data || data.length <= 1) return { status: "error", message: "No data" };
  const headers = data[0];
  const idx = {
    lot: headers.findIndex((h: any) => /lot/i.test(String(h).trim())),
    desc: headers.findIndex((h: any) =>
      /material.*desc|description/i.test(String(h).trim()),
    ),
    dr: headers.findIndex((h: any) =>
      /dr date|shipping date/i.test(String(h).trim()),
    ),
    exp: headers.findIndex((h: any) =>
      /exp date|expired/i.test(String(h).trim()),
    ),
  };
  if (idx.lot === -1) return { status: "error", message: "Lot column missing" };
  const calcMonths = (start: any, end: any) => {
    try {
      if (!start || !end || start === "" || end === "") return "";
      const d1 = new Date(start);
      const d2 = new Date(end);
      if (isNaN(d1.getTime()) || isNaN(d2.getTime())) return "";
      return Math.round(
        (d2.getTime() - d1.getTime()) / (1000 * 3600 * 24) / 30.416,
      );
    } catch (e) {
      return "";
    }
  };

  const todayDate = getJakartaDate();
  todayDate.setHours(0, 0, 0, 0);
  const targetLot = String(lotNo).trim().toUpperCase();
  const foundRow = data.slice(1).find(
    (row) =>
      String(row[idx.lot] || "")
        .trim()
        .toUpperCase() === targetLot,
  );
  if (foundRow) {
    const rawDesc =
      idx.desc !== -1
        ? String(foundRow[idx.desc] || "").trim()
        : "Unknown Material";
    const mapInfo = hybridMap[rawDesc.toLowerCase()] || {
      hybrid: rawDesc,
      crops: "",
    };
    const drDateVal = idx.dr !== -1 ? foundRow[idx.dr] : "";

    return {
      status: "success",
      data: {
        desc: mapInfo.hybrid,
        crops: mapInfo.crops,
        drDate: idx.dr !== -1 ? formatMyDate(drDateVal) : "N/A",
        expDate: idx.exp !== -1 ? formatMyDate(foundRow[idx.exp]) : "N/A",
        aging: calcMonths(drDateVal, todayDate),
      },
    };
  }
  return { status: "error", message: "Lot not found" };
}

let cachedEmployeeList: any[] | null = null;
let lastEmployeeListFetch = 0;
const EMPLOYEE_LIST_CACHE_MS = 120 * 1000; // 120 seconds cache for incredibly fast subsequent loads

async function getEmployeeList(): Promise<any[]> {
  const now = Date.now();
  if (cachedEmployeeList && (now - lastEmployeeListFetch < EMPLOYEE_LIST_CACHE_MS)) {
    return cachedEmployeeList;
  }

  const data = await getSheetValues("employee");
  if (!data || data.length <= 1) return [];
  const headers = data[0];
  const emailIdx = headers.findIndex((h: any) => /email/i.test(String(h).trim()));
  const userIdx = headers.findIndex((h: any) => /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()));
  const idx = {
    name: headers.findIndex((h: any) => /nama|name|pic/i.test(String(h).trim())),
    email: emailIdx !== -1 ? emailIdx : headers.findIndex((h: any) => /email|user/i.test(String(h).trim())),
    username: userIdx !== -1 ? userIdx : -1,
    pos: headers.findIndex((h: any) => /position|jabatan/i.test(String(h).trim())),
    prov: headers.findIndex((h: any) => /province|provinsi/i.test(String(h).trim())),
    area: headers.findIndex((h: any) => /area/i.test(String(h).trim())),
    upline: headers.findIndex((h: any) => /upline|spv|supervisor|atasan|manager/i.test(String(h).trim())),
    password: headers.findIndex((h: any) => /password|pass/i.test(String(h).trim())),
    level: headers.findIndex((h: any) => /level|grade/i.test(String(h).trim())),
    group: headers.findIndex((h: any) => /group|tim|divisi|division/i.test(String(h).trim())),
  };

  const list = data.slice(1).map((row) => {
    const p = idx.pos !== -1 ? row[idx.pos] : "Business Solution";
    return {
      name: idx.name !== -1 ? String(row[idx.name] || "").trim() : "",
      email: idx.email !== -1 ? String(row[idx.email] || "").trim() : "",
      user: idx.username !== -1 ? String(row[idx.username] || "").trim() : "",
      position: normalizePosition(p),
      province: idx.prov !== -1 ? String(row[idx.prov] || "").trim() : "-",
      area: idx.area !== -1 ? String(row[idx.area] || "").trim() : "-",
      upline: idx.upline !== -1 ? String(row[idx.upline] || "").trim() : "",
      password: idx.password !== -1 ? String(row[idx.password] || "").trim() : "",
      level: idx.level !== -1 && row[idx.level] !== "" && row[idx.level] !== undefined ? row[idx.level] : null,
      group: idx.group !== -1 ? String(row[idx.group] || "").trim() : "",
    };
  });

    // Deduplicate employees
  const seenEmployees = new Set();
  const dedupedList = [];
  for (const emp of list) {
    const cleanName = String(emp.name || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const cleanUser = String(emp.user || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const uniqKey = cleanName;
    
    if (!seenEmployees.has(uniqKey) && cleanName !== "") {
      seenEmployees.add(uniqKey);
      dedupedList.push(emp);
    }
  }

  cachedEmployeeList = dedupedList;
  lastEmployeeListFetch = now;
  return dedupedList;
}

async function handleGetUserProfile(user: string) {
  const employees = await getEmployeeList();
  if (employees.length === 0) {
    return {
      status: "error",
      message: "Data employee kosong atau tidak ditemukan",
    };
  }

  const lowerUser = String(user).trim().toLowerCase();

  // Strict matching: username MUST match the 'user' column, the 'email' column, or the local part of the 'email' column (strictly excluding 'name' column)
  const foundEmployee = employees.find((emp) => {
    const rowUser = String(emp.user || "").trim().toLowerCase();
    const rowEmail = String(emp.email || "").trim().toLowerCase();
    const rowUserLocal = rowEmail.includes("@") ? rowEmail.split("@")[0] : rowEmail;

    if (rowUser !== "") {
      return rowUser === lowerUser || rowEmail === lowerUser;
    } else {
      return rowUserLocal === lowerUser || rowEmail === lowerUser;
    }
  });

  if (!foundEmployee) {
    return { status: "error", message: "Username tidak ditemukan" };
  }

  const resolvedUser = foundEmployee.user || (foundEmployee.email ? (foundEmployee.email.includes("@") ? foundEmployee.email.split("@")[0] : foundEmployee.email) : "");

  const profile: any = {
    name: foundEmployee.name,
    email: foundEmployee.email,
    user: resolvedUser,
    position: foundEmployee.position,
    province: foundEmployee.province,
    area: foundEmployee.area,
    password: foundEmployee.password,
    upline: foundEmployee.upline,
    level: foundEmployee.level,
    group: foundEmployee.group,
    subordinates: [],
  };

  const userAliases = new Set([lowerUser]);
  if (profile.name !== "") userAliases.add(profile.name.toLowerCase());
  if (profile.email !== "") userAliases.add(profile.email.toLowerCase());
  if (profile.user !== "") userAliases.add(profile.user.toLowerCase());

  const cleanProfilePos = profile.position ? profile.position.toLowerCase().replace(/\s+/g, "") : "";
  const profileLevelClean = profile.level ? String(profile.level).toLowerCase().trim() : "";
  const isBusinessAnalyst =
    lowerUser === "adityawiratama" ||
    lowerUser.includes("adityawiratama") ||
    cleanProfilePos === "businessanalyst" ||
    cleanProfilePos === "analyst" ||
    profileLevelClean === "admin";
  if (isBusinessAnalyst) {
    if (profileLevelClean !== "admin") {
      profile.position = "Business Analyst";
    }
    const asmSubordinates: string[] = [];
    if (profileLevelClean !== "admin") {
      employees.forEach((emp) => {
        if (emp.name !== "" && emp.position === "Area Sales Manager") {
          if (!asmSubordinates.includes(emp.name)) {
            asmSubordinates.push(emp.name);
          }
        }
      });
    }
    profile.subordinates = asmSubordinates;
    return { status: "success", data: profile };
  }

  const subs: string[] = [];
  const queue = Array.from(userAliases);
  const visited = new Set(queue);

  while (queue.length > 0) {
    const currentUpline = queue.shift();
    employees.forEach((emp) => {
      const empNameRaw = String(emp.name || "").trim();
      const empNameLower = empNameRaw.toLowerCase();
      const empEmailRaw = String(emp.email || "").trim();
      const empEmailLower = empEmailRaw.toLowerCase();
      const empUplineRaw = String(emp.upline || "").trim();
      const empUplineLower = empUplineRaw.toLowerCase();

      if (empUplineLower !== "") {
        const isMatch =
          empUplineLower === currentUpline ||
          empUplineLower.includes(currentUpline!) ||
          currentUpline!.includes(empUplineLower);
        if (isMatch) {
          let addedAny = false;
          if (empNameLower !== "" && !visited.has(empNameLower)) {
            visited.add(empNameLower);
            queue.push(empNameLower);
            addedAny = true;
          }
          if (empEmailLower !== "" && !visited.has(empEmailLower)) {
            visited.add(empEmailLower);
            queue.push(empEmailLower);
            addedAny = true;
          }
          if (addedAny) {
            const displayName = empNameRaw !== "" ? empNameRaw : empEmailRaw;
            if (displayName !== "" && !subs.includes(displayName)) {
              subs.push(displayName);
            }
          }
        }
      }
    });
  }

  profile.subordinates = subs;
  return { status: "success", data: profile };
}

async function handleGetEmployees() {
  const data = await getSheetValues("employee");
  if (!data || data.length <= 1) return { status: "success", data: [] };
  const headers = data[0];
  const emailIdx = headers.findIndex((h: any) => /email/i.test(String(h).trim()));
  const userIdx = headers.findIndex((h: any) => /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()));
  const idx = {
    name: headers.findIndex((h: any) =>
      /nama|name|pic/i.test(String(h).trim()),
    ),
    email: emailIdx !== -1 ? emailIdx : headers.findIndex((h: any) => /email|user/i.test(String(h).trim())),
    username: userIdx !== -1 ? userIdx : -1,
    pos: headers.findIndex((h: any) =>
      /position|jabatan/i.test(String(h).trim()),
    ),
    prov: headers.findIndex((h: any) =>
      /province|provinsi/i.test(String(h).trim()),
    ),
    area: headers.findIndex((h: any) => /area/i.test(String(h).trim())),
    upline: headers.findIndex((h: any) =>
      /upline|spv|supervisor|atasan|manager/i.test(String(h).trim()),
    ),
    password: headers.findIndex((h: any) =>
      /password|pass/i.test(String(h).trim()),
    ),
    level: headers.findIndex((h: any) => /level|grade/i.test(String(h).trim())),
    group: headers.findIndex((h: any) =>
      /group|tim|divisi|division/i.test(String(h).trim()),
    ),
    status: headers.findIndex((h: any) =>
      /status|aktif|active/i.test(String(h).trim()),
    ),
  };

  const result = data
    .slice(1)
    .filter((row) => row[idx.name] !== "" && row[idx.name] !== undefined)
    .map((row) => {
      const p = idx.pos !== -1 ? row[idx.pos] : "Business Solution";
      return {
        name: idx.name !== -1 ? String(row[idx.name] || "").trim() : "",
        email: idx.email !== -1 ? String(row[idx.email] || "").trim() : "",
        user: idx.username !== -1 ? String(row[idx.username] || "").trim() : "",
        position: normalizePosition(p),
        province: idx.prov !== -1 ? String(row[idx.prov] || "").trim() : "-",
        area: idx.area !== -1 ? String(row[idx.area] || "").trim() : "-",
        upline: idx.upline !== -1 ? String(row[idx.upline] || "").trim() : "",
        password:
          idx.password !== -1 ? String(row[idx.password] || "").trim() : "",
        level:
          idx.level !== -1 &&
          row[idx.level] !== "" &&
          row[idx.level] !== undefined
            ? row[idx.level]
            : null,
        group: idx.group !== -1 ? String(row[idx.group] || "").trim() : "",
        status: idx.status !== -1 ? String(row[idx.status] || "").trim() : "Aktif",
      };
    });
    
  // Deduplicate employees by exact clean name or username to prevent key collisions
  const seenEmployees = new Set();
  const dedupedResult = [];
  for (const emp of result) {
    const cleanName = String(emp.name || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const cleanUser = String(emp.user || "").toLowerCase().replace(/[^a-z0-9]/g, "");
    const uniqKey = cleanName;
    
    if (!seenEmployees.has(uniqKey) && cleanName !== "") {
      seenEmployees.add(uniqKey);
      dedupedResult.push(emp);
    }
  }
  
  return { status: "success", data: dedupedResult };
}

async function handleGetInitialData(user: string) {
  try {
    // Fetch user profile, employees, channels, working data, sales data, and access rules in parallel
    const [
      profileJson,
      employeesJson,
      channelsJson,
      workingDataJson,
      drSalesDataJson,
      accessRulesJson,
    ] = await Promise.all([
      handleGetUserProfile(user),
      handleGetEmployees(),
      handleGetChannels(user),
      handleGetWorkingData(user),
      handleGetDrSalesData(user),
      handleGetAccessRules(),
    ]);

    return {
      status: "success",
      data: {
        profile: profileJson.status === "success" ? profileJson.data : null,
        employees: employeesJson.status === "success" ? employeesJson.data : [],
        channels: channelsJson.status === "success" ? channelsJson.data : [],
        workingData:
          workingDataJson.status === "success" ? workingDataJson.data : [],
        drSalesData:
          drSalesDataJson.status === "success" ? drSalesDataJson.data : [],
        accessRules:
          accessRulesJson.status === "success" ? accessRulesJson.data : {},
      },
    };
  } catch (error: any) {
    return { status: "error", message: error.toString() };
  }
}

// POST operations mirroring Apps Script logic exactly
async function handleBatchActivity(body: any) {
  const data = await getSheetValues("working");
  if (!data) throw new Error("Sheet 'working' tidak ditemukan");

  const headers = data[0];
  const getIdx = (patterns: RegExp) =>
    headers.findIndex((h) => patterns.test(String(h).trim()));

  let idxPog = headers.findIndex((h) =>
    /^pog$|^selisih$/i.test(String(h).trim()),
  );
  if (idxPog === -1) {
    headers.push("POG");
    idxPog = headers.length - 1;
    await updateSheetValues("working", data); // Overwrite first to include new header
  }

  const idx = {
    time: getIdx(/^tgl$|^waktu$|^date$|^timestamp$|^tanggal$/i),
    kiosk: getIdx(/^channel$|^kiosk$/i),
    user: getIdx(/^name checker$|^nama checker$|^user$|^pic$|^checker$/i),
    lot: getIdx(/^lot package$|^lot$/i),
    qty: getIdx(/^quantity \(kg\)|^qty$|^stock$|^kg$/i),
    area: getIdx(/^area$|^region$/i),
    desc: getIdx(/^hybrid$|^material$/i),
    exp: getIdx(/^exp date$|^expired$/i),
    agingMonth: getIdx(/^aging \(month\)/i),
    cond: getIdx(/^condition$|^kondisi$/i),
    crops: getIdx(/^crops$/i),
    dr: getIdx(/^shipping date$|^dr date$/i),
    agingExp: getIdx(/^aging to exp$/i),
    cluster: getIdx(/^cluster$/i),
    pog: idxPog,
  };

  const monthIndices = getMonthIndices(headers);
  const updMonthIndices = getUpdMonthIndices(headers);

  const currentMonthRowsMap: Record<string, { index: number; date: Date }> = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const k = idx.kiosk !== -1 ? cleanForMatch(row[idx.kiosk]) : "";
    const l = idx.lot !== -1 ? cleanForMatch(row[idx.lot]) : "";
    const h = idx.desc !== -1 ? cleanForMatch(row[idx.desc]) : "";
    const u = idx.user !== -1 ? cleanForMatch(row[idx.user]) : "";
    const key = `${k}_${l}_${h}_${u}`;
    if (k && l) {
      const rowDate =
        idx.time !== -1 && row[idx.time]
          ? parseGasDate(row[idx.time])
          : new Date(0);
      if (
        !currentMonthRowsMap[key] ||
        rowDate.getTime() > currentMonthRowsMap[key].date.getTime()
      ) {
        currentMonthRowsMap[key] = { index: i + 1, date: rowDate };
      }
    }
  }

  const pad = (n: number) => String(n).padStart(2, "0");
  const now = getJakartaDate();
  const timestamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

  const empData = (await getSheetValues("employee")) || [];

  if (body.items && body.items.length > 0) {
    for (const item of body.items) {
      const k = cleanForMatch(item.kiosk);
      const l = cleanForMatch(item.lot);
      const h = cleanForMatch(item.hybrid);
      const u = cleanForMatch(item.user || body.user);
      const key = `${k}_${l}_${h}_${u}`;
      const existingMatch = currentMonthRowsMap[key];
      const existingRow = existingMatch ? existingMatch.index : null;
      let agingExpVal = "";
      if (item.expired && item.expired !== "N/A") {
        const expD = new Date(item.expired);
        if (!isNaN(expD.getTime())) {
          const today = getJakartaDate();
          today.setHours(0, 0, 0, 0);
          agingExpVal = String(
            Math.round(
              (expD.getTime() - today.getTime()) / (1000 * 3600 * 24) / 30.416,
            ),
          );
        }
      }
      let clusterVal = "";
      if (item.aging && item.aging !== "-" && !isNaN(Number(item.aging))) {
        const aVal = Number(item.aging);
        if (aVal <= 2) clusterVal = "0-2";
        else if (aVal <= 4) clusterVal = "2-4";
        else if (aVal <= 6) clusterVal = "4-6";
        else if (aVal <= 9) clusterVal = "6-9";
        else if (aVal <= 12) clusterVal = "9-12";
        else clusterVal = ">12";
      }
      const itemMonthIdx = getMonthIndexFromDateString(timestamp);
      const monthColIdx = monthIndices[itemMonthIdx];

      let prevMonthStock = 0;
      if (existingRow) {
        const prevMonthIdx = (itemMonthIdx - 1 + 12) % 12;
        const prevMonthColIdx = monthIndices[prevMonthIdx];
        if (prevMonthColIdx !== -1) {
          prevMonthStock = Number(data[existingRow - 1][prevMonthColIdx]) || 0;
        } else {
          prevMonthStock = Number(item.originalStock) || 0;
        }
      } else {
        prevMonthStock = Number(item.originalStock) || 0;
      }

      if (item.condition === "habis" || Number(item.stock) === 0) {
        item.stock = 0;
      }
      const pogVal = prevMonthStock - (Number(item.stock) || 0);

      if (existingRow) {
        const rowIndex = existingRow - 1;
        if (monthColIdx !== -1) {
          data[rowIndex][monthColIdx] = item.stock;
        }
        const updMonthColIdx = updMonthIndices[itemMonthIdx];
        if (updMonthColIdx !== -1) {
          data[rowIndex][updMonthColIdx] = "sales";
        }

        if (idx.qty !== -1) data[rowIndex][idx.qty] = item.stock;
        if (idx.cond !== -1) data[rowIndex][idx.cond] = item.condition;
        if (idx.time !== -1) data[rowIndex][idx.time] = timestamp;
        if (idx.user !== -1) data[rowIndex][idx.user] = item.user || body.user;
        const resolvedArea =
          getUserProvince(item.user || body.user, empData) || body.area || "";
        if (idx.area !== -1 && resolvedArea)
          data[rowIndex][idx.area] = resolvedArea;
        if (idx.agingExp !== -1 && agingExpVal !== "")
          data[rowIndex][idx.agingExp] = agingExpVal;
        if (idx.cluster !== -1 && clusterVal !== "")
          data[rowIndex][idx.cluster] = clusterVal;
        if (idx.pog !== -1) data[rowIndex][idx.pog] = pogVal;
      } else {
        const newRow = new Array(headers.length).fill("");
        if (idx.time !== -1) newRow[idx.time] = timestamp;
        if (idx.kiosk !== -1) newRow[idx.kiosk] = item.kiosk;
        if (idx.user !== -1) newRow[idx.user] = item.user || body.user;
        if (idx.lot !== -1) newRow[idx.lot] = String(item.lot).toUpperCase();

        if (monthColIdx !== -1) {
          newRow[monthColIdx] = item.stock;
        }
        const updMonthColIdx = updMonthIndices[itemMonthIdx];
        if (updMonthColIdx !== -1) {
          newRow[updMonthColIdx] = "sales";
        }

        if (idx.qty !== -1) newRow[idx.qty] = item.stock;
        const resolvedArea =
          getUserProvince(item.user || body.user, empData) || body.area || "";
        if (idx.area !== -1) newRow[idx.area] = resolvedArea;
        if (idx.desc !== -1) newRow[idx.desc] = item.hybrid;
        if (idx.crops !== -1) newRow[idx.crops] = item.crops || "";
        if (idx.dr !== -1) newRow[idx.dr] = item.drDate || "";
        if (idx.exp !== -1) newRow[idx.exp] = item.expired;
        if (idx.agingMonth !== -1) newRow[idx.agingMonth] = item.aging;
        if (idx.cond !== -1) newRow[idx.cond] = item.condition;
        if (idx.agingExp !== -1) newRow[idx.agingExp] = agingExpVal;
        if (idx.cluster !== -1) newRow[idx.cluster] = clusterVal;
        if (idx.pog !== -1) newRow[idx.pog] = pogVal;

        for (let m = 0; m < 12; m++) {
          const colIdx = monthIndices[m];
          if (colIdx !== -1 && colIdx !== monthColIdx) {
            newRow[colIdx] = 0;
          }
        }
        data.push(newRow);
      }
    }
    await updateSheetValues("working", data);
  }
  return { status: "success" };
}

async function handleConsolidateDatabase(body: any) {
  // Fetch required sheets in parallel for high performance
  const [data, drData, empDataRaw] = await Promise.all([
    getSheetValues("working"),
    getSheetValues("dr"),
    getSheetValues("employee"),
  ]);

  if (!data || data.length <= 1) {
    return { status: "success", message: "Tidak ada data untuk dikonsolidasi" };
  }

  const empData = empDataRaw || [];
  const headers = data[0];
  const getIdx = (patterns: RegExp) =>
    headers.findIndex((h) => patterns.test(String(h).trim()));

  let idxPog = headers.findIndex((h) =>
    /^pog$|^selisih$/i.test(String(h).trim()),
  );
  if (idxPog === -1) {
    headers.push("POG");
    idxPog = headers.length - 1;
  }

  const idx = {
    time: getIdx(/^tgl$|^waktu$|^date$|^timestamp$|^tanggal$/i),
    kiosk: getIdx(/^channel$|^kiosk$/i),
    user: getIdx(/^name checker$|^nama checker$|^user$|^pic$|^checker$/i),
    lot: getIdx(/^lot package$|^lot$/i),
    qty: getIdx(/^quantity \(kg\)|^qty$|^stock$|^kg$/i),
    area: getIdx(/^area$|^region$/i),
    desc: getIdx(/^hybrid$|^material$/i),
    exp: getIdx(/^exp date$|^expired$/i),
    agingMonth: getIdx(/^aging \(month\)/i),
    cond: getIdx(/^condition$|^kondisi$/i),
    crops: getIdx(/^crops$/i),
    dr: getIdx(/^shipping date$|^dr date$/i),
    agingExp: getIdx(/^aging to exp$/i),
    cluster: getIdx(/^cluster$/i),
    pog: idxPog,
  };

  const monthIndices = getMonthIndices(headers);
  const updMonthIndices = getUpdMonthIndices(headers);

  const isValidVal = (v: any) => {
    return (
      v !== undefined &&
      v !== null &&
      String(v).trim() !== "" &&
      String(v).trim().toUpperCase() !== "N/A" &&
      String(v).trim() !== "-"
    );
  };

  const lotLookup: Record<string, { drDate: string; expDate: string }> = {};
  if (drData && drData.length > 1) {
    const drHeaders = drData[0];
    const drIdx = {
      lot: drHeaders.findIndex((h: any) => /lot/i.test(String(h).trim())),
      dr: drHeaders.findIndex((h: any) =>
        /dr date|shipping date/i.test(String(h).trim()),
      ),
      exp: drHeaders.findIndex((h: any) =>
        /exp date|expired/i.test(String(h).trim()),
      ),
    };
    if (drIdx.lot !== -1) {
      for (let j = 1; j < drData.length; j++) {
        const drRow = drData[j];
        const lNo = String(drRow[drIdx.lot] || "")
          .trim()
          .toUpperCase();
        if (lNo && !lotLookup[lNo]) {
          lotLookup[lNo] = {
            drDate:
              drIdx.dr !== -1 && drRow[drIdx.dr]
                ? formatMyDate(drRow[drIdx.dr])
                : "",
            expDate:
              drIdx.exp !== -1 && drRow[drIdx.exp]
                ? formatMyDate(drRow[drIdx.exp])
                : "",
          };
        }
      }
    }
  }

  const grouped: Record<string, any> = {};

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && idx.kiosk !== -1 && !row[idx.kiosk]) continue;

    const kioskVal =
      idx.kiosk !== -1 ? String(row[idx.kiosk] || "").trim() : "";
    const lotVal = idx.lot !== -1 ? String(row[idx.lot] || "").trim() : "";
    const descVal = idx.desc !== -1 ? String(row[idx.desc] || "").trim() : "";
    const userVal = idx.user !== -1 ? String(row[idx.user] || "").trim() : "";

    const groupKey = `${cleanForMatch(kioskVal)}_${cleanForMatch(lotVal)}_${cleanForMatch(descVal)}_${cleanForMatch(userVal)}`;
    const timestampStr = idx.time !== -1 ? row[idx.time] : "";
    const rowDate = timestampStr ? parseGasDate(timestampStr) : new Date(0);
    const rowMonthIdx = getMonthIndexFromDateString(timestampStr);

    const monthValsSrc = Array(12).fill(0);
    monthIndices.forEach((colIdx, mIdx) => {
      if (colIdx !== -1 && colIdx < row.length && row[colIdx] !== "") {
        monthValsSrc[mIdx] = Number(row[colIdx]) || 0;
      }
    });

    const updValsSrc = Array(12).fill("");
    updMonthIndices.forEach((colIdx, mIdx) => {
      if (
        colIdx !== -1 &&
        colIdx < row.length &&
        row[colIdx] !== undefined &&
        row[colIdx] !== null &&
        row[colIdx] !== ""
      ) {
        updValsSrc[mIdx] = String(row[colIdx]).trim();
      }
    });

    const totalMonthVals = monthValsSrc.reduce((a, b) => a + b, 0);
    const qtyVal = idx.qty !== -1 ? Number(row[idx.qty]) || 0 : 0;
    if (totalMonthVals === 0 && qtyVal > 0) {
      monthValsSrc[rowMonthIdx] = qtyVal;
    }

    const lotUpper = String(lotVal).trim().toUpperCase();
    let currentExp =
      idx.exp !== -1 && colIndexInBounds(idx.exp, row) ? row[idx.exp] : "";
    let currentDr =
      idx.dr !== -1 && colIndexInBounds(idx.dr, row) ? row[idx.dr] : "";

    if (!isValidVal(currentExp) && lotLookup[lotUpper]) {
      currentExp = lotLookup[lotUpper].expDate;
    }
    if (!isValidVal(currentDr) && lotLookup[lotUpper]) {
      currentDr = lotLookup[lotUpper].drDate;
    }

    if (!grouped[groupKey]) {
      grouped[groupKey] = {
        kiosk: kioskVal,
        lot: lotVal,
        desc: descVal,
        user: userVal,
        timestamp: rowDate,
        originalTimestampStr: timestampStr,
        area:
          getUserProvince(userVal, empData) ||
          (idx.area !== -1 && colIndexInBounds(idx.area, row)
            ? row[idx.area]
            : ""),
        crops:
          idx.crops !== -1 && colIndexInBounds(idx.crops, row)
            ? row[idx.crops]
            : "",
        exp: currentExp,
        dr: currentDr,
        agingMonth:
          idx.agingMonth !== -1 && colIndexInBounds(idx.agingMonth, row)
            ? row[idx.agingMonth]
            : "",
        cond:
          idx.cond !== -1 && colIndexInBounds(idx.cond, row)
            ? row[idx.cond]
            : "tetap",
        agingExp:
          idx.agingExp !== -1 && colIndexInBounds(idx.agingExp, row)
            ? row[idx.agingExp]
            : "",
        cluster:
          idx.cluster !== -1 && colIndexInBounds(idx.cluster, row)
            ? row[idx.cluster]
            : "",
        pog:
          idx.pog !== -1 && colIndexInBounds(idx.pog, row)
            ? Number(row[idx.pog]) || 0
            : 0,
        monthlyQty: monthValsSrc,
        updMonthVals: updValsSrc,
        rawIndex: i,
      };
    } else {
      const g = grouped[groupKey];
      for (let m = 0; m < 12; m++) {
        g.monthlyQty[m] += monthValsSrc[m];
      }
      if (!g.updMonthVals) g.updMonthVals = Array(12).fill("");
      for (let m = 0; m < 12; m++) {
        const v1 = String(g.updMonthVals[m] || "")
          .trim()
          .toLowerCase();
        const v2 = String(updValsSrc[m] || "")
          .trim()
          .toLowerCase();
        if (v1 === "sales" || v2 === "sales") {
          g.updMonthVals[m] = "sales";
        } else if (v1 === "admin" || v2 === "admin") {
          g.updMonthVals[m] = "admin";
        } else {
          g.updMonthVals[m] = "";
        }
      }

      const isNewer = rowDate.getTime() > g.timestamp.getTime();
      if (isNewer) {
        g.timestamp = rowDate;
        g.originalTimestampStr = timestampStr;
      }

      const updateField = (gKey: string, rowVal: any) => {
        if (isNewer) {
          if (isValidVal(rowVal)) g[gKey] = rowVal;
        } else {
          if (isValidVal(rowVal) && !isValidVal(g[gKey])) g[gKey] = rowVal;
        }
      };

      if (idx.area !== -1 && colIndexInBounds(idx.area, row))
        updateField("area", getUserProvince(userVal, empData) || row[idx.area]);
      if (idx.crops !== -1 && colIndexInBounds(idx.crops, row))
        updateField("crops", row[idx.crops]);
      if (idx.exp !== -1) updateField("exp", currentExp);
      if (idx.dr !== -1) updateField("dr", currentDr);
      if (idx.agingMonth !== -1 && colIndexInBounds(idx.agingMonth, row))
        updateField("agingMonth", row[idx.agingMonth]);
      if (idx.agingExp !== -1 && colIndexInBounds(idx.agingExp, row))
        updateField("agingExp", row[idx.agingExp]);
      if (idx.cluster !== -1 && colIndexInBounds(idx.cluster, row))
        updateField("cluster", row[idx.cluster]);

      if (idx.cond !== -1 && colIndexInBounds(idx.cond, row)) {
        const rowVal = row[idx.cond];
        if (isNewer && isValidVal(rowVal) && rowVal !== "tetap") {
          g.cond = rowVal;
        } else if (
          !isNewer &&
          isValidVal(rowVal) &&
          rowVal !== "tetap" &&
          (!g.cond || g.cond === "tetap")
        ) {
          g.cond = rowVal;
        }
      }

      if (idx.pog !== -1 && colIndexInBounds(idx.pog, row)) {
        g.pog += Number(row[idx.pog]) || 0;
      }
    }
  }

  const todayDate = getJakartaDate();
  todayDate.setHours(0, 0, 0, 0);
  const curMonthIdx = todayDate.getMonth();
  const prevMonthIdx = (curMonthIdx - 1 + 12) % 12;

  const newSheetValues = [headers];

  Object.values(grouped).forEach((g: any) => {
    if (!g.monthlyQty) g.monthlyQty = Array(12).fill(0);
    if (!g.updMonthVals) g.updMonthVals = Array(12).fill("");
    if (g.updMonthVals[curMonthIdx] !== "sales") {
      g.monthlyQty[curMonthIdx] = g.monthlyQty[prevMonthIdx] || 0;
      g.updMonthVals[curMonthIdx] = "admin";
    }

    const curStock = Number(g.monthlyQty[curMonthIdx]) || 0;
    const prevStock = Number(g.monthlyQty[prevMonthIdx]) || 0;
    g.pog = prevStock - curStock;

    if (isValidVal(g.dr) && g.dr !== "N/A") {
      g.dr = formatMyDate(parseGasDate(g.dr));
    }
    if (isValidVal(g.exp) && g.exp !== "N/A") {
      g.exp = formatMyDate(parseGasDate(g.exp));
    }

    if (isValidVal(g.dr) && g.dr !== "N/A") {
      const drD = parseGasDate(g.dr);
      if (drD && !isNaN(drD.getTime()) && drD.getTime() !== 0) {
        const calcAge = Math.round(
          (todayDate.getTime() - drD.getTime()) / (1000 * 3600 * 24) / 30.416,
        );
        g.agingMonth = calcAge >= 0 ? calcAge : 0;
      }
    }

    if (isValidVal(g.exp) && g.exp !== "N/A") {
      const expD = parseGasDate(g.exp);
      if (expD && !isNaN(expD.getTime()) && expD.getTime() !== 0) {
        g.agingExp = Math.round(
          (expD.getTime() - todayDate.getTime()) / (1000 * 3600 * 24) / 30.416,
        );
      }
    }

    if (
      g.agingMonth !== "" &&
      g.agingMonth !== undefined &&
      g.agingMonth !== null &&
      !isNaN(Number(g.agingMonth))
    ) {
      const aVal = Number(g.agingMonth);
      if (aVal <= 2) g.cluster = "0-2";
      else if (aVal <= 4) g.cluster = "2-4";
      else if (aVal <= 6) g.cluster = "4-6";
      else if (aVal <= 9) g.cluster = "6-9";
      else if (aVal <= 12) g.cluster = "9-12";
      else g.cluster = ">12";
    }

    const totalQty = g.monthlyQty.reduce((a: number, b: number) => a + b, 0);
    const currentMonthQty = g.monthlyQty[curMonthIdx];
    const prevMonthQty = g.monthlyQty[prevMonthIdx];

    if (totalQty === 0) {
      g.cond = "habis";
    } else {
      if (g.cond === "habis") g.cond = "tetap";
      if (g.cond !== "new" && g.cond !== "baru" && g.cond !== "baru (new)") {
        if (currentMonthQty < prevMonthQty) {
          g.cond = "berkurang";
        } else if (currentMonthQty > prevMonthQty) {
          g.cond = "bertambah";
        } else {
          g.cond = "tetap";
        }
      }
    }

    for (let m = 0; m < 12; m++) {
      const uVal = String(g.updMonthVals[m] || "")
        .trim()
        .toLowerCase();
      if (uVal === "sales") {
        g.updMonthVals[m] = "sales";
      } else if (uVal === "admin") {
        g.updMonthVals[m] = "admin";
      } else {
        g.updMonthVals[m] = "";
      }
    }

    const newRow = new Array(headers.length).fill("");
    const pad = (n: number) => String(n).padStart(2, "0");
    const now = getJakartaDate();
    const timestamp = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;

    if (idx.time !== -1) newRow[idx.time] = timestamp;
    if (idx.kiosk !== -1) newRow[idx.kiosk] = g.kiosk;
    if (idx.user !== -1) newRow[idx.user] = g.user;
    if (idx.lot !== -1) newRow[idx.lot] = String(g.lot).toUpperCase();
    if (idx.qty !== -1) newRow[idx.qty] = g.monthlyQty[curMonthIdx];

    if (idx.area !== -1) newRow[idx.area] = g.area;
    if (idx.desc !== -1) newRow[idx.desc] = g.desc;
    if (idx.crops !== -1) newRow[idx.crops] = g.crops;
    if (idx.dr !== -1) newRow[idx.dr] = g.dr;
    if (idx.exp !== -1) newRow[idx.exp] = g.exp;
    if (idx.agingMonth !== -1) newRow[idx.agingMonth] = g.agingMonth;
    if (idx.cond !== -1) newRow[idx.cond] = g.cond;
    if (idx.agingExp !== -1) newRow[idx.agingExp] = g.agingExp;
    if (idx.cluster !== -1) newRow[idx.cluster] = g.cluster;
    if (idx.pog !== -1) newRow[idx.pog] = g.pog;

    monthIndices.forEach((colIdx, mIdx) => {
      if (colIdx !== -1) {
        newRow[colIdx] =
          g.updMonthVals[mIdx] !== "" ? Number(g.monthlyQty[mIdx]) || 0 : "";
      }
    });

    updMonthIndices.forEach((colIdx, mIdx) => {
      if (colIdx !== -1) {
        newRow[colIdx] = g.updMonthVals[mIdx] || "";
      }
    });

    newSheetValues.push(newRow);
  });

  await updateSheetValues("working", newSheetValues);
  return { status: "success", message: "Konsolidasi berhasil dilakukan" };
}

function colIndexInBounds(idx: number, row: any[]) {
  return idx !== -1 && idx < row.length;
}

async function handleAddPartner(body: any) {
  const data = await getSheetValues("channel");
  if (!data) throw new Error("Sheet 'channel' tidak ditemukan");
  const headers = data[0];
  const idx = {
    pic: headers.findIndex((h: any) =>
      /pic|user|nama|analyst|solution/i.test(String(h).trim()),
    ),
    channel: headers.findIndex((h: any) =>
      /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()),
    ),
    cat: headers.findIndex((h: any) =>
      /kategori|category|klasifikasi|^cat$/i.test(String(h).trim()),
    ),
    upline: headers.findIndex((h: any) =>
      /upline|spv|supervisor/i.test(String(h).trim()),
    ),
    province: headers.findIndex((h: any) =>
      /provinsi|province/i.test(String(h).trim()),
    ),
    area: headers.findIndex((h: any) =>
      /^area$/i.test(String(h).trim()),
    ),
    group: headers.findIndex((h: any) =>
      /group|tim|divisi|division/i.test(String(h).trim()),
    ),
  };

  if (idx.channel === -1) throw new Error("Kolom nama partner tidak ditemukan");

  // Prevent duplicate partners!
  if (body.name) {
    const cleanNewName = cleanForMatch(body.name);
    const existingIndex = data.findIndex(
      (row, idxVal) =>
        idxVal > 0 &&
        cleanForMatch(row[idx.channel]) === cleanNewName,
    );
    if (existingIndex !== -1) {
      // Instead of throwing an error, update the existing row and return success
      let userProvince = body.province || "";
      let userArea = "";

      if (body.pic && data.length > 1) {
        const cleanPic = cleanForMatch(body.pic);
        const existingPicRow = data.find(
          (row, idxVal) =>
            idxVal > 0 &&
            idxVal !== existingIndex &&
            idx.pic !== -1 &&
            cleanForMatch(row[idx.pic]) === cleanPic
        );
        if (existingPicRow) {
          if (idx.province !== -1 && existingPicRow[idx.province]) {
            userProvince = String(existingPicRow[idx.province]).trim();
          }
          if (idx.area !== -1 && existingPicRow[idx.area]) {
            userArea = String(existingPicRow[idx.area]).trim();
          }
        }
      }

      if (idx.pic !== -1 && body.pic !== undefined) {
        data[existingIndex][idx.pic] = body.pic;
      }
      if (idx.province !== -1 && userProvince) {
        data[existingIndex][idx.province] = userProvince;
      }
      if (idx.area !== -1 && userArea) {
        data[existingIndex][idx.area] = userArea;
      }
      if (idx.cat !== -1 && body.category !== undefined && body.category !== "") {
        data[existingIndex][idx.cat] = body.category;
      }
      if (idx.group !== -1 && body.group !== undefined && body.group !== "") {
        data[existingIndex][idx.group] = body.group;
      }

      const successUpdate = await updateSheetValues("channel", data);
      if (!successUpdate) {
        throw new Error("Gagal memperbarui data partner yang sudah ada di database.");
      }
      return {
        status: "success",
        id: existingIndex + 1,
        message: `Partner "${body.name}" sudah ada di database, data berhasil diperbarui.`
      };
    }
  }

  let userProvince = body.province || "";
  let userArea = "";

  // Lookup in existing channel data to resolve PIC's Province and Area if possible
  if (body.pic && data.length > 1) {
    const cleanPic = cleanForMatch(body.pic);
    const existingPicRow = data.find(
      (row, idxVal) =>
        idxVal > 0 &&
        idx.pic !== -1 &&
        cleanForMatch(row[idx.pic]) === cleanPic
    );
    if (existingPicRow) {
      if (idx.province !== -1 && existingPicRow[idx.province]) {
        userProvince = String(existingPicRow[idx.province]).trim();
      }
      if (idx.area !== -1 && existingPicRow[idx.area]) {
        userArea = String(existingPicRow[idx.area]).trim();
      }
    }
  }

  const newRow = new Array(headers.length).fill("");
  if (idx.channel !== -1) newRow[idx.channel] = body.name || "";
  if (idx.cat !== -1) newRow[idx.cat] = body.category || "";
  if (idx.pic !== -1) newRow[idx.pic] = body.pic || "";
  if (idx.province !== -1) newRow[idx.province] = userProvince || "";
  if (idx.area !== -1) newRow[idx.area] = userArea || "";
  if (idx.group !== -1) newRow[idx.group] = body.group || "";

  const successAppend = await appendSheetRow("channel", newRow);
  if (!successAppend) {
    throw new Error("Gagal menyimpan partner baru ke database.");
  }
  
  // Return success with the exact sheet row number as the new ID
  const newId = data.length + 1;
  return { 
    status: "success", 
    id: newId, 
    message: `Partner "${body.name}" berhasil ditambahkan` 
  };
}

function findPartnerRowIndex(data: any[][], body: any, idx: any): number {
  const rowNum = Number(body.id);
  const targetName = body.originalName || body.name;
  if (!targetName) return -1;

  const cleanTargetName = cleanForMatch(targetName);

  // 1. Try direct index matching
  if (!isNaN(rowNum) && rowNum > 1 && rowNum <= data.length) {
    const potentialRow = data[rowNum - 1];
    if (potentialRow && idx.channel !== -1) {
      const currentClean = cleanForMatch(potentialRow[idx.channel]);
      if (currentClean === cleanTargetName) {
        return rowNum - 1;
      }
    }
  }

  // 2. If index didn't match (due to shifts), do a multi-criteria search
  const reqProvince = body.originalProvince || body.province ? cleanForMatch(body.originalProvince || body.province) : "";
  const reqGroup = body.originalGroup || body.group ? cleanForMatch(body.originalGroup || body.group) : "";
  const reqPic = body.originalPic || body.pic ? cleanForMatch(body.originalPic || body.pic) : "";
  const reqCat = body.originalCategory || body.category ? cleanForMatch(body.originalCategory || body.category) : "";

  let bestIndex = -1;
  let bestScore = -1;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row || idx.channel === -1) continue;

    const rowNameClean = cleanForMatch(row[idx.channel]);
    if (rowNameClean !== cleanTargetName) continue;

    // We have a name match. Let's calculate a similarity score based on other fields
    let score = 0;

    if (idx.province !== -1 && reqProvince) {
      const rowProv = cleanForMatch(row[idx.province]);
      if (rowProv === reqProvince) {
        score += 10;
      }
    }
    if (idx.group !== -1 && reqGroup) {
      const rowGroup = cleanForMatch(row[idx.group]);
      if (rowGroup === reqGroup) {
        score += 10;
      }
    }
    if (idx.pic !== -1 && reqPic) {
      const rowPic = cleanForMatch(row[idx.pic]);
      if (rowPic === reqPic) {
        score += 10;
      }
    }
    if (idx.cat !== -1 && reqCat) {
      const rowCat = cleanForMatch(row[idx.cat]);
      if (rowCat === reqCat) {
        score += 5;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestIndex = r;
    }
  }

  if (bestIndex !== -1) {
    return bestIndex;
  }

  // Last resort: find the first row matching the name
  return data.findIndex(
    (row, idxVal) =>
      idxVal > 0 &&
      cleanForMatch(row[idx.channel]) === cleanTargetName
  );
}

async function handleUpdatePartner(body: any) {
  console.log("handleUpdatePartner body:", body);
  const data = await getSheetValues("channel");
  if (!data) throw new Error("Sheet 'channel' tidak ditemukan");
  const headers = data[0];
  const idx = {
    pic: headers.findIndex((h: any) =>
      /pic|user|nama|analyst|solution/i.test(String(h).trim()),
    ),
    channel: headers.findIndex((h: any) =>
      /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()),
    ),
    cat: headers.findIndex((h: any) =>
      /kategori|category|klasifikasi|^cat$/i.test(String(h).trim()),
    ),
    upline: headers.findIndex((h: any) =>
      /upline|spv|supervisor/i.test(String(h).trim()),
    ),
    province: headers.findIndex((h: any) =>
      /provinsi|province/i.test(String(h).trim()),
    ),
    area: headers.findIndex((h: any) =>
      /^area$/i.test(String(h).trim()),
    ),
    group: headers.findIndex((h: any) =>
      /group|tim|divisi|division/i.test(String(h).trim()),
    ),
  };

  const rowIndex = findPartnerRowIndex(data, body, idx);

  if (rowIndex > 0 && rowIndex < data.length) {
    let userProvince = body.province || "";
    let userArea = "";

    // Lookup in existing channel data to resolve PIC's Province and Area if possible
    if (body.pic && data.length > 1) {
      const cleanPic = cleanForMatch(body.pic);
      const existingPicRow = data.find(
        (row, idxVal) =>
          idxVal > 0 &&
          idxVal !== rowIndex &&
          idx.pic !== -1 &&
          cleanForMatch(row[idx.pic]) === cleanPic
      );
      if (existingPicRow) {
        if (idx.province !== -1 && existingPicRow[idx.province]) {
          userProvince = String(existingPicRow[idx.province]).trim();
        }
        if (idx.area !== -1 && existingPicRow[idx.area]) {
          userArea = String(existingPicRow[idx.area]).trim();
        }
      }
    }

    if (idx.pic !== -1 && body.pic !== undefined) {
      data[rowIndex][idx.pic] = body.pic;
    }
    if (idx.province !== -1) {
      data[rowIndex][idx.province] = body.province || userProvince || data[rowIndex][idx.province] || "";
    }
    if (idx.area !== -1) {
      data[rowIndex][idx.area] = userArea || data[rowIndex][idx.area] || "";
    }
    if (idx.channel !== -1 && body.name !== undefined && body.name !== "") {
      data[rowIndex][idx.channel] = body.name;
    }
    if (idx.cat !== -1 && body.category !== undefined && body.category !== "") {
      data[rowIndex][idx.cat] = body.category;
    }
    if (idx.group !== -1 && body.group !== undefined && body.group !== "") {
      data[rowIndex][idx.group] = body.group;
    }
    const successUpdate = await updateSheetValues("channel", data);
    if (!successUpdate) {
      throw new Error("Gagal menyimpan perubahan partner ke database.");
    }
  } else {
    console.warn("Partner row not found for update, attempting to add instead");
    return await handleAddPartner(body);
  }
  return { status: "success", message: `Partner ${body.name} berhasil diperbarui` };
}

async function handleDeletePartner(body: any) {
  const data = await getSheetValues("channel");
  if (!data) throw new Error("Sheet 'channel' tidak ditemukan");
  const headers = data[0];
  const idx = {
    pic: headers.findIndex((h: any) =>
      /pic|user|nama|analyst|solution/i.test(String(h).trim()),
    ),
    channel: headers.findIndex((h: any) =>
      /channel|kiosk|nama toko|toko|name|distributor|partner|mitra/i.test(String(h).trim()),
    ),
    cat: headers.findIndex((h: any) =>
      /kategori|category|klasifikasi|^cat$/i.test(String(h).trim()),
    ),
    upline: headers.findIndex((h: any) =>
      /upline|spv|supervisor/i.test(String(h).trim()),
    ),
    province: headers.findIndex((h: any) =>
      /provinsi|province/i.test(String(h).trim()),
    ),
    area: headers.findIndex((h: any) =>
      /^area$/i.test(String(h).trim()),
    ),
    group: headers.findIndex((h: any) =>
      /group|tim|divisi|division/i.test(String(h).trim()),
    ),
  };

  if (idx.channel === -1) throw new Error("Kolom nama partner tidak ditemukan di sheet");

  const rowIndex = findPartnerRowIndex(data, body, idx);

  if (rowIndex > 0 && rowIndex < data.length) {
    const deletedName = data[rowIndex][idx.channel];
    data.splice(rowIndex, 1);
    const success = await updateSheetValues("channel", data);
    if (!success) throw new Error("Gagal menyimpan perubahan ke Google Sheets");
    console.log(`[Delete] Successfully deleted partner: ${deletedName} at row ${rowIndex + 1}`);
    return { status: "success", message: `Partner ${deletedName} berhasil dihapus` };
  }

  return { 
    status: "error", 
    message: rowIndex === -1 
      ? `Data partner "${body.name}" tidak ditemukan` 
      : "Indeks baris tidak valid untuk penghapusan"
  };
}

async function handleUpdateEmployee(body: any) {
  const data = await getSheetValues("employee");
  if (!data) throw new Error("Sheet 'employee' tidak ditemukan");
  const headers = data[0];
  const getIdx = (patterns: RegExp) =>
    headers.findIndex((h) => patterns.test(String(h).trim()));
  const emailIdx = headers.findIndex((h: any) => /email/i.test(String(h).trim()));
  const userIdx = headers.findIndex((h: any) => /^user$|^username$|^user\s*name$/i.test(String(h).trim().toLowerCase()));
  const idx = {
    name: getIdx(/nama|name|pic/i),
    email: emailIdx !== -1 ? emailIdx : getIdx(/email|user/i),
    username: userIdx !== -1 ? userIdx : -1,
    pos: getIdx(/position|jabatan/i),
    prov: getIdx(/province|provinsi/i),
    area: getIdx(/area/i),
    upline: getIdx(/upline|spv|supervisor|atasan|manager/i),
    password: getIdx(/password|pass/i),
    level: getIdx(/level|grade/i),
    group: getIdx(/group|tim|divisi|division/i),
  };

  if (idx.name === -1) throw new Error("Name column not found");

  const targetClean = cleanForMatch(body.originalName);
  let targetRow = -1;
  if (body.originalName) {
    for (let i = 1; i < data.length; i++) {
      if (cleanForMatch(data[i][idx.name]) === targetClean) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow !== -1) {
    const rowIndex = targetRow - 1;
    if (body.name !== undefined) data[rowIndex][idx.name] = body.name;
    if (idx.username !== -1 && body.user !== undefined) {
      data[rowIndex][idx.username] = body.user;
    } else if (idx.email !== -1 && body.email !== undefined) {
      data[rowIndex][idx.email] = body.email;
    }
    if (idx.pos !== -1 && body.position !== undefined)
      data[rowIndex][idx.pos] = body.position;
    if (idx.prov !== -1 && body.province !== undefined)
      data[rowIndex][idx.prov] = body.province;
    if (idx.area !== -1 && body.area !== undefined)
      data[rowIndex][idx.area] = body.area;
    if (idx.upline !== -1 && body.upline !== undefined)
      data[rowIndex][idx.upline] = body.upline;
    if (idx.password !== -1 && body.password !== undefined)
      data[rowIndex][idx.password] = body.password;
    if (idx.level !== -1 && body.level !== undefined)
      data[rowIndex][idx.level] = body.level;
    if (idx.group !== -1 && body.group !== undefined)
      data[rowIndex][idx.group] = body.group;
    await updateSheetValues("employee", data);
  } else {
    const newRow = new Array(headers.length).fill("");
    if (idx.name !== -1 && body.name !== undefined)
      newRow[idx.name] = body.name;
    if (idx.username !== -1 && body.user !== undefined) {
      newRow[idx.username] = body.user;
    } else if (idx.email !== -1 && body.email !== undefined) {
      newRow[idx.email] = body.email;
    }
    if (idx.pos !== -1 && body.position !== undefined)
      newRow[idx.pos] = body.position;
    if (idx.prov !== -1 && body.province !== undefined)
      newRow[idx.prov] = body.province;
    if (idx.area !== -1 && body.area !== undefined)
      newRow[idx.area] = body.area;
    if (idx.upline !== -1 && body.upline !== undefined)
      newRow[idx.upline] = body.upline;
    if (idx.password !== -1 && body.password !== undefined)
      newRow[idx.password] = body.password;
    if (idx.level !== -1 && body.level !== undefined)
      newRow[idx.level] = body.level;
    if (idx.group !== -1 && body.group !== undefined)
      newRow[idx.group] = body.group;
    await appendSheetRow("employee", newRow);
  }
  cachedEmployeeList = null;
  return { status: "success" };
}

async function handleDeleteEmployee(body: any) {
  const data = await getSheetValues("employee");
  if (!data) throw new Error("Sheet 'employee' tidak ditemukan");
  const headers = data[0];
  const nameIdx = headers.findIndex((h: any) =>
    /nama|name|pic/i.test(String(h).trim()),
  );
  if (nameIdx === -1) throw new Error("Kolom nama employee tidak ditemukan");

  const targetClean = cleanForMatch(body.name);
  let targetRow = -1;
  
  // Search for the employee by name
  for (let i = 1; i < data.length; i++) {
    if (cleanForMatch(data[i][nameIdx]) === targetClean) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow !== -1) {
    const deletedName = data[targetRow - 1][nameIdx];
    data.splice(targetRow - 1, 1);
    const success = await updateSheetValues("employee", data);
    if (!success) throw new Error("Gagal menyimpan perubahan ke Google Sheets");
    
    console.log(`[Delete] Successfully deleted employee: ${deletedName} at row ${targetRow}`);
    cachedEmployeeList = null;
    return { status: "success", message: `Employee ${deletedName} berhasil dihapus` };
  } else {
    return { status: "error", message: `Employee dengan nama "${body.name}" tidak ditemukan` };
  }
}

async function handleGetAccessRules() {
  try {
    const empData = await getSheetValues("employee");
    const uniquePositions: string[] = [];
    if (empData && empData.length > 1) {
      const headers = empData[0];
      const posIdx = headers.findIndex((h: any) => /position|jabatan/i.test(String(h).trim()));
      const nameIdx = headers.findIndex((h: any) => /nama|name|pic/i.test(String(h).trim()));
      for (let i = 1; i < empData.length; i++) {
        const row = empData[i];
        if (nameIdx !== -1 && !row[nameIdx]) continue;
        const rawPos = posIdx !== -1 ? row[posIdx] : "";
        const normalized = normalizePosition(rawPos);
        if (normalized && !uniquePositions.includes(normalized)) {
          uniquePositions.push(normalized);
        }
      }
    }
    // Always guarantee defaults are present
    const defaults = ["Business Analyst", "Vegetables Sales Manager", "Area Sales Manager", "Sales Agronomist", "Business Solution"];
    defaults.forEach(d => {
      if (!uniquePositions.includes(d)) {
        uniquePositions.push(d);
      }
    });

    const data = await getSheetValues("access");
    const rules: Record<string, any> = {};
    const existingPositions: string[] = [];

    if (data && data.length > 1) {
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const position = String(row[0]).trim();
        if (!position) continue;
        existingPositions.push(position);
        rules[position] = {
          home: row[1] === true || String(row[1]).toUpperCase() === "TRUE",
          partner: row[2] === true || String(row[2]).toUpperCase() === "TRUE",
          stock: row[3] === true || String(row[3]).toUpperCase() === "TRUE",
          pog: row[4] === true || String(row[4]).toUpperCase() === "TRUE",
          overview: row[5] === true || String(row[5]).toUpperCase() === "TRUE",
          temp: row[6] === true || String(row[6]).toUpperCase() === "TRUE",
          access: row[7] === true || String(row[7]).toUpperCase() === "TRUE",
        };
      }
    }

    const missingPositions = uniquePositions.filter((p) => !existingPositions.includes(p));
    if (missingPositions.length > 0) {
      for (const p of missingPositions) {
        let home = "TRUE";
        let partner = "TRUE";
        let stock = "TRUE";
        let pog = "TRUE";
        let overview = "FALSE";
        let temp = "FALSE";
        let access = "FALSE";

        if (p === "Business Analyst") {
          overview = "TRUE";
          temp = "TRUE";
          access = "TRUE";
        }

        const newRow = [p, home, partner, stock, pog, overview, temp, access];
        await appendSheetRow("access", newRow);

        rules[p] = {
          home: home === "TRUE",
          partner: partner === "TRUE",
          stock: stock === "TRUE",
          pog: pog === "TRUE",
          overview: overview === "TRUE",
          temp: temp === "TRUE",
          access: access === "TRUE",
        };
      }
    }

    return { status: "success", data: rules };
  } catch (error: any) {
    return { status: "error", message: error.toString() };
  }
}

async function handleSaveAccessRules(body: any) {
  try {
    const rules = body.rules || {};
    const headers = ["position", "home", "partner", "stock", "pog", "overview", "temp", "access"];
    const rows: any[][] = [headers];

    for (const position in rules) {
      const rule = rules[position];
      rows.push([
        position,
        rule.home ? "TRUE" : "FALSE",
        rule.partner ? "TRUE" : "FALSE",
        rule.stock ? "TRUE" : "FALSE",
        rule.pog ? "TRUE" : "FALSE",
        rule.overview ? "TRUE" : "FALSE",
        rule.temp ? "TRUE" : "FALSE",
        rule.access ? "TRUE" : "FALSE",
      ]);
    }

    await updateSheetValues("access", rows);
    return { status: "success" };
  } catch (error: any) {
    return { status: "error", message: error.toString() };
  }
}

// Extremely fast overview caching variables
let cachedOverviewData: any = null;
let lastOverviewFetchTime: number = 0;
let isFetchingOverview: boolean = false;

// Main API Router Route
app.all("/api", async (req, res) => {
  const action = (req.query.action || req.body?.action) as string;
  const user = (req.query.user || req.body?.user) as string;
  const lot = (req.query.lot || req.body?.lot) as string;

  console.log(
    `[API Call] Method: ${req.method}, Action: ${action}, User: ${user}`,
  );

  try {
    let result: any = null;

    if (req.method === "GET") {
      if (action === "getWorkingData")
        result = await handleGetWorkingData(user);
      else if (action === "getDrSalesData")
        result = await handleGetDrSalesData(user);
      else if (action === "getChannels") result = await handleGetChannels(user);
      else if (action === "getLotInfo") result = await handleGetLotInfo(lot);
      else if (action === "getUserProfile")
        result = await handleGetUserProfile(user);
      else if (action === "getEmployees") result = await handleGetEmployees();
      else if (action === "getInitialData")
        result = await handleGetInitialData(user);
      else if (action === "getOverviewData") {
        const fetchUrl = `https://script.google.com/macros/s/AKfycbxVCPzllBhLd8J-1UWeDJteFshrCs2M2PtehcXj7mHpW3PWfAcXe1d69NSkE9j3LYnM7A/exec?action=getOverviewData`;
        const cachePath = path.join(process.cwd(), "local_overview_cache.json");

        // 1. If we don't have it in memory, try to load from local file cache immediately
        if (!cachedOverviewData) {
          if (fs.existsSync(cachePath)) {
            try {
              const cacheText = fs.readFileSync(cachePath, "utf8");
              cachedOverviewData = JSON.parse(cacheText);
              console.log("[getOverviewData] Loaded cache file into memory successfully.");
            } catch (cacheErr) {
              console.error("[getOverviewData] Error loading cache file:", cacheErr);
              // Clean up the corrupted cache file immediately so it doesn't cause errors next time
              try {
                if (fs.existsSync(cachePath)) {
                  fs.unlinkSync(cachePath);
                  console.log("[getOverviewData] Deleted corrupted cache file successfully.");
                }
              } catch (unlinkErr) {
                console.error("[getOverviewData] Failed to delete corrupted cache file:", unlinkErr);
              }
            }
          } else {
            const apiTestPath = path.join(process.cwd(), "api_test.json");
            if (fs.existsSync(apiTestPath)) {
              try {
                const apiTestText = fs.readFileSync(apiTestPath, "utf8");
                cachedOverviewData = JSON.parse(apiTestText);
                console.log("[getOverviewData] Loaded fallback api_test.json into memory successfully.");
              } catch (fallbackErr) {
                console.warn("[getOverviewData] Error loading fallback api_test.json:", fallbackErr);
              }
            }
          }
        }

        // 2. Set the result to the cached data if we have it
        if (cachedOverviewData) {
          result = cachedOverviewData;
        }

        // 3. Trigger background refresh if it is stale (older than 10 minutes) AND we are not already fetching
        const now = Date.now();
        if (!isFetchingOverview && (now - lastOverviewFetchTime > 10 * 60 * 1000 || !cachedOverviewData)) {
          isFetchingOverview = true;
          console.log("[getOverviewData] Cache is empty or stale. Triggering asynchronous background refresh from Google Apps Script...");

          // Running as non-blocking background task
          (async () => {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 120000); // Plenty of time for Google
            try {
              const resp = await fetch(fetchUrl, { signal: controller.signal });
              clearTimeout(timeoutId);

              if (resp.ok) {
                const text = await resp.text();
                const parsed = JSON.parse(text);
                if (parsed && parsed.status === "success") {
                  cachedOverviewData = parsed;
                  lastOverviewFetchTime = Date.now();
                  console.log("[getOverviewData] Background refresh succeeded. Cache updated.");
                  
                  // Atomic write to avoid corruption in case of crashes or interruptions
                  const tempPath = cachePath + ".tmp";
                  fs.writeFile(tempPath, text, "utf8", (err) => {
                    if (err) {
                      console.error("[getOverviewData] Error writing background temp cache file:", err);
                      return;
                    }
                    fs.rename(tempPath, cachePath, (renameErr) => {
                      if (renameErr) {
                        console.error("[getOverviewData] Error renaming background cache file:", renameErr);
                      }
                    });
                  });
                } else {
                  console.warn("[getOverviewData] Background fetch returned invalid status or format.");
                }
              } else {
                console.warn(`[getOverviewData] Background fetch failed with HTTP status: ${resp.status}`);
              }
            } catch (err: any) {
              clearTimeout(timeoutId);
              console.warn(`[getOverviewData] Background refresh failed: ${err.message || err}`);
            } finally {
              isFetchingOverview = false;
            }
          })();
        }

        // 4. If we STILL don't have result (e.g. first boot and no cache file exists), wait for the initial fetch synchronously (safe fallback)
        if (!result) {
          console.log("[getOverviewData] First boot and no cache file. Waiting for Google Apps Script synchronously...");
          const controller = new AbortController();
          const timeoutId = setTimeout(() => controller.abort(), 120000);
          try {
            const resp = await fetch(fetchUrl, { signal: controller.signal });
            clearTimeout(timeoutId);
            if (!resp.ok) throw new Error(`HTTP status ${resp.status}`);
            const text = await resp.text();
            const parsed = JSON.parse(text);
            if (parsed && parsed.status === "success") {
              cachedOverviewData = parsed;
              lastOverviewFetchTime = Date.now();
              
              // Atomic write to avoid corruption
              const tempPath = cachePath + ".tmp";
              fs.writeFileSync(tempPath, text, "utf8");
              fs.renameSync(tempPath, cachePath);
              
              result = parsed;
            } else {
              throw new Error("Invalid response format");
            }
          } catch (e: any) {
            clearTimeout(timeoutId);
            console.error("[getOverviewData] Synchronous initial fetch failed:", e.message || e);
            result = { status: "error", message: `Fetch failed: ${e.message}` };
          }
        }
      }
      else if (action === "getAccessRules")
        result = await handleGetAccessRules();
      else {
        return res
          .status(400)
          .json({ status: "error", message: "Unknown action" });
      }
    } else if (req.method === "POST") {
      if (action === "batchActivity")
        result = await handleBatchActivity(req.body);
      else if (action === "consolidateDatabase")
        result = await handleConsolidateDatabase(req.body);
      else if (action === "addPartner")
        result = await handleAddPartner(req.body);
      else if (action === "updatePartner")
        result = await handleUpdatePartner(req.body);
      else if (action === "deletePartner")
        result = await handleDeletePartner(req.body);
      else if (action === "updateEmployee")
        result = await handleUpdateEmployee(req.body);
      else if (action === "deleteEmployee")
        result = await handleDeleteEmployee(req.body);
      else if (action === "saveAccessRules")
        result = await handleSaveAccessRules(req.body);
      else {
        return res
          .status(400)
          .json({ status: "error", message: "Unknown action" });
      }
    } else {
      return res
        .status(405)
        .json({ status: "error", message: "Method not allowed" });
    }

    return res.json(result);
  } catch (error: any) {
    const errorStr = error.toString();
    if (errorStr.includes("fallback to AppScript") || errorStr.includes("API not configured")) {
      console.warn(`[API Proxy Fallback] Action: ${action}, Error: ${errorStr}`);
      return res.status(503).json({ status: "error", message: "API not configured, fallback to Apps Script" });
    }
    console.log(`[API Validation/Execution Error] Action: ${action}, Error: ${errorStr}`);
    return res.json({ status: "error", message: error.message || errorStr });
  }
});

// Vite Middleware & SPA serving
async function startServer() {
  initLocalDb();
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
