const SPREADSHEET_ID = "1xhzm71s7Is5NhXAFZfpflAgKOLlwR16_-Yog-SUCM9M";

function doGet(e) {
  const action = e.parameter.action;
  
  try {
    if (action === "getOverviewData") {
      return getOverviewData();
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Unknown action"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getOverviewData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("working");
  
  if (!sheet) {
    throw new Error("Sheet 'working' tidak ditemukan.");
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  const headers = data[0];
  const rows = data.slice(1);
  
  const getIdx = (headerName) => {
    return headers.findIndex(h => h.toString().toLowerCase().trim() === headerName.toLowerCase().trim());
  };
  
  // Mapping index berdasarkan header yang diberikan
  const idx = {
    territory: getIdx("territory"),
    sa: getIdx("SA"),
    bs: getIdx("BS"),
    activity: getIdx("Activity"),
    month: getIdx("month"),
    budgetRp: getIdx("Budget (Rp)"),
    
    budAdvMontok: getIdx("bud ADV MONTOK"),
    budAdvJoss: getIdx("bud ADV JOSS"),
    budAdvBejo: getIdx("bud ADV BEJO"),
    budAdvGanesh: getIdx("bud ADV GANESH"),
    budAdvJago: getIdx("bud ADV JAGO"),
    budAdvJalu: getIdx("bud ADV JALU"),
    budAdvRuby: getIdx("bud ADV RUBY"),
    
    budgetActivity: getIdx("budget activity"),
    amountBudget: getIdx("amount budget"),
    budgetFarmerReach: getIdx("budget farmer reach"),
    budgetDirectSales: getIdx("budget direct sales"),
    
    actAdvMontok: getIdx("act ADV MONTOK"),
    actAdvJoss: getIdx("act ADV JOSS"),
    actAdvBejo: getIdx("act ADV BEJO"),
    actAdvGanesh: getIdx("act ADV GANESH"),
    actAdvJago: getIdx("act ADV JAGO"),
    actAdvJalu: getIdx("act ADV JALU"),
    actAdvRuby: getIdx("act ADV RUBY"),
    
    actualActivity: getIdx("actual activity"),
    actualAmount: getIdx("actual amount"),
    farmerReach: getIdx("farmer reach"),
    directSales: getIdx("direct sales")
  };
  
  const result = rows.map(row => {
    return {
      territory: idx.territory !== -1 ? row[idx.territory] : "",
      sa: idx.sa !== -1 ? row[idx.sa] : "",
      bs: idx.bs !== -1 ? row[idx.bs] : "",
      activity: idx.activity !== -1 ? row[idx.activity] : "",
      month: idx.month !== -1 ? row[idx.month] : "",
      budgetRp: idx.budgetRp !== -1 ? row[idx.budgetRp] : 0,
      
      "bud ADV MONTOK": idx.budAdvMontok !== -1 ? row[idx.budAdvMontok] : 0,
      "bud ADV JOSS": idx.budAdvJoss !== -1 ? row[idx.budAdvJoss] : 0,
      "bud ADV BEJO": idx.budAdvBejo !== -1 ? row[idx.budAdvBejo] : 0,
      "bud ADV GANESH": idx.budAdvGanesh !== -1 ? row[idx.budAdvGanesh] : 0,
      "bud ADV JAGO": idx.budAdvJago !== -1 ? row[idx.budAdvJago] : 0,
      "bud ADV JALU": idx.budAdvJalu !== -1 ? row[idx.budAdvJalu] : 0,
      "bud ADV RUBY": idx.budAdvRuby !== -1 ? row[idx.budAdvRuby] : 0,

      "act ADV MONTOK": idx.actAdvMontok !== -1 ? row[idx.actAdvMontok] : 0,
      "act ADV JOSS": idx.actAdvJoss !== -1 ? row[idx.actAdvJoss] : 0,
      "act ADV BEJO": idx.actAdvBejo !== -1 ? row[idx.actAdvBejo] : 0,
      "act ADV GANESH": idx.actAdvGanesh !== -1 ? row[idx.actAdvGanesh] : 0,
      "act ADV JAGO": idx.actAdvJago !== -1 ? row[idx.actAdvJago] : 0,
      "act ADV JALU": idx.actAdvJalu !== -1 ? row[idx.actAdvJalu] : 0,
      "act ADV RUBY": idx.actAdvRuby !== -1 ? row[idx.actAdvRuby] : 0,

      budgetActivity: idx.budgetActivity !== -1 ? row[idx.budgetActivity] : 0,
      amountBudget: idx.amountBudget !== -1 ? row[idx.amountBudget] : 0,
      budgetFarmerReach: idx.budgetFarmerReach !== -1 ? row[idx.budgetFarmerReach] : 0,
      budgetDirectSales: idx.budgetDirectSales !== -1 ? row[idx.budgetDirectSales] : 0,
      actualActivity: idx.actualActivity !== -1 ? row[idx.actualActivity] : 0,
      actualAmount: idx.actualAmount !== -1 ? row[idx.actualAmount] : 0,
      farmerReach: idx.farmerReach !== -1 ? row[idx.farmerReach] : 0,
      directSales: idx.directSales !== -1 ? row[idx.directSales] : 0
    };
  });
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    data: result
  })).setMimeType(ContentService.MimeType.JSON);
}
