// ============================================================
// FINANCE TRACKER - Code.gs (Backend)
// ============================================================

const SPREADSHEET_ID = ''; // Leave blank — script will auto-detect active spreadsheet
const BUDGET_MONTHLY = 70000;

const SHEETS = {
  TRANSACTIONS: 'Transactions',
  CREDIT_CARDS: 'Credit Card Tracker',
  MONTHLY: 'Monthly Summary',
  CATEGORY: 'Category Summary',
  BUDGET: 'Budget Settings',
  KEYWORD_MAP: 'Keyword Mapping',
  CATEGORIES: 'Categories',
  LENDING: 'Lending Tracker'
};

const ACCOUNTS = [
  'HDFC Bank', 'Union Bank', 'SBI',
  'Sapphiro (ICICI)', 'Millennia (HDFC)', 'Amazon (ICICI)',
  'Amazon Pay Wallet', 'UPI Lite', 'Cash'
];

const CATEGORIES = [
  'Groceries', 'Food Orders', 'Eat Out', 'Cafe & Snacks', 'Fuel & Tolls',
  'Transport (Uber/Auto)', 'Medical & Health', 'Pharmacy',
  'Shopping (Clothing)', 'Shopping (General)', 'Electronics & Gadgets', 'Entertainment',
  'Gaming', 'Travel & Holidays', 'Flights', 'Train & Metro', 'Rent & Maintenance',
  'Bills & Utilities', 'Internet & Phone', 'Subscriptions', 'Education & Courses',
  'EMI & Loan', 'Fitness & Gym', 'Gifts & Donations',
  'Vehicle Maintenance & Repairs', 'Business & Work', 'Home Repair', 'Miscellaneous',
  'Salary', 'Freelance', 'Investment Returns', 'Refunds', 'Lend Recovery', 'Other Income',
  'Cashback & Rewards',
  'Mutual Fund', 'Stocks & Zerodha', 'Fixed Deposit',
  'Plot & Property', 'PPF & NPS', 'Other Investment',
  'Personal Lend', 'Business Lend'
];

const TRANSACTION_TYPES = ['Expense', 'Income', 'Transfer', 'Cashback', 'Investment', 'Lend', 'Recover'];

// ============================================================
// WEB APP ENTRY POINT
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Finance Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Handles file upload POST requests
function doPost(e) {
  try {
    const source = e.parameter.source || '';
    const fileData = e.parameters.fileData ? e.parameters.fileData[0] : null;
    const fileName = e.parameter.fileName || 'upload';

    if (!fileData) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'No file data received' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const decoded = Utilities.base64Decode(fileData);
    const isXLS = fileName.endsWith('.xls') || fileName.endsWith('.xlsx');
    
    let rows;
    if (isXLS) {
      rows = _xlsToRows(decoded, fileName);
    } else {
      const csvText = Utilities.newBlob(decoded).getDataAsString();
      rows = Utilities.parseCsv(csvText);
    }

    if (!rows || rows.length < 2) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'No data found in file' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const kwSh = ss.getSheetByName(SHEETS.KEYWORD_MAP);
    const kwData = kwSh.getLastRow() > 1 ? kwSh.getRange(2, 1, kwSh.getLastRow() - 1, 3).getValues() : [];
    const parsed = _parseBySource(rows, source, kwData);

    return ContentService.createTextOutput(JSON.stringify({ success: true, preview: parsed }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: e.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function _xlsToRows(decoded, fileName) {
  let tempFile = null;
  let convertedFile = null;
  try {
    const ext = (fileName || '').endsWith('.xlsx') ? '.xlsx' : '.xls';
    const mimeType = ext === '.xlsx'
      ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      : 'application/vnd.ms-excel';

    const blob = Utilities.newBlob(decoded, mimeType, 'temp_import' + ext);
    tempFile = DriveApp.createFile(blob);

    const token = ScriptApp.getOAuthToken();
    const copyResponse = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + tempFile.getId() + '/copy',
      {
        method: 'POST',
        headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ name: 'temp_converted_sheet', mimeType: 'application/vnd.google-apps.spreadsheet' }),
        muteHttpExceptions: true
      }
    );

    const copyResult = JSON.parse(copyResponse.getContentText());
    if (!copyResult.id) throw new Error('Conversion failed: ' + copyResponse.getContentText());

    const convertedSS = SpreadsheetApp.openById(copyResult.id);
    convertedFile = DriveApp.getFileById(copyResult.id);
    const sheet = convertedSS.getSheets()[0];
    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 4);

    if (lastRow < 1) throw new Error('File appears empty');
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();

  } finally {
    try { if (tempFile) tempFile.setTrashed(true); } catch(e) {}
    try { if (convertedFile) convertedFile.setTrashed(true); } catch(e) {}
  }
}

// ============================================================
// SETUP — Run once to create all sheets
// ============================================================

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupTransactions(ss);
  _setupCreditCards(ss);
  _setupMonthly(ss);
  _setupCategory(ss);
  _setupBudget(ss);
  _setupKeywordMap(ss);
  _setupCategories(ss);
  _setupLending(ss);
  Logger.log('✅ All sheets created successfully! Your Finance Tracker is ready.');
}

function _setupTransactions(ss) {
  let sh = ss.getSheetByName(SHEETS.TRANSACTIONS) || ss.insertSheet(SHEETS.TRANSACTIONS);
  sh.clearContents();
  const headers = ['ID', 'Date', 'Type', 'Category', 'Amount', 'Account', 'Description', 'Notes', 'Cashback?', 'Cashback Amount', 'Linked Tx ID', 'Month', 'Created At'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 80);
  sh.setColumnWidth(7, 220);
  sh.setColumnWidth(8, 180);
}

function _setupCreditCards(ss) {
  let sh = ss.getSheetByName(SHEETS.CREDIT_CARDS) || ss.insertSheet(SHEETS.CREDIT_CARDS);
  sh.clearContents();
  const headers = ['Card Name', 'Bank', 'Credit Limit', 'Current Outstanding', 'Statement Date', 'Due Date', 'Last Payment', 'Reward Points', 'Cashback This Month'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const cards = [
    ['Sapphiro', 'ICICI', '', 0, '', '', 0, 0, 0],
    ['Millennia', 'HDFC', '', 0, '', '', 0, 0, 0],
    ['Amazon Pay', 'ICICI', '', 0, '', '', 0, 0, 0],
  ];
  sh.getRange(2, 1, cards.length, cards[0].length).setValues(cards);
}

function _setupMonthly(ss) {
  let sh = ss.getSheetByName(SHEETS.MONTHLY) || ss.insertSheet(SHEETS.MONTHLY);
  sh.clearContents();
  const headers = ['Month', 'Total Income', 'Total Expense', 'Total Cashback', 'Net Spending', 'Budget', 'Remaining', 'Status'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function _setupCategory(ss) {
  let sh = ss.getSheetByName(SHEETS.CATEGORY) || ss.insertSheet(SHEETS.CATEGORY);
  sh.clearContents();
  const headers = ['Category', 'This Month', 'Last Month', 'All Time Total', 'Tx Count'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const catData = CATEGORIES.map(c => [c, 0, 0, 0, 0]);
  sh.getRange(2, 1, catData.length, 5).setValues(catData);
}

function _setupBudget(ss) {
  let sh = ss.getSheetByName(SHEETS.BUDGET) || ss.insertSheet(SHEETS.BUDGET);
  sh.clearContents();
  sh.getRange(1, 1).setValue('Monthly Budget (₹)').setFontWeight('bold');
  sh.getRange(1, 2).setValue(BUDGET_MONTHLY);
  sh.getRange(3, 1).setValue('Alert at (%)').setFontWeight('bold');
  sh.getRange(3, 2).setValue(80);
  sh.getRange(5, 1).setValue('Last Updated').setFontWeight('bold');
  sh.getRange(5, 2).setValue(new Date());
}

function _setupLending(ss) {
  let sh = ss.getSheetByName(SHEETS.LENDING) || ss.insertSheet(SHEETS.LENDING);
  sh.clearContents();
  const headers = ['Linked TX ID', 'Person', 'Amount', 'Date Lent', 'Due Date', 'Notes', 'Status', 'Repaid Amount'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(2, 150);
  sh.setColumnWidth(6, 200);
}

function _setupKeywordMap(ss) {
  let sh = ss.getSheetByName(SHEETS.KEYWORD_MAP) || ss.insertSheet(SHEETS.KEYWORD_MAP);
  sh.clearContents();
  const headers = ['Keyword', 'Category', 'Type'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  const keywords = [
    ['claude', 'Education & Courses', 'Expense'],
    ['chatgpt', 'Education & Courses', 'Expense'],
    ['openai', 'Education & Courses', 'Expense'],
    ['igst', 'Bills & Utilities', 'Expense'],
    ['markup fee', 'Bills & Utilities', 'Expense'],
    ['swiggy', 'Food Orders', 'Expense'],
    ['zomato', 'Food Orders', 'Expense'],
    ['blinkit', 'Groceries', 'Expense'],
    ['bigbasket', 'Groceries', 'Expense'],
    ['dmart', 'Groceries', 'Expense'],
    ['zepto', 'Groceries', 'Expense'],
    ['uber', 'Transport (Uber/Auto)', 'Expense'],
    ['ola', 'Transport (Uber/Auto)', 'Expense'],
    ['rapido', 'Transport (Uber/Auto)', 'Expense'],
    ['bpcl', 'Fuel & Tolls', 'Expense'],
    ['indian oil', 'Fuel & Tolls', 'Expense'],
    ['iocl', 'Fuel & Tolls', 'Expense'],
    ['hp petrol', 'Fuel & Tolls', 'Expense'],
    ['netflix', 'Entertainment', 'Expense'],
    ['hotstar', 'Entertainment', 'Expense'],
    ['spotify', 'Entertainment', 'Expense'],
    ['prime video', 'Entertainment', 'Expense'],
    ['amazon', 'Shopping (Clothing)', 'Expense'],
    ['flipkart', 'Shopping (Clothing)', 'Expense'],
    ['myntra', 'Shopping (Clothing)', 'Expense'],
    ['electricity', 'Bills & Utilities', 'Expense'],
    ['airtel', 'Internet & Phone', 'Expense'],
    ['jio', 'Internet & Phone', 'Expense'],
    ['vi ', 'Internet & Phone', 'Expense'],
    ['gym', 'Fitness & Gym', 'Expense'],
    ['pharmacy', 'Pharmacy', 'Expense'],
    ['medical', 'Medical & Health', 'Expense'],
    ['hospital', 'Medical & Health', 'Expense'],
    ['credit card bill', 'Transfer', 'Transfer'],
    ['cc bill', 'Transfer', 'Transfer'],
    ['salary', 'Salary', 'Income'],
    ['refund', 'Refunds', 'Recover'],
    ['cashback', 'Cashback & Rewards', 'Cashback'],
    // add at end:
    ['zerodha', 'Stocks & Zerodha', 'Investment'],
    ['mutual fund', 'Mutual Fund', 'Investment'],
    ['sip', 'Mutual Fund', 'Investment'],
    ['fd ', 'Fixed Deposit', 'Investment'],
    ['ppf', 'PPF & NPS', 'Investment'],
    ['nps', 'PPF & NPS', 'Investment'],
  ];
  sh.getRange(2, 1, keywords.length, 3).setValues(keywords);
}

// ============================================================
// DATE PARSING HELPER
// ============================================================

function _parseDate(dateVal) {
  if (dateVal instanceof Date && !isNaN(dateVal)) return dateVal;
  const s = String(dateVal || '').trim();
  if (!s) return new Date();

  let m = s.match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
  if (m) return new Date(2000 + parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));

  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));

  m = s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (m) return new Date(parseInt(m[3]), parseInt(m[2]) - 1, parseInt(m[1]));

  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));

  const d = new Date(s);
  return isNaN(d) ? new Date() : d;
}

// ============================================================
// TRANSACTION CRUD
// ============================================================

function addTransaction(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
    const lastRow = sh.getLastRow();
    const id = 'TX' + String(lastRow).padStart(5, '0');
    const now = new Date();

    const parsedDate = _parseDate(data.date);
    const normalizedDate = Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'yyyy-MM-dd');
    const month = Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'yyyy-MM');

    const row = [
      id, normalizedDate, data.type, data.category,
      parseFloat(data.amount), data.account,
      data.description || '', data.notes || '',
      data.isCashback ? 'Yes' : 'No',
      data.cashbackAmount ? parseFloat(data.cashbackAmount) : 0,
      data.linkedTxId || '', month,
      Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss')
    ];

    sh.appendRow(row);
    _updateSummaries(ss, data, month);
    return { success: true, id };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function getTransactions(limit) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const count = limit || 50;
  const startRow = Math.max(2, lastRow - count + 1);
  const data = sh.getRange(startRow, 1, lastRow - startRow + 1, 13).getValues();
  Logger.log('Total rows fetched: ' + data.length);
  try {
    return data.map(r => {
      let dateStr = '';
      try {
        const d = _parseDate(r[1]);
        dateStr = Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
      } catch(e) {
        dateStr = String(r[1] || '');
      }
      return {
        id: r[0], date: dateStr, type: r[2], category: r[3],
        amount: r[4], account: r[5], description: r[6],
        notes: r[7], isCashback: r[8], cashbackAmount: r[9],
        linkedTxId: r[10], month: String(r[11] || '').substring(0, 7)
      };
    });
  } catch(e) {
    Logger.log('getTransactions map error: ' + e.message);
    return [];
  }
}

function getOpenLends() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
  // Collect all linkedTxIds from Recover/Lend Recovery transactions
  const recoveredIds = new Set(
    data.filter(r => r[2] === 'Recover' && r[3] === 'Lend Recovery' && r[10])
        .map(r => String(r[10]))
  );
  // Return Lend transactions not yet recovered
  return data
    .filter(r => r[2] === 'Lend' && r[0] && !recoveredIds.has(String(r[0])))
    .map(r => {
      let dateStr = '';
      try { dateStr = Utilities.formatDate(_parseDate(r[1]), 'Asia/Kolkata', 'yyyy-MM-dd'); }
      catch(e) { dateStr = String(r[1] || ''); }
      return { id: r[0], date: dateStr, description: r[6], amount: r[4], account: r[5] };
    });
}

function getTransactionsByMonth(month) {
  // month = 'YYYY-MM', 'YYYY' (year only), or '' for recent 200
  if (!month) return getTransactions(200);
  const isYearOnly = /^\d{4}$/.test(month);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
  const result = [];
  data.forEach(r => {
    let txMonth = '';
    try {
      const d = _parseDate(r[1]);
      txMonth = Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM');
    } catch(e) {
      txMonth = String(r[11] || '').substring(0, 7);
    }
    const matches = isYearOnly ? txMonth.startsWith(month) : txMonth === month;
    if (matches) {
      let dateStr = '';
      try {
        const d = _parseDate(r[1]);
        dateStr = Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd');
      } catch(e) { dateStr = String(r[1] || ''); }
      result.push({
        id: r[0], date: dateStr, type: r[2], category: r[3],
        amount: r[4], account: r[5], description: r[6],
        notes: r[7], isCashback: r[8], cashbackAmount: r[9],
        linkedTxId: r[10], month: txMonth
      });
    }
  });
  return result;
}

function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const budgetSh = ss.getSheetByName(SHEETS.BUDGET);

  const budget = budgetSh.getRange(1, 2).getValue() || BUDGET_MONTHLY;
  const alertPct = budgetSh.getRange(3, 2).getValue() || 80;

  const now = new Date();
  const currentMonth = Utilities.formatDate(now, 'Asia/Kolkata', 'yyyy-MM');

  const lastRow = sh.getLastRow();
  let totalExpense = 0, totalIncome = 0, totalCashback = 0, totalInvested = 0, totalLent = 0;
  const categoryMap = {}, monthMap = {}, accountMap = {};

  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
    data.forEach(r => {
      const type = r[2], amount = parseFloat(r[4]) || 0;
      const cat = r[3], acc = r[5];
      const cbAmt = parseFloat(r[9]) || 0;

      let month = r[11];
      try {
        const parsedDate = _parseDate(r[1]);
        if (parsedDate && parsedDate.getFullYear() > 1971) {
          month = Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'yyyy-MM');
        }
      } catch(e) {}

      if (type === 'Expense') {
        totalExpense += amount;
        categoryMap[cat] = (categoryMap[cat] || 0) + (month === currentMonth ? amount : 0);
        monthMap[month] = monthMap[month] || { expense: 0, income: 0 };
        monthMap[month].expense += amount;
        accountMap[acc] = (accountMap[acc] || 0) + amount;
      } else if (type === 'Income') {
        totalIncome += amount;
        monthMap[month] = monthMap[month] || { expense: 0, income: 0 };
        monthMap[month].income += amount;
      } else if (type === 'Cashback') {
        totalCashback += amount;
      } else if (type === 'Investment') {
        totalInvested += amount;
      } else if (type === 'Lend') {
        totalLent += amount;
      }
      if (cbAmt > 0) totalCashback += cbAmt;
    });
  }

  const currentExpense = Object.values(categoryMap).reduce((a, b) => a + b, 0);
  const netSpending = currentExpense - totalCashback;

  return {
    budget, alertPct, currentMonth, currentExpense,
    totalCashback, netSpending,
    remaining: budget - currentExpense,
    percentUsed: Math.round((currentExpense / budget) * 100),
    categoryMap, monthMap, accountMap,
    totalIncome, totalExpense, totalInvested, totalLent
  };
}

function getCategoryDataForYear(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  const categoryMap = {}, accountMap = {};
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
    data.forEach(r => {
      const type = r[2], amount = parseFloat(r[4]) || 0;
      const cat = r[3], acc = r[5];
      let txYear = '';
      try {
        const parsedDate = _parseDate(r[1]);
        if (parsedDate && parsedDate.getFullYear() > 1971)
          txYear = String(parsedDate.getFullYear());
      } catch(e) {}
      if (txYear === String(year) && type === 'Expense') {
        categoryMap[cat] = (categoryMap[cat] || 0) + amount;
        accountMap[acc]  = (accountMap[acc]  || 0) + amount;
      }
    });
  }
  return { categoryMap, accountMap };
}

function getCategoryDataForMonth(month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  const categoryMap = {}, accountMap = {};
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
    data.forEach(r => {
      const type = r[2], amount = parseFloat(r[4]) || 0;
      const cat = r[3], acc = r[5];
      let txMonth = r[11];
      try {
        const parsedDate = _parseDate(r[1]);
        if (parsedDate && parsedDate.getFullYear() > 1971)
          txMonth = Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'yyyy-MM');
      } catch(e) {}
      if (txMonth === month && type === 'Expense') {
        categoryMap[cat] = (categoryMap[cat] || 0) + amount;
        accountMap[acc]  = (accountMap[acc]  || 0) + amount;
      }
    });
  }
  return { categoryMap, accountMap };
}

function getCategoryDataForPeriod(fromMonth) {
  // fromMonth = 'YYYY-MM' cutoff (>= this month) or '' for all time
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  const categoryMap = {};
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
    data.forEach(r => {
      if (r[2] !== 'Expense') return;
      const amount = parseFloat(r[4]) || 0;
      const cat = r[3];
      let txMonth = '';
      try {
        const parsedDate = _parseDate(r[1]);
        if (parsedDate && parsedDate.getFullYear() > 1971)
          txMonth = Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'yyyy-MM');
      } catch(e) { txMonth = String(r[11] || '').substring(0, 7); }
      if (!fromMonth || txMonth >= fromMonth) {
        categoryMap[cat] = (categoryMap[cat] || 0) + amount;
      }
    });
  }
  return { categoryMap };
}

function getAccountDataForPeriod(filter) {
  // filter = 'YYYY-MM' (exact), 'YYYY' (year), 'from:YYYY-MM' (>= cutoff), or '' (all)
  const isFrom     = filter && filter.startsWith('from:');
  const fromMonth  = isFrom ? filter.slice(5) : '';
  const cleanFilter = isFrom ? '' : filter;
  const isYearOnly = /^\d{4}$/.test(cleanFilter);
  const isMonth    = /^\d{4}-\d{2}$/.test(cleanFilter);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  const accountMap = {};
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
    data.forEach(r => {
      if (r[2] !== 'Expense') return;
      const amount = parseFloat(r[4]) || 0;
      const acc = r[5];
      let txMonth = '';
      try {
        const parsedDate = _parseDate(r[1]);
        if (parsedDate && parsedDate.getFullYear() > 1971)
          txMonth = Utilities.formatDate(parsedDate, 'Asia/Kolkata', 'yyyy-MM');
      } catch(e) { txMonth = String(r[11] || '').substring(0, 7); }
      let matches;
      if (isFrom)          matches = txMonth >= fromMonth;
      else if (!cleanFilter) matches = true;
      else if (isYearOnly) matches = txMonth.startsWith(cleanFilter);
      else if (isMonth)    matches = txMonth === cleanFilter;
      else                 matches = true;
      if (matches) accountMap[acc] = (accountMap[acc] || 0) + amount;
    });
  }
  return { accountMap };
}

function getInvestmentTrend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const lastRow = sh.getLastRow();
  const INV_CATS = ['Mutual Fund','Stocks & Zerodha','Fixed Deposit','Plot & Property','PPF & NPS','Other Investment'];
  // monthMap: { 'yyyy-MM': { 'Mutual Fund': 0, ... } }
  const monthMap = {};
  if (lastRow >= 2) {
    const data = sh.getRange(2, 1, lastRow - 1, 13).getValues();
    data.forEach(r => {
      const type = r[2], amount = parseFloat(r[4]) || 0, cat = r[3];
      if (type !== 'Investment') return;
      let month = r[11];
      try {
        const d = _parseDate(r[1]);
        if (d && d.getFullYear() > 1971)
          month = Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM');
      } catch(e) {}
      if (!monthMap[month]) monthMap[month] = {};
      monthMap[month][cat] = (monthMap[month][cat] || 0) + amount;
    });
  }
  return { monthMap, categories: INV_CATS };
}

function _updateSummaries(ss, data, month) {
  const msh = ss.getSheetByName(SHEETS.MONTHLY);
  const mData = msh.getLastRow() > 1 ? msh.getRange(2, 1, msh.getLastRow() - 1, 8).getValues() : [];
  let mRow = mData.findIndex(r => r[0] === month);
  const budget = ss.getSheetByName(SHEETS.BUDGET).getRange(1, 2).getValue() || BUDGET_MONTHLY;

  if (mRow === -1) {
    msh.appendRow([month, 0, 0, 0, 0, budget, budget, '✅ On Track']);
    mRow = msh.getLastRow() - 2;
  }

  const rowIdx = mRow + 2;
  const amt = parseFloat(data.amount);

  // Investment and Lend excluded from budget summaries intentionally
  if (data.type === 'Income') {
    msh.getRange(rowIdx, 2).setValue((msh.getRange(rowIdx, 2).getValue() || 0) + amt);
  } else if (data.type === 'Expense') {
    msh.getRange(rowIdx, 3).setValue((msh.getRange(rowIdx, 3).getValue() || 0) + amt);
  } else if (data.type === 'Cashback') {
    msh.getRange(rowIdx, 4).setValue((msh.getRange(rowIdx, 4).getValue() || 0) + amt);
  }

  const expense = msh.getRange(rowIdx, 3).getValue();
  const cashback = msh.getRange(rowIdx, 4).getValue();
  const net = expense - cashback;
  const remaining = budget - expense;
  const status = expense > budget ? '🔴 Over Budget' : expense > budget * 0.8 ? '🟡 Alert' : '✅ On Track';
  msh.getRange(rowIdx, 5, 1, 4).setValues([[net, budget, remaining, status]]);
}

// ============================================================
// LENDING TRACKER
// ============================================================

function getLendingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.LENDING);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues().map((r, i) => ({
    rowIdx: i,
    linkedTxId: r[0], person: r[1],
    amount: r[2], dateLent: r[3],
    dueDate: r[4], notes: r[5],
    status: r[6], repaidAmount: r[7]
  }));
}

function addLendingEntry(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.LENDING);
    sh.appendRow([
      data.linkedTxId || '',
      data.person || '',
      parseFloat(data.amount) || 0,
      data.date || '',
      data.dueDate || '',
      data.notes || '',
      'Pending',
      0
    ]);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function updateLendingStatus(rowIdx, repaidAmount, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.LENDING);
    sh.getRange(rowIdx + 2, 7).setValue(status);
    sh.getRange(rowIdx + 2, 8).setValue(parseFloat(repaidAmount) || 0);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// CONFIG GETTERS
// ============================================================

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.CATEGORIES);

  let categories = CATEGORIES;
  let categoryTypeMap = {};

  if (sh && sh.getLastRow() > 1) {
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
    categories = data
      .filter(r => String(r[1]).trim().toLowerCase() === 'yes' && String(r[0]).trim())
      .map(r => String(r[0]).trim());
    data.forEach(r => {
      if (String(r[0]).trim()) {
        categoryTypeMap[String(r[0]).trim()] = String(r[2]).trim() || 'Expense';
      }
    });
  } else {
    // fallback — build from hardcoded
    CATEGORIES.forEach(c => { categoryTypeMap[c] = _getCategoryType(c); });
  }

  return { accounts: ACCOUNTS, categories, types: TRANSACTION_TYPES, categoryTypeMap };
}

function getBudgetSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.BUDGET);
  return { budget: sh.getRange(1, 2).getValue(), alertPct: sh.getRange(3, 2).getValue() };
}

function updateBudget(amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.BUDGET);
  sh.getRange(1, 2).setValue(amount);
  sh.getRange(5, 2).setValue(new Date());
  return { success: true };
}

// ============================================================
// BULK IMPORT
// ============================================================

function importHTML(htmlContent, source) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const kwSh = ss.getSheetByName(SHEETS.KEYWORD_MAP);
    const kwData = kwSh.getLastRow() > 1 ? kwSh.getRange(2, 1, kwSh.getLastRow() - 1, 3).getValues() : [];
    const parsed = _parseBHIMHTML(htmlContent, source, kwData);
    Logger.log('BHIM parsed: ' + parsed.length + ' transactions');
    return { success: true, preview: parsed };
  } catch(e) {
    Logger.log('importHTML error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function _parseBHIMHTML(html, source, kwData) {
  const results = [];
  const seen = new Set();

  const tbodyMatch = html.match(/<tbody>([\s\S]*?)<\/tbody>/i);
  if (!tbodyMatch) return results;

  const rowMatches = tbodyMatch[1].match(/<tr>([\s\S]*?)<\/tr>/gi);
  if (!rowMatches) return results;

  const account = source === 'bhim_upilite' ? 'UPI Lite' : 'Union Bank';

  rowMatches.forEach(row => {
    const cells = [];
    const cellMatches = row.match(/<td>([\s\S]*?)<\/td>/gi);
    if (!cellMatches || cellMatches.length < 10) return;
    cellMatches.forEach(cell => { cells.push(cell.replace(/<[^>]+>/g, '').trim()); });

    const dateStr = cells[0];
    const receiver = cells[5];
    const refId = cells[6];
    const amount = parseFloat(cells[8]) || 0;
    const drCr = cells[9];
    const status = cells[10];

    if (!dateStr || !amount) return;
    if (status && status.toUpperCase() === 'FAILED') return;
    if (seen.has(refId)) return;
    seen.add(refId);

    let description = receiver;
    const bracketMatch = receiver.match(/\(([^)]+)\)\s*$/);
    if (bracketMatch) description = bracketMatch[1].trim();

    results.push({
      date: dateStr,
      type: drCr === 'CR' ? 'Income' : 'Expense',
      amount, account, description,
      notes: 'Ref: ' + refId,
      category: _matchCategory(description, kwData) || 'Miscellaneous',
      isCashback: false, cashbackAmount: 0
    });
  });

  return results;
}

function importCSV(csvData, source) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const kwSh = ss.getSheetByName(SHEETS.KEYWORD_MAP);
    const kwData = kwSh.getLastRow() > 1 ? kwSh.getRange(2, 1, kwSh.getLastRow() - 1, 3).getValues() : [];
    const rows = Utilities.parseCsv(csvData);
    if (rows.length < 2) return { success: false, error: 'Empty CSV' };
    const parsed = _parseBySource(rows, source, kwData);
    return { success: true, preview: parsed };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function importXLS(base64Data, source) {
  try {
    const decoded = Utilities.base64Decode(base64Data);
    const rows = _xlsToRows(decoded, 'import.xls');
    if (!rows || rows.length < 2) return { success: false, error: 'No data found in file' };

    const debugRows = rows.slice(0, 20).map(r => r.map(c => {
      if (c instanceof Date) return 'DATE:' + c.toISOString();
      return String(c).substring(0, 30);
    }));
    Logger.log('XLS rows preview: ' + JSON.stringify(debugRows));

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const kwSh = ss.getSheetByName(SHEETS.KEYWORD_MAP);
    const kwData = kwSh.getLastRow() > 1 ? kwSh.getRange(2, 1, kwSh.getLastRow() - 1, 3).getValues() : [];
    const parsed = _parseBySource(rows, source, kwData);
    Logger.log('Parsed ' + parsed.length + ' transactions');
    return { success: true, preview: parsed };
  } catch(e) {
    Logger.log('importXLS error: ' + e.message);
    return { success: false, error: e.message };
  }
}

function confirmImport(transactions) {
  let imported = 0, failed = 0;
  transactions.forEach(tx => {
    const result = addTransaction(tx);
    if (result.success) imported++;
    else failed++;
  });
  return { success: true, imported, failed };
}

function _parseBySource(rows, source, kwData) {
  const headers = rows[0].map(h => h.toLowerCase().trim());
  const results = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every(c => !String(c).trim())) continue;

    let tx = { type: 'Expense', category: 'Miscellaneous', notes: source };

    if (source === 'gpay' || source === 'bhim') {
      const dateIdx = headers.findIndex(h => h.includes('date'));
      const descIdx = headers.findIndex(h => h.includes('description') || h.includes('narration'));
      const amtIdx = headers.findIndex(h => h.includes('amount'));
      tx.date = row[dateIdx] || '';
      tx.description = row[descIdx] || '';
      tx.amount = Math.abs(parseFloat(row[amtIdx]) || 0);
      if (parseFloat(row[amtIdx]) > 0) tx.type = 'Income';
      tx.account = source === 'gpay' ? 'HDFC Bank' : 'SBI';

    } else if (source === 'hdfc_bank') {
      const firstCell = row[0];
      let dateStr = '';
      if (firstCell instanceof Date && !isNaN(firstCell)) {
        dateStr = Utilities.formatDate(firstCell, 'Asia/Kolkata', 'yyyy-MM-dd');
      } else {
        dateStr = String(firstCell || '').trim();
      }
      if (!dateStr || dateStr.includes('*') || dateStr.toLowerCase().includes('date')
          || dateStr.toLowerCase().includes('opening') || dateStr.toLowerCase().includes('statement')) continue;
      if (!/^\d{2}\/\d{2}\/\d{2,4}/.test(dateStr) && !(firstCell instanceof Date)) continue;

      const withdrawal = parseFloat(String(row[4] || '').replace(/[^0-9.]/g, '')) || 0;
      const deposit = parseFloat(String(row[5] || '').replace(/[^0-9.]/g, '')) || 0;
      if (!withdrawal && !deposit) continue;

      tx.date = dateStr;
      tx.description = String(row[1] || '').trim();
      tx.amount = withdrawal || deposit;
      tx.type = deposit > 0 && !withdrawal ? 'Income' : 'Expense';
      tx.account = 'HDFC Bank';
      if (row[2]) tx.notes = 'Ref: ' + String(row[2]).trim();

    } else if (source === 'sbi') {
      const firstCell = row[0];
      let dateStr = '';
      if (firstCell instanceof Date && !isNaN(firstCell)) {
        dateStr = Utilities.formatDate(firstCell, 'Asia/Kolkata', 'yyyy-MM-dd');
      } else {
        dateStr = String(firstCell || '').trim();
      }
      if (!dateStr || dateStr.includes('*') || dateStr.toLowerCase().includes('date')
          || dateStr.toLowerCase().includes('opening') || dateStr.toLowerCase().includes('statement')) continue;
      if (!/^\d{2}\/\d{2}\/\d{2,4}/.test(dateStr) && !(firstCell instanceof Date)) continue;

      const withdrawal = parseFloat(String(row[4] || '').replace(/[^0-9.]/g, '')) || 0;
      const deposit = parseFloat(String(row[5] || '').replace(/[^0-9.]/g, '')) || 0;
      if (!withdrawal && !deposit) continue;

      tx.date = dateStr;
      tx.description = String(row[1] || '').trim();
      tx.amount = withdrawal || deposit;
      tx.type = deposit > 0 && !withdrawal ? 'Income' : 'Expense';
      tx.account = 'SBI';
      if (row[2]) tx.notes = 'Ref: ' + String(row[2]).trim();

    } else if (source === 'union') {
      tx.date = row[0] || '';
      tx.description = row[1] || '';
      const debit = parseFloat(row[2]) || 0;
      const credit = parseFloat(row[3]) || 0;
      tx.amount = debit || credit;
      tx.type = credit > 0 ? 'Income' : 'Expense';
      tx.account = 'Union Bank';

    } else if (source === 'icici_cc') {
      Logger.log('ROW ' + i + ': ' + JSON.stringify(row.map((c, idx) => idx + ':' + String(c).substring(0, 25))));
      const firstCell = row[1];
      let dateStr = '';
      if (firstCell instanceof Date && !isNaN(firstCell)) {
        dateStr = Utilities.formatDate(firstCell, 'Asia/Kolkata', 'yyyy-MM-dd');
      } else {
        dateStr = String(firstCell || '').trim();
      }
      if (!/^\d{2}-\d{2}-\d{4}/.test(dateStr)) continue;

      const amtRaw = String(row[9] || '').trim();
      if (!amtRaw) continue;
      const isCr = amtRaw.toLowerCase().includes('cr');
      const amtNum = parseFloat(amtRaw.replace(/[^0-9.]/g, '')) || 0;
      if (!amtNum) continue;

      tx.date = dateStr;
      tx.description = String(row[5] || '').trim();
      tx.amount = amtNum;
      tx.type = isCr ? 'Cashback' : 'Expense';
      tx.account = 'Sapphiro (ICICI)';
      if (row[13]) tx.notes = 'Ref: ' + String(row[13]).trim();

    } else if (source === 'hdfc_cc') {
      tx.date = row[0] || '';
      tx.description = row[1] || '';
      tx.amount = Math.abs(parseFloat(row[2]) || 0);
      tx.type = (row[3] || '').toLowerCase().includes('cr') ? 'Cashback' : 'Expense';
      tx.account = 'Millennia (HDFC)';

    } else if (source === 'amazon_pay') {
      tx.date = row[0] || '';
      tx.description = row[1] || '';
      tx.amount = Math.abs(parseFloat(row[2]) || 0);
      tx.type = (row[3] || '').toLowerCase().includes('credit') ? 'Cashback' : 'Expense';
      tx.account = 'Amazon Pay Wallet';
    }

    tx.category = _matchCategory(tx.description, kwData) || 'Miscellaneous';
    if (!tx.amount || tx.amount === 0) continue;
    results.push(tx);
  }

  return results;
}

function _matchCategory(description, kwData) {
  const desc = (description || '').toLowerCase();
  for (const kw of kwData) {
    if (desc.includes((kw[0] || '').toLowerCase())) return kw[1];
  }
  return null;
}


function _setupCategories(ss) {
  let sh = ss.getSheetByName(SHEETS.CATEGORIES) || ss.insertSheet(SHEETS.CATEGORIES);
  sh.clearContents();
  const headers = ['Category Name', 'Active (Yes/No)', 'Type'];
  sh.getRange(1, 1, 1, 3).setValues([headers]).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  const catData = CATEGORIES.map(c => [c, 'Yes', _getCategoryType(c)]);
  sh.getRange(2, 1, catData.length, 3).setValues(catData);
}


function _getCategoryType(cat) {
  const investCats = ['Mutual Fund','Stocks & Zerodha','Fixed Deposit','Plot & Property','PPF & NPS','Other Investment'];
  const lendCats   = ['Personal Lend','Business Lend'];
  const incomeCats = ['Salary','Freelance','Investment Returns','Other Income'];
  const recoverCats = ['Refunds', 'Lend Recovery'];
  const cashbackCats = ['Cashback & Rewards'];
  if (investCats.includes(cat))  return 'Investment';
  if (lendCats.includes(cat))    return 'Lend';
  if (recoverCats.includes(cat)) return 'Recover';
  if (incomeCats.includes(cat))  return 'Income';
  if (cashbackCats.includes(cat))return 'Cashback';
  return 'Expense';
}
