/**
 * Sf AL AsdiKKAA — Google Sheets backend + formatted workbook setup
 *
 * The first 10 Requests columns and first 4 Donations columns remain compatible
 * with the HTML app. Extra columns are admin-only organization fields.
 * Public visitors can create requests and record contributions; they cannot edit
 * or delete existing records through this backend.
 */
const SHEET_DASHBOARD = 'Dashboard';
const SHEET_REQUESTS = 'Requests';
const SHEET_GIFTS = 'Donations';
const REQUEST_HEADERS = ['id','title','name','desc','date','target','owner','phone','photos','createdAt'];
const GIFT_HEADERS = ['id','rid','amount','ts'];
const REQUEST_ADMIN_HEADERS = ['Status','Category','Priority','Admin notes'];
const GIFT_ADMIN_HEADERS = ['Admin review','Contribution method'];
const COLORS = {
  navy:'#16324F', blue:'#2F6690', teal:'#3A7D7C', paleBlue:'#EAF2F8',
  paleTeal:'#E7F3F1', paleGreen:'#E2F0D9', paleAmber:'#FFF2CC',
  paleRed:'#FCE8E6', white:'#FFFFFF', text:'#23313F', muted:'#637381', border:'#D7E0E7'
};

function doGet(e) {
  // JSON API mode for the Vercel-hosted HTML. Keep the original HTML mode available too.
  const action = e && e.parameter ? String(e.parameter.action || '') : '';
  if (action === 'getSharedData') {
    return jsonOutput_(getSharedData());
  }
  ensureSheets_(false);
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('الأصدقاء')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** JSON API endpoint used by the Vercel serverless proxy. */
function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents || '{}');
    let result;
    if (body.action === 'createRequest') result = createRequest(body.payload);
    else if (body.action === 'recordContribution') result = recordContribution(body.payload);
    else throw new Error('Unsupported API action.');
    return jsonOutput_({ ok: true, result: result });
  } catch (err) {
    return jsonOutput_({ ok: false, error: String(err && err.message || err) });
  }
}

function jsonOutput_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run once manually from the Apps Script editor to create/style the workbook. */
function setupSheets() {
  ensureSheets_(true);
  return 'Workbook ready: Dashboard, Requests and Donations are formatted.';
}

function ensureSheets_(formatWorkbook) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open Apps Script from the target spreadsheet: Extensions > Apps Script.');
  const requests = getOrCreateSheet_(ss, SHEET_REQUESTS);
  const donations = getOrCreateSheet_(ss, SHEET_GIFTS);
  const dashboard = getOrCreateSheet_(ss, SHEET_DASHBOARD);
  ensureHeaders_(requests, REQUEST_HEADERS);
  ensureHeaders_(donations, GIFT_HEADERS);
  ensureAdminHeaders_(requests, REQUEST_ADMIN_HEADERS, 11);
  ensureAdminHeaders_(donations, GIFT_ADMIN_HEADERS, 5);
  if (formatWorkbook) {
    formatRequests_(requests);
    formatDonations_(donations);
    formatDashboard_(dashboard);
    ss.setActiveSheet(dashboard);
    ss.moveActiveSheet(1);
  }
}

function getOrCreateSheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }

function ensureHeaders_(sheet, headers) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  headers.forEach((h, i) => { if (!current[i]) sheet.getRange(1, i + 1).setValue(h); });
}

function ensureAdminHeaders_(sheet, headers, startCol) {
  const current = sheet.getRange(1, startCol, 1, headers.length).getDisplayValues()[0];
  headers.forEach((h, i) => {
    // Earlier setup drafts sometimes used an unused "done" column at Requests!K1.
    if (!current[i] || (startCol === 11 && i === 0 && current[i] === 'done')) {
      sheet.getRange(1, startCol + i).setValue(h);
    }
  });
}

function formatHeader_(sheet, lastCol) {
  sheet.getRange(1, 1, 1, lastCol)
    .setBackground(COLORS.navy).setFontColor(COLORS.white).setFontWeight('bold')
    .setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 36);
}

function applyBanding_(sheet, lastCol) {
  sheet.getBandings().forEach(b => b.remove());
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1000), lastCol)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_BLUE, true, false);
  formatHeader_(sheet, lastCol);
}

function ensureFilter_(sheet, lastCol) {
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getMaxRows(), lastCol).createFilter();
}

function formatRequests_(sheet) {
  const lastCol = REQUEST_HEADERS.length + REQUEST_ADMIN_HEADERS.length;
  applyBanding_(sheet, lastCol);
  ensureFilter_(sheet, lastCol);
  sheet.setTabColor(COLORS.blue);
  [150,210,150,320,115,125,150,145,180,145,125,150,110,260]
    .forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('D:D').setWrap(true).setVerticalAlignment('top');
  sheet.getRange('E:E').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('F:F').setNumberFormat('#,##0 "EGP"');
  sheet.getRange('H:H').setNumberFormat('@');
  sheet.getRange('I:I').setWrap(true);
  sheet.getRange('J:J').setNumberFormat('dd mmm yyyy hh:mm');
  sheet.getRange('N:N').setWrap(true).setVerticalAlignment('top');

  const n = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2,11,n,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Open','In progress','Completed','On hold'], true).setAllowInvalid(false).build());
  sheet.getRange(2,12,n,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Medical','Education','Living expenses','Emergency','Other'], true).setAllowInvalid(false).build());
  sheet.getRange(2,13,n,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Normal','High','Urgent'], true).setAllowInvalid(false).build());

  const status = sheet.getRange(2,11,n,1), priority = sheet.getRange(2,13,n,1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Open').setBackground(COLORS.paleBlue).setFontColor(COLORS.blue).setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('In progress').setBackground(COLORS.paleAmber).setFontColor('#7F6000').setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Completed').setBackground(COLORS.paleGreen).setFontColor('#38761D').setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('On hold').setBackground(COLORS.paleRed).setFontColor('#990000').setRanges([status]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Urgent').setBackground(COLORS.paleRed).setFontColor('#990000').setRanges([priority]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('High').setBackground(COLORS.paleAmber).setFontColor('#7F6000').setRanges([priority]).build()
  ]);
  sheet.getRange(1,1,Math.max(sheet.getLastRow(),2),lastCol).setFontFamily('Arial').setFontColor(COLORS.text);
  formatHeader_(sheet, lastCol);
}

function formatDonations_(sheet) {
  const lastCol = GIFT_HEADERS.length + GIFT_ADMIN_HEADERS.length;
  applyBanding_(sheet, lastCol);
  ensureFilter_(sheet, lastCol);
  sheet.setTabColor(COLORS.teal);
  [165,165,145,155,150,180].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('A:B').setNumberFormat('@');
  sheet.getRange('C:C').setNumberFormat('#,##0 "EGP"');
  sheet.getRange('D:D').setNumberFormat('dd mmm yyyy hh:mm');
  const n = Math.max(sheet.getMaxRows() - 1, 1);
  sheet.getRange(2,5,n,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Recorded','Needs review','Disputed'], true).setAllowInvalid(false).build());
  sheet.getRange(2,6,n,1).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Cash','Bank transfer','Other','Not specified'], true).setAllowInvalid(false).build());
  const review = sheet.getRange(2,5,n,1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Recorded').setBackground(COLORS.paleGreen).setFontColor('#38761D').setRanges([review]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Needs review').setBackground(COLORS.paleAmber).setFontColor('#7F6000').setRanges([review]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Disputed').setBackground(COLORS.paleRed).setFontColor('#990000').setRanges([review]).build()
  ]);
  sheet.getRange(1,1,Math.max(sheet.getLastRow(),2),lastCol).setFontFamily('Arial').setFontColor(COLORS.text);
  formatHeader_(sheet, lastCol);
}

function formatDashboard_(sheet) {
  // The dashboard is a generated summary tab; the two data tabs are not cleared.
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setTabColor(COLORS.teal);
  [190,145,30,190,145,30,190,145].forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  sheet.getRange('A1:H1').merge().setValue('SF AL ASDIKKAA  |  COMMUNITY SUPPORT DASHBOARD')
    .setBackground(COLORS.navy).setFontColor(COLORS.white).setFontFamily('Arial')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 48);
  [['A3:B3','TOTAL REQUESTS'],['C3:D3','OPEN / ACTIVE REQUESTS'],['E3:F3','TOTAL REQUEST TARGETS (EGP)'],['G3:H3','CONTRIBUTIONS RECORDED (EGP)']]
    .forEach(([range, label]) => sheet.getRange(range).merge().setValue(label)
      .setBackground(COLORS.blue).setFontColor(COLORS.white).setFontWeight('bold')
      .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true));
  sheet.getRange('A4:B5').merge().setFormula('=COUNTA(Requests!A2:A)');
  sheet.getRange('C4:D5').merge().setFormula('=COUNTIFS(Requests!A2:A,"<>",Requests!K2:K,"<>Completed")');
  sheet.getRange('E4:F5').merge().setFormula('=SUM(Requests!F2:F)');
  sheet.getRange('G4:H5').merge().setFormula('=SUM(Donations!C2:C)');
  ['A4:B5','C4:D5','E4:F5','G4:H5'].forEach(r => sheet.getRange(r)
    .setBackground(COLORS.paleBlue).setFontColor(COLORS.navy).setFontWeight('bold')
    .setFontSize(20).setHorizontalAlignment('center').setVerticalAlignment('middle'));
  sheet.getRange('E4:F5').setNumberFormat('#,##0 "EGP"');
  sheet.getRange('G4:H5').setNumberFormat('#,##0 "EGP"');
  sheet.setRowHeight(3, 38); sheet.setRowHeights(4, 2, 30);
  sheet.getRange('A7:H7').merge().setValue('HOW TO USE THIS WORKBOOK')
    .setBackground(COLORS.teal).setFontColor(COLORS.white).setFontWeight('bold').setFontSize(11);
  const guide = [
    ['1','Requests','Every public request is stored here. Columns A:J are app data; K:N are admin tracking fields.'],
    ['2','Donations','Every contribution entered in the app is recorded here. This is not payment verification.'],
    ['3','Dropdowns','Use Status, Category and Priority on Requests, and Admin review / Contribution method on Donations.'],
    ['4','Filters','Use the filter icons in each header to sort or narrow records. The header row stays frozen.'],
    ['5','Privacy','Anyone with access to the deployed app can submit entries. Keep private information out of descriptions.']
  ];
  sheet.getRange(8,1,guide.length,3).setValues(guide);
  for (let r = 8; r <= 12; r++) {
    sheet.getRange(r,3,1,6).merge();
    sheet.getRange(r,3).setWrap(true).setVerticalAlignment('middle');
    sheet.setRowHeight(r,42);
  }
  sheet.getRange('A8:A12').setFontWeight('bold').setHorizontalAlignment('center').setBackground(COLORS.paleTeal);
  sheet.getRange('B8:B12').setFontWeight('bold').setFontColor(COLORS.navy);
  sheet.getRange('A8:H12').setFontFamily('Arial').setFontColor(COLORS.text)
    .setBorder(true,true,true,true,true,true,COLORS.border,SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('A14:H14').merge().setValue('Tip: Run setupSheets() again after major changes to restore formatting and dropdowns. It rebuilds the Dashboard summary.')
    .setFontColor(COLORS.muted).setFontStyle('italic').setWrap(true);
  sheet.setRowHeight(14,36);
  sheet.setFrozenRows(1);
}

function getSharedData() {
  ensureSheets_(false);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requests = rowsAsObjects_(ss.getSheetByName(SHEET_REQUESTS), REQUEST_HEADERS).map(r => {
    let photos = [];
    try { photos = r.photos ? JSON.parse(r.photos) : []; } catch (_) { photos = []; }
    return {
      id:String(r.id || ''), title:String(r.title || ''), name:String(r.name || ''),
      desc:String(r.desc || ''), date:normalizeDate_(r.date), target:Number(r.target) || 0,
      owner:String(r.owner || ''), phone:String(r.phone || ''),
      photos:Array.isArray(photos) ? photos : [], createdAt:Number(r.createdAt) || 0
    };
  }).filter(r => r.id);
  const gifts = rowsAsObjects_(ss.getSheetByName(SHEET_GIFTS), GIFT_HEADERS).map(g => ({
    id:String(g.id || ''), rid:String(g.rid || ''), amount:Number(g.amount) || 0, ts:Number(g.ts) || 0
  })).filter(g => g.id && g.rid && g.amount > 0);
  return { requests, gifts };
}

/** Public create-only endpoint. Updating an existing request is refused. */
function createRequest(input) {
  ensureSheets_(false);
  const r = validateRequest_(input);
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_REQUESTS);
    if (findRowById_(sh, r.id) > 0) throw new Error('That request ID already exists. Please refresh and try again.');
    r.createdAt = Date.now();
    sh.appendRow([r.id,r.title,r.name,r.desc,r.date,r.target,r.owner,r.phone,JSON.stringify(r.photos),r.createdAt]);
    sh.getRange(sh.getLastRow(), 11).setValue('Open');
    return { ok:true };
  } finally { lock.releaseLock(); }
}

/** Records a contribution entered by a visitor; it does not process or verify payment. */
function recordContribution(input) {
  ensureSheets_(false);
  if (!input || typeof input !== 'object') throw new Error('Invalid contribution.');
  const rid = clean_(input.rid,100), id = clean_(input.id,100), amount = Number(input.amount);
  if (!rid || !id || !Number.isSafeInteger(amount) || amount < 1 || amount > 300000) throw new Error('Enter a valid contribution amount.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const reqSh = ss.getSheetByName(SHEET_REQUESTS), giftSh = ss.getSheetByName(SHEET_GIFTS);
    const row = findRowById_(reqSh, rid);
    if (row < 2) throw new Error('This request no longer exists.');
    const req = rowsAsObjects_(reqSh, REQUEST_HEADERS).find(r => String(r.id) === rid);
    if (!req) throw new Error('This request no longer exists.');
    const target = Number(req.target) || 0, allGifts = rowsAsObjects_(giftSh, GIFT_HEADERS);
    if (allGifts.some(g => String(g.id) === id)) throw new Error('This contribution was already recorded.');
    const raised = allGifts.filter(g => String(g.rid) === rid).reduce((sum,g) => sum + (Number(g.amount) || 0),0);
    const remaining = Math.max(0,target-raised);
    if (remaining < 1 || amount > remaining) throw new Error('That amount is greater than the remaining goal. Refresh the page and try again.');
    giftSh.appendRow([id,rid,amount,Date.now()]);
    giftSh.getRange(giftSh.getLastRow(), 5).setValue('Needs review');
    return { ok:true };
  } finally { lock.releaseLock(); }
}

function validateRequest_(input) {
  if (!input || typeof input !== 'object') throw new Error('Invalid request.');
  const r = {
    id:clean_(input.id,100), title:clean_(input.title,70), name:clean_(input.name,40),
    desc:clean_(input.desc,280), date:clean_(input.date,10), target:Number(input.target),
    owner:clean_(input.owner,40), phone:clean_(input.phone,20),
    photos:Array.isArray(input.photos) ? input.photos.slice(0,3) : []
  };
  if (!r.id || r.title.length < 3 || r.name.length < 2 || r.desc.length < 5 || r.owner.length < 2) throw new Error('Please complete the required request fields.');
  if (!Number.isSafeInteger(r.target) || r.target < 1 || r.target > 300000) throw new Error('The target must be between 1 and 300,000 EGP.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.date) || isNaN(Date.parse(r.date + 'T00:00:00Z'))) throw new Error('Enter a valid request date.');
  if (r.phone && !/^[\d+\s()\-]{7,20}$/.test(r.phone)) throw new Error('Enter a valid phone number.');
  r.photos = r.photos.map(p => {
    if (typeof p !== 'string' || p.length > 14000 || !/^data:image\/(jpeg|png|webp);base64,/.test(p)) throw new Error('One of the photos is invalid or too large.');
    return p;
  });
  if (JSON.stringify(r.photos).length > 48000) throw new Error('The photos are too large. Please use fewer or smaller photos.');
  return r;
}

function normalizeDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(value || '');
}
function clean_(value,maxLen) {
  return String(value == null ? '' : value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').trim().slice(0,maxLen);
}
function rowsAsObjects_(sheet,headers) {
  const last = sheet.getLastRow(); if (last < 2) return [];
  return sheet.getRange(2,1,last-1,headers.length).getValues().map(row => {
    const obj = {}; headers.forEach((h,i) => obj[h] = row[i]); return obj;
  });
}
function findRowById_(sheet,id) {
  const last = sheet.getLastRow(); if (last < 2) return -1;
  const found = sheet.getRange(2,1,last-1,1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}
