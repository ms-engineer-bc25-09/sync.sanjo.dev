function getSheetByName_(sheetName) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(sheetName);

  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }

  return sheet;
}

function getSheetHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();

  if (lastColumn <= 0) {
    throw new Error('ヘッダー行が見つかりません');
  }

  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

function appendRowByHeaderMap_(sheetName, rowObject) {
  const sheet = getSheetByName_(sheetName);
  const headers = getSheetHeaders_(sheet);

  const row = headers.map(function (header) {
    const value = rowObject[header];
    return value === undefined || value === null ? '' : value;
  });

  sheet.appendRow(row);
  Logger.log('appendRowByHeaderMap_ success: ' + sheetName);
}

function appendInternalLogRow_(data) {
  const rowObject = {};

  rowObject[COLUMNS.INTERNAL_LOG.ID] = data.id || '';
  rowObject[COLUMNS.INTERNAL_LOG.CREATED_AT] = data.createdAt || new Date();
  rowObject[COLUMNS.INTERNAL_LOG.SOURCE] = data.source || '';
  rowObject[COLUMNS.INTERNAL_LOG.STATUS] = data.status || '';
  rowObject[COLUMNS.INTERNAL_LOG.CUSTOMER_NAME] = data.customerName || '';
  rowObject[COLUMNS.INTERNAL_LOG.CONTACT_NAME] = data.contactName || '';
  rowObject[COLUMNS.INTERNAL_LOG.PROJECT_NAME] = data.projectName || '';
  rowObject[COLUMNS.INTERNAL_LOG.DRAWING_NUMBER] = data.drawingNumber || '';
  rowObject[COLUMNS.INTERNAL_LOG.MATERIAL] = data.material || '';
  rowObject[COLUMNS.INTERNAL_LOG.SIZE] = data.size || '';
  rowObject[COLUMNS.INTERNAL_LOG.QUANTITY] = data.quantity || '';
  rowObject[COLUMNS.INTERNAL_LOG.DUE_DATE] = data.dueDate || '';
  rowObject[COLUMNS.INTERNAL_LOG.NOTES] = data.notes || '';
  rowObject[COLUMNS.INTERNAL_LOG.RAW_TEXT] = data.rawText || '';
  rowObject[COLUMNS.INTERNAL_LOG.LINE_USER_ID] = data.lineUserId || '';
  rowObject[COLUMNS.INTERNAL_LOG.AI_EXTRACTED_JSON] =
    data.aiExtractedJson || '';
  rowObject[COLUMNS.INTERNAL_LOG.SIMILAR_CASE] = data.similarCase || '';
  rowObject[COLUMNS.INTERNAL_LOG.PAST_UNIT_PRICE] = data.pastUnitPrice || '';
  rowObject[COLUMNS.INTERNAL_LOG.SUGGESTED_PRICE] = data.suggestedPrice || '';
  rowObject[COLUMNS.INTERNAL_LOG.SPREADSHEET_UPDATED_AT] =
    data.spreadsheetUpdatedAt || new Date();

  appendRowByHeaderMap_(SHEET_NAMES.INTERNAL_LOG, rowObject);
}

function appendLedgerRow_(data) {
  const rowObject = {};

  rowObject[COLUMNS.LEDGER.RECEIVED_AT] = data.receivedAt || new Date();
  rowObject[COLUMNS.LEDGER.STATUS] = data.status || '';
  rowObject[COLUMNS.LEDGER.SOURCE] = data.source || '';
  rowObject[COLUMNS.LEDGER.CUSTOMER_NAME] = data.customerName || '';
  rowObject[COLUMNS.LEDGER.CONTACT_NAME] = data.contactName || '';
  rowObject[COLUMNS.LEDGER.EMAIL] = data.email || '';
  rowObject[COLUMNS.LEDGER.PHONE] = data.phone || '';
  rowObject[COLUMNS.LEDGER.INQUIRY] = data.inquiry || '';
  rowObject[COLUMNS.LEDGER.DUE_DATE] = data.dueDate || '';
  rowObject[COLUMNS.LEDGER.MATERIAL] = data.material || '';
  rowObject[COLUMNS.LEDGER.SIZE_THICKNESS] = data.sizeThickness || '';
  rowObject[COLUMNS.LEDGER.QUANTITY] = data.quantity || '';
  rowObject[COLUMNS.LEDGER.NOTES] = data.notes || '';
  rowObject[COLUMNS.LEDGER.DRAWING_URL] = data.drawingUrl || '';
  rowObject[COLUMNS.LEDGER.RAW_JSON] = data.rawJson || '';

  appendRowByHeaderMap_(SHEET_NAMES.LEDGER, rowObject);
}

function buildInquiryText_(projectName, drawingNumber, fallbackText) {
  const name = (projectName || '').trim();
  const drawing = (drawingNumber || '').trim();
  const fallback = (fallbackText || '').trim();

  if (name && drawing) return name + ' / ' + drawing;
  if (name) return name;
  if (drawing) return drawing;
  if (fallback) return fallback.length > 60 ? fallback.slice(0, 60) : fallback;

  return '内容確認中';
}

function saveLineGeminiResultToSheets_(params) {
  const now = new Date();
  const result = params.geminiResult || {};
  const source = params.source || 'LINE';
  const status = params.status || '未対応';
  const rawText = params.rawText || '';
  const lineUserId = params.lineUserId || '';
  const drawingUrl = params.drawingUrl || '';
  const rawJson = JSON.stringify(result);

  appendInternalLogRow_({
    id: '',
    createdAt: now,
    source: source,
    status: status,
    customerName: result.customer_name || '',
    contactName: result.contact_name || '',
    projectName: result.project_name || '',
    drawingNumber: result.drawing_number || '',
    material: result.material || '',
    size: result.size_thickness || '',
    quantity: result.quantity || '',
    dueDate: result.desired_due_date || '',
    notes: result.notes || '',
    rawText: rawText,
    lineUserId: lineUserId,
    aiExtractedJson: rawJson,
    similarCase: '',
    pastUnitPrice: '',
    suggestedPrice: '',
    spreadsheetUpdatedAt: now,
  });

  appendLedgerRow_({
    receivedAt: now,
    status: status,
    source: source,
    customerName: result.customer_name || '',
    contactName: result.contact_name || '',
    email: '',
    phone: '',
    inquiry: buildInquiryText_(
      result.project_name || '',
      result.drawing_number || '',
      rawText
    ),
    dueDate: result.desired_due_date || '',
    material: result.material || '',
    sizeThickness: result.size_thickness || '',
    quantity: result.quantity || '',
    notes: result.notes || '',
    drawingUrl: drawingUrl,
    rawJson: rawJson,
  });

  Logger.log('saveLineGeminiResultToSheets_ success');
}
