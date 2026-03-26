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
  const normalizedRowObject = normalizeRowObjectKeys_(rowObject);

  const row = headers.map(function (header) {
    const value = normalizedRowObject[normalizeHeaderName_(header)];
    return value === undefined || value === null ? '' : value;
  });

  sheet.appendRow(row);
  Logger.log('appendRowByHeaderMap_ success: ' + sheetName);
}

function updateRowByHeaderMap_(sheetName, rowNumber, rowObject) {
  const sheet = getSheetByName_(sheetName);
  const headers = getSheetHeaders_(sheet);
  const normalizedRowObject = normalizeRowObjectKeys_(rowObject);

  if (!rowNumber || rowNumber < 2) {
    throw new Error('更新対象の行番号が不正です: ' + rowNumber);
  }

  headers.forEach(function (header, index) {
    const normalizedHeader = normalizeHeaderName_(header);
    const hasNormalizedKey = Object.prototype.hasOwnProperty.call(
      normalizedRowObject,
      normalizedHeader
    );

    if (!hasNormalizedKey) {
      return;
    }

    const value = normalizedRowObject[normalizedHeader];
    sheet
      .getRange(rowNumber, index + 1)
      .setValue(value === undefined || value === null ? '' : value);
  });

  Logger.log(
    'updateRowByHeaderMap_ success: ' + sheetName + ' row=' + rowNumber
  );
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
  const sizeThicknessValue = data.sizeThickness || '';

  rowObject[COLUMNS.LEDGER.ID] = data.id || '';
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
  rowObject[COLUMNS.LEDGER.SIZE_THICKNESS] = sizeThicknessValue;
  rowObject[COLUMNS.LEDGER.QUANTITY] = data.quantity || '';
  rowObject[COLUMNS.LEDGER.NOTES] = data.notes || '';
  rowObject[COLUMNS.LEDGER.DRAWING_URL] = data.drawingUrl || '';
  rowObject[COLUMNS.LEDGER.RAW_JSON] = data.rawJson || '';

  appendRowByHeaderMap_(SHEET_NAMES.LEDGER, rowObject);
  forceWriteLedgerSizeThicknessToLastRow_(sizeThicknessValue);
}

function createLedgerEntry_(data) {
  const ledgerId = String(data.id || '').trim() || Utilities.getUuid();

  appendLedgerRow_(
    Object.assign({}, data, {
      id: ledgerId,
    })
  );

  return ledgerId;
}

function appendImageLedgerRow_(data) {
  const rowObject = {};

  rowObject[COLUMNS.IMAGE_LEDGER.ID] = data.id || '';
  rowObject[COLUMNS.IMAGE_LEDGER.RECEIVED_AT] = data.receivedAt || new Date();
  rowObject[COLUMNS.IMAGE_LEDGER.SOURCE] = data.source || '';
  rowObject[COLUMNS.IMAGE_LEDGER.PROJECT_TYPE] = data.projectType || '';
  rowObject[COLUMNS.IMAGE_LEDGER.CUSTOMER_NAME] = data.customerName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.CONTACT_NAME] = data.contactName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.EMAIL] = data.email || '';
  rowObject[COLUMNS.IMAGE_LEDGER.PHONE] = data.phone || '';
  rowObject[COLUMNS.IMAGE_LEDGER.PROJECT_NAME] = data.projectName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.DUE_DATE] = data.dueDate || '';
  rowObject[COLUMNS.IMAGE_LEDGER.MATERIAL] = data.material || '';
  rowObject[COLUMNS.IMAGE_LEDGER.SIZE_THICKNESS] = data.sizeThickness || '';
  rowObject[COLUMNS.IMAGE_LEDGER.QUANTITY] = data.quantity || '';
  rowObject[COLUMNS.IMAGE_LEDGER.NOTES] = data.notes || '';
  rowObject[COLUMNS.IMAGE_LEDGER.ORIGINAL_FILE_NAME] =
    data.originalFileName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.ORIGINAL_IMAGE_URL] =
    data.originalImageUrl || '';
  rowObject[COLUMNS.IMAGE_LEDGER.DRAWING_URL] = data.drawingUrl || '';
  rowObject[COLUMNS.IMAGE_LEDGER.STATUS] = data.status || '';
  rowObject[COLUMNS.IMAGE_LEDGER.LEDGER_ID] = data.ledgerId || '';

  appendRowByHeaderMap_(SHEET_NAMES.IMAGE_LEDGER, rowObject);
}

function upsertImageLedgerRow_(data) {
  const rowObject = {};
  const rowId = String(data.id || '').trim();

  rowObject[COLUMNS.IMAGE_LEDGER.ID] = data.id || '';
  rowObject[COLUMNS.IMAGE_LEDGER.RECEIVED_AT] = data.receivedAt || new Date();
  rowObject[COLUMNS.IMAGE_LEDGER.SOURCE] = data.source || '';
  rowObject[COLUMNS.IMAGE_LEDGER.PROJECT_TYPE] = data.projectType || '';
  rowObject[COLUMNS.IMAGE_LEDGER.CUSTOMER_NAME] = data.customerName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.CONTACT_NAME] = data.contactName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.EMAIL] = data.email || '';
  rowObject[COLUMNS.IMAGE_LEDGER.PHONE] = data.phone || '';
  rowObject[COLUMNS.IMAGE_LEDGER.PROJECT_NAME] = data.projectName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.DUE_DATE] = data.dueDate || '';
  rowObject[COLUMNS.IMAGE_LEDGER.MATERIAL] = data.material || '';
  rowObject[COLUMNS.IMAGE_LEDGER.SIZE_THICKNESS] = data.sizeThickness || '';
  rowObject[COLUMNS.IMAGE_LEDGER.QUANTITY] = data.quantity || '';
  rowObject[COLUMNS.IMAGE_LEDGER.NOTES] = data.notes || '';
  rowObject[COLUMNS.IMAGE_LEDGER.ORIGINAL_FILE_NAME] =
    data.originalFileName || '';
  rowObject[COLUMNS.IMAGE_LEDGER.ORIGINAL_IMAGE_URL] =
    data.originalImageUrl || '';
  rowObject[COLUMNS.IMAGE_LEDGER.DRAWING_URL] = data.drawingUrl || '';
  rowObject[COLUMNS.IMAGE_LEDGER.STATUS] = data.status || '';
  rowObject[COLUMNS.IMAGE_LEDGER.LEDGER_ID] = data.ledgerId || '';

  if (!rowId) {
    appendRowByHeaderMap_(SHEET_NAMES.IMAGE_LEDGER, rowObject);
    return;
  }

  const existingRowNumber = findRowNumberById_(
    SHEET_NAMES.IMAGE_LEDGER,
    COLUMNS.IMAGE_LEDGER.ID,
    rowId
  );

  if (existingRowNumber) {
    updateRowByHeaderMap_(SHEET_NAMES.IMAGE_LEDGER, existingRowNumber, rowObject);
    return;
  }

  appendRowByHeaderMap_(SHEET_NAMES.IMAGE_LEDGER, rowObject);
}

function appendImageInternalLogRow_(data) {
  const rowObject = {};

  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ID] = data.id || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.CREATED_AT] =
    data.createdAt || new Date();
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.UPDATED_AT] =
    data.updatedAt || data.createdAt || new Date();
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.SOURCE] = data.source || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.PROJECT_TYPE] = data.projectType || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.STATUS] = data.status || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.CUSTOMER_NAME] = data.customerName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.CONTACT_NAME] = data.contactName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.PROJECT_NAME] = data.projectName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ORIGINAL_FILE_NAME] =
    data.originalFileName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ORIGINAL_IMAGE_URL] =
    data.originalImageUrl || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.SAVED_IMAGE_URL] =
    data.savedImageUrl || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.OCR_TEXT] = data.ocrText || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.AI_EXTRACTED_JSON] =
    data.aiExtractedJson || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.VALIDATION_RESULT] =
    data.validationResult || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.PROCESSING_STATUS] =
    data.processingStatus || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ERROR_MESSAGE] = data.errorMessage || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.LEDGER_ID] = data.ledgerId || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.NOTES] = data.notes || '';

  appendRowByHeaderMap_(SHEET_NAMES.IMAGE_INTERNAL_LOG, rowObject);
}

function upsertImageInternalLogRow_(data) {
  const rowObject = {};
  const rowId = String(data.id || '').trim();

  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ID] = data.id || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.CREATED_AT] =
    data.createdAt || new Date();
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.UPDATED_AT] =
    data.updatedAt || data.createdAt || new Date();
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.SOURCE] = data.source || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.PROJECT_TYPE] = data.projectType || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.STATUS] = data.status || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.CUSTOMER_NAME] = data.customerName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.CONTACT_NAME] = data.contactName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.PROJECT_NAME] = data.projectName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ORIGINAL_FILE_NAME] =
    data.originalFileName || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ORIGINAL_IMAGE_URL] =
    data.originalImageUrl || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.SAVED_IMAGE_URL] =
    data.savedImageUrl || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.OCR_TEXT] = data.ocrText || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.AI_EXTRACTED_JSON] =
    data.aiExtractedJson || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.VALIDATION_RESULT] =
    data.validationResult || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.PROCESSING_STATUS] =
    data.processingStatus || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.ERROR_MESSAGE] = data.errorMessage || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.LEDGER_ID] = data.ledgerId || '';
  rowObject[COLUMNS.IMAGE_INTERNAL_LOG.NOTES] = data.notes || '';

  if (!rowId) {
    appendRowByHeaderMap_(SHEET_NAMES.IMAGE_INTERNAL_LOG, rowObject);
    return;
  }

  const existingRowNumber = findRowNumberById_(
    SHEET_NAMES.IMAGE_INTERNAL_LOG,
    COLUMNS.IMAGE_INTERNAL_LOG.ID,
    rowId
  );

  if (existingRowNumber) {
    updateRowByHeaderMap_(
      SHEET_NAMES.IMAGE_INTERNAL_LOG,
      existingRowNumber,
      rowObject
    );
    return;
  }

  appendRowByHeaderMap_(SHEET_NAMES.IMAGE_INTERNAL_LOG, rowObject);
}

function findRowNumberById_(sheetName, idHeaderName, rowId) {
  const normalizedRowId = String(rowId || '').trim();

  if (!normalizedRowId) {
    return null;
  }

  const sheet = getSheetByName_(sheetName);
  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) {
    return null;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const indexMap = buildHeaderIndexMap_(headers);
  const idColumnIndex = indexMap[normalizeHeaderName_(idHeaderName)];

  if (idColumnIndex === undefined || idColumnIndex === null) {
    return null;
  }

  for (let i = rows.length - 1; i >= 0; i -= 1) {
    const candidateId = String(rows[i][idColumnIndex] || '').trim();
    if (candidateId === normalizedRowId) {
      return i + 2;
    }
  }

  return null;
}

function getRowObjectByRowNumber_(sheetName, rowNumber) {
  const sheet = getSheetByName_(sheetName);
  const headers = getSheetHeaders_(sheet);

  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    throw new Error('取得対象の行番号が不正です: ' + rowNumber);
  }

  const row = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const result = {};

  headers.forEach(function (header, index) {
    result[String(header).trim()] = row[index];
  });

  return result;
}

function reflectImageLedgerRowToLedger_(imageLedgerRowNumber, ledgerId) {
  const rowNumber = parseInt(imageLedgerRowNumber, 10);
  const normalizedLedgerId = String(ledgerId || '').trim();

  if (!normalizedLedgerId) {
    throw new Error('ledgerId が指定されていません');
  }

  const imageLedgerRow = getRowObjectByRowNumber_(
    SHEET_NAMES.IMAGE_LEDGER,
    rowNumber
  );
  const existingLedgerId = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.LEDGER_ID] || ''
  ).trim();
  const existingStatus = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.STATUS] || ''
  ).trim();

  if (existingLedgerId) {
    throw new Error(
      'この画像案件はすでに案件台帳へ反映済みです: ' + existingLedgerId
    );
  }

  if (existingStatus === '案件台帳反映済') {
    throw new Error('この画像案件はすでに反映済みステータスです');
  }

  const matchedImageInternalLog =
    findMatchingImageInternalLogRow_(imageLedgerRow);
  const extractedJson = parseImageInternalLogExtractedJson_(
    matchedImageInternalLog
  );
  const sizeThickness =
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.SIZE_THICKNESS] ||
    extractedJson.size_thickness ||
    '';

  appendLedgerRow_({
    id: normalizedLedgerId,
    receivedAt: imageLedgerRow[COLUMNS.IMAGE_LEDGER.RECEIVED_AT] || new Date(),
    status: '未対応',
    source: imageLedgerRow[COLUMNS.IMAGE_LEDGER.SOURCE] || '',
    customerName: imageLedgerRow[COLUMNS.IMAGE_LEDGER.CUSTOMER_NAME] || '',
    contactName: imageLedgerRow[COLUMNS.IMAGE_LEDGER.CONTACT_NAME] || '',
    email: imageLedgerRow[COLUMNS.IMAGE_LEDGER.EMAIL] || '',
    phone: imageLedgerRow[COLUMNS.IMAGE_LEDGER.PHONE] || '',
    inquiry: imageLedgerRow[COLUMNS.IMAGE_LEDGER.PROJECT_NAME] || '',
    dueDate: imageLedgerRow[COLUMNS.IMAGE_LEDGER.DUE_DATE] || '',
    material: imageLedgerRow[COLUMNS.IMAGE_LEDGER.MATERIAL] || '',
    sizeThickness: sizeThickness,
    quantity: imageLedgerRow[COLUMNS.IMAGE_LEDGER.QUANTITY] || '',
    notes: imageLedgerRow[COLUMNS.IMAGE_LEDGER.NOTES] || '',
    drawingUrl: imageLedgerRow[COLUMNS.IMAGE_LEDGER.DRAWING_URL] || '',
    rawJson:
      matchedImageInternalLog[COLUMNS.IMAGE_INTERNAL_LOG.AI_EXTRACTED_JSON] ||
      '',
  });

  updateRowByHeaderMap_(SHEET_NAMES.IMAGE_LEDGER, rowNumber, {
    [COLUMNS.IMAGE_LEDGER.STATUS]: '案件台帳反映済',
    [COLUMNS.IMAGE_LEDGER.LEDGER_ID]: normalizedLedgerId,
  });

  updateMatchingImageInternalLogs_(imageLedgerRow, normalizedLedgerId);
  updateSupabaseProjectLedgerIdByImageLedger_(
    imageLedgerRow,
    normalizedLedgerId
  );

  Logger.log(
    'reflectImageLedgerRowToLedger_ success: imageRow=' +
      rowNumber +
      ' ledgerId=' +
      normalizedLedgerId
  );
}

function reflectImageLedgerRowToLedger(imageLedgerRowNumber, ledgerId) {
  return reflectImageLedgerRowToLedger_(imageLedgerRowNumber, ledgerId);
}

function testReflectImageLedgerRowToLedger() {
  // 必要に応じて対象行番号と案件台帳IDを変更して実行する
  reflectImageLedgerRowToLedger(12, 'KAN-TEST-0012');
}

function updateMatchingImageInternalLogs_(imageLedgerRow, ledgerId) {
  const matchedRowInfo = findMatchingImageInternalLogRowInfo_(imageLedgerRow);

  if (!matchedRowInfo) {
    return;
  }

  updateRowByHeaderMap_(
    SHEET_NAMES.IMAGE_INTERNAL_LOG,
    matchedRowInfo.rowNumber,
    {
      [COLUMNS.IMAGE_INTERNAL_LOG.UPDATED_AT]: new Date(),
      [COLUMNS.IMAGE_INTERNAL_LOG.STATUS]: '案件台帳反映済',
      [COLUMNS.IMAGE_INTERNAL_LOG.PROCESSING_STATUS]: '台帳反映済',
      [COLUMNS.IMAGE_INTERNAL_LOG.LEDGER_ID]: ledgerId,
    }
  );
}

function findMatchingImageInternalLogRow_(imageLedgerRow) {
  const matchedRowInfo = findMatchingImageInternalLogRowInfo_(imageLedgerRow);
  return matchedRowInfo ? matchedRowInfo.rowObject : {};
}

function findMatchingImageInternalLogRowInfo_(imageLedgerRow) {
  const sheet = getSheetByName_(SHEET_NAMES.IMAGE_INTERNAL_LOG);
  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) {
    return null;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const indexMap = buildHeaderIndexMap_(headers);
  const source = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.SOURCE] || ''
  ).trim();
  const projectName = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.PROJECT_NAME] || ''
  ).trim();
  const originalFileName = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.ORIGINAL_FILE_NAME] || ''
  ).trim();
  const drawingUrl = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.DRAWING_URL] || ''
  ).trim();

  for (let i = rows.length - 1; i >= 0; i -= 1) {
    const row = rows[i];
    const rowSource = String(
      getCellValueByHeader_(row, indexMap, COLUMNS.IMAGE_INTERNAL_LOG.SOURCE) ||
        ''
    ).trim();
    const rowProjectName = String(
      getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.IMAGE_INTERNAL_LOG.PROJECT_NAME
      ) || ''
    ).trim();
    const rowOriginalFileName = String(
      getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.IMAGE_INTERNAL_LOG.ORIGINAL_FILE_NAME
      ) || ''
    ).trim();
    const rowSavedImageUrl = String(
      getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.IMAGE_INTERNAL_LOG.SAVED_IMAGE_URL
      ) || ''
    ).trim();

    if (
      rowSource === source &&
      rowProjectName === projectName &&
      rowOriginalFileName === originalFileName &&
      (!drawingUrl || rowSavedImageUrl === drawingUrl)
    ) {
      return {
        rowNumber: i + 2,
        rowObject: buildRowObjectFromValues_(headers, row),
      };
    }
  }

  return null;
}

function buildRowObjectFromValues_(headers, row) {
  const result = {};

  headers.forEach(function (header, index) {
    result[String(header).trim()] = row[index];
  });

  return result;
}

function parseImageInternalLogExtractedJson_(imageInternalLogRow) {
  const rawJson =
    imageInternalLogRow[COLUMNS.IMAGE_INTERNAL_LOG.AI_EXTRACTED_JSON] || '';

  if (!rawJson) {
    return {};
  }

  try {
    return JSON.parse(String(rawJson));
  } catch (error) {
    Logger.log('parseImageInternalLogExtractedJson_ error: ' + error.message);
    return {};
  }
}

function forceWriteLedgerSizeThicknessToLastRow_(sizeThicknessValue) {
  if (!sizeThicknessValue) {
    return;
  }

  const sheet = getSheetByName_(SHEET_NAMES.LEDGER);
  const headers = getSheetHeaders_(sheet);
  const indexMap = buildHeaderIndexMap_(headers);
  const headerName = normalizeHeaderName_(COLUMNS.LEDGER.SIZE_THICKNESS);
  const sizeThicknessColumnIndex = indexMap[headerName];
  const lastRow = sheet.getLastRow();

  if (
    sizeThicknessColumnIndex === undefined ||
    sizeThicknessColumnIndex === null ||
    lastRow < 2
  ) {
    return;
  }

  sheet
    .getRange(lastRow, sizeThicknessColumnIndex + 1)
    .setValue(sizeThicknessValue);

  Logger.log(
    'forceWriteLedgerSizeThicknessToLastRow_ success: row=' +
      lastRow +
      ' column=' +
      (sizeThicknessColumnIndex + 1) +
      ' value=' +
      sizeThicknessValue
  );
}

function debugLedgerSheetHeaders() {
  const sheet = getSheetByName_(SHEET_NAMES.LEDGER);
  const headers = getSheetHeaders_(sheet);

  const diagnostics = headers.map(function (header, index) {
    const raw = String(header);
    return {
      index: index + 1,
      raw: raw,
      trimmed: raw.trim(),
      length: raw.length,
      codes: raw.split('').map(function (char) {
        return char.charCodeAt(0);
      }),
    };
  });

  Logger.log('debugLedgerSheetHeaders: ' + JSON.stringify(diagnostics));
}

function debugLastLedgerRow() {
  const sheet = getSheetByName_(SHEET_NAMES.LEDGER);
  const headers = getSheetHeaders_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    Logger.log('debugLastLedgerRow: no data rows');
    return;
  }

  const row = sheet.getRange(lastRow, 1, 1, headers.length).getValues()[0];
  const result = {};

  headers.forEach(function (header, index) {
    result[String(header).trim()] = row[index];
  });

  Logger.log(
    'debugLastLedgerRow row=' + lastRow + ' data=' + JSON.stringify(result)
  );
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

  const ledgerId = createLedgerEntry_({
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
  return ledgerId;
}

function findLatestInternalLogByLineUserId_(lineUserId) {
  if (!lineUserId) return null;

  const sheet = getSheetByName_(SHEET_NAMES.INTERNAL_LOG);
  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) return null;

  const headers = values[0];
  const rows = values.slice(1);
  const indexMap = buildHeaderIndexMap_(headers);

  for (let i = rows.length - 1; i >= 0; i -= 1) {
    const row = rows[i];
    const rowLineUserId = getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.INTERNAL_LOG.LINE_USER_ID
    );

    if (String(rowLineUserId || '').trim() !== String(lineUserId).trim()) {
      continue;
    }

    return {
      createdAt: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.CREATED_AT
      ),
      customerName: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.CUSTOMER_NAME
      ),
      contactName: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.CONTACT_NAME
      ),
      projectName: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.PROJECT_NAME
      ),
      drawingNumber: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.DRAWING_NUMBER
      ),
      material: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.MATERIAL
      ),
      size: getCellValueByHeader_(row, indexMap, COLUMNS.INTERNAL_LOG.SIZE),
      quantity: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.QUANTITY
      ),
      dueDate: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.DUE_DATE
      ),
      notes: getCellValueByHeader_(row, indexMap, COLUMNS.INTERNAL_LOG.NOTES),
      rawText: getCellValueByHeader_(
        row,
        indexMap,
        COLUMNS.INTERNAL_LOG.RAW_TEXT
      ),
    };
  }

  return null;
}

function findSimilarCaseFromSampleSheet_(baseProject) {
  const sheet = getSheetByName_(SHEET_NAMES.SIMILAR_CASES);
  const values = sheet.getDataRange().getValues();

  if (!values || values.length < 2) {
    return null;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const indexMap = buildHeaderIndexMap_(headers);

  const normalizedDrawingNumber = normalizeTextForMatch_(
    baseProject.drawingNumber
  );
  const normalizedCustomerName = normalizeTextForMatch_(
    baseProject.customerName
  );
  const projectKeywords = extractProjectKeywords_(baseProject.projectName);

  if (normalizedDrawingNumber) {
    for (let i = rows.length - 1; i >= 0; i -= 1) {
      const row = rows[i];
      const sampleDrawingNumber = normalizeTextForMatch_(
        getCellValueByHeader_(
          row,
          indexMap,
          COLUMNS.SIMILAR_CASES.DRAWING_NUMBER
        )
      );

      if (
        sampleDrawingNumber &&
        sampleDrawingNumber === normalizedDrawingNumber
      ) {
        return {
          reason: '図面番号一致',
          project: buildSimilarCaseProjectFromRow_(row, indexMap),
        };
      }
    }
  }

  if (normalizedCustomerName) {
    for (let i = rows.length - 1; i >= 0; i -= 1) {
      const row = rows[i];
      const sampleCustomerName = normalizeTextForMatch_(
        getCellValueByHeader_(
          row,
          indexMap,
          COLUMNS.SIMILAR_CASES.CUSTOMER_NAME
        )
      );

      if (
        sampleCustomerName &&
        (sampleCustomerName.indexOf(normalizedCustomerName) >= 0 ||
          normalizedCustomerName.indexOf(sampleCustomerName) >= 0)
      ) {
        return {
          reason: '顧客名で検索',
          project: buildSimilarCaseProjectFromRow_(row, indexMap),
        };
      }
    }
  }

  if (projectKeywords.length > 0) {
    for (let i = rows.length - 1; i >= 0; i -= 1) {
      const row = rows[i];
      const sampleProjectName = normalizeTextForMatch_(
        getCellValueByHeader_(row, indexMap, COLUMNS.SIMILAR_CASES.PROJECT_NAME)
      );

      if (!sampleProjectName) continue;

      for (let j = 0; j < projectKeywords.length; j += 1) {
        if (sampleProjectName.indexOf(projectKeywords[j]) >= 0) {
          return {
            reason: '案件名のキーワード一致',
            project: buildSimilarCaseProjectFromRow_(row, indexMap),
          };
        }
      }
    }
  }

  return null;
}

function buildSimilarCaseProjectFromRow_(row, indexMap) {
  return {
    projectId: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.PROJECT_ID
    ),
    receivedAt: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.RECEIVED_AT
    ),
    customerName: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.CUSTOMER_NAME
    ),
    contactName: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.CONTACT_NAME
    ),
    projectName: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.PROJECT_NAME
    ),
    drawingNumber: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.DRAWING_NUMBER
    ),
    processingType: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.PROCESS_TYPE
    ),
    material: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.MATERIAL
    ),
    size: getCellValueByHeader_(row, indexMap, COLUMNS.SIMILAR_CASES.SIZE),
    quantity: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.QUANTITY
    ),
    dueDate: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.DUE_DATE
    ),
    pastUnitPrice: getCellValueByHeader_(
      row,
      indexMap,
      COLUMNS.SIMILAR_CASES.PAST_UNIT_PRICE
    ),
    notes: getCellValueByHeader_(row, indexMap, COLUMNS.SIMILAR_CASES.NOTES),
  };
}

function buildHeaderIndexMap_(headers) {
  const map = {};

  for (let i = 0; i < headers.length; i += 1) {
    map[normalizeHeaderName_(headers[i])] = i;
  }

  return map;
}

function getCellValueByHeader_(row, indexMap, headerName) {
  const index = indexMap[normalizeHeaderName_(headerName)];

  if (index === undefined || index === null) {
    return '';
  }

  return row[index];
}

function normalizeHeaderName_(value) {
  return String(value || '')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/[・･\/／・\-－―ー_]/g, '')
    .replace(/[　\s]+/g, '')
    .trim();
}

function normalizeRowObjectKeys_(rowObject) {
  const normalized = {};

  Object.keys(rowObject || {}).forEach(function (key) {
    normalized[normalizeHeaderName_(key)] = rowObject[key];
  });

  return normalized;
}

function normalizeTextForMatch_(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[　\s]+/g, '');
}

function extractProjectKeywords_(projectName) {
  const normalized = String(projectName || '')
    .replace(/[　]/g, ' ')
    .trim();

  if (!normalized) return [];

  return normalized
    .split(/\s+/)
    .map(function (keyword) {
      return normalizeTextForMatch_(keyword);
    })
    .filter(function (keyword) {
      return keyword && keyword.length >= 2;
    });
}
