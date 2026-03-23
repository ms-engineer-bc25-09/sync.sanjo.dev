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
    map[String(headers[i]).trim()] = i;
  }

  return map;
}

function getCellValueByHeader_(row, indexMap, headerName) {
  const index = indexMap[String(headerName).trim()];

  if (index === undefined || index === null) {
    return '';
  }

  return row[index];
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
