function doPost(e) {
  try {
    const body =
      e && e.postData && e.postData.contents ? e.postData.contents : '{}';

    Logger.log('doPost body: ' + body);

    const data = JSON.parse(body);

    // LINEルート
    if (Array.isArray(data.events) && data.events.length > 0) {
      return handleLine_(data);
    }

    // Tallyルート
    if (typeof handleTally_ === 'function') {
      Logger.log('doPost: tally route');
      handleTally_(data);
    }

    // 検証や空POSTでも200を返す
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('doPost error: ' + error.message);
    Logger.log('doPost error stack: ' + error.stack);

    return ContentService.createTextOutput(
      JSON.stringify({
        ok: false,
        error: error.message,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleLine_(data) {
  Logger.log('handleLine_ called: ' + JSON.stringify(data));

  try {
    const events = data.events || [];

    for (const event of events) {
      try {
        if (event.type !== 'message') continue;

        const messageType = event.message?.type || '';

        if (messageType === 'text') {
          handleLineTextMessage_(event);
          continue;
        }

        if (messageType === 'image') {
          handleLineImageMessage_(event);
          continue;
        }

        Logger.log('skip unsupported message type: ' + messageType);
      } catch (eventError) {
        Logger.log('handleLine_ event error: ' + eventError.message);
        Logger.log('handleLine_ event error stack: ' + eventError.stack);

        const replyToken = event.replyToken || '';
        if (replyToken) {
          try {
            replyLineMessage_(replyToken, [
              {
                type: 'text',
                text:
                  '処理中にエラーが発生しました。\n' +
                  'お手数ですが、もう一度送ってください。',
              },
            ]);
          } catch (replyError) {
            Logger.log('error reply failed: ' + replyError.message);
          }
        }
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('handleLine_ error: ' + error.message);
    Logger.log('handleLine_ error stack: ' + error.stack);

    return ContentService.createTextOutput(
      JSON.stringify({
        ok: false,
        error: error.message,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleLineTextMessage_(event) {
  const lineUserId = event.source?.userId || '';
  const replyToken = event.replyToken || '';
  const text = event.message?.text || '';

  Logger.log('handleLineTextMessage_ start');
  Logger.log('lineUserId: ' + lineUserId);
  Logger.log('text: ' + text);

  if (!text) {
    Logger.log('handleLineTextMessage_: empty text');
    return;
  }

  const geminiResult = callGemini_(text);

  Logger.log('text gemini result: ' + JSON.stringify(geminiResult));

  saveLineGeminiResultToSheets_({
    geminiResult: geminiResult,
    source: 'LINE',
    status: '未対応',
    rawText: text,
    lineUserId: lineUserId,
    drawingUrl: '',
  });

  if (replyToken) {
    const replyText = buildLineTextReplyMessage_(
      geminiResult,
      '内容を受け取りました。'
    );

    replyLineMessage_(replyToken, [
      {
        type: 'text',
        text: replyText,
      },
    ]);
  }
}

function handleLineImageMessage_(event) {
  const lineUserId = event.source?.userId || '';
  const replyToken = event.replyToken || '';
  const messageId = event.message?.id || '';

  Logger.log('handleLineImageMessage_ start');
  Logger.log('lineUserId: ' + lineUserId);
  Logger.log('messageId: ' + messageId);

  if (!messageId) {
    throw new Error('画像メッセージの messageId が取得できません');
  }

  const imageData = getLineImageContent_(messageId);
  const geminiResult = callGeminiImage_(imageData.base64, imageData.mimeType);
  let drawingUrl = '';

  Logger.log('image gemini result: ' + JSON.stringify(geminiResult));

  try {
    drawingUrl = uploadBlobToSupabase_(imageData.blob, {
      source: 'line',
      receivedAt: new Date(),
      fileName: imageData.blob ? imageData.blob.getName() : '',
    });
    Logger.log('handleLineImageMessage_ uploaded file to supabase: ' + drawingUrl);
  } catch (error) {
    Logger.log('handleLineImageMessage_ file upload error: ' + error.message);
    Logger.log('handleLineImageMessage_ file upload error stack: ' + error.stack);
  }

  const rawText =
    'LINE画像メッセージを受信\n' +
    'messageId: ' +
    messageId +
    '\n' +
    'contentType: ' +
    imageData.mimeType +
    '\n' +
    'fileName: ' +
    (imageData.blob ? imageData.blob.getName() : 'line-image-' + messageId) +
    '\n' +
    'size(bytes): ' +
    (imageData.blob ? imageData.blob.getBytes().length : 0);

  const ledgerRow = buildLedgerRowFromGeminiResult_({
    geminiResult: geminiResult,
    rawText: rawText,
    source: 'LINE',
    status: '未対応',
    drawingUrl: drawingUrl,
  });
  const internalLogRow = buildInternalLogRowFromGeminiResult_({
    geminiResult: geminiResult,
    ledgerRow: ledgerRow,
    rawText: rawText,
    lineUserId: lineUserId,
  });

  appendInternalLogRow_(internalLogRow);
  appendLedgerRow_(ledgerRow);

  saveProjectToSupabase_({
    geminiResult: geminiResult,
    ledgerRow: ledgerRow,
    lineUserId: lineUserId,
  });

  if (replyToken) {
    const customerName = geminiResult.customer_name || '未入力';
    const projectName = geminiResult.project_name || '未入力';
    const material = geminiResult.material || '未入力';
    const quantity = geminiResult.quantity || '未入力';
    const dueDate = geminiResult.desired_due_date || '未入力';

    const replyText =
      '画像を受け取りました。\n' +
      '以下の内容で記録しました。\n\n' +
      '・顧客名：' +
      customerName +
      '\n' +
      '・案件名：' +
      projectName +
      '\n' +
      '・材質：' +
      material +
      '\n' +
      '・数量：' +
      quantity +
      '\n' +
      '・希望納期：' +
      dueDate +
      '\n\n' +
      '不足があれば補足をテキストで送ってください。\n\n' +
      '修正は案件台帳から直接できます。\n' +
      SPREADSHEET_URL;

    replyLineMessage_(replyToken, [
      {
        type: 'text',
        text: replyText,
      },
    ]);
  }
}

function testGemini() {
  const result = callGemini_('テストです。SUS304で10個、来週まで');
  Logger.log('testGemini result: ' + JSON.stringify(result));
}

function buildLedgerRowFromGeminiResult_(options) {
  const result = options.geminiResult || {};
  const rawText = (options.rawText || '').trim();
  const projectName = (result.project_name || '').trim();
  const drawingNumber = (result.drawing_number || '').trim();
  const inquiry = projectName || drawingNumber || rawText || '画像受信案件';

  let rawJson = '';

  try {
    rawJson = JSON.stringify(result || {});
  } catch (error) {
    rawJson = '';
  }

  return {
    receivedAt: new Date(),
    status: options.status || '',
    source: options.source || '',
    customerName: result.customer_name || '',
    contactName: result.contact_name || '',
    email: '',
    phone: '',
    inquiry: inquiry,
    dueDate: result.desired_due_date || '',
    material: result.material || '',
    sizeThickness: result.size_thickness || '',
    quantity: result.quantity || '',
    notes: result.notes || '',
    drawingUrl: options.drawingUrl || '',
    rawJson: rawJson,
  };
}

function buildInternalLogRowFromGeminiResult_(options) {
  const result = options.geminiResult || {};
  const ledgerRow = options.ledgerRow || {};
  const now = ledgerRow.receivedAt || new Date();

  return {
    id: '',
    createdAt: now,
    source: ledgerRow.source || '',
    status: ledgerRow.status || '',
    customerName: ledgerRow.customerName || '',
    contactName: ledgerRow.contactName || '',
    projectName: result.project_name || '',
    drawingNumber: result.drawing_number || '',
    material: ledgerRow.material || '',
    size: ledgerRow.sizeThickness || '',
    quantity: ledgerRow.quantity || '',
    dueDate: ledgerRow.dueDate || '',
    notes: ledgerRow.notes || '',
    rawText: options.rawText || '',
    lineUserId: options.lineUserId || '',
    aiExtractedJson: ledgerRow.rawJson || '',
    similarCase: '',
    pastUnitPrice: '',
    suggestedPrice: '',
    spreadsheetUpdatedAt: now,
  };
}

function saveProjectToSupabase_(options) {
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    Logger.log('saveProjectToSupabase_ skipped: missing Supabase config');
    return;
  }

  const url = SUPABASE_URL.replace(/\/$/, '') + '/rest/v1/projects';
  const payload = buildSupabaseProjectPayload_(options);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: 'Bearer ' + SUPABASE_KEY,
        Prefer: 'return=minimal',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    const body = response.getContentText();

    Logger.log('saveProjectToSupabase_ status: ' + statusCode);

    if (statusCode < 200 || statusCode >= 300) {
      Logger.log('saveProjectToSupabase_ response: ' + body);
    }
  } catch (error) {
    Logger.log('saveProjectToSupabase_ error: ' + error.message);
    Logger.log('saveProjectToSupabase_ error stack: ' + error.stack);
  }
}

function uploadTallyFileToSupabase_(fileUrl, options) {
  if (!fileUrl) {
    throw new Error('fileUrl が指定されていません');
  }

  if (!SUPABASE_URL || !SUPABASE_KEY) {
    throw new Error('Supabase config が設定されていません');
  }

  const fileResponse = UrlFetchApp.fetch(fileUrl, {
    method: 'get',
    muteHttpExceptions: true,
  });
  const fileStatusCode = fileResponse.getResponseCode();

  Logger.log('uploadTallyFileToSupabase_ file fetch status: ' + fileStatusCode);

  if (fileStatusCode < 200 || fileStatusCode >= 300) {
    throw new Error(
      'Tally添付ファイル取得エラー: ' + fileResponse.getContentText()
    );
  }

  const blob = fileResponse.getBlob();
  const mergedOptions = Object.assign({}, options, {
    fileName: fileUrl,
  });

  return uploadBlobToSupabase_(blob, mergedOptions);
}

function uploadBlobToSupabase_(blob, options) {
  if (!blob) {
    throw new Error('blob が指定されていません');
  }

  if (!SUPABASE_URL || !SUPABASE_KEY) {
    throw new Error('Supabase config が設定されていません');
  }

  const contentType = blob.getContentType() || 'application/octet-stream';
  const objectPath = buildSupabaseStoragePath_(options, contentType);
  const uploadUrl =
    SUPABASE_URL.replace(/\/$/, '') +
    '/storage/v1/object/' +
    SUPABASE_STORAGE_BUCKET +
    '/' +
    objectPath;

  const response = UrlFetchApp.fetch(uploadUrl, {
    method: 'post',
    contentType: contentType,
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: 'Bearer ' + SUPABASE_KEY,
      'x-upsert': 'true',
    },
    payload: blob.getBytes(),
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  Logger.log('uploadTallyFileToSupabase_ upload status: ' + statusCode);
  Logger.log('uploadTallyFileToSupabase_ upload response: ' + body);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Supabase Storageアップロードエラー: ' + body);
  }

  return buildSupabasePublicFileUrl_(SUPABASE_STORAGE_BUCKET, objectPath);
}

function buildSupabaseStoragePath_(options, contentType) {
  const receivedAt = options?.receivedAt || new Date();
  const datePath = Utilities.formatDate(receivedAt, 'Asia/Tokyo', 'yyyy/MM/dd');
  const source = (options?.source || 'misc').toLowerCase();
  const extension = inferFileExtension_(options?.fileName || '', contentType);

  return [
    source,
    datePath,
    Utilities.getUuid() + extension,
  ].join('/');
}

function inferFileExtension_(fileUrl, contentType) {
  const urlMatch = String(fileUrl || '').match(/\.([a-zA-Z0-9]+)(?:\?|$)/);

  if (urlMatch) {
    return '.' + urlMatch[1].toLowerCase();
  }

  const mimeType = String(contentType || '').toLowerCase();

  if (mimeType === 'image/jpeg') return '.jpg';
  if (mimeType === 'image/png') return '.png';
  if (mimeType === 'application/pdf') return '.pdf';
  if (mimeType === 'image/webp') return '.webp';

  return '.bin';
}

function buildSupabasePublicFileUrl_(bucketName, objectPath) {
  const encodedPath = String(objectPath || '')
    .split('/')
    .map(function (segment) {
      return encodeURIComponent(segment);
    })
    .join('/');

  return (
    SUPABASE_URL.replace(/\/$/, '') +
    '/storage/v1/object/public/' +
    encodeURIComponent(bucketName) +
    '/' +
    encodedPath
  );
}

function buildSupabaseProjectPayload_(options) {
  const result = options.geminiResult || {};
  const ledgerRow = options.ledgerRow || {};

  return {
    source_type: normalizeSupabaseSourceType_(ledgerRow.source),
    received_at: toIsoStringOrNull_(ledgerRow.receivedAt),
    status: normalizeSupabaseStatus_(ledgerRow.status),
    customer_name: ledgerRow.customerName || '',
    contact_name: ledgerRow.contactName || '',
    project_name: result.project_name || '',
    drawing_number: result.drawing_number || '',
    material: ledgerRow.material || '',
    size_text: ledgerRow.sizeThickness || '',
    quantity: parseSupabaseQuantity_(ledgerRow.quantity),
    due_date: parseSupabaseDueDate_(ledgerRow.dueDate),
    current_response_note: '',
    note: ledgerRow.notes || '',
    ai_summary: '',
    line_user_id: options.lineUserId || '',
    spreadsheet_row_no: null,
  };
}

function normalizeSupabaseSourceType_(source) {
  const value = (source || '').toLowerCase();

  if (value === 'line') return 'line';
  if (value === 'tally') return 'tally';
  if (value === 'phone') return 'phone';
  if (value === 'fax') return 'fax';
  if (value === 'email') return 'email';
  if (value === 'talk') return 'talk';

  return 'line';
}

function normalizeSupabaseStatus_(status) {
  if (status === '未対応') return 'received';
  return 'received';
}

function toIsoStringOrNull_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    return value.toISOString();
  }

  return null;
}

function parseSupabaseQuantity_(value) {
  const text = (value || '').toString();
  const match = text.match(/\d+/);

  if (!match) return null;

  const quantity = parseInt(match[0], 10);
  return isNaN(quantity) ? null : quantity;
}

function parseSupabaseDueDate_(value) {
  const text = (value || '').trim();

  if (!text) return null;

  const normalized = text.replace(/\//g, '-');

  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    return null;
  }

  return normalized;
}

function testSaveLineGeminiResultToSheets() {
  const sample = {
    customer_name: 'テスト工業',
    contact_name: '',
    project_name: 'ブラケット試作',
    drawing_number: '',
    processing: '',
    material: 'SUS304',
    size_thickness: '',
    quantity: '10個',
    desired_due_date: '来週まで',
    notes: '急ぎ',
  };

  saveLineGeminiResultToSheets_({
    geminiResult: sample,
    source: 'LINE',
    status: '未対応',
    rawText:
      'テスト工業です。ブラケット試作をお願いします。SUS304で10個、来週まで。急ぎです。',
    lineUserId: 'TEST_LINE_USER_ID',
    drawingUrl: '',
  });

  Logger.log('内部ログ・案件台帳への書き込み完了');
}

function testHandleLine() {
  const mockData = {
    events: [
      {
        type: 'message',
        message: {
          type: 'text',
          text: 'テスト工業です。ブラケット試作をお願いします。SUS304で10個、来週まで。急ぎです。',
        },
        source: {
          type: 'user',
          userId: 'TEST_LINE_USER_ID',
        },
        replyToken: 'TEST_REPLY_TOKEN',
      },
    ],
  };

  const result = handleLine_(mockData);
  Logger.log('testHandleLine result: ' + result.getContent());
}

function testHandleLineImage() {
  const mockData = {
    events: [
      {
        type: 'message',
        message: {
          type: 'image',
          id: 'TEST_IMAGE_MESSAGE_ID',
        },
        source: {
          type: 'user',
          userId: 'TEST_LINE_USER_ID',
        },
        replyToken: 'TEST_REPLY_TOKEN',
      },
    ],
  };

  const result = handleLine_(mockData);
  Logger.log('testHandleLineImage result: ' + result.getContent());
}

function buildLineTextReplyMessage_(result, prefix) {
  const lines = [];

  // ① 冒頭文（テキスト or 画像で切り替え）
  lines.push(prefix || '内容を受け取りました。');
  lines.push('以下の内容で記録しました。');
  lines.push('');

  // ② 項目表示
  if (hasReplyValue_(result.customer_name)) {
    lines.push('・顧客名：' + result.customer_name);
  }

  if (hasReplyValue_(result.project_name)) {
    lines.push('・案件名：' + result.project_name);
  }

  if (hasReplyValue_(result.material)) {
    lines.push('・材質：' + result.material);
  }

  if (hasReplyValue_(result.quantity)) {
    lines.push('・数量：' + result.quantity);
  }

  if (hasReplyValue_(result.desired_due_date)) {
    lines.push('・希望納期：' + result.desired_due_date);
  }

  // 必要なら追加表示（任意）
  if (hasReplyValue_(result.size_thickness)) {
    lines.push('・サイズ・板厚：' + result.size_thickness);
  }

  if (hasReplyValue_(result.drawing_number)) {
    lines.push('・図面番号：' + result.drawing_number);
  }

  // ③ フッター
  lines.push('');
  lines.push('不足があれば補足をテキストで送ってください。');
  lines.push('');
  lines.push('修正は案件台帳から直接できます。');
  lines.push(SPREADSHEET_URL);

  return lines.join('\n');
}

function hasReplyValue_(value) {
  if (value === null || value === undefined) return false;

  const normalized = String(value).trim();

  if (!normalized) return false;
  if (normalized === '不明') return false;
  if (normalized === '（案件名未設定）') return false;

  return true;
}
