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

    // Supabase画像Webhookルート
    if (isSupabaseImageWebhookPayload_(data)) {
      Logger.log('doPost: supabase image route');
      handleSupabaseImageWebhook_(data);
      return ContentService.createTextOutput(
        JSON.stringify({ ok: true })
      ).setMimeType(ContentService.MimeType.JSON);
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
        if (event.type === 'postback') {
          handleLinePostback_(event);
          continue;
        }

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

function handleLinePostback_(event) {
  const lineUserId = event.source?.userId || '';
  const replyToken = event.replyToken || '';
  const data = (event.postback?.data || '').trim();

  Logger.log('handleLinePostback_ start');
  Logger.log('lineUserId: ' + lineUserId);
  Logger.log('postback data: ' + data);

  if (!data) {
    Logger.log('handleLinePostback_: data is empty');
    return;
  }

  if (data === 'action=past_case_search') {
    if (replyToken) {
      replyLineMessage_(replyToken, handlePastCaseSearchEntry_(lineUserId));
    }
    return;
  }

  Logger.log('skip unsupported postback data: ' + data);
}

function handleLineTextMessage_(event) {
  const lineUserId = event.source?.userId || '';
  const replyToken = event.replyToken || '';
  const text = (event.message?.text || '').trim();

  Logger.log('handleLineTextMessage_ start');
  Logger.log('lineUserId: ' + lineUserId);
  Logger.log('text: ' + text);

  if (!text) {
    Logger.log('handleLineTextMessage_: empty text');
    return;
  }

  if (isPastCaseSearchEntryText_(text)) {
    if (replyToken) {
      replyLineMessage_(replyToken, handlePastCaseSearchEntry_(lineUserId));
    }
    return;
  }

  const lineUserState = getLineUserState_(lineUserId);
  if (
    lineUserState &&
    (lineUserState.mode === 'search_waiting_type' ||
      lineUserState.mode === 'search_waiting_value')
  ) {
    if (replyToken) {
      replyLineMessage_(
        replyToken,
        handlePastCaseSearchInput_(lineUserId, text, lineUserState)
      );
    }
    return;
  }

  const richMenuGuideReply = buildRichMenuGuideReply_(text);
  if (richMenuGuideReply) {
    if (replyToken) {
      replyLineMessage_(replyToken, richMenuGuideReply);
    }
    return;
  }

  if (text === '類似') {
    if (replyToken) {
      const similarReplyText =
        buildSimilarCaseReplyFromLatestProject_(lineUserId);
      replyLineMessage_(replyToken, [
        {
          type: 'text',
          text: similarReplyText,
        },
      ]);
    }
    return;
  }

  const geminiResult = callGemini_(text);

  Logger.log('text gemini result: ' + JSON.stringify(geminiResult));

  if (!geminiResult) {
    if (replyToken) {
      replyLineMessage_(
        replyToken,
        '受付内容の読み取りに失敗しました。内容を変えてもう一度送ってください。'
      );
    }
    return;
  }

  let parsedResult;

  try {
    parsedResult =
      typeof geminiResult === 'string'
        ? JSON.parse(geminiResult)
        : geminiResult;
  } catch (parseError) {
    Logger.log('Gemini JSON parse error: ' + parseError.message);
    Logger.log('Gemini raw result: ' + geminiResult);

    if (replyToken) {
      replyLineMessage_(
        replyToken,
        '内容をうまく整理できませんでした。\n\n顧客名・案件名・数量などを\n少し補足して送っていただけますか？'
      );
    }
    return;
  }

  saveLineGeminiResultToSheets_({
    geminiResult: parsedResult,
    source: 'LINE',
    status: '未対応',
    rawText: text,
    lineUserId: lineUserId,
    drawingUrl: '',
  });

  if (replyToken) {
    const replyText = buildLineRegistrationReplyMessage_(parsedResult);

    replyLineMessage_(replyToken, [
      {
        type: 'text',
        text: replyText,
      },
    ]);
  }
}

function buildRichMenuGuideReply_(text) {
  const guides = {
    図面FAXを登録してください:
      '図面画像を送ってください。\n受信後、内容を読み取って案件登録します。',
    図面FAXを登録します:
      '図面画像を送ってください。\n受信後、内容を読み取って案件登録します。',
    書類を撮る:
      '書類の画像を送ってください。\n受信後、内容を読み取って案件登録します。',
    テキストで案件を登録してください:
      '案件内容をそのままテキストで送ってください。\n顧客名案件名材質数量希望納期が入っていると整理しやすいです。',
    テキストで登録します:
      '案件内容をそのままテキストで送ってください。\n顧客名案件名材質数量希望納期が入っていると整理しやすいです。',
    メモを残す:
      '案件内容をそのままテキストで送ってください。\n顧客名案件名材質数量希望納期が入っていると整理しやすいです。',
    音声で案件を登録してください:
      '音声登録は、現在準備中です。',
    音声で登録します:
      '音声登録は、現在準備中です。',
    声で残す:
      '音声登録は、現在準備中です。',
  };

  return guides[text] || null;
}

function buildLineRegistrationReplyMessage_(project) {
  const customerName = project.customer_name || '登録なし';
  const projectName = project.project_name || '登録なし';
  const material = project.material || '登録なし';
  const quantity = project.quantity || '登録なし';
  const dueDate = formatLineDate_(
    project.desired_due_date || project.due_date || ''
  );

  return [
    '案件を受付しました。',
    '',
    '・顧客名：' + customerName,
    '・案件名：' + projectName,
    '・材質：' + material,
    '・数量：' + quantity,
    '・希望納期：' + (dueDate === '-' ? '登録なし' : dueDate),
    '',
    '内容の変更があれば、',
    '項目名と一緒に送ってください。',
    '',
    '「類似」と送ると、過去の類似案件を検索できます。',
    '',
    '案件台帳はこちら：',
    SPREADSHEET_URL,
  ].join('\n');
}

function buildSimilarCaseReplyFromLatestProject_(lineUserId) {
  if (!lineUserId) {
    return [
      '類似する案件は見つかりませんでした。',
      '',
      '先に案件内容を登録してから「類似」と送ってください。',
    ].join('\n');
  }

  try {
    const latestProject = findLatestInternalLogByLineUserId_(lineUserId);

    if (!latestProject) {
      return [
        '類似する案件は見つかりませんでした。',
        '',
        '先に案件内容を登録してから「類似」と送ってください。',
      ].join('\n');
    }

    const similarCaseResult = findSimilarCaseFromSampleSheet_(latestProject);

    if (!similarCaseResult || !similarCaseResult.project) {
      return buildNoSimilarCaseReply_();
    }

    return buildSimilarCaseReply_(
      similarCaseResult.project,
      similarCaseResult.reason
    );
  } catch (error) {
    Logger.log(
      'buildSimilarCaseReplyFromLatestProject_ error: ' + error.message
    );
    Logger.log(
      'buildSimilarCaseReplyFromLatestProject_ error stack: ' + error.stack
    );

    return [
      '類似案件検索でエラーが発生しました。',
      '',
      '時間をおいてもう一度「類似」と送ってください。',
    ].join('\n');
  }
}

function buildSimilarCaseReply_(project, reason) {
  return [
    '類似案件が見つかりました。（' + reason + '）',
    '',
    '顧客名：' + (project.customerName || '登録なし'),
    '案件名：' + (project.projectName || '登録なし'),
    '図面番号：' + (project.drawingNumber || '-'),
    '前回単価：' + (project.pastUnitPrice || '登録なし'),
    '受付日：' + formatLineDate_(project.receivedAt),
    '',
    '見積の参考にしてください。',
    '台帳はこちら：',
    SPREADSHEET_URL,
  ].join('\n');
}

function buildNoSimilarCaseReply_() {
  return [
    '類似案件は見つかりませんでした。',
    '',
    '今回の検索条件：',
    '・図面番号一致',
    '・顧客名で検索',
    '・案件名のキーワード一致',
    '',
    '台帳はこちら：',
    SPREADSHEET_URL,
  ].join('\n');
}

function isPastCaseSearchEntryText_(text) {
  return (
    text === '過去案件検索' || text === '過去案件検索ができる予定です'
  );
}

function handlePastCaseSearchEntry_(lineUserId) {
  if (!lineUserId) {
    return '過去案件検索を開始できませんでした。しばらくしてからもう一度送ってください。';
  }

  setLineUserState_(lineUserId, 'search_waiting_type', {
    flow: 'past_case_search',
  });

  return [
    {
      type: 'text',
      text: '過去案件検索を受け付けました。\n\nどの条件で探しますか？',
      quickReply: {
        items: [
          buildSearchTypeQuickReplyItem_('図面番号'),
          buildSearchTypeQuickReplyItem_('顧客名'),
          buildSearchTypeQuickReplyItem_('案件名'),
        ],
      },
    },
  ];
}

function handlePastCaseSearchInput_(lineUserId, text, lineUserState) {
  try {
    const mode = lineUserState?.mode || '';
    const payload = lineUserState?.payload || {};

    if (mode === 'search_waiting_type') {
      const searchType = getPastCaseSearchType_(text);

      if (!searchType) {
        return [
          {
            type: 'text',
            text: '図面番号、顧客名、案件名のいずれかを選んでください。',
            quickReply: {
              items: [
                buildSearchTypeQuickReplyItem_('図面番号'),
                buildSearchTypeQuickReplyItem_('顧客名'),
                buildSearchTypeQuickReplyItem_('案件名'),
              ],
            },
          },
        ];
      }

      setLineUserState_(lineUserId, 'search_waiting_value', {
        flow: 'past_case_search',
        searchType: searchType,
      });

      return getPastCaseSearchPromptByType_(searchType);
    }

    if (mode !== 'search_waiting_value') {
      clearLineUserState_(lineUserId);
      return '検索状態を確認できませんでした。もう一度お試しください。';
    }

    const similarCaseResult = findSimilarCaseFromSampleSheet_(
      buildPastCaseSearchQuery_(payload.searchType, text)
    );

    clearLineUserState_(lineUserId);

    if (!similarCaseResult || !similarCaseResult.project) {
      return buildNoSimilarCaseReply_();
    }

    return buildSimilarCaseReply_(
      similarCaseResult.project,
      similarCaseResult.reason
    );
  } catch (error) {
    Logger.log('handlePastCaseSearchInput_ error: ' + error.message);
    Logger.log('handlePastCaseSearchInput_ error stack: ' + error.stack);
    clearLineUserState_(lineUserId);

    return [
      '過去案件検索でエラーが発生しました。',
      '',
      '時間をおいてもう一度送ってください。',
    ].join('\n');
  }
}

function buildSearchTypeQuickReplyItem_(label) {
  return {
    type: 'action',
    action: {
      type: 'message',
      label: label,
      text: label,
    },
  };
}

function getPastCaseSearchType_(text) {
  const normalized = String(text || '').trim();

  if (normalized === '図面番号') return 'drawing_number';
  if (normalized === '顧客名') return 'customer_name';
  if (normalized === '案件名') return 'project_name';

  return '';
}

function getPastCaseSearchPromptByType_(searchType) {
  if (searchType === 'drawing_number') {
    return '図面番号を送ってください。';
  }

  if (searchType === 'customer_name') {
    return '顧客名を送ってください。';
  }

  return '案件名を送ってください。';
}

function buildPastCaseSearchQuery_(searchType, text) {
  const query = String(text || '').trim();

  if (searchType === 'drawing_number') {
    return {
      drawingNumber: query,
      customerName: '',
      projectName: '',
    };
  }

  if (searchType === 'customer_name') {
    return {
      drawingNumber: '',
      customerName: query,
      projectName: '',
    };
  }

  return {
    drawingNumber: '',
    customerName: '',
    projectName: query,
  };
}

function formatLineDate_(value) {
  if (!value) {
    return '-';
  }

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return '-';
    return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy/MM/dd');
  }

  const text = String(value).trim();

  if (!text) {
    return '-';
  }

  if (/^\d{4}\/\d{2}\/\d{2}$/.test(text)) {
    return text;
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text.replace(/-/g, '/');
  }

  const parsedDate = new Date(text);
  if (isNaN(parsedDate.getTime())) {
    return text;
  }

  return Utilities.formatDate(parsedDate, 'Asia/Tokyo', 'yyyy/MM/dd');
}

function handleLineImageMessage_(event) {
  const lineUserId = event.source?.userId || '';
  const replyToken = event.replyToken || '';
  const messageId = event.message?.id || '';
  const now = new Date();

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
      receivedAt: now,
      fileName: imageData.blob ? imageData.blob.getName() : '',
    });
    Logger.log(
      'handleLineImageMessage_ uploaded file to supabase: ' + drawingUrl
    );
  } catch (error) {
    Logger.log('handleLineImageMessage_ file upload error: ' + error.message);
    Logger.log(
      'handleLineImageMessage_ file upload error stack: ' + error.stack
    );
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

  const imageLedgerRow = buildImageLedgerRowFromGeminiResult_({
    geminiResult: geminiResult,
    receivedAt: now,
    rawText: rawText,
    source: 'LINE',
    projectType: '受付メモ',
    status: '未処理',
    originalFileName: imageData.blob
      ? imageData.blob.getName()
      : 'line-image-' + messageId,
    originalImageUrl: '',
    drawingUrl: drawingUrl,
  });
  const imageInternalLogRow = buildImageInternalLogRowFromGeminiResult_({
    geminiResult: geminiResult,
    imageLedgerRow: imageLedgerRow,
    createdAt: now,
    savedImageUrl: drawingUrl,
    processingStatus: 'AI抽出済',
    validationResult: '未確認',
    errorMessage: '',
    ocrText: '',
    notes: rawText,
  });

  upsertImageInternalLogRow_(imageInternalLogRow);
  upsertImageLedgerRow_(imageLedgerRow);

  const savedProject = saveProjectToSupabase_({
    geminiResult: geminiResult,
    ledgerRow: imageLedgerRow,
    flowType: 'image',
    projectType: imageLedgerRow.projectType,
    originalFileName: imageLedgerRow.originalFileName,
    originalImageUrl: imageLedgerRow.originalImageUrl,
    savedImageUrl: drawingUrl,
    drawingUrl: imageLedgerRow.drawingUrl,
    ocrText: imageInternalLogRow.ocrText,
    aiExtractedJson: imageInternalLogRow.aiExtractedJson,
    validationResult: imageInternalLogRow.validationResult,
    processingStatus: imageInternalLogRow.processingStatus,
    errorMessage: imageInternalLogRow.errorMessage,
    lineUserId: lineUserId,
  });

  if (savedProject && savedProject.id) {
    saveProjectItemsToSupabase_(savedProject.id, geminiResult, {
      parentDueDate: imageLedgerRow.dueDate,
    });
  }

  if (replyToken) {
    const customerName = geminiResult.customer_name || '登録なし';
    const projectName = geminiResult.project_name || '登録なし';
    const material = geminiResult.material || '登録なし';
    const quantity = geminiResult.quantity || '登録なし';
    const dueDate = formatLineDate_(geminiResult.desired_due_date || '');

    const replyText =
      '画像を受け付けました。\n' +
      '以下の内容で登録しました。\n\n' +
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
      (dueDate === '-' ? '登録なし' : dueDate) +
      '\n\n' +
      '内容の変更があれば、\n' +
      '項目名と一緒に送ってください。\n' +
      'その内容で登録情報を上書きします。\n\n' +
      '過去の類似案件も確認できます。\n' +
      '「類似」と送ると検索します。\n\n' +
      '案件台帳はこちら：\n' +
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
    id: options.id || '',
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

function buildImageLedgerRowFromGeminiResult_(options) {
  const result = options.geminiResult || {};
  const primaryItem = getPrimaryGeminiItem_(result);
  const rawText = (options.rawText || '').trim();
  const projectName = (result.project_name || '').trim();
  const itemName = (primaryItem.item_name || '').trim();
  const drawingNumber = (result.drawing_number || '').trim();
  const fallbackProjectName =
    projectName || itemName || drawingNumber || rawText || '画像受信案件';

  return {
    id: options.id || '',
    receivedAt: options.receivedAt || new Date(),
    source: options.source || '',
    projectType: options.projectType || '',
    customerName: result.customer_name || primaryItem.customer_name || '',
    contactName: result.contact_name || primaryItem.contact_name || '',
    email: options.email || '',
    phone: options.phone || '',
    projectName: fallbackProjectName,
    dueDate: result.desired_due_date || primaryItem.due_date || '',
    material: result.material || primaryItem.material || '',
    sizeThickness:
      result.size_thickness || primaryItem.size_thickness || '',
    quantity: result.quantity || primaryItem.quantity || '',
    notes: result.notes || primaryItem.note || '',
    originalFileName: options.originalFileName || '',
    originalImageUrl: options.originalImageUrl || '',
    drawingUrl: options.drawingUrl || '',
    status: options.status || '',
    ledgerId: options.ledgerId || '',
    rawJson: safeStringify_(result),
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

function buildImageInternalLogRowFromGeminiResult_(options) {
  const result = options.geminiResult || {};
  const imageLedgerRow = options.imageLedgerRow || {};
  const createdAt = options.createdAt || imageLedgerRow.receivedAt || new Date();

  return {
    id: options.id || '',
    createdAt: createdAt,
    updatedAt: options.updatedAt || createdAt,
    source: imageLedgerRow.source || '',
    projectType: imageLedgerRow.projectType || '',
    status: imageLedgerRow.status || '',
    customerName: imageLedgerRow.customerName || '',
    contactName: imageLedgerRow.contactName || '',
    projectName: result.project_name || imageLedgerRow.projectName || '',
    originalFileName: imageLedgerRow.originalFileName || '',
    originalImageUrl: imageLedgerRow.originalImageUrl || '',
    savedImageUrl: options.savedImageUrl || imageLedgerRow.drawingUrl || '',
    ocrText: options.ocrText || '',
    aiExtractedJson:
      options.aiExtractedJson || imageLedgerRow.rawJson || safeStringify_(result),
    validationResult: options.validationResult || '',
    processingStatus: options.processingStatus || '',
    errorMessage: options.errorMessage || '',
    ledgerId: imageLedgerRow.ledgerId || '',
    notes: options.notes || imageLedgerRow.notes || '',
  };
}

function getPrimaryGeminiItem_(result) {
  const items = Array.isArray(result?.items) ? result.items : [];

  if (!items.length) {
    return {};
  }

  return items[0] || {};
}

function safeStringify_(value) {
  try {
    return JSON.stringify(value || {});
  } catch (error) {
    Logger.log('safeStringify_ error: ' + error.message);
    return '';
  }
}

function handleSupabaseImageWebhook_(data) {
  const record = getSupabaseWebhookProjectRecord_(data);

  if (!record) {
    throw new Error('Supabase webhook payload に record がありません');
  }

  const eventType = getSupabaseWebhookEventType_(data);
  if (eventType === 'DELETE') {
    Logger.log('handleSupabaseImageWebhook_ skipped: delete event');
    return;
  }

  const imageUrl = extractSupabaseImageUrl_(record);
  if (!imageUrl) {
    throw new Error('Supabase webhook payload に画像URLがありません');
  }

  const imageData = fetchImageFromUrl_(imageUrl);
  const geminiResult = callGeminiImage_(imageData.base64, imageData.mimeType);
  const receivedAt = parseDateValue_(record.received_at) || new Date();
  const rawText = buildSupabaseImageWebhookRawText_(record, imageData, imageUrl);
  const imageLedgerRow = buildImageLedgerRowFromGeminiResult_({
    id: record.id || '',
    geminiResult: geminiResult,
    receivedAt: receivedAt,
    rawText: rawText,
    source: normalizeSupabaseWebhookSource_(record.source_type || record.source),
    projectType: record.project_type || '受付メモ',
    status: '未処理',
    email: record.email || '',
    phone: record.phone || '',
    originalFileName: record.original_file_name || imageData.fileName || '',
    originalImageUrl: record.original_image_url || '',
    drawingUrl: record.saved_image_url || record.drawing_url || imageUrl,
    ledgerId: record.ledger_id || '',
  });
  const imageInternalLogRow = buildImageInternalLogRowFromGeminiResult_({
    geminiResult: geminiResult,
    imageLedgerRow: imageLedgerRow,
    createdAt: receivedAt,
    updatedAt: parseDateValue_(record.updated_at) || receivedAt,
    savedImageUrl: imageLedgerRow.drawingUrl,
    processingStatus: record.processing_status || 'AI抽出済',
    validationResult: record.validation_result || '未確認',
    errorMessage: record.error_message || '',
    ocrText: record.ocr_text || '',
    aiExtractedJson:
      safeStringify_(geminiResult) || safeStringify_(record.ai_extracted_json),
    notes: rawText,
    id: record.id || '',
  });

  upsertImageInternalLogRow_(imageInternalLogRow);
  upsertImageLedgerRow_(imageLedgerRow);

  Logger.log(
    'handleSupabaseImageWebhook_ success: projectId=' +
      String(record.id || '') +
      ' fileName=' +
      String(imageLedgerRow.originalFileName || '')
  );
}

function isSupabaseImageWebhookPayload_(data) {
  const record = getSupabaseWebhookProjectRecord_(data);

  if (!record) {
    return false;
  }

  const eventType = getSupabaseWebhookEventType_(data);
  if (eventType === 'DELETE') {
    return false;
  }

  const flowType = String(record.flow_type || record.flowType || '').toLowerCase();
  if (flowType && flowType !== 'image') {
    return false;
  }

  return Boolean(extractSupabaseImageUrl_(record));
}

function getSupabaseWebhookProjectRecord_(data) {
  const candidates = [
    data?.record,
    data?.new,
    data?.data?.record,
    data?.event?.data?.record,
    data,
  ];

  for (let i = 0; i < candidates.length; i += 1) {
    const candidate = candidates[i];
    if (candidate && typeof candidate === 'object' && !Array.isArray(candidate)) {
      if (
        candidate.saved_image_url ||
        candidate.original_image_url ||
        candidate.drawing_url ||
        candidate.flow_type ||
        candidate.project_type
      ) {
        return candidate;
      }
    }
  }

  return null;
}

function getSupabaseWebhookEventType_(data) {
  return String(
    data?.type || data?.eventType || data?.event_type || data?.operation || ''
  ).toUpperCase();
}

function extractSupabaseImageUrl_(record) {
  return String(
    record?.saved_image_url ||
      record?.drawing_url ||
      record?.original_image_url ||
      ''
  ).trim();
}

function fetchImageFromUrl_(fileUrl) {
  const response = UrlFetchApp.fetch(fileUrl, {
    method: 'get',
    muteHttpExceptions: true,
  });
  const statusCode = response.getResponseCode();

  Logger.log('fetchImageFromUrl_ status: ' + statusCode);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('画像取得エラー: ' + response.getContentText());
  }

  const blob = response.getBlob();
  const fileName = inferFileNameFromUrl_(fileUrl);
  if (fileName) {
    blob.setName(fileName);
  }

  return {
    blob: blob,
    mimeType: blob.getContentType() || 'image/jpeg',
    base64: Utilities.base64Encode(blob.getBytes()),
    fileName: blob.getName() || fileName || '',
  };
}

function inferFileNameFromUrl_(fileUrl) {
  const match = String(fileUrl || '').match(/\/([^\/?#]+)(?:\?|#|$)/);
  return match ? decodeURIComponent(match[1]) : '';
}

function parseDateValue_(value) {
  if (!value) return null;

  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function normalizeSupabaseWebhookSource_(value) {
  const source = String(value || '').trim().toLowerCase();

  if (!source) {
    return 'Supabase';
  }

  if (source === 'line') return 'LINE';
  if (source === 'tally') return 'Tally';
  if (source === 'fax') return 'FAX';
  if (source === 'email') return 'Email';
  if (source === 'phone') return 'Phone';

  return String(value);
}

function buildSupabaseImageWebhookRawText_(record, imageData, imageUrl) {
  return [
    'Supabase画像Webhookを受信',
    'projectId: ' + String(record.id || ''),
    'sourceType: ' + String(record.source_type || record.source || ''),
    'projectType: ' + String(record.project_type || ''),
    'fileName: ' + String(record.original_file_name || imageData.fileName || ''),
    'imageUrl: ' + imageUrl,
    'contentType: ' + String(imageData.mimeType || ''),
    'size(bytes): ' + String(imageData.blob ? imageData.blob.getBytes().length : 0),
  ].join('\n');
}

function saveProjectToSupabase_(options) {
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    Logger.log('saveProjectToSupabase_ skipped: missing Supabase config');
    return null;
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
        Prefer: 'return=representation',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    const body = response.getContentText();

    Logger.log('saveProjectToSupabase_ status: ' + statusCode);

    if (statusCode < 200 || statusCode >= 300) {
      Logger.log('saveProjectToSupabase_ response: ' + body);
      return null;
    }

    const rows = JSON.parse(body || '[]');
    if (Array.isArray(rows) && rows.length > 0) {
      return rows[0];
    }

    return rows || null;
  } catch (error) {
    Logger.log('saveProjectToSupabase_ error: ' + error.message);
    Logger.log('saveProjectToSupabase_ error stack: ' + error.stack);
    return null;
  }
}

function saveProjectItemsToSupabase_(projectId, geminiResult, options) {
  if (!projectId || !SUPABASE_URL || !SUPABASE_KEY) {
    return;
  }

  const items = buildSupabaseProjectItemsPayloads_(
    projectId,
    geminiResult,
    options
  );

  if (!items.length) {
    Logger.log('saveProjectItemsToSupabase_ skipped: no items');
    return;
  }

  const url = SUPABASE_URL.replace(/\/$/, '') + '/rest/v1/project_items';

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: 'Bearer ' + SUPABASE_KEY,
        Prefer: 'return=minimal',
      },
      payload: JSON.stringify(items),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    Logger.log('saveProjectItemsToSupabase_ status: ' + statusCode);

    if (statusCode < 200 || statusCode >= 300) {
      Logger.log(
        'saveProjectItemsToSupabase_ response: ' + response.getContentText()
      );
    }
  } catch (error) {
    Logger.log('saveProjectItemsToSupabase_ error: ' + error.message);
    Logger.log('saveProjectItemsToSupabase_ error stack: ' + error.stack);
  }
}

function updateSupabaseProjectLedgerIdByImageLedger_(imageLedgerRow, ledgerId) {
  if (!SUPABASE_URL || !SUPABASE_KEY) {
    Logger.log(
      'updateSupabaseProjectLedgerIdByImageLedger_ skipped: missing Supabase config'
    );
    return;
  }

  const normalizedLedgerId = String(ledgerId || '').trim();
  const drawingUrl = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.DRAWING_URL] || ''
  ).trim();
  const originalFileName = String(
    imageLedgerRow[COLUMNS.IMAGE_LEDGER.ORIGINAL_FILE_NAME] || ''
  ).trim();

  if (!normalizedLedgerId) {
    throw new Error('ledgerId が指定されていません');
  }

  const query = buildSupabaseProjectUpdateQuery_(drawingUrl, originalFileName);

  if (!query) {
    Logger.log(
      'updateSupabaseProjectLedgerIdByImageLedger_ skipped: no match query'
    );
    return;
  }

  const url =
    SUPABASE_URL.replace(/\/$/, '') + '/rest/v1/projects?' + query;

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'patch',
      contentType: 'application/json',
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: 'Bearer ' + SUPABASE_KEY,
        Prefer: 'return=minimal',
      },
      payload: JSON.stringify({
        ledger_id: normalizedLedgerId,
        updated_at: new Date().toISOString(),
      }),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    Logger.log(
      'updateSupabaseProjectLedgerIdByImageLedger_ status: ' + statusCode
    );

    if (statusCode < 200 || statusCode >= 300) {
      Logger.log(
        'updateSupabaseProjectLedgerIdByImageLedger_ response: ' +
          response.getContentText()
      );
    }
  } catch (error) {
    Logger.log(
      'updateSupabaseProjectLedgerIdByImageLedger_ error: ' + error.message
    );
    Logger.log(
      'updateSupabaseProjectLedgerIdByImageLedger_ error stack: ' + error.stack
    );
  }
}

function buildSupabaseProjectUpdateQuery_(drawingUrl, originalFileName) {
  if (drawingUrl) {
    return 'saved_image_url=eq.' + encodeURIComponent(drawingUrl);
  }

  if (originalFileName) {
    return 'original_file_name=eq.' + encodeURIComponent(originalFileName);
  }

  return '';
}

function buildSupabaseProjectItemsPayloads_(projectId, geminiResult, options) {
  const normalizedItems = normalizeSupabaseProjectItems_(
    geminiResult,
    options
  );

  return normalizedItems.map(function (item, index) {
    return {
      project_id: projectId,
      item_no: index + 1,
      item_name: item.item_name || null,
      drawing_number: item.drawing_number || null,
      processing: item.processing || null,
      material: item.material || null,
      size_text: item.size_thickness || null,
      quantity: parseSupabaseDecimalOrNull_(item.quantity),
      unit_price: parseSupabaseAmount_(item.unit_price),
      amount: parseSupabaseAmount_(item.amount),
      due_date: parseSupabaseDueDate_(item.due_date || options?.parentDueDate),
      note: item.note || null,
      ai_extracted_json: item,
    };
  });
}

function normalizeSupabaseProjectItems_(geminiResult, options) {
  const items = Array.isArray(geminiResult?.items) ? geminiResult.items : [];

  if (items.length > 0) {
    return items;
  }

  const fallbackItem = {
    item_name: geminiResult?.project_name || '',
    drawing_number: geminiResult?.drawing_number || '',
    processing: geminiResult?.processing || '',
    material: geminiResult?.material || '',
    size_thickness: geminiResult?.size_thickness || '',
    quantity: geminiResult?.quantity || '',
    unit_price: '',
    amount: '',
    due_date: geminiResult?.desired_due_date || options?.parentDueDate || '',
    note: geminiResult?.notes || '',
  };

  const hasValue = Object.keys(fallbackItem).some(function (key) {
    return Boolean(fallbackItem[key]);
  });

  return hasValue ? [fallbackItem] : [];
}

function getLineUserState_(lineUserId) {
  if (!lineUserId || !SUPABASE_URL || !SUPABASE_KEY) {
    return null;
  }

  const url =
    SUPABASE_URL.replace(/\/$/, '') +
    '/rest/v1/line_user_states?line_user_id=eq.' +
    encodeURIComponent(lineUserId) +
    '&select=line_user_id,mode,payload,created_at,updated_at&limit=1';

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: 'Bearer ' + SUPABASE_KEY,
      },
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    const body = response.getContentText();

    if (statusCode < 200 || statusCode >= 300) {
      Logger.log('getLineUserState_ status: ' + statusCode);
      Logger.log('getLineUserState_ response: ' + body);
      return null;
    }

    const rows = JSON.parse(body || '[]');

    if (!Array.isArray(rows) || rows.length === 0) {
      return null;
    }

    return rows[0];
  } catch (error) {
    Logger.log('getLineUserState_ error: ' + error.message);
    Logger.log('getLineUserState_ error stack: ' + error.stack);
    return null;
  }
}

function setLineUserState_(lineUserId, mode, payload) {
  if (!lineUserId || !mode || !SUPABASE_URL || !SUPABASE_KEY) {
    return;
  }

  const url = SUPABASE_URL.replace(/\/$/, '') + '/rest/v1/line_user_states';
  const now = new Date().toISOString();
  const body = {
    line_user_id: lineUserId,
    mode: mode,
    payload: payload || {},
    created_at: now,
    updated_at: now,
  };

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: 'Bearer ' + SUPABASE_KEY,
        Prefer: 'resolution=merge-duplicates,return=minimal',
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode < 200 || statusCode >= 300) {
      Logger.log('setLineUserState_ status: ' + statusCode);
      Logger.log('setLineUserState_ response: ' + response.getContentText());
    }
  } catch (error) {
    Logger.log('setLineUserState_ error: ' + error.message);
    Logger.log('setLineUserState_ error stack: ' + error.stack);
  }
}

function clearLineUserState_(lineUserId) {
  if (!lineUserId || !SUPABASE_URL || !SUPABASE_KEY) {
    return;
  }

  const url =
    SUPABASE_URL.replace(/\/$/, '') +
    '/rest/v1/line_user_states?line_user_id=eq.' +
    encodeURIComponent(lineUserId);

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: 'Bearer ' + SUPABASE_KEY,
      },
      muteHttpExceptions: true,
    });

    const statusCode = response.getResponseCode();
    if (statusCode < 200 || statusCode >= 300) {
      Logger.log('clearLineUserState_ status: ' + statusCode);
      Logger.log('clearLineUserState_ response: ' + response.getContentText());
    }
  } catch (error) {
    Logger.log('clearLineUserState_ error: ' + error.message);
    Logger.log('clearLineUserState_ error stack: ' + error.stack);
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

  return [source, datePath, Utilities.getUuid() + extension].join('/');
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
  const flowType = options.flowType || inferSupabaseFlowType_(options);
  const drawingUrl = options.drawingUrl || ledgerRow.drawingUrl || '';
  const aiExtractedJson =
    options.aiExtractedJson ||
    ledgerRow.rawJson ||
    safeStringify_(result);

  return {
    flow_type: flowType,
    source_type: normalizeSupabaseSourceType_(ledgerRow.source),
    received_at: toIsoStringOrNull_(ledgerRow.receivedAt),
    status: normalizeSupabaseStatus_(ledgerRow.status),
    project_code: options.projectCode || null,
    quote_no: result.quote_no || options.quoteNo || null,
    project_type: options.projectType || ledgerRow.projectType || null,
    customer_name: ledgerRow.customerName || '',
    contact_name: ledgerRow.contactName || '',
    project_name: result.project_name || '',
    drawing_number: result.drawing_number || '',
    material: ledgerRow.material || '',
    size_text: ledgerRow.sizeThickness || '',
    quantity: parseSupabaseQuantity_(ledgerRow.quantity),
    due_date: parseSupabaseDueDate_(ledgerRow.dueDate),
    total_amount: parseSupabaseAmount_(result.total_amount || options.totalAmount),
    current_response_note: '',
    note: ledgerRow.notes || '',
    ai_summary: '',
    original_file_name:
      options.originalFileName || ledgerRow.originalFileName || null,
    original_image_url:
      options.originalImageUrl || ledgerRow.originalImageUrl || null,
    saved_image_url: options.savedImageUrl || drawingUrl || null,
    drawing_url: drawingUrl || null,
    ocr_text: options.ocrText || null,
    ai_extracted_json: parseSupabaseJsonOrNull_(aiExtractedJson),
    validation_result: options.validationResult || null,
    processing_status: options.processingStatus || null,
    error_message: options.errorMessage || null,
    ledger_id: options.ledgerId || ledgerRow.ledgerId || null,
    line_user_id: options.lineUserId || '',
    spreadsheet_row_no: null,
  };
}

function inferSupabaseFlowType_(options) {
  if (options.flowType) return options.flowType;

  if (
    options.projectType ||
    options.originalFileName ||
    options.originalImageUrl ||
    options.savedImageUrl ||
    options.ocrText ||
    options.validationResult ||
    options.processingStatus ||
    options.errorMessage
  ) {
    return 'image';
  }

  return 'normal';
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
  const match = text.match(/\d+(?:\.\d+)?/);

  if (!match) return null;

  const quantity = parseFloat(match[0]);
  return isNaN(quantity) ? null : quantity;
}

function parseSupabaseDueDate_(value) {
  const text = (value || '').trim();

  if (!text) return null;

  const normalized = text.replace(/まで|迄|期限|納期|希望/gi, '').trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    return normalized;
  }

  if (/^\d{4}\/\d{2}\/\d{2}$/.test(normalized)) {
    return normalized.replace(/\//g, '-');
  }

  const monthDayMatch = normalized.match(/^(\d{1,2})[\/-](\d{1,2})$/);
  if (monthDayMatch) {
    const now = new Date();
    const year = now.getFullYear();
    const month = padSupabaseDatePart_(monthDayMatch[1]);
    const day = padSupabaseDatePart_(monthDayMatch[2]);
    return year + '-' + month + '-' + day;
  }

  return null;
}

function parseSupabaseJsonOrNull_(value) {
  if (!value) return null;

  if (typeof value === 'object') {
    return value;
  }

  try {
    return JSON.parse(String(value));
  } catch (error) {
    Logger.log('parseSupabaseJsonOrNull_ error: ' + error.message);
    return null;
  }
}

function parseSupabaseAmount_(value) {
  if (value === null || value === undefined || value === '') return null;

  const text = String(value).replace(/[^0-9.\-]/g, '');
  if (!text) return null;

  const amount = parseFloat(text);
  return isNaN(amount) ? null : amount;
}

function parseSupabaseDecimalOrNull_(value) {
  if (value === null || value === undefined || value === '') return null;

  const text = String(value).replace(/[^0-9.\-]/g, '');
  if (!text) return null;

  const decimal = parseFloat(text);
  return isNaN(decimal) ? null : decimal;
}

function padSupabaseDatePart_(value) {
  return ('0' + String(value)).slice(-2);
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

function testHandleLinePostback() {
  const mockData = {
    events: [
      {
        type: 'postback',
        postback: {
          data: 'action=past_case_search',
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
  Logger.log('testHandleLinePostback result: ' + result.getContent());
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
    lines.push('・希望納期：' + formatLineDate_(result.desired_due_date));
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
