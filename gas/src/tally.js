function handleTally_(data) {
  Logger.log('handleTally_ start');
  Logger.log('handleTally_ payload: ' + JSON.stringify(data));

  const now = new Date();
  const answers = normalizeTallyAnswers_(data);
  const fields = data?.data?.fields || data?.fields || [];
  const freeText = getTallyFreeText_(fields);
  const parsed = parseFreeText_(freeText);

  answers.inquiry = answers.inquiry || parsed.project_name || '';
  answers.dueDate = answers.dueDate || parsed.due_date || '';
  answers.material = answers.material || parsed.material || '';
  answers.quantity = answers.quantity || parsed.quantity || '';
  answers.notes = answers.notes || parsed.notes || '';

  let drawingUrl = answers.fileUrl || '';
  let uploadedImageData = null;

  if (answers.fileUrl) {
    try {
      const uploadResult = uploadTallyFileToSupabase_(answers.fileUrl, {
        receivedAt: now,
        source: 'tally',
      });
      drawingUrl = uploadResult.publicUrl;
      uploadedImageData = buildImageDataFromBlob_(
        uploadResult.blob,
        uploadResult.fileName
      );
      Logger.log('handleTally_ uploaded file to supabase: ' + drawingUrl);
    } catch (error) {
      Logger.log('handleTally_ file upload error: ' + error.message);
      Logger.log('handleTally_ file upload error stack: ' + error.stack);
    }
  }

  Logger.log('handleTally_ normalized answers: ' + JSON.stringify(answers));
  Logger.log('handleTally_ parsed free text: ' + JSON.stringify(parsed));

  if (isEmptyTallySubmission_(answers, freeText, parsed)) {
    Logger.log('handleTally_ skipped: empty submission');
    return ContentService.createTextOutput(
      JSON.stringify({ ok: true, skipped: 'empty_tally_submission' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const structuredResult = buildTallyStructuredResult_(answers, parsed);
  const structuredJson = safeStringify_(structuredResult);

  if (drawingUrl) {
    const imageData = uploadedImageData || fetchImageFromUrl_(drawingUrl);
    const geminiResult = callGeminiImage_(imageData.base64, imageData.mimeType);
    const mergedGeminiResult = mergeTallyAnswersIntoGeminiResult_(
      geminiResult,
      answers,
      parsed
    );
    const rawText = buildTallyImageRawText_(answers, freeText, drawingUrl);
    const originalFileName = imageData.fileName || inferFileNameFromUrl_(drawingUrl);
    const ledgerRow = buildLedgerRowFromGeminiResult_({
      geminiResult: mergedGeminiResult,
      receivedAt: now,
      rawText: rawText,
      source: 'Tally',
      status: 'AI登録済',
      drawingUrl: drawingUrl,
    });
    const ledgerId = createLedgerEntry_(ledgerRow);
    const internalLogRow = buildInternalLogRowFromGeminiResult_({
      geminiResult: mergedGeminiResult,
      ledgerRow: Object.assign({}, ledgerRow, {
        id: ledgerId,
      }),
      rawText: rawText,
      lineUserId: '',
    });
    const savedProject = saveProjectToSupabase_({
      geminiResult: mergedGeminiResult,
      ledgerRow: Object.assign({}, ledgerRow, {
        projectType: 'Web見積依頼',
        originalFileName: originalFileName,
        originalImageUrl: answers.fileUrl || '',
        drawingUrl: drawingUrl,
        ledgerId: ledgerId,
      }),
      flowType: 'image',
      projectType: 'Web見積依頼',
      originalFileName: originalFileName,
      originalImageUrl: answers.fileUrl || '',
      savedImageUrl: drawingUrl,
      drawingUrl: drawingUrl,
      ocrText: '',
      aiExtractedJson: safeStringify_(mergedGeminiResult),
      validationResult: '必要時確認',
      processingStatus: '案件台帳登録済',
      errorMessage: '',
      ledgerId: ledgerId,
      lineUserId: '',
    });
    const imageRecordId = savedProject && savedProject.id ? savedProject.id : '';
    const imageLedgerRow = buildImageLedgerRowFromGeminiResult_({
      id: imageRecordId,
      geminiResult: mergedGeminiResult,
      receivedAt: now,
      rawText: rawText,
      source: 'Tally',
      projectType: 'Web見積依頼',
      status: 'AI登録済',
      email: answers.email,
      phone: answers.phone,
      originalFileName: originalFileName,
      originalImageUrl: answers.fileUrl || '',
      drawingUrl: drawingUrl,
      ledgerId: ledgerId,
    });
    const imageInternalLogRow = buildImageInternalLogRowFromGeminiResult_({
      id: imageRecordId,
      geminiResult: mergedGeminiResult,
      imageLedgerRow: imageLedgerRow,
      createdAt: now,
      savedImageUrl: drawingUrl,
      processingStatus: '案件台帳登録済',
      validationResult: '必要時確認',
      errorMessage: '',
      ocrText: '',
      notes: rawText,
    });

    appendInternalLogRow_(internalLogRow);
    upsertImageInternalLogRow_(imageInternalLogRow);
    upsertImageLedgerRow_(imageLedgerRow);

    if (savedProject && savedProject.id) {
      saveProjectItemsToSupabase_(savedProject.id, mergedGeminiResult, {
        parentDueDate: imageLedgerRow.dueDate,
      });
    }

    notifyTallyInquiry_(buildTallyNotificationAnswers_(answers, mergedGeminiResult));
    Logger.log('handleTally_ notify finished');

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const rawText = buildTallyImageRawText_(answers, freeText, '');
  const ledgerRow = buildLedgerRowFromGeminiResult_({
    geminiResult: structuredResult,
    receivedAt: now,
    rawText: rawText,
    source: 'Tally',
    status: 'AI登録済',
    drawingUrl: '',
  });

  ledgerRow.email = answers.email || '';
  ledgerRow.phone = answers.phone || '';
  ledgerRow.rawJson = structuredJson;

  const ledgerId = createLedgerEntry_(ledgerRow);
  const internalLogRow = buildInternalLogRowFromGeminiResult_({
    geminiResult: structuredResult,
    ledgerRow: Object.assign({}, ledgerRow, {
      id: ledgerId,
    }),
    rawText: rawText,
    lineUserId: '',
  });

  appendInternalLogRow_(internalLogRow);

  Logger.log('handleTally_ internal log appended');

  const savedProject = saveProjectToSupabase_({
    geminiResult: structuredResult,
    ledgerRow: Object.assign({}, ledgerRow, {
      projectType: 'Web見積依頼',
      ledgerId: ledgerId,
    }),
    flowType: 'normal',
    projectType: 'Web見積依頼',
    aiExtractedJson: structuredJson,
    validationResult: '必要時確認',
    processingStatus: '案件台帳登録済',
    errorMessage: '',
    ledgerId: ledgerId,
    lineUserId: '',
  });

  if (savedProject && savedProject.id) {
    saveProjectItemsToSupabase_(savedProject.id, structuredResult, {
      parentDueDate: ledgerRow.dueDate,
    });
  }

  Logger.log('handleTally_ ledger appended');

  notifyTallyInquiry_(buildTallyNotificationAnswers_(answers, structuredResult));
  Logger.log('handleTally_ notify finished');

  return ContentService.createTextOutput(
    JSON.stringify({ ok: true }),
  ).setMimeType(ContentService.MimeType.JSON);
}

function mergeTallyAnswersIntoGeminiResult_(geminiResult, answers, parsed) {
  const result = Object.assign({}, geminiResult || {});
  const primaryItem =
    Array.isArray(result.items) && result.items.length > 0
      ? Object.assign({}, result.items[0])
      : {};

  result.customer_name = result.customer_name || answers.companyName || '';
  result.contact_name = result.contact_name || answers.contactName || '';
  result.project_name =
    result.project_name || answers.inquiry || parsed.project_name || '';
  result.material = result.material || answers.material || parsed.material || '';
  result.size_thickness =
    result.size_thickness || answers.sizeThickness || '';
  result.quantity = result.quantity || answers.quantity || parsed.quantity || '';
  result.desired_due_date =
    result.desired_due_date || answers.dueDate || parsed.due_date || '';
  result.notes = result.notes || answers.notes || parsed.notes || '';

  primaryItem.customer_name =
    primaryItem.customer_name || result.customer_name || '';
  primaryItem.contact_name = primaryItem.contact_name || result.contact_name || '';
  primaryItem.item_name =
    primaryItem.item_name || result.project_name || answers.inquiry || '';
  primaryItem.material = primaryItem.material || result.material || '';
  primaryItem.size_thickness =
    primaryItem.size_thickness || result.size_thickness || '';
  primaryItem.quantity = primaryItem.quantity || result.quantity || '';
  primaryItem.due_date =
    primaryItem.due_date || result.desired_due_date || '';
  primaryItem.note = primaryItem.note || result.notes || '';

  result.items = [primaryItem];

  return result;
}

function buildTallyImageRawText_(answers, freeText, drawingUrl) {
  return [
    'Tally画像フォームを受信',
    'companyName: ' + String(answers.companyName || ''),
    'contactName: ' + String(answers.contactName || ''),
    'email: ' + String(answers.email || ''),
    'phone: ' + String(answers.phone || ''),
    'inquiry: ' + String(answers.inquiry || ''),
    'dueDate: ' + String(answers.dueDate || ''),
    'fileUrl: ' + String(answers.fileUrl || ''),
    'savedImageUrl: ' + String(drawingUrl || ''),
    'freeText: ' + String(freeText || ''),
  ].join('\n');
}

function buildTallyNotificationAnswers_(answers, geminiResult) {
  return Object.assign({}, answers, {
    inquiry: geminiResult.project_name || answers.inquiry || '',
    dueDate: geminiResult.desired_due_date || answers.dueDate || '',
  });
}

function notifyTallyInquiry_(answers) {
  if (!LINE_NOTIFY_USER_ID) {
    Logger.log('LINE_NOTIFY_USER_ID 未設定のため通知スキップ');
    return;
  }

  const message = buildTallyNotifyMessage_(answers);

  try {
    pushLineMessage_(LINE_NOTIFY_USER_ID, [
      {
        type: 'text',
        text: message,
      },
    ]);
  } catch (error) {
    Logger.log('notifyTallyInquiry_ error: ' + error.message);
    Logger.log('notifyTallyInquiry_ error stack: ' + error.stack);
  }
}

function buildTallyNotifyMessage_(answers) {
  const inquiry = hasTallyValue_(answers.inquiry) ? answers.inquiry : '未入力';
  const dueDate = hasTallyValue_(answers.dueDate)
    ? answers.dueDate
    : '未入力（要確認）';

  return [
    '新規問い合わせ（見積依頼）が届きました。',
    '',
    '会社名：' + normalizeTallyValue_(answers.companyName),
    '担当者名：' + normalizeTallyValue_(answers.contactName),
    '案件名：' + inquiry,
    '希望納期：' + dueDate,
    '',
    '詳細は案件台帳を確認してください。',
    SPREADSHEET_URL,
  ].join('\n');
}

function normalizeTallyValue_(value) {
  return hasTallyValue_(value) ? String(value).trim() : '未入力';
}

function hasTallyValue_(value) {
  if (value === null || value === undefined) return false;
  return String(value).trim() !== '';
}

function getFieldValueByLabel_(fields, label) {
  if (!Array.isArray(fields) || !label) return '';

  const field = fields.find(function (item) {
    const fieldLabel = item?.label || item?.title || '';
    return fieldLabel === label || fieldLabel.includes(label);
  });

  return field ? extractFieldValue_(field) : '';
}

function parseFreeText_(text) {
  const result = {};

  if (!text || !String(text).trim()) return result;

  const lines = String(text).split(/\r?\n/);

  lines.forEach(function (line) {
    const trimmed = line.trim();
    if (!trimmed) return;

    const parts = trimmed.split(/[：:]/);
    if (parts.length < 2) return;

    const key = (parts[0] || '').trim();
    const value = parts.slice(1).join('：').trim();
    if (!value) return;

    if (key.includes('案件名')) {
      result.project_name = value;
      return;
    }

    if (key.includes('材質')) {
      result.material = value;
      return;
    }

    if (key.includes('数量')) {
      result.quantity = value;
      return;
    }

    if (key.includes('希望納期')) {
      result.due_date = value;
      return;
    }

    if (key.includes('備考')) {
      result.notes = value;
    }
  });

  return result;
}

function normalizeTallyAnswers_(payload) {
  const fields = payload?.data?.fields || payload?.fields || [];

  const result = {
    companyName: '',
    contactName: '',
    email: '',
    phone: '',
    inquiry: '',
    dueDate: '',
    material: '',
    sizeThickness: '',
    quantity: '',
    notes: '',
    fileUrl: '',
  };

  for (const field of fields) {
    const label = field.label || field.title || '';
    const value = extractFieldValue_(field);

    if (label.includes('会社名')) result.companyName = value;
    else if (label.includes('ご担当者名') || label.includes('氏名')) {
      result.contactName = value;
    } else if (label.includes('メールアドレス')) result.email = value;
    else if (label.includes('電話番号') || label.includes('ご連絡先電話番号')) {
      result.phone = value;
    } else if (label.includes('ご相談内容') || label.includes('具体的な内容')) {
      result.inquiry = value;
    } else if (label.includes('希望納期')) result.dueDate = value;
    else if (label.includes('材質')) result.material = value;
    else if (label.includes('サイズ・板厚') || label.includes('サイズ板厚')) {
      result.sizeThickness = value;
    } else if (label.includes('数量')) result.quantity = value;
    else if (label.includes('補足事項')) result.notes = value;
    else if (
      label.includes('図面・参考資料') ||
      label.includes('図面参考資料') ||
      label.includes('図面や写真のアップロード')
    ) {
      result.fileUrl = value;
    }
  }

  if (!result.companyName && result.contactName) {
    result.companyName = result.contactName;
  }

  return result;
}

function buildTallyStructuredResult_(answers, parsed) {
  return mergeTallyAnswersIntoGeminiResult_({}, answers, parsed);
}

function isEmptyTallySubmission_(answers, freeText, parsed) {
  const values = [
    answers.companyName,
    answers.contactName,
    answers.email,
    answers.phone,
    answers.inquiry,
    answers.dueDate,
    answers.material,
    answers.sizeThickness,
    answers.quantity,
    answers.notes,
    answers.fileUrl,
    freeText,
    parsed.project_name,
    parsed.due_date,
    parsed.material,
    parsed.quantity,
    parsed.notes,
  ];

  return !values.some(function (value) {
    return hasTallyValue_(value);
  });
}

function getTallyFreeText_(fields) {
  return (
    getFieldValueByLabel_(fields, '具体的な内容を教えてください') ||
    getFieldValueByLabel_(fields, '具体的な内容') ||
    ''
  );
}
