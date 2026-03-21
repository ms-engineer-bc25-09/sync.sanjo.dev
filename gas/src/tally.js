function handleTally_(data) {
  Logger.log('handleTally_ start');
  Logger.log('handleTally_ payload: ' + JSON.stringify(data));

  const now = new Date();
  const answers = normalizeTallyAnswers_(data);
  const fields = data?.data?.fields || data?.fields || [];
  const freeText = getFieldValueByLabel_(
    fields,
    '具体的な内容を教えてください（必須）',
  );
  const parsed = parseFreeText_(freeText);
  const rawJson = JSON.stringify(data);

  answers.inquiry = answers.inquiry || parsed.project_name || '';
  answers.dueDate = answers.dueDate || parsed.due_date || '';
  answers.material = answers.material || parsed.material || '';
  answers.quantity = answers.quantity || parsed.quantity || '';
  answers.notes = answers.notes || parsed.notes || '';

  Logger.log('handleTally_ normalized answers: ' + JSON.stringify(answers));
  Logger.log('handleTally_ parsed free text: ' + JSON.stringify(parsed));

  appendInternalLogRow_({
    id: '',
    createdAt: now,
    source: 'Tally',
    status: '未対応',
    customerName: answers.companyName,
    contactName: answers.contactName,
    projectName: answers.inquiry,
    drawingNumber: '',
    material: answers.material,
    size: answers.sizeThickness,
    quantity: answers.quantity,
    dueDate: answers.dueDate,
    notes: answers.notes,
    rawText: answers.inquiry,
    lineUserId: '',
    aiExtractedJson: rawJson,
    similarCase: '',
    pastUnitPrice: '',
    suggestedPrice: '',
    spreadsheetUpdatedAt: now,
  });

  Logger.log('handleTally_ internal log appended');

  appendLedgerRow_({
    receivedAt: now,
    status: '未対応',
    source: 'Tally',
    customerName: answers.companyName,
    contactName: answers.contactName,
    email: answers.email,
    phone: answers.phone,
    inquiry: answers.inquiry,
    dueDate: answers.dueDate,
    material: answers.material,
    sizeThickness: answers.sizeThickness,
    quantity: answers.quantity,
    notes: answers.notes,
    drawingUrl: answers.fileUrl,
    rawJson: rawJson,
  });

  Logger.log('handleTally_ ledger appended');

  notifyTallyInquiry_(answers);
  Logger.log('handleTally_ notify finished');

  return ContentService.createTextOutput(
    JSON.stringify({ ok: true }),
  ).setMimeType(ContentService.MimeType.JSON);
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
    return fieldLabel === label;
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
    companyName: "",
    contactName: "",
    email: "",
    phone: "",
    inquiry: "",
    dueDate: "",
    material: "",
    sizeThickness: "",
    quantity: "",
    notes: "",
    fileUrl: "",
  };

  for (const field of fields) {
    const label = field.label || field.title || "";
    const value = extractFieldValue_(field);

    if (label.includes("会社名")) result.companyName = value;
    else if (label.includes("ご担当者名")) result.contactName = value;
    else if (label.includes("メールアドレス")) result.email = value;
    else if (label.includes("電話番号")) result.phone = value;
    else if (label.includes("ご相談内容")) result.inquiry = value;
    else if (label.includes("希望納期")) result.dueDate = value;
    else if (label.includes("材質")) result.material = value;
    else if (label.includes("サイズ・板厚")) result.sizeThickness = value;
    else if (label.includes("数量")) result.quantity = value;
    else if (label.includes("補足事項")) result.notes = value;
    else if (label.includes("図面・参考資料")) result.fileUrl = value;
  }

  return result;
}
