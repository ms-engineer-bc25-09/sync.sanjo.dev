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

  Logger.log('image gemini result: ' + JSON.stringify(geminiResult));

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

  saveLineGeminiResultToSheets_({
    geminiResult: geminiResult,
    source: 'LINE',
    status: '未対応',
    rawText: rawText,
    lineUserId: lineUserId,
    drawingUrl: '',
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
      '不足があれば補足をテキストで送ってください。';

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
