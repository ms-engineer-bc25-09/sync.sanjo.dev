function doPost(e) {
  try {
    const body =
      e && e.postData && e.postData.contents ? e.postData.contents : '{}';

    Logger.log('doPost body: ' + body);

    const data = JSON.parse(body);

    if (Array.isArray(data.events) && data.events.length > 0) {
      Logger.log('LINE route');
      return handleLine_(data);
    }

    Logger.log('TALLY route');
    return handleTally_(data);
  } catch (error) {
    Logger.log('doPost error: ' + error.message);

    return ContentService.createTextOutput(
      JSON.stringify({
        ok: false,
        error: error.message,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function testGemini() {
  const result = callGemini_('テストです。SUS304で10個、来週まで');
  Logger.log('testGemini result: ' + JSON.stringify(result));
}

function testWriteToLedger() {
  const sample = {
    customer_name: 'テスト工業',
    contact_name: '',
    project_name: 'ブラケット試作',
    drawing_number: '',
    material: 'SUS304',
    size_thickness: '',
    quantity: '10個',
    desired_due_date: '来週まで',
    notes: '急ぎ',
  };

  writeToLedger_(new Date(), sample);
  Logger.log('案件台帳へ書き込み完了');
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
