function doPost(e) {
  try {
    const body = e && e.postData && e.postData.contents
      ? e.postData.contents
      : "{}";

    Logger.log("doPost body: " + body);

    const data = JSON.parse(body);

    if (Array.isArray(data.events) && data.events.length > 0) {
      Logger.log("LINE route");
      return handleLine_(data);
    }

    Logger.log("TALLY route");
    return handleTally_(data);
  } catch (error) {
    Logger.log("doPost error: " + error.message);

    return ContentService.createTextOutput(
      JSON.stringify({
        ok: false,
        error: error.message
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function testGemini() {
  const result = callGemini_("テストです。SUS304で10個、来週まで");
  Logger.log("testGemini result: " + JSON.stringify(result));
}

function testWriteToLedger() {
  const sample = {
    customer_name: "テスト工業",
    contact_name: "",
    project_name: "ブラケット試作",
    drawing_number: "",
    material: "SUS304",
    size_thickness: "",
    quantity: "10個",
    desired_due_date: "来週まで",
    notes: "急ぎ"
  };

  writeToLedger_(new Date(), sample);
  Logger.log("案件台帳へ書き込み完了");
}

function testHandleLine() {
  const mockData = {
    events: [
      {
        type: "message",
        message: {
          type: "text",
          text: "テスト工業です。ブラケット試作をお願いします。SUS304で10個、来週まで。急ぎです。"
        },
        source: {
          type: "user",
          userId: "TEST_LINE_USER_ID"
        },
        replyToken: "TEST_REPLY_TOKEN"
      }
    ]
  };

  const result = handleLine_(mockData);
  Logger.log("testHandleLine result: " + result.getContent()); 
}

function insertSampleData() {
  const sheet = SpreadsheetApp
    .openById(SHEET_ID)
    .getSheetByName(SHEET_NAMES.LEDGER);

  const rows = [
    ['2026/03/01 10:00:00', '対応済', 'LINE', '三条製作所', '佐藤', '', '', 'SUSブラケット加工', '2026/03/10', 'SUS304', '50×100×t3', '20', '急ぎ対応', '', ''],
    ['2026/03/02 11:00:00', '対応済', 'LINE', '三条製作所', '佐藤', '', '', 'アルミカバー製作', '2026/03/15', 'A5052', '200×300×t2', '5', '通常納期', '', ''],
    ['2026/03/03 09:00:00', '対応済', 'Tally', '田中鉄工所', '田中', '', '', '鉄プレート切断', '2026/03/12', 'SS400', '100×200×t6', '50', '', '', ''],
    ['2026/03/04 14:00:00', '対応済', 'LINE', '田中鉄工所', '田中', '', '', 'SUSシャフト加工', '2026/03/20', 'SUS303', 'φ20×L200', '10', '寸法公差あり', '', ''],
    ['2026/03/05 10:00:00', '対応済', 'FAX', '山本製作所', '山本', '', '', 'アルミブラケット', '2026/03/18', 'A6061', '80×80×t4', '30', '', '', ''],
    ['2026/03/06 15:00:00', '対応済', 'LINE', '山本製作所', '山本', '', '', '鉄ケース溶接', '2026/03/25', 'SS400', '150×200×H100', '3', '溶接仕上げ', '', ''],
    ['2026/03/07 09:30:00', '対応済', 'Tally', '新潟精工', '鈴木', '', '', 'SUSプレート穴あけ', '2026/03/14', 'SUS304', '200×200×t5', '100', '穴径φ10', '', ''],
    ['2026/03/08 11:00:00', '対応済', 'LINE', '新潟精工', '鈴木', '', '', 'アルミシャフト', '2026/03/22', 'A2017', 'φ15×L150', '8', '', '', ''],
    ['2026/03/09 13:00:00', '対応済', 'FAX', '燕三条工業', '渡辺', '', '', '鉄ブラケット溶接', '2026/03/16', 'SS400', '60×120×t4', '15', '黒皮材', '', ''],
    ['2026/03/10 10:00:00', '対応済', 'LINE', '燕三条工業', '渡辺', '', '', 'SUSカバー板金', '2026/03/28', 'SUS316', '300×400×t1.5', '2', '食品工場向け', '', ''],
    ['2026/03/11 09:00:00', '対応済', 'Tally', '加茂精密', '中村', '', '', 'アルミプレート加工', '2026/03/20', 'A5052', '120×240×t10', '25', '', '', ''],
    ['2026/03/12 14:00:00', '対応済', 'LINE', '加茂精密', '中村', '', '', 'SUSネジ加工', '2026/03/19', 'SUS303', 'M10×L30', '200', '右ネジ', '', ''],
    ['2026/03/13 10:00:00', '対応済', 'FAX', '見附工作所', '小林', '', '', '鉄シャフト旋削', '2026/03/24', 'S45C', 'φ30×L500', '4', '焼き入れ処理', '', ''],
    ['2026/03/14 11:00:00', '対応済', 'LINE', '見附工作所', '小林', '', '', 'アルミケース切削', '2026/03/30', 'A6061', '100×150×H80', '1', '試作品', '', ''],
  ];

  for (const row of rows) {
    sheet.appendRow(row);
  }

  Logger.log('サンプルデータ投入完了: ' + rows.length + '件');
}



