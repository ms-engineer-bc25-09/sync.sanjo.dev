function callGemini_(text) {
  const url =
  "https://generativelanguage.googleapis.com/v1beta/models/" +
  GEMINI_MODEL +
  ":generateContent?key=" +
  GEMINI_API_KEY;

  const prompt = buildGeminiPrompt_(text);

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt
          }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log("Gemini status: " + statusCode);
  Logger.log("Gemini response: " + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Gemini呼び出しエラー: " + responseText);
  }

  const json = JSON.parse(responseText);
  const rawText =
    json.candidates?.[0]?.content?.parts?.[0]?.text || "";

  if (!rawText) {
    throw new Error("Geminiの返り値が空です");
  }

  let parsed;
  try {
    parsed = JSON.parse(rawText);
  } catch (error) {
    Logger.log("Gemini raw text parse error: " + error.message);
    Logger.log("Gemini raw text: " + rawText);
    throw new Error("GeminiのJSON解析に失敗しました");
  }

  return normalizeGeminiResult_(parsed, text);
}

function buildGeminiPrompt_(text) {
  return [
    "あなたは町工場の見積依頼メモから情報を抽出するアシスタントです。",
    "入力文から案件情報を抽出し、必ずJSONのみで返してください。",
    "説明文、補足、コードブロック、前置き、後置きは不要です。",
    "",
    "出力ルール:",
    "- 必ず1つのJSONオブジェクトだけ返す",
    "- キー名は必ず snake_case にする",
    "- 不明な値は空文字 \"\" にする",
    "- quantity はできるだけ元文に忠実に文字列で入れる（例: \"10個\"）",
    "- desired_due_date は元文の表現をそのまま入れる（例: \"来週まで\"）",
    "- notes には特記事項や補足を書く",
    "",
    "出力するJSONのキーは必ず以下のみを使うこと:",
    "{",
    '  "customer_name": "",',
    '  "contact_name": "",',
    '  "project_name": "",',
    '  "drawing_number": "",',
    '  "material": "",',
    '  "size_thickness": "",',
    '  "quantity": "",',
    '  "desired_due_date": "",',
    '  "notes": ""',
    "}",
    "",
    "入力文:",
    text
  ].join("\n");
}

function normalizeGeminiResult_(parsed, originalText) {
  return {
    customer_name: toSafeString_(parsed.customer_name),
    contact_name: toSafeString_(parsed.contact_name),
    project_name: toSafeString_(parsed.project_name),
    drawing_number: toSafeString_(parsed.drawing_number),
    material: toSafeString_(parsed.material),
    size_thickness: toSafeString_(parsed.size_thickness),
    quantity: toSafeString_(parsed.quantity),
    desired_due_date: toSafeString_(parsed.desired_due_date),
    notes: toSafeString_(parsed.notes) || originalText
  };
}

function toSafeString_(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}