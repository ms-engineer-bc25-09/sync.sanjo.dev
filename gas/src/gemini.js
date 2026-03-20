function callGemini_(text) {
  if (!text || !String(text).trim()) {
    throw new Error('Geminiに渡すテキストが空です');
  }

  const url = buildGeminiApiUrl_();
  const prompt = buildGeminiPrompt_(text);

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt,
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
    },
  };

  const responseText = fetchGemini_(url, payload);
  const rawText = extractGeminiText_(responseText);
  const parsed = parseGeminiJson_(rawText);

  return normalizeGeminiResult_(parsed, text);
}

function callGeminiImage_(base64Image, mimeType) {
  if (!base64Image || !String(base64Image).trim()) {
    throw new Error('Geminiに渡す画像データが空です');
  }

  const url = buildGeminiApiUrl_();
  const prompt = buildGeminiImagePrompt_();

  const payload = {
    contents: [
      {
        parts: [
          {
            text: prompt,
          },
          {
            inlineData: {
              mimeType: mimeType || 'image/jpeg',
              data: base64Image,
            },
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json',
    },
  };

  const responseText = fetchGemini_(url, payload);
  const rawText = extractGeminiText_(responseText);
  const parsed = parseGeminiJson_(rawText);

  return normalizeGeminiResult_(parsed, '');
}

function buildGeminiApiUrl_() {
  if (!GEMINI_MODEL) {
    throw new Error('GEMINI_MODEL が設定されていません');
  }

  if (!GEMINI_API_KEY) {
    throw new Error('GEMINI_API_KEY が設定されていません');
  }

  return (
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    GEMINI_MODEL +
    ':generateContent?key=' +
    GEMINI_API_KEY
  );
}

function fetchGemini_(url, payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('Gemini status: ' + statusCode);
  Logger.log('Gemini response: ' + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Gemini呼び出しエラー: ' + responseText);
  }

  return responseText;
}

function extractGeminiText_(responseText) {
  let json;

  try {
    json = JSON.parse(responseText);
  } catch (error) {
    Logger.log('Gemini response JSON parse error: ' + error.message);
    throw new Error('GeminiのレスポンスJSON解析に失敗しました');
  }

  const rawText = json.candidates?.[0]?.content?.parts?.[0]?.text || '';

  if (!rawText) {
    Logger.log('Gemini empty text response: ' + responseText);
    throw new Error('Geminiの返り値が空です');
  }

  return rawText;
}

function parseGeminiJson_(rawText) {
  try {
    return JSON.parse(rawText);
  } catch (error) {
    Logger.log('Gemini raw text parse error: ' + error.message);
    Logger.log('Gemini raw text: ' + rawText);
  }

  const cleaned = String(rawText)
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/\s*```$/i, '')
    .trim();

  try {
    return JSON.parse(cleaned);
  } catch (error) {
    Logger.log('Gemini cleaned text parse error: ' + error.message);
    Logger.log('Gemini cleaned text: ' + cleaned);
    throw new Error('GeminiのJSON解析に失敗しました');
  }
}

function buildGeminiPrompt_(text) {
  return [
    'あなたは町工場の見積依頼メモから案件情報を抽出するアシスタントです。',
    '入力文から必要な情報を抽出し、必ずJSONのみで返してください。',
    '説明文、補足、コードブロック、前置き、後置きは不要です。',
    '',
    '出力ルール:',
    '- 必ず1つのJSONオブジェクトだけ返す',
    '- キー名は必ず snake_case にする',
    '- 不明な値は空文字 "" にする',
    '- project_name は「案件名」として抽出する',
    '- drawing_number は図面番号だけを入れる',
    '- size_thickness にはサイズと板厚をまとめて入れる（例: \"300×200×t3mm\"）',
    '- quantity はできるだけ元文に忠実に文字列で入れる（例: \"10個\"）',
    '- desired_due_date は元文の表現をそのまま入れる（例: \"来週まで\"）',
    '- processing には加工分類や加工内容を入れる（例: \"レーザー加工\", \"曲げ加工\", \"レーザー加工＋曲げ\"）',
    '- notes には補足事項や特記事項を書く',
    '',
    '出力するJSONのキーは必ず以下のみを使うこと:',
    '{',
    '  "customer_name": "",',
    '  "contact_name": "",',
    '  "project_name": "",',
    '  "drawing_number": "",',
    '  "processing": "",',
    '  "material": "",',
    '  "size_thickness": "",',
    '  "quantity": "",',
    '  "desired_due_date": "",',
    '  "notes": ""',
    '}',
    '',
    '入力文:',
    text,
  ].join('\n');
}

function buildGeminiImagePrompt_() {
  return [
    'あなたは町工場向けの見積受付アシスタントです。',
    '添付された画像（図面、FAX、手書きメモ、見積依頼書など）を読み取り、案件情報を抽出してください。',
    '必ずJSONのみで返してください。',
    '説明文、補足、コードブロック、前置き、後置きは不要です。',
    '',
    '出力ルール:',
    '- 必ず1つのJSONオブジェクトだけ返す',
    '- キー名は必ず snake_case にする',
    '- 不明な値は空文字 "" にする',
    '- project_name は「案件名」として抽出する',
    '- drawing_number は図面番号だけを入れる',
    '- size_thickness にはサイズと板厚をまとめて入れる（例: \"300×200×t3mm\"）',
    '- quantity は画像中の表記に忠実に文字列で入れる（例: \"10個\"）',
    '- desired_due_date は画像中の表現をそのまま入れる',
    '- processing には加工分類や加工内容を入れる',
    '- notes には補足事項や特記事項を書く',
    '- 読み取れない項目は推測しすぎず空文字にする',
    '',
    '出力するJSONのキーは必ず以下のみを使うこと:',
    '{',
    '  "customer_name": "",',
    '  "contact_name": "",',
    '  "project_name": "",',
    '  "drawing_number": "",',
    '  "processing": "",',
    '  "material": "",',
    '  "size_thickness": "",',
    '  "quantity": "",',
    '  "desired_due_date": "",',
    '  "notes": ""',
    '}',
  ].join('\n');
}

function normalizeGeminiResult_(parsed, originalText) {
  const notes = toSafeString_(parsed.notes);
  const fallbackText = toSafeString_(originalText);

  return {
    customer_name: toSafeString_(parsed.customer_name),
    contact_name: toSafeString_(parsed.contact_name),
    project_name: toSafeString_(parsed.project_name),
    drawing_number: toSafeString_(parsed.drawing_number),
    processing: toSafeString_(parsed.processing),
    material: toSafeString_(parsed.material),
    size_thickness: normalizeSizeThickness_(parsed.size_thickness),
    quantity: toSafeString_(parsed.quantity),
    desired_due_date: toSafeString_(parsed.desired_due_date),
    notes: notes || fallbackText,
  };
}

function normalizeSizeThickness_(value) {
  return toSafeString_(value);
}

function toSafeString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}
