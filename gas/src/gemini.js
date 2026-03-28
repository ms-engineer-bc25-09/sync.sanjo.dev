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
    '- size_thickness にはサイズと板厚をまとめて入れる',
    '- size_thickness は可能なら「300×200 t3.0mm」のように整える',
    '- サイズは「100×200」のように × を使い、数字を単純連結しない',
    '- 板厚は「t2.0」「t3.0mm」のように t を付けて表記する',
    '- サイズのみ分かる場合はサイズだけ、板厚のみ分かる場合は板厚だけを入れる',
    '- quantity はできるだけ元文に忠実に文字列で入れる（例: \"10個\"）',
    '- desired_due_date は元文の表現をそのまま入れる（例: \"来週まで\"）',
    '- processing には加工分類や加工内容を入れる（例: \"レーザー加工\", \"曲げ加工\", \"レーザー加工＋曲げ\"）',
    '- notes には補足事項や特記事項を書く',
    '',
    '出力するJSONのキーは必ず以下のみを使うこと:',
    '{',
    '  "customer_name": "",',
    '  "contact_name": "",',
    '  "email": "",',
    '  "phone": "",',
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
    '1画像に複数の明細が含まれる場合は、必ず items 配列に分けてください。',
    '必ずJSONのみで返してください。',
    '説明文、補足、コードブロック、前置き、後置きは不要です。',
    '',
    '出力ルール:',
    '- 必ず1つのJSONオブジェクトだけ返す',
    '- キー名は必ず snake_case にする',
    '- 不明な値は空文字 "" にする',
    '- project_name は「案件名」として抽出する',
    '- drawing_number は図面番号だけを入れる',
    '- size_thickness にはサイズと板厚をまとめて入れる',
    '- size_thickness は可能なら「300×200 t3.0mm」のように整える',
    '- サイズは「100×200」のように × を使い、数字を単純連結しない',
    '- 板厚は「t2.0」「t3.0mm」のように t を付けて表記する',
    '- サイズのみ分かる場合はサイズだけ、板厚のみ分かる場合は板厚だけを入れる',
    '- quantity は画像中の表記に忠実に文字列で入れる（例: \"10個\"）',
    '- desired_due_date は画像中の表現をそのまま入れる',
    '- processing には加工分類や加工内容を入れる',
    '- notes には補足事項や特記事項を書く',
    '- email にはメールアドレスを入れる',
    '- phone には電話番号を入れる',
    '- quote_no には見積Noを入れる（例: "S-0214"）',
    '- total_amount には合計金額を数字または数字入り文字列で入れる',
    '- items は明細ごとの配列にする',
    '- items が1件だけでも配列にする',
    '- items[n].item_name は明細の品名や部品名',
    '- items[n].quantity はその明細の数量を入れる',
    '- items[n].unit_price と items[n].amount は画像にあれば入れる',
    '- 明細ごとの納期があれば items[n].due_date に入れる',
    '- 読み取れない項目は推測しすぎず空文字にする',
    '',
    '出力するJSONのキーは必ず以下のみを使うこと:',
    '{',
    '  "customer_name": "",',
    '  "contact_name": "",',
    '  "email": "",',
    '  "phone": "",',
    '  "project_name": "",',
    '  "quote_no": "",',
    '  "total_amount": "",',
    '  "drawing_number": "",',
    '  "processing": "",',
    '  "material": "",',
    '  "size_thickness": "",',
    '  "quantity": "",',
    '  "desired_due_date": "",',
    '  "notes": "",',
    '  "items": [',
    '    {',
    '      "item_name": "",',
    '      "drawing_number": "",',
    '      "processing": "",',
    '      "material": "",',
    '      "size_thickness": "",',
    '      "quantity": "",',
    '      "unit_price": "",',
    '      "amount": "",',
    '      "due_date": "",',
    '      "note": ""',
    '    }',
    '  ]',
    '}',
  ].join('\n');
}

function normalizeGeminiResult_(parsed, originalText) {
  const notes = toSafeString_(parsed.notes);
  const fallbackText = toSafeString_(originalText);

  return {
    customer_name: toSafeString_(parsed.customer_name),
    contact_name: toSafeString_(parsed.contact_name),
    email: toSafeString_(parsed.email),
    phone: toSafeString_(parsed.phone),
    project_name: toSafeString_(parsed.project_name),
    quote_no: toSafeString_(parsed.quote_no),
    total_amount: toSafeString_(parsed.total_amount),
    drawing_number: toSafeString_(parsed.drawing_number),
    processing: toSafeString_(parsed.processing),
    material: toSafeString_(parsed.material),
    size_thickness: normalizeSizeThickness_(parsed.size_thickness),
    quantity: toSafeString_(parsed.quantity),
    desired_due_date: toSafeString_(parsed.desired_due_date),
    notes: notes || fallbackText,
    items: normalizeGeminiItems_(parsed.items),
  };
}

function normalizeGeminiItems_(items) {
  if (!Array.isArray(items)) return [];

  return items
    .map(function (item) {
      return {
        item_name: toSafeString_(item?.item_name),
        drawing_number: toSafeString_(item?.drawing_number),
        processing: toSafeString_(item?.processing),
        material: toSafeString_(item?.material),
        size_thickness: normalizeSizeThickness_(item?.size_thickness),
        quantity: toSafeString_(item?.quantity),
        unit_price: toSafeString_(item?.unit_price),
        amount: toSafeString_(item?.amount),
        due_date: toSafeString_(item?.due_date),
        note: toSafeString_(item?.note),
      };
    })
    .filter(function (item) {
      return (
        item.item_name ||
        item.drawing_number ||
        item.processing ||
        item.material ||
        item.size_thickness ||
        item.quantity ||
        item.unit_price ||
        item.amount ||
        item.due_date ||
        item.note
      );
    });
}

function normalizeSizeThickness_(value) {
  let normalized = toSafeString_(value);

  if (!normalized) return '';

  normalized = normalized.replace(/[✕xX]/g, '×');
  normalized = normalized.replace(/\s+/g, ' ').trim();
  normalized = normalized.replace(/mm\s*t\s*=?\s*/gi, ' t');
  normalized = normalized.replace(/mmt\s*=?\s*/gi, ' t');
  normalized = normalized.replace(/\bt\s*=\s*/gi, 't');
  normalized = normalized.replace(/\bt\s+/gi, 't');
  normalized = normalized.replace(/×\s*t/gi, ' t');
  normalized = normalized.replace(/(\d)\s*t(\d)/gi, '$1 t$2');

  const compact = normalized.replace(/\s+/g, '');
  const joinedSizeThicknessMatch = compact.match(
    /^(\d{2,4})(\d{2,4})mm?t(\d+(?:\.\d+)?)$/i
  );

  if (joinedSizeThicknessMatch) {
    return (
      joinedSizeThicknessMatch[1] +
      '×' +
      joinedSizeThicknessMatch[2] +
      ' t' +
      joinedSizeThicknessMatch[3] +
      'mm'
    );
  }

  const joinedSizeWithThicknessMatch = compact.match(
    /^(\d{2,4})(\d{2,4})t(\d+(?:\.\d+)?)$/i
  );

  if (joinedSizeWithThicknessMatch) {
    return (
      joinedSizeWithThicknessMatch[1] +
      '×' +
      joinedSizeWithThicknessMatch[2] +
      ' t' +
      joinedSizeWithThicknessMatch[3]
    );
  }

  const joinedSizeMatch = compact.match(/^(\d{2,4})(\d{2,4})mm$/i);

  if (joinedSizeMatch) {
    return joinedSizeMatch[1] + '×' + joinedSizeMatch[2] + ' mm';
  }

  return normalized;
}

function toSafeString_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}
