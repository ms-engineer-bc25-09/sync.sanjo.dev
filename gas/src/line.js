function handleLine_(data) {
  Logger.log('handleLine_ called: ' + JSON.stringify(data));
  try {
    const events = data.events || [];

    for (const event of events) {
      try {
        if (event.type !== 'message') continue;
        if (event.message?.type !== 'text') continue;

        const lineUserId = event.source?.userId || '';
        const text = event.message?.text || '';
        const now = new Date();

        if (!text) continue;

        // 1. Geminiで抽出
        const geminiResult = callGemini_(text);
        if (!geminiResult) {
          Logger.log('Gemini result is empty.');
          continue;
        }

        // 2. 文字列JSON / オブジェクト の両方に対応
        let parsed;
        try {
          parsed =
            typeof geminiResult === 'string'
              ? JSON.parse(geminiResult)
              : geminiResult;
        } catch (parseError) {
          Logger.log('Gemini JSON parse error: ' + parseError.message);
          Logger.log('Gemini raw result: ' + geminiResult);

          if (event.replyToken) {
            replyLineMessage_(event.replyToken, [
              {
                type: 'text',
                text: '受付に失敗しました。内容を少し変えてもう一度送ってください。',
              },
            ]);
          }
          continue;
        }

        Logger.log(
          'Gemini parsed(before normalize): ' + JSON.stringify(parsed)
        );

        // 3. 表記ゆれを軽く正規化
        parsed = normalizeParsedData_(parsed);

        Logger.log('Gemini parsed(after normalize): ' + JSON.stringify(parsed));

        // 4. 保存用データへ整形
        const formatted = formatProjectData_(parsed, text);

        Logger.log('Formatted project data: ' + JSON.stringify(formatted));

        // 5. デモ用バリデーション
        const validation = validateProjectDataForDemo_(formatted);

        if (!validation.isValid) {
          Logger.log(
            'Project data rejected by demo validation: ' +
              JSON.stringify(validation)
          );

          if (event.replyToken) {
            replyLineMessage_(event.replyToken, [
              {
                type: 'text',
                text:
                  '情報が少ないため受付できませんでした。\n' +
                  '「案件名・材質・数量」のうち、少なくとも2つが分かるように送ってください。\n' +
                  '例）SUS304のブラケットを10個、来週まで',
              },
            ]);
          }
          continue;
        }

        // 6. 内部ログへ保存
        writeToInternalLog_(now, lineUserId, text, formatted);

        // 7. 案件台帳へ保存
        writeToLedger_(now, formatted);

<<<<<<< HEAD
        // 8. Supabaseへ保存
        writeToSupabase_(now, lineUserId, formatted);

        // 9. LINE返信
        if (event.replyToken) {
          const projectName = formatted.project_name || '-';
          const material = formatted.material || '-';
          const quantity = formatted.quantity || '-';
=======
        // 8. Supabaseへ保存（デバッグ付き）
        let supabaseResult = 'ok';

        try {
          writeToSupabase_(now, lineUserId, formatted);
        } catch (e) {
          supabaseResult = 'ERROR: ' + e.message;
          Logger.log('Supabase debug error: ' + e.message);
        }

        // 9. LINE返信
        if (event.replyToken) {
          const material = formatted.material || '-';
          const quantity = formatted.quantity || '-';
          const projectName = formatted.project_name || '-';
>>>>>>> 9afcb1e (fix: LINEルート整形・バリデーション追加)

          replyLineMessage_(event.replyToken, [
            {
              type: 'text',
              text:
                '受付しました。\n' +
                `案件名：${projectName}\n` +
                `材質：${material}\n` +
                `数量：${quantity}`,
            },
          ]);
        }
      } catch (eventError) {
        Logger.log('Event error: ' + eventError.message);
        Logger.log('Event payload: ' + JSON.stringify(event));

        if (event.replyToken) {
          replyLineMessage_(event.replyToken, [
            {
              type: 'text',
              text: '処理中にエラーが発生しました。もう一度送信してください。',
            },
          ]);
        }
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log('handleLine_ error: ' + error.message);

    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function formatProjectData_(data, rawText) {
  return {
    customer_name: data.customer_name || '不明',
    contact_name: data.contact_name || '',
    project_name: data.project_name || '（案件名未設定）',
    drawing_number: data.drawing_number || '',
    material: data.material || '不明',
    size_thickness: data.size_thickness || data.size_text || '',
    quantity: data.quantity || '不明',
    desired_due_date: data.desired_due_date || '未定',
    notes: buildNotes_(data.notes, rawText),
  };
}

function buildNotes_(notes, rawText) {
  if (notes && rawText) {
    return notes + '\n\n【元メッセージ】\n' + rawText;
  }
  if (notes) {
    return notes;
  }
  if (rawText) {
    return '【元メッセージ】\n' + rawText;
  }
  return '';
}

function validateProjectDataForDemo_(data) {
  const checks = {
    project_name: hasMeaningfulValue_(data.project_name, [
      '（案件名未設定）',
      '不明',
      '',
    ]),
    material: hasMeaningfulValue_(data.material, ['不明', '']),
    quantity: hasMeaningfulValue_(data.quantity, ['不明', '']),
  };

  const score = Object.values(checks).filter(Boolean).length;

  return {
    isValid: score >= 2,
    score: score,
    checks: checks,
  };
}

function hasMeaningfulValue_(value, invalidValues) {
  if (value === null || value === undefined) return false;

  const normalized = String(value).trim();
  if (!normalized) return false;

  return !invalidValues.includes(normalized);
}

function normalizeParsedData_(parsed) {
  const normalized = Object.assign({}, parsed);

  normalized.material = normalizeMaterial_(parsed.material);
  normalized.quantity = normalizeQuantity_(parsed.quantity);

  return normalized;
}

function normalizeMaterial_(material) {
  if (!material) return '';

  const value = String(material).trim();
<<<<<<< HEAD
=======

  // 空白やハイフンを軽く吸収
>>>>>>> 9afcb1e (fix: LINEルート整形・バリデーション追加)
  const compact = value.replace(/[\s\-－_]/g, '').toUpperCase();

  if (/^SUS\d+$/.test(compact)) {
    return compact;
  }

  return value.toUpperCase();
}

function normalizeQuantity_(quantity) {
  if (!quantity) return '';

  return String(quantity).trim();
}

function writeToInternalLog_(now, lineUserId, text, parsed) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(
    SHEET_NAMES.INTERNAL_LOG
  );

  sheet.appendRow([
<<<<<<< HEAD
    '',
    formatDate_(now),
    'LINE',
    '未対応',
    parsed.customer_name || '',
    parsed.contact_name || '',
    parsed.project_name || '',
    parsed.drawing_number || '',
    parsed.material || '',
    parsed.size_thickness || parsed.size_text || '',
    parsed.quantity || '',
    parsed.desired_due_date || '',
    parsed.notes || '',
    text,
    lineUserId,
    JSON.stringify(parsed),
    '',
    '',
    '',
    formatDate_(now),
=======
    '', // ID
    formatDate_(now), // 作成日時
    'LINE', // 受付経路
    '未対応', // ステータス
    parsed.customer_name || '', // 顧客名
    parsed.contact_name || '', // 担当者名
    parsed.project_name || '', // 案件名
    parsed.drawing_number || '', // 図面番号
    parsed.material || '', // 材質
    parsed.size_thickness || parsed.size_text || '', // サイズ
    parsed.quantity || '', // 数量
    parsed.desired_due_date || '', // 希望納期
    parsed.notes || '', // 備考
    text, // 元テキスト
    lineUserId, // LINEユーザーID
    JSON.stringify(parsed), // 保存用整形後JSON
    '', // 類似案件候補
    '', // 過去単価
    '', // 今回提示単価
    formatDate_(now), // スプレッドシート更新日時
>>>>>>> 9afcb1e (fix: LINEルート整形・バリデーション追加)
  ]);
}

function writeToSupabase_(now, lineUserId, parsed) {
  const url = SUPABASE_URL + '/rest/v1/projects';

  const quantityValue = parsed.quantity
    ? parseInt(String(parsed.quantity).replace(/[^\d]/g, ''), 10)
    : null;

  const payload = {
    source_type: 'line',
    status: 'received',
    customer_name: parsed.customer_name || null,
    contact_name: parsed.contact_name || null,
    project_name: parsed.project_name || null,
    drawing_number: parsed.drawing_number || null,
    material: parsed.material || null,
    size_text: parsed.size_thickness || parsed.size_text || null,
    quantity: Number.isNaN(quantityValue) ? null : quantityValue,
    due_date: null,
    note: parsed.notes || null,
    line_user_id: lineUserId || null,
    received_at: formatDateForSupabase_(now),
  };

  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      apikey: SUPABASE_KEY,
      Authorization: 'Bearer ' + SUPABASE_KEY,
      Prefer: 'return=minimal',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('Supabase status: ' + statusCode);
  Logger.log('Supabase response: ' + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('Supabase書き込みエラー: ' + responseText);
  }
}

function formatDateForSupabase_(date) {
  return date.toISOString();
}

function replyLineMessage_(replyToken, messages) {
  const url = 'https://api.line.me/v2/bot/message/reply';

  const payload = {
    replyToken: replyToken,
    messages: messages,
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  Logger.log('LINE reply status: ' + response.getResponseCode());
  Logger.log('LINE reply response: ' + response.getContentText());
}
