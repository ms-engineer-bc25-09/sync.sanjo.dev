function handleLine_(data) {
  Logger.log("handleLine_ called: " + JSON.stringify(data));
  try {
    const events = data.events || [];

    for (const event of events) {
      try {
        if (event.type !== "message") continue;
        if (event.message?.type !== "text") continue;

        const lineUserId = event.source?.userId || "";
        const text = event.message?.text || "";
        const now = new Date();

        if (!text) continue;

        // 1. まず内部ログへ保存
        writeToInternalLog_(now, lineUserId, text);

        // 2. Geminiで抽出
        const geminiResult = callGemini_(text);
        if (!geminiResult) {
          Logger.log("Gemini result is empty.");
          continue;
        }

        // 3. 文字列JSON / オブジェクト の両方に対応
        let parsed;
        try {
          parsed =
            typeof geminiResult === "string"
              ? JSON.parse(geminiResult)
              : geminiResult;
        } catch (parseError) {
          Logger.log("Gemini JSON parse error: " + parseError.message);
          Logger.log("Gemini raw result: " + geminiResult);
          continue;
        }

        Logger.log("Gemini parsed: " + JSON.stringify(parsed));

        // 4. 必要なら元本文をnotesへ補完
        if (!parsed.notes) {
          parsed.notes = text;
        }

        // 5. 案件台帳へ保存
        writeToLedger_(now, parsed);

        // 6. Supabaseへ保存
        writeToSupabase_(now, lineUserId, parsed);

        // 7. LINE返信（必要なら有効化）
        // replyLineMessage_(event.replyToken, [
        //   {
        //     type: "text",
        //     text: "受付しました。案件台帳へ保存しました。"
        //   }
        // ]);
      } catch (eventError) {
        Logger.log("Event error: " + eventError.message);
        Logger.log("Event payload: " + JSON.stringify(event));

        // 失敗しても他のイベント処理は継続
        // 必要なら返信を有効化
        // if (event.replyToken) {
        //   replyLineMessage_(event.replyToken, [
        //     {
        //       type: "text",
        //       text: "受付処理中にエラーが発生しました。"
        //     }
        //   ]);
        // }
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log("handleLine_ error: " + error.message);

    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function writeToInternalLog_(now, lineUserId, text) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(
    SHEET_NAMES.INTERNAL_LOG
  );

  sheet.appendRow([
    formatDate_(now),
    "LINE",
    lineUserId,
    text
  ]);
}

function writeToSupabase_(now, lineUserId, parsed) {
  const url = SUPABASE_URL + "/rest/v1/projects";

  const quantityValue = parsed.quantity
    ? parseInt(String(parsed.quantity).replace(/[^\d]/g, ""), 10)
    : null;

  const payload = {
    source_type: "line",
    status: "received",

    // Gemini返り値は snake_case 前提
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

    // 必要なら受付日時も保存
    received_at: formatDateForSupabase_(now)
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_KEY,
      Authorization: "Bearer " + SUPABASE_KEY,
      Prefer: "return=minimal"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log("Supabase status: " + statusCode);
  Logger.log("Supabase response: " + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Supabase書き込みエラー: " + responseText);
  }
}

function formatDateForSupabase_(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// 必要なら使う
// function replyLineMessage_(replyToken, messages) {
//   const url = "https://api.line.me/v2/bot/message/reply";
//
//   const payload = {
//     replyToken: replyToken,
//     messages: messages
//   };
//
//   UrlFetchApp.fetch(url, {
//     method: "post",
//     contentType: "application/json",
//     headers: {
//       Authorization: "Bearer " + LINE_CHANNEL_ACCESS_TOKEN
//     },
//     payload: JSON.stringify(payload),
//     muteHttpExceptions: true
//   });
// }