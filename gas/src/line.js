function handleLine_(data) {
  try {
    const events = data.events || [];

    for (const event of events) {
      if (event.type !== "message") continue;
      if (event.message?.type !== "text") continue;

      const lineUserId = event.source?.userId || "";
      const text = event.message?.text || "";
      const now = new Date();

      writeToInternalLog_(now, lineUserId, text);

      const parsed = callGemini_(text);
      if (!parsed) continue;

      writeToLedger_(now, parsed);
      writeToSupabase_(now, lineUserId, parsed);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ ok: true }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: error.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function writeToInternalLog_(now, lineUserId, text) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(
    SHEET_NAMES.INTERNAL_LOG,
  );

  sheet.appendRow([formatDate_(now), "LINE", lineUserId, text]);
}

function writeToSupabase_(now, lineUserId, parsed) {
  const url = SUPABASE_URL + "/rest/v1/projects";

  const payload = {
    source_type: "line",
    status: "received",
    customer_name: parsed.customerName || null,
    contact_name: parsed.contactName || null,
    project_name: parsed.projectName || null,
    drawing_number: parsed.drawingNumber || null,
    material: parsed.material || null,
    size_text: parsed.sizeThickness || null,
    quantity: parsed.quantity ? parseInt(parsed.quantity) : null,
    due_date: parsed.dueDate || null,
    note: parsed.notes || null,
    line_user_id: lineUserId || null,
  };

  const options = {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: SUPABASE_KEY,
      Authorization: "Bearer " + SUPABASE_KEY,
      Prefer: "return=minimal",
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  };

  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 201) {
    throw new Error("Supabase書き込みエラー: " + response.getContentText());
  }
}
