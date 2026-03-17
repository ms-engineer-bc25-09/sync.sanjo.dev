function handleTally_(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(
    SHEET_NAMES.LEDGER,
  );

  const now = new Date();
  const answers = normalizeTallyAnswers_(data);

  const row = [
    formatDate_(now), // 受付日時
    "未対応", // ステータス
    "Tally", // 受付経路
    answers.companyName, // 顧客名
    answers.contactName, // 担当者名
    answers.email, // メールアドレス
    answers.phone, // 電話番号
    answers.inquiry, // 案件内容
    answers.dueDate, // 希望納期
    answers.material, // 材質
    answers.sizeThickness, // サイズ・板厚
    answers.quantity, // 数量
    answers.notes, // 補足事項
    answers.fileUrl, // 図面・参考資料
    JSON.stringify(data), // raw_json
  ];

  sheet.appendRow(row);

  return ContentService.createTextOutput(
    JSON.stringify({ ok: true }),
  ).setMimeType(ContentService.MimeType.JSON);
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
