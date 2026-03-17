function writeToLedger_(now, parsed) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(
    SHEET_NAMES.LEDGER
  );

  if (!sheet) {
    throw new Error("案件台帳シートが見つかりません");
  }

  sheet.appendRow([
    formatDate_(now),                   // 受付日時
    "未対応",                           // ステータス
    "LINE",                             // 受付経路
    parsed.customer_name || "",         // 顧客名
    parsed.contact_name || "",          // 担当者名
    "",                                 // メールアドレス
    "",                                 // 電話番号
    parsed.project_name || "",          // 案件内容
    parsed.desired_due_date || "",      // 希望納期
    parsed.material || "",              // 材質
    parsed.size_thickness || "",        // サイズ・板厚
    parsed.quantity || "",              // 数量
    parsed.notes || "",                 // 補足事項
    "",                                 // 図面URL
    JSON.stringify(parsed)              // raw_json
  ]);
}