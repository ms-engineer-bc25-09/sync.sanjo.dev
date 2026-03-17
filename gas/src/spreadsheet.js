function writeToLedger_(now, parsed) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(
    SHEET_NAMES.LEDGER,
  );

  sheet.appendRow([
    formatDate_(now),
    "未対応",
    "LINE",
    parsed.customerName || "",
    parsed.contactName || "",
    "",
    "",
    parsed.inquiry || "",
    parsed.dueDate || "",
    parsed.material || "",
    parsed.sizeThickness || "",
    parsed.quantity || "",
    parsed.notes || "",
    "",
    "",
  ]);
}
