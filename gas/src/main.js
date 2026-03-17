function doPost(e) {
  try {
    const body = e.postData?.contents || "{}";
    const data = JSON.parse(body);

    if (data.events) {
      return handleLine_(data);
    } else {
      return handleTally_(data);
    }
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ ok: false, error: error.message }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
