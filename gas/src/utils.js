function formatDate_(date) {
  return Utilities.formatDate(date, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
}

function extractFieldValue_(field) {
  if (typeof field.value === "string") return field.value;
  if (typeof field.answer === "string") return field.answer;

  if (Array.isArray(field.value)) {
    return field.value
      .map((v) => {
        if (typeof v === "string") return v;
        if (v?.url) return v.url;
        return JSON.stringify(v);
      })
      .join(", ");
  }

  if (Array.isArray(field.answer)) {
    return field.answer
      .map((v) => {
        if (typeof v === "string") return v;
        if (v?.url) return v.url;
        return JSON.stringify(v);
      })
      .join(", ");
  }

  if (field.value?.url) return field.value.url;
  if (field.answer?.url) return field.answer.url;
  if (field.value != null) return JSON.stringify(field.value);
  if (field.answer != null) return JSON.stringify(field.answer);

  return "";
}
