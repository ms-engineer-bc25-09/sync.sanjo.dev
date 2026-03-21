function replyLineMessage_(replyToken, messages) {
  if (!replyToken) {
    Logger.log('replyLineMessage_: replyToken is empty');
    return;
  }

  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN が設定されていません');
  }

  const normalizedMessages = Array.isArray(messages)
    ? messages
    : [
        {
          type: 'text',
          text: String(messages || ''),
        },
      ];

  const url = 'https://api.line.me/v2/bot/message/reply';

  const payload = {
    replyToken: replyToken,
    messages: normalizedMessages,
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

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('LINE reply status: ' + statusCode);
  Logger.log('LINE reply response: ' + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('LINE返信エラー: ' + responseText);
  }
}

function pushLineMessage_(toUserId, messages) {
  if (!toUserId) {
    Logger.log('pushLineMessage_: toUserId is empty');
    return;
  }

  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN が設定されていません');
  }

  const normalizedMessages = Array.isArray(messages)
    ? messages
    : [
        {
          type: 'text',
          text: String(messages || ''),
        },
      ];

  const url = 'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: toUserId,
    messages: normalizedMessages,
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

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('LINE push status: ' + statusCode);
  Logger.log('LINE push response: ' + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('LINE pushエラー: ' + responseText);
  }
}

function getLineImageContent_(messageId) {
  if (!messageId) {
    throw new Error('messageId が指定されていません');
  }

  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN が設定されていません');
  }

  const url =
    'https://api-data.line.me/v2/bot/message/' + messageId + '/content';

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN,
    },
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  Logger.log('LINE image fetch status: ' + statusCode);

  if (statusCode < 200 || statusCode >= 300) {
    const errorText = response.getContentText();
    Logger.log('LINE image fetch error response: ' + errorText);
    throw new Error('LINE画像取得エラー: ' + errorText);
  }

  const blob = response.getBlob();
  if (!blob) {
    throw new Error('LINE画像Blobの取得に失敗しました');
  }

  blob.setName('line-image-' + messageId);

  const mimeType = blob.getContentType() || 'image/jpeg';
  const bytes = blob.getBytes();
  const base64 = Utilities.base64Encode(bytes);

  Logger.log('LINE image mimeType: ' + mimeType);
  Logger.log('LINE image bytes length: ' + bytes.length);
  Logger.log('LINE image base64 length: ' + base64.length);

  return {
    blob: blob,
    mimeType: mimeType,
    base64: base64,
  };
}
