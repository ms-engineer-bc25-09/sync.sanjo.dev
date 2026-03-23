function buildMessagingApiRichMenuConfig_() {
  return {
    size: {
      width: 2500,
      height: 1686,
    },
    selected: true,
    name: 'sanjo-default-richmenu',
    chatBarText: 'メニュー',
    areas: [
      {
        bounds: {
          x: 0,
          y: 0,
          width: 1667,
          height: 843,
        },
        action: {
          type: 'message',
          text: '書類を撮る',
        },
      },
      {
        bounds: {
          x: 1667,
          y: 0,
          width: 833,
          height: 843,
        },
        action: {
          type: 'message',
          text: 'メモを残す',
        },
      },
      {
        bounds: {
          x: 0,
          y: 843,
          width: 833,
          height: 843,
        },
        action: {
          type: 'message',
          text: '声で残す',
        },
      },
      {
        bounds: {
          x: 833,
          y: 843,
          width: 834,
          height: 843,
        },
        action: {
          type: 'postback',
          data: 'past_case_search',
        },
      },
      {
        bounds: {
          x: 1667,
          y: 843,
          width: 833,
          height: 843,
        },
        action: {
          type: 'uri',
          uri: SPREADSHEET_URL,
        },
      },
    ],
  };
}

function createAndSetDefaultMessagingApiRichMenu_() {
  const richMenuId = createMessagingApiRichMenu_();

  uploadMessagingApiRichMenuImage_(richMenuId);
  setDefaultMessagingApiRichMenu_(richMenuId);

  Logger.log('createAndSetDefaultMessagingApiRichMenu_ richMenuId: ' + richMenuId);
  return richMenuId;
}

function createMessagingApiRichMenu_() {
  const payload = buildMessagingApiRichMenuConfig_();
  const response = callLineMessagingApi_('/v2/bot/richmenu', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
  });

  const richMenuId = response.richMenuId || '';

  if (!richMenuId) {
    throw new Error('richMenuId を取得できませんでした');
  }

  return richMenuId;
}

function uploadMessagingApiRichMenuImage_(richMenuId) {
  if (!richMenuId) {
    throw new Error('richMenuId が指定されていません');
  }

  if (!LINE_RICH_MENU_IMAGE_FILE_ID) {
    throw new Error(
      'LINE_RICH_MENU_IMAGE_FILE_ID が設定されていません'
    );
  }

  const file = DriveApp.getFileById(LINE_RICH_MENU_IMAGE_FILE_ID);
  const blob = file.getBlob();
  const contentType = blob.getContentType();

  if (contentType !== 'image/png' && contentType !== 'image/jpeg') {
    throw new Error('リッチメニュー画像は PNG または JPEG が必要です');
  }

  callLineMessagingApi_('/v2/bot/richmenu/' + richMenuId + '/content', {
    method: 'post',
    contentType: contentType,
    payload: blob.getBytes(),
  });
}

function setDefaultMessagingApiRichMenu_(richMenuId) {
  if (!richMenuId) {
    throw new Error('richMenuId が指定されていません');
  }

  callLineMessagingApi_('/v2/bot/user/all/richmenu/' + richMenuId, {
    method: 'post',
  });
}

function getDefaultMessagingApiRichMenu_() {
  return callLineMessagingApi_('/v2/bot/user/all/richmenu', {
    method: 'get',
  });
}

function deleteMessagingApiRichMenu_(richMenuId) {
  if (!richMenuId) {
    throw new Error('richMenuId が指定されていません');
  }

  callLineMessagingApi_('/v2/bot/richmenu/' + richMenuId, {
    method: 'delete',
  });
}

function callLineMessagingApi_(path, options) {
  if (!LINE_CHANNEL_ACCESS_TOKEN) {
    throw new Error('LINE_CHANNEL_ACCESS_TOKEN が設定されていません');
  }

  const response = UrlFetchApp.fetch('https://api.line.me' + path, {
    method: options.method || 'get',
    contentType: options.contentType,
    headers: {
      Authorization: 'Bearer ' + LINE_CHANNEL_ACCESS_TOKEN,
    },
    payload: options.payload,
    muteHttpExceptions: true,
  });

  const statusCode = response.getResponseCode();
  const responseText = response.getContentText();

  Logger.log('LINE Messaging API status: ' + statusCode);
  Logger.log('LINE Messaging API response: ' + responseText);

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error('LINE Messaging API error: ' + responseText);
  }

  if (!responseText) {
    return {};
  }

  return JSON.parse(responseText);
}
