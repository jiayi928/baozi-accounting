// ── 包子記帳 LINE Bot 代理 (Google Apps Script) ──────────
// 部署步驟：
//   1. 到 https://script.google.com 建立新專案，貼上此檔案
//   2. 點「部署」→「新增部署」→ 類型選「網頁應用程式」
//   3. 執行身分：「我自己」；誰可以存取：「所有人」
//   4. 複製部署 URL，填入 js/config.js 的 GAS_PROXY_URL
//   5. 將部署 URL 填到 LINE Developers → Webhook URL

const CHANNEL_ACCESS_TOKEN = 'mCVsMtHT30VvCl4NHRhj9Igv5DUBRWWIbEbhlyWPkJhPcBX472p6dbiZlQUmhGnsR7dg5+H4pYKPFudD18crmqftMZyhMYp4H9qNPPjKrcQjkGfeARIFWt4pFxvT2373hEqs2vkCvyb7a5e7fE6PMgdB04t89/1O/w1cDnyilFU=';
const AUTH_KEY = 'baozi2025';

function doPost(e) {
  try {
    const json = JSON.parse(e.postData.contents);

    // LINE Webhook 事件（LINE 伺服器送來）
    if (json.events) {
      json.events.forEach(handleLineEvent);
      return ContentService.createTextOutput('OK');
    }

    // PWA 發送通知請求
    if (json.userId && json.message) {
      if (json.authKey !== AUTH_KEY) {
        return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      pushMessage(json.userId, json.message);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch(err) {
    Logger.log(err);
  }
  return ContentService.createTextOutput('OK');
}

function doGet(e) {
  return ContentService.createTextOutput('LINE Bot Proxy is running.');
}

function handleLineEvent(event) {
  if (event.type === 'follow') {
    replyMessage(event.replyToken,
      '歡迎加入包子記帳！🥟\n\n您的 LINE User ID 是：\n\n' + event.source.userId +
      '\n\n請複製上方 ID，貼到 App 設定頁面的「LINE User ID」欄位，即可收到財務提醒通知。');
  } else if (event.type === 'message' && event.message.type === 'text') {
    replyMessage(event.replyToken,
      '您的 LINE User ID 是：\n\n' + event.source.userId +
      '\n\n請複製此 ID，貼到 App 設定頁面的「LINE User ID」欄位。');
  }
}

function replyMessage(replyToken, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
}

function pushMessage(userId, text) {
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
}
