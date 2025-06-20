/**
 * LINEbot用Googleカレンダー連携システム
 * 機能：
 * 1. スプレッドシートUIからカレンダーID設定
 * 2. 今日のイベント取得
 * 3. LINE Messaging APIでの送信
 * 4. 手動送信・自動送信（朝8時）
 * 5. Webhook対応
 */

// 設定値
const CONFIG = {
  SHEET_NAME: 'カレンダー設定',
  CALENDAR_ID_CELL: 'B1',
  LINE_CHANNEL_ACCESS_TOKEN_CELL: 'B2',
  LINE_USER_ID_CELL: 'B3',
  WEBHOOK_VERIFY_TOKEN_CELL: 'B4'
};

/**
 * スプレッドシート開いた時の初期化処理
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 カレンダーLINEbot')
    .addItem('📝 設定シートを初期化', 'initializeSettingSheet')
    .addSeparator()
    .addItem('📤 今日のイベントを手動送信', 'sendTodayEventsManually')
    .addItem('🌅 明日のイベントを手動送信', 'sendTomorrowEventsManually')
    .addItem('📊 週間イベントを手動送信', 'sendWeekEventsManually')
    .addSeparator()
    .addItem('⏰ 朝8時の自動送信を設定', 'setupMorningTrigger')
    .addItem('❌ 自動送信を停止', 'deleteMorningTrigger')
    .addItem('📋 現在のトリガー一覧を表示', 'showCurrentTriggers')
    .addSeparator()
    .addItem('🔗 Webhook URLを表示', 'showWebhookUrl')
    .addItem('🔍 設定状況を確認', 'showSettingsStatus')
    .addToUi();
  
  // 設定シートが存在しない場合は自動作成
  const sheet = getOrCreateSettingSheet();
  if (sheet) {
    console.log('設定シートを確認/作成しました');
  }
}

/**
 * 設定シートを取得または作成
 */
function getOrCreateSettingSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);
    initializeSettingSheet();
  }
  
  return sheet;
}

/**
 * 設定シートの初期化
 */
function initializeSettingSheet() {
  const sheet = getOrCreateSettingSheet();
  
  // ヘッダー設定
  sheet.getRange('A1').setValue('GoogleカレンダーID');
  sheet.getRange('A2').setValue('LINE Channel Access Token');
  sheet.getRange('A3').setValue('LINE User ID');
  sheet.getRange('A4').setValue('Webhook Verify Token');
  
  // 説明を追加
  sheet.getRange('C1').setValue('例: your-calendar@gmail.com');
  sheet.getRange('C2').setValue('LINE Developers コンソールから取得');
  sheet.getRange('C3').setValue('送信先のLINE User ID');
  sheet.getRange('C4').setValue('Webhook検証用のトークン（任意）');
  
  // セルの書式設定
  sheet.getRange('A1:A4').setFontWeight('bold');
  sheet.getRange('B1:B4').setBackground('#f0f0f0');
  
  // 列幅調整
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 250);
  
  SpreadsheetApp.getUi().alert('設定シートを初期化しました。\nB列に必要な情報を入力してください。');
}

/**
 * 設定値を取得
 */
function getSettings() {
  const sheet = getOrCreateSettingSheet();
  
  return {
    calendarId: sheet.getRange(CONFIG.CALENDAR_ID_CELL).getValue(),
    lineChannelAccessToken: sheet.getRange(CONFIG.LINE_CHANNEL_ACCESS_TOKEN_CELL).getValue(),
    lineUserId: sheet.getRange(CONFIG.LINE_USER_ID_CELL).getValue(),
    webhookVerifyToken: sheet.getRange(CONFIG.WEBHOOK_VERIFY_TOKEN_CELL).getValue()
  };
}

/**
 * 今日のカレンダーイベントを取得
 */
function getTodayEvents(calendarId) {
  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error('指定されたカレンダーIDが見つかりません');
    }
    
    const today = new Date();
    const startOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const endOfDay = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1);
    
    const events = calendar.getEvents(startOfDay, endOfDay);
    
    return events.map(event => ({
      title: event.getTitle(),
      startTime: event.getStartTime(),
      endTime: event.getEndTime(),
      description: event.getDescription() || '',
      location: event.getLocation() || ''
    }));
  } catch (error) {
    console.error('カレンダーイベント取得エラー:', error);
    throw error;
  }
}

/**
 * LINEメッセージを作成
 */
function createLineMessage(events) {
  const today = new Date();
  const dateStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy年MM月dd日(E)');
  
  if (events.length === 0) {
    return `📅 ${dateStr}\n\n今日の予定はありません。`;
  }
  
  let message = `📅 ${dateStr}の予定\n\n`;
  
  events.forEach((event, index) => {
    const startTime = Utilities.formatDate(event.startTime, Session.getScriptTimeZone(), 'HH:mm');
    const endTime = Utilities.formatDate(event.endTime, Session.getScriptTimeZone(), 'HH:mm');
    
    message += `${index + 1}. ${event.title}\n`;
    message += `⏰ ${startTime} - ${endTime}\n`;
    
    if (event.location) {
      message += `📍 ${event.location}\n`;
    }
    
    if (event.description) {
      message += `📝 ${event.description}\n`;
    }
    
    message += '\n';
  });
  
  return message.trim();
}

/**
 * LINEメッセージを送信
 */
function sendLineMessage(message, accessToken, userId) {
  const url = 'https://api.line.me/v2/bot/message/push';
  
  // 送信前にパラメータをログ出力
  console.log('=== LINE送信パラメータ ===');
  console.log('User ID:', userId);
  console.log('Access Token前半:', accessToken ? accessToken.substring(0, 20) + '...' : 'なし');
  console.log('Message:', message.substring(0, 100) + (message.length > 100 ? '...' : ''));
  
  const payload = {
    to: userId,
    messages: [{
      type: 'text',
      text: message
    }]
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${accessToken}`
    },
    payload: JSON.stringify(payload)
  };
  
  try {
    console.log('送信ペイロード:', JSON.stringify(payload));
    const response = UrlFetchApp.fetch(url, options);
    const responseData = response.getContentText();
    
    console.log('Response Code:', response.getResponseCode());
    console.log('Response Data:', responseData);
    
    if (response.getResponseCode() !== 200) {
      throw new Error(`LINE API Error: ${response.getResponseCode()} - ${responseData}`);
    }
    
    console.log('LINEメッセージ送信成功');
    return true;
  } catch (error) {
    console.error('LINEメッセージ送信エラー:', error);
    throw error;
  }
}

/**
 * 今日のイベントを手動送信
 */
function sendTodayEventsManually() {
  try {
    const settings = getSettings();
    
    // 設定値の検証
    if (!settings.calendarId || !settings.lineChannelAccessToken || !settings.lineUserId) {
      SpreadsheetApp.getUi().alert(
        'エラー: 設定が不完全です\n\n' +
        '以下を設定シートで確認してください:\n' +
        '- GoogleカレンダーID\n' +
        '- LINE Channel Access Token\n' +
        '- LINE User ID'
      );
      return;
    }
    
    // イベント取得
    const events = getTodayEvents(settings.calendarId);
    
    // メッセージ作成
    const message = createLineMessage(events);
    
    // LINE送信
    sendLineMessage(message, settings.lineChannelAccessToken, settings.lineUserId);
    
    SpreadsheetApp.getUi().alert(`✅ 送信完了\n\n今日の予定 ${events.length}件をLINEに送信しました。`);
    
  } catch (error) {
    console.error('手動送信エラー:', error);
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n\n${error.message}`);
  }
}

/**
 * 朝8時の自動送信トリガーを設定
 */
function setupMorningTrigger() {
  try {
    // 既存のトリガーを削除
    deleteMorningTrigger();
    
    // 新しいトリガーを作成（毎日朝8時）
    ScriptApp.newTrigger('sendTodayEventsAutomatically')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();
    
    SpreadsheetApp.getUi().alert('✅ 自動送信を設定しました\n\n毎日朝8時に今日の予定をLINEに送信します。');
    
  } catch (error) {
    console.error('トリガー設定エラー:', error);
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n\n${error.message}`);
  }
}

/**
 * 朝8時の自動送信トリガーを削除
 */
function deleteMorningTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === 'sendTodayEventsAutomatically') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    
    console.log('既存の自動送信トリガーを削除しました');
    
  } catch (error) {
    console.error('トリガー削除エラー:', error);
  }
}

/**
 * 自動送信実行関数（トリガーから呼ばれる）
 */
function sendTodayEventsAutomatically() {
  try {
    const settings = getSettings();
    
    // 設定値の検証
    if (!settings.calendarId || !settings.lineChannelAccessToken || !settings.lineUserId) {
      console.error('自動送信: 設定が不完全です');
      return;
    }
    
    // イベント取得
    const events = getTodayEvents(settings.calendarId);
    
    // メッセージ作成
    const message = createLineMessage(events);
    
    // LINE送信
    sendLineMessage(message, settings.lineChannelAccessToken, settings.lineUserId);
    
    console.log(`自動送信完了: 今日の予定 ${events.length}件をLINEに送信しました`);
    
  } catch (error) {
    console.error('自動送信エラー:', error);
  }
}

/**
 * Webhook URLを表示
 */
function showWebhookUrl() {
  const scriptId = ScriptApp.getScriptId();
  const webhookUrl = `https://script.google.com/macros/s/${scriptId}/exec`;
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Webhook URL',
    `以下のURLをLINE DevelopersのWebhook URLに設定してください:\n\n${webhookUrl}\n\n※このURLはGASをWebアプリとしてデプロイした後に有効になります。`,
    ui.ButtonSet.OK
  );
}

/**
 * Webhook処理（POSTリクエスト）
 */
function doPost(e) {
  try {
    // postDataが存在しない場合（検証リクエスト等）は200を返す
    if (!e.postData || !e.postData.contents) {
      console.log('Webhook検証リクエストを受信');
      return ContentService.createTextOutput('OK')
        .setMimeType(ContentService.MimeType.TEXT)
        .setStatusCode(200);
    }
    
    const settings = getSettings();
    
    // Webhook検証（設定されている場合）
    if (settings.webhookVerifyToken) {
      const receivedToken = e.parameter.token || e.postData?.contents ? JSON.parse(e.postData.contents).token : null;
      if (receivedToken !== settings.webhookVerifyToken) {
        return ContentService.createTextOutput('Unauthorized')
          .setMimeType(ContentService.MimeType.TEXT)
          .setStatusCode(401);
      }
    }
    
    // LINEからのWebhookデータを解析
    const data = JSON.parse(e.postData.contents);
    console.log('受信データ:', JSON.stringify(data));
    
    if (data.events && data.events.length > 0) {
      data.events.forEach(event => {
        console.log('イベントタイプ:', event.type);
        if (event.type === 'message' && event.message.type === 'text') {
          handleTextMessage(event);
        }
      });
    }
    
    return ContentService.createTextOutput('OK')
      .setMimeType(ContentService.MimeType.TEXT)
      .setStatusCode(200);
    
  } catch (error) {
    console.error('Webhook処理エラー:', error);
    return ContentService.createTextOutput('OK')
      .setMimeType(ContentService.MimeType.TEXT)
      .setStatusCode(200);
  }
}

/**
 * テキストメッセージの処理
 */
function handleTextMessage(event) {
  try {
    const settings = getSettings();
    const userId = event.source.userId;
    const messageText = event.message.text.trim();
    
    // User ID取得用の特別なコマンド
    if (messageText === 'userid' || messageText === 'ユーザーID' || messageText === 'test') {
      const idMessage = `📋 あなたのUser ID:\n${userId}\n\nこのIDをスプレッドシートのB3セルに入力してください。`;
      sendLineMessage(idMessage, settings.lineChannelAccessToken, userId);
      console.log(`User ID取得: ${userId}`);
      return;
    }
    
    let responseMessage = '';
    
    if (messageText === '今日の予定' || messageText === '予定' || messageText === '今日') {
      // 今日のカレンダーイベントを取得して送信
      const events = getTodayEvents(settings.calendarId);
      responseMessage = createLineMessage(events);
      
    } else if (messageText === '明日の予定' || messageText === '明日') {
      // 明日のカレンダーイベントを取得して送信
      const events = getTomorrowEvents(settings.calendarId);
      responseMessage = createTomorrowLineMessage(events);
      
    } else if (messageText === '今週の予定' || messageText === '今週' || messageText === '週間') {
      // 週間のカレンダーイベントを取得して送信
      const events = getWeekEvents(settings.calendarId);
      responseMessage = createWeekLineMessage(events);
      
    } else if (messageText === 'ヘルプ' || messageText === 'help' || messageText === '？' || messageText === '?') {
      responseMessage = '📅 カレンダーbot使い方\n\n' +
                       '「今日の予定」「予定」「今日」: 今日の予定を表示\n' +
                       '「明日の予定」「明日」: 明日の予定を表示\n' +
                       '「今週の予定」「今週」「週間」: 今週の予定を表示\n' +
                       '「userid」「test」: User IDを表示\n' +
                       '「ヘルプ」「help」「？」: この使い方を表示';
      
    } else {
      responseMessage = 'こんにちは！📅\n\n' +
                       '以下のコマンドが使用できます：\n\n' +
                       '📝 「今日の予定」: 今日の予定を表示\n' +
                       '🌅 「明日の予定」: 明日の予定を表示\n' +
                       '📊 「今週の予定」: 今週の予定を表示\n' +
                       '🆔 「userid」: User IDを取得\n' +
                       '❓ 「ヘルプ」: 詳しい使い方を表示\n\n' +
                       'お気軽にお使いください！';
    }
    
    // レスポンスを送信
    sendLineMessage(responseMessage, settings.lineChannelAccessToken, userId);
    
  } catch (error) {
    console.error('テキストメッセージ処理エラー:', error);
    
    // エラー時のフォールバック応答
    const errorMessage = 'すみません、エラーが発生しました。\n少し時間をおいてから再度お試しください。';
    try {
      sendLineMessage(errorMessage, settings.lineChannelAccessToken, userId);
    } catch (fallbackError) {
      console.error('フォールバック応答送信エラー:', fallbackError);
    }
  }
}

/**
 * Webhook処理（GETリクエスト）
 */
function doGet(e) {
  console.log('GET リクエストを受信:', JSON.stringify(e));
  return ContentService.createTextOutput('LINEbot Webhook Endpoint')
    .setMimeType(ContentService.MimeType.TEXT)
    .setStatusCode(200);
}

/**
 * テスト用関数
 */
function testGetTodayEvents() {
  const settings = getSettings();
  if (!settings.calendarId) {
    console.log('カレンダーIDが設定されていません');
    return;
  }
  
  try {
    const events = getTodayEvents(settings.calendarId);
    console.log('今日のイベント:', events);
    
    const message = createLineMessage(events);
    console.log('作成されたメッセージ:', message);
    
  } catch (error) {
    console.error('テストエラー:', error);
  }
} 