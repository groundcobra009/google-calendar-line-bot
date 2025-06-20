/**
 * ユーティリティ関数とヘルパー関数
 * 追加機能や便利な機能を提供
 */

/**
 * 明日のカレンダーイベントを取得
 */
function getTomorrowEvents(calendarId) {
  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error('指定されたカレンダーIDが見つかりません');
    }
    
    const tomorrow = new Date();
    tomorrow.setDate(tomorrow.getDate() + 1);
    const startOfDay = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate());
    const endOfDay = new Date(tomorrow.getFullYear(), tomorrow.getMonth(), tomorrow.getDate() + 1);
    
    const events = calendar.getEvents(startOfDay, endOfDay);
    
    return events.map(event => ({
      title: event.getTitle(),
      startTime: event.getStartTime(),
      endTime: event.getEndTime(),
      description: event.getDescription() || '',
      location: event.getLocation() || ''
    }));
  } catch (error) {
    console.error('明日のカレンダーイベント取得エラー:', error);
    throw error;
  }
}

/**
 * 週間のカレンダーイベントを取得
 */
function getWeekEvents(calendarId) {
  try {
    const calendar = CalendarApp.getCalendarById(calendarId);
    if (!calendar) {
      throw new Error('指定されたカレンダーIDが見つかりません');
    }
    
    const today = new Date();
    const startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay()); // 日曜日
    
    const endOfWeek = new Date(startOfWeek);
    endOfWeek.setDate(startOfWeek.getDate() + 7); // 次の日曜日
    
    const events = calendar.getEvents(startOfWeek, endOfWeek);
    
    return events.map(event => ({
      title: event.getTitle(),
      startTime: event.getStartTime(),
      endTime: event.getEndTime(),
      description: event.getDescription() || '',
      location: event.getLocation() || ''
    }));
  } catch (error) {
    console.error('週間カレンダーイベント取得エラー:', error);
    throw error;
  }
}

/**
 * 明日の予定LINEメッセージを作成
 */
function createTomorrowLineMessage(events) {
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const dateStr = Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy年MM月dd日(E)');
  
  if (events.length === 0) {
    return `📅 ${dateStr}\n\n明日の予定はありません。`;
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
 * 週間予定LINEメッセージを作成
 */
function createWeekLineMessage(events) {
  if (events.length === 0) {
    return `📅 今週の予定\n\n今週の予定はありません。`;
  }
  
  // 日付ごとにイベントをグループ化
  const eventsByDate = {};
  
  events.forEach(event => {
    const dateKey = Utilities.formatDate(event.startTime, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!eventsByDate[dateKey]) {
      eventsByDate[dateKey] = [];
    }
    eventsByDate[dateKey].push(event);
  });
  
  let message = `📅 今週の予定\n\n`;
  
  // 日付順にソート
  const sortedDates = Object.keys(eventsByDate).sort();
  
  sortedDates.forEach(dateKey => {
    const dateStr = Utilities.formatDate(new Date(dateKey), Session.getScriptTimeZone(), 'MM月dd日(E)');
    message += `■ ${dateStr}\n`;
    
    eventsByDate[dateKey].forEach(event => {
      const startTime = Utilities.formatDate(event.startTime, Session.getScriptTimeZone(), 'HH:mm');
      const endTime = Utilities.formatDate(event.endTime, Session.getScriptTimeZone(), 'HH:mm');
      
      message += `  • ${event.title} (${startTime}-${endTime})\n`;
    });
    
    message += '\n';
  });
  
  return message.trim();
}

/**
 * 明日の予定を手動送信
 */
function sendTomorrowEventsManually() {
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
    
    // 明日のイベント取得
    const events = getTomorrowEvents(settings.calendarId);
    
    // メッセージ作成
    const message = createTomorrowLineMessage(events);
    
    // LINE送信
    sendLineMessage(message, settings.lineChannelAccessToken, settings.lineUserId);
    
    SpreadsheetApp.getUi().alert(`✅ 送信完了\n\n明日の予定 ${events.length}件をLINEに送信しました。`);
    
  } catch (error) {
    console.error('明日の予定手動送信エラー:', error);
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n\n${error.message}`);
  }
}

/**
 * 週間予定を手動送信
 */
function sendWeekEventsManually() {
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
    
    // 週間イベント取得
    const events = getWeekEvents(settings.calendarId);
    
    // メッセージ作成
    const message = createWeekLineMessage(events);
    
    // LINE送信
    sendLineMessage(message, settings.lineChannelAccessToken, settings.lineUserId);
    
    SpreadsheetApp.getUi().alert(`✅ 送信完了\n\n今週の予定 ${events.length}件をLINEに送信しました。`);
    
  } catch (error) {
    console.error('週間予定手動送信エラー:', error);
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n\n${error.message}`);
  }
}

/**
 * 設定情報の妥当性チェック
 */
function validateSettings() {
  try {
    const settings = getSettings();
    const results = {
      calendarId: false,
      lineChannelAccessToken: false,
      lineUserId: false,
      errors: []
    };
    
    // カレンダーIDチェック
    if (settings.calendarId) {
      try {
        const calendar = CalendarApp.getCalendarById(settings.calendarId);
        if (calendar) {
          results.calendarId = true;
        } else {
          results.errors.push('カレンダーIDが無効です');
        }
      } catch (error) {
        results.errors.push(`カレンダーアクセスエラー: ${error.message}`);
      }
    } else {
      results.errors.push('カレンダーIDが設定されていません');
    }
    
    // LINE設定チェック
    if (settings.lineChannelAccessToken) {
      results.lineChannelAccessToken = true;
    } else {
      results.errors.push('LINE Channel Access Tokenが設定されていません');
    }
    
    if (settings.lineUserId) {
      results.lineUserId = true;
    } else {
      results.errors.push('LINE User IDが設定されていません');
    }
    
    return results;
    
  } catch (error) {
    console.error('設定妥当性チェックエラー:', error);
    return {
      calendarId: false,
      lineChannelAccessToken: false,
      lineUserId: false,
      errors: [`検証エラー: ${error.message}`]
    };
  }
}

/**
 * 設定状況を表示
 */
function showSettingsStatus() {
  try {
    const validation = validateSettings();
    
    let message = '📊 設定状況\n\n';
    message += `📅 カレンダーID: ${validation.calendarId ? '✅ 正常' : '❌ 問題あり'}\n`;
    message += `🔑 LINE Access Token: ${validation.lineChannelAccessToken ? '✅ 設定済み' : '❌ 未設定'}\n`;
    message += `👤 LINE User ID: ${validation.lineUserId ? '✅ 設定済み' : '❌ 未設定'}\n`;
    
    if (validation.errors.length > 0) {
      message += '\n⚠️ エラー:\n';
      validation.errors.forEach(error => {
        message += `• ${error}\n`;
      });
    }
    
    SpreadsheetApp.getUi().alert('設定状況', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('設定状況表示エラー:', error);
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n\n${error.message}`);
  }
}

/**
 * カレンダーイベントをキーワードでフィルタリング
 */
function filterEventsByKeyword(events, keyword) {
  if (!keyword || keyword.trim() === '') {
    return events;
  }
  
  const lowerKeyword = keyword.toLowerCase();
  
  return events.filter(event => {
    return event.title.toLowerCase().includes(lowerKeyword) ||
           event.description.toLowerCase().includes(lowerKeyword) ||
           event.location.toLowerCase().includes(lowerKeyword);
  });
}

/**
 * ログ記録用関数
 */
function logActivity(action, details = '') {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const logMessage = `[${timestamp}] ${action}`;
  
  if (details) {
    console.log(`${logMessage}: ${details}`);
  } else {
    console.log(logMessage);
  }
  
  // スプレッドシートにログを記録する場合は以下のコメントアウトを解除
  /*
  try {
    const sheet = getOrCreateLogSheet();
    sheet.appendRow([timestamp, action, details]);
  } catch (error) {
    console.error('ログ記録エラー:', error);
  }
  */
}

/**
 * ログシートを取得または作成（オプション）
 */
function getOrCreateLogSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName('活動ログ');
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet('活動ログ');
    sheet.getRange('A1').setValue('日時');
    sheet.getRange('B1').setValue('アクション');
    sheet.getRange('C1').setValue('詳細');
    
    // ヘッダーのフォーマット
    sheet.getRange('A1:C1').setFontWeight('bold');
    sheet.getRange('A1:C1').setBackground('#e0e0e0');
    
    // 列幅調整
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 300);
  }
  
  return sheet;
}

/**
 * 現在のトリガー一覧を表示
 */
function showCurrentTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    if (triggers.length === 0) {
      SpreadsheetApp.getUi().alert('⏰ トリガー設定', '現在設定されているトリガーはありません。', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    let message = '⏰ 現在のトリガー設定\n\n';
    
    triggers.forEach((trigger, index) => {
      const handlerFunction = trigger.getHandlerFunction();
      const triggerSource = trigger.getTriggerSource();
      
      message += `${index + 1}. ${handlerFunction}\n`;
      
      if (triggerSource === ScriptApp.TriggerSource.CLOCK) {
        message += `   種類: 時間ベース\n`;
        // 詳細な時間情報は取得できないため、基本情報のみ表示
      } else {
        message += `   種類: ${triggerSource}\n`;
      }
      
      message += '\n';
    });
    
    SpreadsheetApp.getUi().alert('トリガー設定', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('トリガー一覧表示エラー:', error);
    SpreadsheetApp.getUi().alert(`❌ エラーが発生しました\n\n${error.message}`);
  }
} 