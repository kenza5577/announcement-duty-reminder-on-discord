/**
 * @fileoverview Discordにリマインダーを送信するGoogle Apps Script
 */

/**
 * スクリプトの初期設定を行います。
 * この関数は、スクリプトプロパティにDiscordのWebhook URLを保存するために、手動で一度だけ実行してください。
 */
function setConfiguration() {
  const properties = PropertiesService.getScriptProperties();
  // TODO: 'YOUR_DISCORD_WEBHOOK_URL'を実際のWebhook URLに置き換えてください。
  properties.setProperty('DISCORD_WEBHOOK_URL', 'YOUR_DISCORD_WEBHOOK_URL');
  // 初期の担当者インデックスを設定
  properties.setProperty('nextUserIndex', '0');
  Logger.log('設定が完了しました。');
}

/**
 * メインの実行関数。トリガーによって毎日この関数が呼び出されます。
 */
function main() {
  try {
    if (isReminderDay()) {
      Logger.log('本日はリマインド日です。処理を開始します。');
      const userData = getNextUser();

      if (userData && userData.userId) {
        const success = sendDiscordNotification(userData.userId);
        if (success) {
          // 通知が成功した場合のみ、次の担当者のインデックスを保存
          const properties = PropertiesService.getScriptProperties();
          properties.setProperty('nextUserIndex', userData.newIndex.toString());
          Logger.log(`通知成功。次の担当者インデックスを ${userData.newIndex} に更新しました。`);
        } else {
          Logger.log('Discordへの通知に失敗しました。インデックスは更新されません。');
        }
      } else {
        Logger.log('担当者が見つかりませんでした。スプレッドシートの設定を確認してください。');
      }
    } else {
      Logger.log('本日はリマインド日ではありません。');
    }
  } catch (e) {
    // エラーが発生した場合、ログに記録
    Logger.log(`エラーが発生しました: ${e.message}\nスタックトレース: ${e.stack}`);
  }
}

/**
 * スプレッドシートから直接スケジュールを読み込み、今日がリマインド対象日か判定する
 * @return {boolean} - リマインド対象日であればtrue
 */
function isReminderDay() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config');
    if (!sheet) {
      Logger.log('シート "Config" が見つかりません。');
      return false;
    }

    // C列の2行目からE列の最後までデータを取得し、開始日(C列)が空の行は除外する
    const scheduleData = sheet.getRange("C2:E" + sheet.getLastRow()).getValues().filter(row => row[0] !== "");

    if (scheduleData.length === 0) {
      Logger.log('スケジュールデータが見つかりませんでした。');
      return false;
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0); // 時刻をリセットして日付だけで比較

    // ループで各スケジュールをチェック
    for (const row of scheduleData) {
      const startDate = new Date(row[0]); // C列: 開始日
      const endDate = new Date(row[1]);   // D列: 終了日
      const frequency = parseInt(row[2], 10); // E列: 頻度

      // データが無効な行（日付ではない、頻度が0以下など）はスキップ
      if (isNaN(startDate.getTime()) || isNaN(endDate.getTime()) || isNaN(frequency) || frequency <= 0) {
        continue;
      }
      
      startDate.setHours(0, 0, 0, 0);
      endDate.setHours(0, 0, 0, 0);

      // 今日が期間内にあるかチェック
      if (today >= startDate && today <= endDate) {
        // 開始日からの経過日数を計算 (初日を0日目とする)
        const diffTime = today.getTime() - startDate.getTime();
        const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
        
        // 頻度に基づいて今日が通知日か判定 (経過日数を頻度で割った余りが0なら通知日)
        if (diffDays % frequency === 0) {
          Logger.log('リマインド対象日です。');
          return true; // 条件に一致するスケジュールが1つでもあればtrueを返して終了
        }
      }
    }
  } catch (e) {
    Logger.log('スケジュール判定中にエラーが発生しました: ' + e.message);
    return false; // エラー発生時は通知しない
  }
  return false; // すべてのスケジュールが条件に一致しなかった
}

/**
 * スプレッドシートから次の担当者を取得し、インデックスを更新します。
 * @return {{userId: string, newIndex: number}|null} 担当者のIDと次のインデックスを含むオブジェクト、またはnull。
 */
function getNextUser() {
  const properties = PropertiesService.getScriptProperties();
  const currentIndex = parseInt(properties.getProperty('nextUserIndex') || '0', 10);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) {
    Logger.log('シート "Config" が見つかりません。');
    return null;
  }

  const userIds = sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
   .map(row => row)
   .filter(id => id.toString().trim()!== ''); // 空白のセルを除外

  if (userIds.length === 0) {
    Logger.log('担当者リストが空です。');
    return null;
  }

  // 剰余演算子でインデックスが範囲外になるのを防ぎ、ローテーションを実現
  const userIndex = currentIndex % userIds.length;
  const userId = userIds[userIndex];
  const newIndex = (currentIndex + 1) % userIds.length;

  return { userId: userId, newIndex: newIndex };
}

/**
 * DiscordにWebhook経由で通知を送信します。
 * @param {string} userId メンションするユーザーのDiscord ID。
 * @return {boolean} 送信が成功すればtrue、失敗すればfalse。
 */
function sendDiscordNotification(userId) {
  const properties = PropertiesService.getScriptProperties();
  const webhookUrl = properties.getProperty('DISCORD_WEBHOOK_URL');

  if (!webhookUrl) {
    Logger.log('Webhook URLが設定されていません。setConfiguration()を実行してください。');
    return false;
  }

  // Discordのメンション形式 <@USER_ID> を使用
  const messageContent = `<@${userId}> 本日の告知ツイートをお願いします。`;

  const payload = {
    content: messageContent,
    username: '告知リマインダー', // Webhookのデフォルト名を上書き
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true, // HTTPエラーでスクリプトが停止しないようにする
  };

  const response = UrlFetchApp.fetch(webhookUrl, options);
  const responseCode = response.getResponseCode();

  if (responseCode >= 200 && responseCode < 300) {
    Logger.log('Discordへの通知に成功しました。');
    return true;
  } else {
    Logger.log(`Discordへの通知に失敗しました。レスポンスコード: ${responseCode}`);
    Logger.log(`レスポンス内容: ${response.getContentText()}`);
    return false;
  }
}
