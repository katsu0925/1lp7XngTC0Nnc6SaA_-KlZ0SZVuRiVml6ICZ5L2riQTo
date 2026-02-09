/**
 * ★完全統合版：トリガー復旧
 * ユーザー様の記憶にある設定 ＋ ログに残っていた設定 をすべて復元します
 */
function FULL_RESTORE_ALL() {
  // 1. 既存トリガーをクリア（重複防止）
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==================================================
  // A. ご提示いただいた設定 (Month, Daily, Hourly, OnChange)
  // ==================================================

  // 毎月1日 0:15 (月次確定処理)
  ScriptApp.newTrigger('freezePastMonthsValues_batch')
    .timeBased().onMonthDay(1).atHour(0).nearMinute(15).inTimezone('GMT+9').create();

  // 毎日 0:30 (閾値判定)
  ScriptApp.newTrigger('stampByThreshold')
    .timeBased().everyDays(1).atHour(0).nearMinute(30).inTimezone('GMT+9').create();

  // 1時間ごと (画像掃除)
  ScriptApp.newTrigger('cleanupOldAITaskImages')
    .timeBased().everyHours(1).inTimezone('GMT+9').create();

  // 1時間ごと (画像監視)
  ScriptApp.newTrigger('monitorAndDeleteUnusedImages')
    .timeBased().everyHours(1).inTimezone('GMT+9').create();

  // 変更時 (メーラー・在庫連携)
  ScriptApp.newTrigger('handleChange_Mailer')
    .forSpreadsheet(ss).onChange().create();

  ScriptApp.newTrigger('handleChange_Inventory')
    .forSpreadsheet(ss).onChange().create();

  ScriptApp.newTrigger('processPendingKeywordRows')
    .timeBased().everyMinutes(1).create();

  // ==================================================
  // B. ログから判明した「足りない設定」を追加
  // ==================================================

  // 1時間ごと (リスト更新・分析) ※ログで頻繁に動いていました
  ScriptApp.newTrigger('generateCompletionList')
    .timeBased().everyHours(1).inTimezone('GMT+9').create();

  ScriptApp.newTrigger('buildWorkAnalysis')
    .timeBased().everyHours(1).inTimezone('GMT+9').create();

  // 毎日 3時〜4時 (報酬計算・欠番確認)
  ScriptApp.newTrigger('updateRewardsNoFormula')
    .timeBased().everyDays(1).atHour(3).inTimezone('GMT+9').create();

  ScriptApp.newTrigger('出力_欠番確認')
    .timeBased().everyDays(1).atHour(3).inTimezone('GMT+9').create();

  // 毎日 4時〜5時 (在庫日数計算)
  ScriptApp.newTrigger('recalcZaikoNissu')
    .timeBased().everyDays(1).atHour(4).inTimezone('GMT+9').create();

  // ==================================================
  
  Browser.msgBox('全てのトリガー（ご提示分＋ログ分析分）を復旧しました。');
}