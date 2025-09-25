function setupTriggers() {
  const lock = LockService.getScriptLock();
  lock.tryLock(3000);
  try {
    // 既存トリガーを一旦削除
    ScriptApp.getProjectTriggers().forEach((t) => {
      const fn = t.getHandlerFunction();
      if (
        fn === 'cron_1min' ||
        fn === 'run_ProcessInbox_S2002' ||
        fn === 'run_ProcessInbox_S2010_P1' ||
        fn === 'run_ProcessInbox_AllForms'
      ) {
        ScriptApp.deleteTrigger(t);
      }
    });

    // 共通フォーム Intake ルーター
    ScriptApp.newTrigger('run_ProcessInbox_AllForms').timeBased().everyMinutes(1).create();

    // Gドライブ整理（既存）
    ScriptApp.newTrigger('cron_1min').timeBased().everyMinutes(5).create();

    Logger.log('Triggers set: run_ProcessInbox_AllForms, cron_1min');
  } finally {
    lock.releaseLock();
  }
}

function removeAllTriggers() {
  ScriptApp.getProjectTriggers().forEach((t) => ScriptApp.deleteTrigger(t));
  Logger.log('all triggers removed.');
}

/**
 * トリガーの健全性チェック（存在しなければ自動再作成）
 * まずエディタから1回実行して認可（Gmail/Drive/Spreadsheet）を付与してください。
 */
function ensureTriggers() {
  const have = new Set(ScriptApp.getProjectTriggers().map((t) => t.getHandlerFunction()));
  let created = 0;
  if (!have.has('run_ProcessInbox_AllForms')) {
    ScriptApp.newTrigger('run_ProcessInbox_AllForms').timeBased().everyMinutes(1).create();
    created++;
  }
  if (!have.has('cron_1min')) {
    ScriptApp.newTrigger('cron_1min').timeBased().everyMinutes(5).create();
    created++;
  }
  Logger.log('ensureTriggers: have=%s created=%s', JSON.stringify(Array.from(have)), created);
}

/** 現在のトリガー一覧をログ出力（デバッグ用） */
function debug_listTriggers() {
  const list = ScriptApp.getProjectTriggers().map((t) => ({
    fn: t.getHandlerFunction(),
    type: t.getEventType && t.getEventType(),
  }));
  Logger.log(JSON.stringify(list));
}

/** 認可テスト用に1回だけS2002処理を実行 */
function debug_run_S2002_once() {
  run_ProcessInbox_S2002();
}
