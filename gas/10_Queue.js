/** The Form response store is the durable queue; submit triggers are wakeups. */
/** One-time recovery for the submission whose trigger failed at 14:45 JST. */
function setupAndRecoverSubmittedTest() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    // Include one minute before the observed trigger time (submission precedes trigger).
    const recoverySince = '2026-10-16T05:44:00.000Z';
    const current = props.getProperty('QUEUE_SINCE');
    if (!current || new Date(current).getTime() > new Date(recoverySince).getTime()) {
      props.setProperty('QUEUE_SINCE', recoverySince);
    }
  } finally { lock.releaseLock(); }
  setupProcessingQueue();
  const result = processFormQueue();
  console.log('RECOVERY_RESULT: ' + JSON.stringify(result));
  return result;
}

function setupProcessingQueue() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    setupFormIntake();
    const props = PropertiesService.getScriptProperties();
    if (!props.getProperty('QUEUE_SINCE')) props.setProperty('QUEUE_SINCE', new Date().toISOString());
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'processFormQueue')) {
      ScriptApp.newTrigger('processFormQueue').timeBased().everyMinutes(1).create();
    }
    console.log('QUEUE_SETUP_OK: 先着順・1件ずつ処理。フォーム送信時と1分ごとのトリガーで待機案件を確認します。');
  } finally { lock.releaseLock(); }
}

function orderedResponses_(responses) {
  return responses.slice().sort((a, b) => a.getTimestamp().getTime() - b.getTimestamp().getTime() || a.getId().localeCompare(b.getId()));
}

function processFormQueue() {
  const lock = LockService.getScriptLock();
  // Busy is normal, not a conversion error. Form responses remain durable.
  if (!lock.tryLock(1000)) return {status: 'BUSY', message: '前の案件が処理中です。提出済みの案件は待機しています。'};
  try {
    const props = PropertiesService.getScriptProperties();
    const since = props.getProperty('QUEUE_SINCE');
    if (!since) throw new Error('setupProcessingQueueを一度実行してください。');
    if (props.getProperty('QUEUE_BLOCKED')) return {status: 'BLOCKED', reason: props.getProperty('QUEUE_BLOCKED')};
    const responses = orderedResponses_(FormApp.openById(CONFIG.formId).getResponses(new Date(since)));
    for (const response of responses) {
      const key = 'QUEUE_' + response.getId();
      const stored = props.getProperty(key);
      const state = stored ? JSON.parse(stored) : null;
      if (state && state.status === 'DONE') continue;
      if (state && state.status === 'RUNNING') {
        // Timeout / process crash: do not overwrite shared workspace or resend.
        props.setProperty('QUEUE_BLOCKED', response.getId());
        console.error('QUEUE_BLOCKED: 前回実行の完了を確認できません: ' + response.getId());
        reportQueueIssue_(response, '前回実行が時間切れ・強制終了した可能性があり、完了状態が不明です。後続を停止しています。');
        return {status: 'BLOCKED', responseId: response.getId()};
      }
      props.setProperty(key, JSON.stringify({status: 'RUNNING', startedAt: new Date().toISOString()}));
      try {
        const result = processResponseLocked_(response);
        if (!result || !['OK', 'NG', 'ERROR', 'ARCHIVE_ERROR'].includes(result.status)) throw new Error('案件の完了結果がありません。');
        props.setProperty(key, JSON.stringify({status: 'DONE', jobId: result.jobId, resultStatus: result.status, notification: result.notification || null, finishedAt: new Date().toISOString()}));
        console.log('QUEUE_DONE: ' + response.getId() + ' ' + result.status);
        // One response per invocation keeps execution time bounded.
        return result;
      } catch (error) {
        props.setProperty('QUEUE_BLOCKED', response.getId());
        console.error('QUEUE_BLOCKED: ' + error.message);
        reportQueueIssue_(response, error.message);
        throw error;
      }
    }
    return {status: 'IDLE'};
  } finally { lock.releaseLock(); }
}
