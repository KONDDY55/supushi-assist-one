/** Run manually once after adding send_mail scope. Does not send mail. */
function authorizeResultMail() {
  console.log('MAIL_AUTH_OK: メール送信権限を確認しました。残り宛先数: ' + MailApp.getRemainingDailyQuota());
}

/** Independent customer and owner delivery paths. Never throws send failures. */
function notifyCompletion_(job, result) {
  try {
    if (job.email) notifyResult_(job, result);
    else result.notification = {status: 'SKIPPED', reason: '通知先メールアドレスがありません。'};
  } catch (error) { result.notification = {status: 'FAILED', reason: error.message}; }
  if (result.status !== 'OK' || result.notification.status === 'FAILED') {
    try { notifyOwnerError_(job, result); }
    catch (error) {
      result.ownerNotification = {status: 'FAILED', reason: error.message};
      console.error('OWNER_NOTIFICATION_FAILED: ' + result.jobId + ' ' + error.message);
    }
  }
}

function ownerRemedy_(result) {
  const reason = (result.reason || '') + ' ' + (result.notification && result.notification.reason || '');
  const steps = [];
  if (result.status === 'NG') steps.push('フォーム回答のシート名・開始列・終了列、行番号・文字コードと提出原本を確認。正解CSVと出力を照合。非対応関数が原因なら個別開発の範囲を整理し、ユーザーとオンラインミーティングを調整。');
  if (result.status === 'ARCHIVE_ERROR') steps.push('出力CSVは作成済み。archivedFileIdsと05_logを確認して未移管ファイルだけ移動。変換は再実行しない。');
  if (/書き込み|貼り付け/.test(reason)) steps.push('入力シートの実値・数式・貼付範囲とflushの順序を確認。関数非対応とは断定しない。');
  if (/権限|permission|authorization|access/i.test(reason)) steps.push('GAS実行アカウントの認可と対象フォーム・Drive・スプシのアクセス権を確認。公開共有で回避しない。');
  if (/上限|quota|limit/i.test(reason)) steps.push('Googleの実行・メール送信上限と復旧時刻を確認。メールだけ失敗した場合は通知のみ再試行。');
  if (/完了.*不明|timeout|時間切れ|強制終了/.test(reason)) steps.push('Apps Scriptの実行履歴、RUN状態、CSV保存先、移管済みID、MAIL状態を照合。完了と送信の有無を確認するまでQUEUE_BLOCKEDを解除しない。');
  if (result.notification && result.notification.status === 'FAILED') steps.push('ユーザー向けメール送信失敗。MAIL_案件IDを確認し、SENDINGは送信結果不明なので自動再送しない。未送信と確認できたらretryConfiguredNotificationで通知のみ再試行。');
  if (!steps.length) steps.push('Apps Scriptの実行ログと診断JSONを確認。接続・API・設定を調査し、標準処理で解決困難なら個別開発とオンラインミーティングを案内。');
  return steps.map((text, i) => (i + 1) + '. ' + text).join('\n');
}

function notifyOwnerError_(job, result) {
  const props = PropertiesService.getScriptProperties();
  const key = 'OWNER_MAIL_' + result.jobId;
  const stored = props.getProperty(key);
  if (stored) {
    const state = JSON.parse(stored);
    if (state.status === 'SENT') { result.ownerNotification = state; return; }
    if (state.status === 'SENDING') throw new Error('オーナー通知の送信結果が不明です。重複防止のため再送を止めました。');
  }
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('メール送信枠不足でオーナー通知を送信できません。');
  const body = [
    'エラーが発生しました。', '案件ID: ' + result.jobId,
    '回答ID: ' + (job.responseId || 'なし'), 'ユーザー: ' + (job.email || '未取得'),
    '判定: ' + result.status, '発生工程: ' + result.step,
    '原因: ' + String(result.reason || '変換は成功。ユーザー向け通知でエラー').slice(0, 1000),
    '対応区分: ' + (errorGuidance_(result).responsibility === 'USER' ? 'ユーザーによる修正・再提出' : 'スプシアシストOneによる確認・対応'),
    'ユーザーへの案内: ' + errorGuidance_(result).remedy,
    'ユーザー通知: ' + JSON.stringify(result.notification || {}),
    '出力ファイルID: ' + (result.outputFileId || '未作成'),
    '移管済みID: ' + JSON.stringify(result.archivedFileIds || []),
    'バックアップID: ' + (result.backupId || 'なし'), '',
    '【対処方法】', ownerRemedy_(result), '',
    '実行履歴: https://script.google.com/home/projects/' + CONFIG.scriptId + '/executions',
    '診断ログ: https://drive.google.com/drive/folders/' + CONFIG.logFolderId
  ].join('\n');
  const state = {status: 'SENDING', recipient: CONFIG.ownerEmail, startedAt: new Date().toISOString()};
  props.setProperty(key, JSON.stringify(state));
  MailApp.sendEmail({to: CONFIG.ownerEmail, subject: '【要対応】スプシアシストOne エラー通知（' + result.jobId + '）', body: body, name: 'スプシアシストOne'});
  state.status = 'SENT'; state.sentAt = new Date().toISOString();
  props.setProperty(key, JSON.stringify(state));
  result.ownerNotification = state;
}

function reportQueueIssue_(response, reason) {
  const props = PropertiesService.getScriptProperties();
  const responseId = response.getId();
  let job = {id: 'queue-' + responseId, responseId: responseId, email: response.getRespondentEmail()};
  const intake = props.getProperty('RESPONSE_' + responseId);
  if (intake && !intake.startsWith('failed:')) {
    try { job = JSON.parse(DriveApp.getFileById(intake).getBlob().getDataAsString('UTF-8')); }
    catch (error) { console.error('QUEUE_JOB_READ_FAILED: ' + error.message); }
  }
  const result = {jobId: job.id, status: 'ERROR', step: '待ち行列の停止・復旧確認', reason: reason, finishedAt: new Date().toISOString()};
  // Separate incident ID preserves the original conversion and its mail state.
  const incident = Object.assign({}, job, {id: 'queue-incident-' + responseId});
  result.jobId = incident.id;
  notifyCompletion_(incident, result);
  props.setProperty('QUEUE_INCIDENT_' + responseId, JSON.stringify(result));
  console.log(JSON.stringify(result));
  return result;
}

/** Only explicit, actionable input defects are assigned to the customer. */
function errorGuidance_(result) {
  const reason = String(result.reason || '');
  const rules = [
    [/変換条件の入力不備|出力行範囲が逆|BOMはUTF-8のみ/, '原因欄に記載した質問と条件を確認してください。開始行・終了行は1以上の整数で、終了行は開始行以上にしてください。文字コードとBOMの組み合わせも確認してください。'],
    [/ファイルを1つ提出してください/, '現Excel・前CSV・正解の後CSVを、それぞれ1ファイルずつ提出してください。'],
    [/文字コード.*(?:自動判定|確認)/, '前CSVと正解の後CSVを、Excelの「名前を付けて保存」で「CSV UTF-8」形式に保存し直してください。フォームで文字コードを回答する必要はありません。'],
    [/列はA|Excelの列上限|開始列が終了列|コピー元と貼り付け先の列数/, '開始列・終了列をA、H、AAなどの列記号で入力し、開始列が終了列より前になるよう確認してください。コピー元と貼り付け先の列数を揃えてください。'],
    [/指定された入力／出力シート名がExcelに存在しません/, '提出Excel内のシート名を確認し、フォームの貼り付け先・出力元シート名を完全に一致させてください。'],
    [/コピー元範囲がCSVの外/, '前CSVの実際の列数を確認し、その範囲内でコピー開始列・終了列を指定してください。'],
    [/CSV拡張子ではありません|CSVの構文が不正|CSVの列数が行ごとに異なります|CSVが空です/, '内容のあるCSVファイルを提出してください。拡張子を.csvとし、各行の列数と引用符・区切り文字を確認してCSV形式で保存し直してください。'],
    [/CSVが10 MBを超え|10 MB以内の.xlsxを提出/, '現Excelは.xlsx形式、CSVは.csv形式で、それぞれ10 MB以内のファイルを提出してください。'],
    [/Excelを読み取れません。パスワード保護|Excelブックの構造がありません/, 'Excelが正常に開けるか確認し、パスワード保護のない.xlsx形式で保存し直してください。'],
    [/CSVデータが10,000行以上/, 'CSVのデータ行を9,999行以内に分割し、対応するExcelと正解CSVも同じ対象データで揃えてください。']
  ];
  if (result.status === 'NG') {
    for (const [pattern, remedy] of rules) {
      if (pattern.test(reason)) return {responsibility: 'USER', remedy: remedy};
    }
  }
  return {responsibility: 'OWNER', remedy: 'スプシアシストOneにて原因を確認し、対応いたします。確認・対応状況について、担当者から追ってご連絡いたします。現時点で追加の操作は不要です。個別開発やオンラインミーティングが必要な場合も、当方からご案内いたします。'};
}

function buildResultMail_(result) {
  const prefix = 'スプシアシストOne｜';
  if (result.status === 'OK') return {
    subject: prefix + 'CSV変換完了（' + result.jobId + '）',
    body: 'CSV変換が正常に完了しました。\n変換後CSVを添付します。\n提出された正解CSVとの一致を確認しました。\n\n案件ID: ' + result.jobId + '\n\n添付ファイルをご確認ください。このメールへの返信でご質問をお寄せいただけます。',
    attach: true
  };
  const cause = String(result.reason || '詳細を確認中です。').replace(/[\r\n]+/g, ' ').slice(0, 1000);
  const guidance = errorGuidance_(result);
  const userAction = guidance.responsibility === 'USER';
  const next = userAction
    ? guidance.remedy + '\n修正後、必要なファイルと入力内容を揃えて、以下のフォームから再度提出してください。\n' + CONFIG.formResponseUrl
    : guidance.remedy + (result.status === 'ERROR' ? '\n今回のエラーだけでExcel関数が非対応とは判断していません。' : '');
  return {
    subject: prefix + (userAction ? '提出内容の修正・再提出のお願い' : '処理エラー・当方にて対応いたします') + '（' + result.jobId + '）',
    body: 'CSV変換処理が正常完了しなかったため、変換後CSVは添付していません。\n\n案件ID: ' + result.jobId + '\n判定: ' + result.status + '\n発生した工程: ' + result.step + '\n原因: ' + cause + '\n\n【' + (userAction ? '修正方法と再提出のお願い' : '今後の対応') + '】\n' + next,
    attach: false
  };
}

function notifyResult_(job, result) {
  if (!['OK', 'NG', 'ERROR', 'ARCHIVE_ERROR'].includes(result.status)) throw new Error('処理が未完了のため通知できません。');
  // Resolve recipient from the original Google Form response, not editable job JSON.
  if (!job.responseId || job.responseId.startsWith('manual-demo-')) throw new Error('通知先をフォーム回答から確認できません。');
  const email = FormApp.openById(CONFIG.formId).getResponse(job.responseId).getRespondentEmail();
  if (!email || !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(email) || email !== job.email) throw new Error('自動収集メールアドレスと案件の通知先が一致しません。');
  const props = PropertiesService.getScriptProperties();
  const key = 'MAIL_' + job.id;
  const previous = props.getProperty(key);
  if (previous) {
    const state = JSON.parse(previous);
    if (state.status === 'SENT') { result.notification = state; return; }
    if (state.status === 'SENDING') throw new Error('前回のメール送信結果が不明です。重複防止のため自動再送せず、送信状況を確認してください。');
  }
  const content = buildResultMail_(result);
  const message = {to: email, subject: content.subject, body: content.body, name: 'スプシアシストOne'};
  if (content.attach) {
    if (!result.outputFileId) throw new Error('添付する変換後CSVがありません。');
    const file = DriveApp.getFileById(result.outputFileId);
    if (file.getSize() > 20 * 1024 * 1024) throw new Error('添付CSVがメール送信の試作上限20 MBを超えています。');
    message.attachments = [file.getBlob().setName(file.getName())];
  }
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('メール送信上限に達しました。変換処理は再実行せず、通知のみ再試行してください。');
  const state = {status: 'SENDING', recipient: email, startedAt: new Date().toISOString()};
  props.setProperty(key, JSON.stringify(state));
  // An exception during send may still mean delivery occurred. Retain SENDING.
  MailApp.sendEmail(message);
  state.status = 'SENT';
  state.sentAt = new Date().toISOString();
  props.setProperty(key, JSON.stringify(state));
  result.notification = state;
}

/** Explicit notification-only retry; does not convert or move files. */
function retryConfiguredNotification() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('JOB_FILE_ID');
    if (!id) throw new Error('JOB_FILE_IDを設定してください。');
    const job = JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8'));
    const saved = props.getProperty('RUN_' + job.id);
    if (!saved) throw new Error('処理結果がありません。');
    const result = JSON.parse(saved);
    notifyCompletion_(job, result);
    props.setProperty('RUN_' + job.id, JSON.stringify(result));
    console.log(JSON.stringify(result.notification));
  } finally { lock.releaseLock(); }
}
