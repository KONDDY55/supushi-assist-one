/** Form submission stages files, then converts and sends the result. */
function setupFormIntake() {
  const form = FormApp.openById(CONFIG.formId);
  if (!form.collectsEmail()) throw new Error('メールアドレス自動収集を有効にしてください。');
  const titles = form.getItems().map(item => ({id: String(item.getId()), title: item.getTitle(), type: String(item.getType())}));
  const required = [
    '現Excelをアップロードしてください（.xlsx）', '前CSVをアップロードしてください（.csv）', '後CSVをアップロードしてください（.csv）',
    '前CSVの貼り付け先となる現Excelのシート名', '後CSVの出力元となる現Excelのシート名',
    '前CSVからコピーする開始列', '前CSVからコピーする終了列',
    '前CSVを現Excelに貼り付ける開始列', '前CSVを現Excelに貼り付ける終了列',
    '現Excelから後CSVへ出力する開始列', '現Excelから後CSVへ出力する終了列'
  ];
  required.forEach((title, index) => {
    const matches = titles.filter(item => item.title === title);
    if (matches.length !== 1) throw new Error('質問がないか重複しています: ' + title);
    if (index < 3 && matches[0].type !== 'FILE_UPLOAD') throw new Error('ファイル提出形式ではありません: ' + title);
  });
  [CONFIG.sourceFolderId, CONFIG.inputFolderIds.excel, CONFIG.inputFolderIds.beforeCsv, CONFIG.inputFolderIds.afterCsv, CONFIG.outputFolderId, CONFIG.logFolderId]
    .forEach(id => DriveApp.getFolderById(id).getName());
  console.log(JSON.stringify(titles, null, 2));
  PropertiesService.getScriptProperties().setProperty('FORM_ITEMS', JSON.stringify(titles));
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'onIntakeSubmit')) {
    ScriptApp.newTrigger('onIntakeSubmit').forForm(form).onFormSubmit().create();
  }
  console.log('SETUP_OK: 質問11項目・フォルダ6件を確認し、フォーム送信トリガーを設定しました。');
}

function onIntakeSubmit(event) {
  if (!event || !event.response) throw new Error('フォーム本体の送信トリガーで実行してください。');
  try { return processFormQueue(); }
  catch (error) {
    reportQueueIssue_(event.response, error.message);
    throw error;
  }
}

/** Invoked by the FIFO worker while holding Script Lock. */
function processResponseLocked_(response) {
  let job;
  let intakeFailure;
  try {
    const id = response.getId();
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('RESPONSE_' + id)) throw new Error('受付済み案件の状態を確認してください: ' + id);
    job = {id: Utilities.getUuid(), email: response.getRespondentEmail(), responseId: id};
    const answers = {};
    const mapping = JSON.parse(props.getProperty('FORM_ITEMS') || '[]');
    if (!mapping.length) throw new Error('setupFormIntakeを先に実行してください。');
    response.getItemResponses().forEach(answer => {
      const configured = mapping.find(item => item.id === String(answer.getItem().getId()));
      if (configured) answers[configured.title] = answer.getResponse();
    });
    const oneFile = title => {
      const value = answers[title];
      if (!Array.isArray(value) || value.length !== 1) diagnosticNg_('ファイルを1つ提出してください: ' + title);
      return value[0];
    };
    const idShort = job.id;
    Object.assign(job, {
      excelFileId: oneFile('現Excelをアップロードしてください（.xlsx）'),
      beforeCsvFileId: oneFile('前CSVをアップロードしてください（.csv）'),
      expectedCsvFileId: oneFile('後CSVをアップロードしてください（.csv）'),
      inputSheet: answers['前CSVの貼り付け先となる現Excelのシート名'],
      outputSheet: answers['後CSVの出力元となる現Excelのシート名'],
      sourceStart: answers['前CSVからコピーする開始列'], sourceEnd: answers['前CSVからコピーする終了列'],
      targetStart: answers['前CSVを現Excelに貼り付ける開始列'], targetEnd: answers['前CSVを現Excelに貼り付ける終了列'],
      outputStart: answers['現Excelから後CSVへ出力する開始列'], outputEnd: answers['現Excelから後CSVへ出力する終了列']
    });
    applyDemoProfile_(job);
    const sourceKeys = ['excelFileId', 'beforeCsvFileId', 'expectedCsvFileId'];
    const sourceFolders = [CONFIG.inputFolderIds.excel, CONFIG.inputFolderIds.beforeCsv, CONFIG.inputFolderIds.afterCsv];
    sourceKeys.forEach((key, i) => {
      const original = DriveApp.getFileById(job[key]);
      job[key] = original.makeCopy(idShort + '-' + original.getName(), DriveApp.getFolderById(sourceFolders[i])).getId();
    });
    const file = DriveApp.getFolderById(CONFIG.sourceFolderId).createFile('job-' + idShort + '.json', JSON.stringify(job, null, 2), MimeType.PLAIN_TEXT);
    props.setProperty('RESPONSE_' + id, file.getId());
    console.log('受付済み・自動変換開始: ' + file.getId());
  } catch (error) {
    intakeFailure = error;
  }
  if (intakeFailure) return reportIntakeFailure_(job, intakeFailure);
  try {
    const result = runConversionLocked_(job);
    console.log(JSON.stringify(result));
    return result;
  } catch (error) {
    // Infrastructure failure before conversion's own result handler.
    return reportIntakeFailure_(job, error);
  }
}

function applyDemoProfile_(job) {
  const ids = [job.excelFileId, job.beforeCsvFileId, job.expectedCsvFileId];
  const hashes = ids.map(id => {
    const blob = DriveApp.getFileById(id).getBlob();
    if (blob.getBytes().length > CONFIG.maxBytes) diagnosticNg_('提出ファイルが10 MBを超えています。');
    return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, blob.getBytes())
      .map(byte => (byte & 255).toString(16).padStart(2, '0')).join('');
  });
  if (hashes.join(',') === 'd12af3b1eac81d4711740f248eb46b90,9b5ecd1335ebd94494b60dbabf508c78,498b20893f432a4c8192dddb8ea7a97e') {
    Object.assign(job, {
      profile: 'verified-demo-20261016',
      csvStartRow: 1, pasteStartRow: 1, outputStartRow: 1, outputEndRow: 11,
      beforeHasHeader: true, expectedHasHeader: true,
      beforeEncoding: 'UTF-8', expectedEncoding: 'UTF-8', outputEncoding: 'UTF-8',
      outputNewline: 'CRLF', outputBom: true
    });
  } else {
    diagnosticNg_('開始行・出力行・文字コードが未設定です。現在は検証済みデモ3ファイルにのみ自動設定できます。担当者と変換条件を確認してください。');
  }
}

function reportIntakeFailure_(job, error) {
  if (!job) throw error;
  const result = {jobId: job.id, status: error.name === 'DiagnosticNG' ? 'NG' : 'ERROR', step: 'フォーム受付・自動変換開始', reason: error.message, finishedAt: new Date().toISOString()};
  const props = PropertiesService.getScriptProperties();
  // Do not replace a completed conversion result if final logging failed.
  const saved = props.getProperty('RUN_' + job.id);
  if (saved) { console.error(error.message); return JSON.parse(saved); }
  props.setProperty('RUN_' + job.id, JSON.stringify(result));
  notifyCompletion_(job, result);
  props.setProperty('RUN_' + job.id, JSON.stringify(result));
  props.setProperty('RESPONSE_' + job.responseId, 'failed:' + job.id);
  DriveApp.getFolderById(CONFIG.logFolderId).createFile('diagnosis-' + job.id + '.json', JSON.stringify(result, null, 2), MimeType.PLAIN_TEXT);
  console.log(JSON.stringify(result));
  return result;
}

function runConfiguredJob() {
  const id = PropertiesService.getScriptProperties().getProperty('JOB_FILE_ID');
  if (!id) throw new Error('Script PropertiesにJOB_FILE_IDを設定してください。');
  return runConversion(JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')));
}
