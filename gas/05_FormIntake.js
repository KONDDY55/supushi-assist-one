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
  ].concat(formConditionFields_().map(field => field.title));
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
  console.log('SETUP_OK: 質問' + required.length + '項目・フォルダ6件を確認し、フォーム送信トリガーを設定しました。');
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
      const item = answer.getItem();
      const configured = mapping.find(entry => entry.id === String(item.getId()));
      answers[configured ? configured.title : item.getTitle()] = answer.getResponse();
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
    Object.assign(job, parseFormConditions_(answers));
    // Validate a copy: validateJob_ normalizes column letters for conversion.
    validateJob_(JSON.parse(JSON.stringify(job)));
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

/** Additional conditions are explicit for every customer; no file fingerprint gate. */
function formConditionFields_() {
  const yesNo = ['あり', 'なし'];
  const encodings = ['UTF-8', 'Shift_JIS'];
  return [
    {key: 'csvStartRow', title: '前CSVからコピーする開始行', help: '1以上の整数。ヘッダーも貼り付ける場合は1、1行目のヘッダーを除く場合は2。ここから最終行までコピーします。'},
    {key: 'pasteStartRow', title: '現Excelに貼り付ける開始行', help: '1以上の整数。ヘッダーを残して2行目からデータを貼る場合は2。指定列の開始行以降の既存値はクリアします。数式のある範囲には貼り付けできません。'},
    {key: 'outputStartRow', title: '現Excelから後CSVへ出力する開始行', help: '1以上の整数。出力にヘッダーを含める場合は、その行から指定してください。'},
    {key: 'outputEndRow', title: '現Excelから後CSVへ出力する終了行', help: '1以上の整数。最後に出力する行番号を指定します。出力範囲の行数・列数・内容を、提出する後CSV全体と一致させてください。'},
    {key: 'beforeHasHeader', title: '前CSVの1行目にヘッダーはありますか？', choices: yesNo, help: '列名の行がある場合は「あり」。開始行はこの回答とは別に指定してください。'},
    {key: 'expectedHasHeader', title: '正解の後CSVの1行目にヘッダーはありますか？', choices: yesNo, help: '照合時にヘッダーを自動で除外しません。出力範囲にも同じヘッダーを含めてください。'},
    {key: 'beforeEncoding', title: '前CSVの文字コード', choices: encodings, help: '提出する前CSVの保存時の文字コードを選択してください。'},
    {key: 'expectedEncoding', title: '正解の後CSVの文字コード', choices: encodings, help: '提出する後CSVの保存時の文字コードを選択してください。'},
    {key: 'outputEncoding', title: '変換後CSVの文字コード', choices: encodings, help: 'メールで受け取るCSVの文字コードを選択してください。'},
    {key: 'outputNewline', title: '変換後CSVの改行コード', choices: ['CRLF', 'LF'], help: 'WindowsのExcel向けは通常CRLFを選択します。'},
    {key: 'outputBom', title: '変換後CSVにBOMを付けますか？', choices: yesNo, help: 'UTF-8をExcelで開く場合は「あり」を推奨します。Shift_JISは必ず「なし」を選択してください。'}
  ];
}

function parseFormConditions_(answers) {
  const conditions = {};
  formConditionFields_().forEach(field => {
    const text = String(answers[field.title] === undefined ? '' : answers[field.title]).trim();
    if (!text) diagnosticNg_('変換条件の入力不備: 「' + field.title + '」が未回答です。フォームを開き直し、入力して再提出してください。');
    if (!field.choices) {
      if (!/^[1-9][0-9]*$/.test(text) || !Number.isSafeInteger(Number(text))) diagnosticNg_('変換条件の入力不備: 「' + field.title + '」は1以上の整数で指定してください。');
      conditions[field.key] = Number(text);
    } else {
      if (!field.choices.includes(text)) diagnosticNg_('変換条件の入力不備: 「' + field.title + '」は ' + field.choices.join(' / ') + ' から選択してください。');
      conditions[field.key] = ['beforeHasHeader', 'expectedHasHeader', 'outputBom'].includes(field.key) ? text === 'あり' : text;
    }
  });
  if (conditions.outputEndRow < conditions.outputStartRow) diagnosticNg_('変換条件の入力不備: 出力終了行は出力開始行以上にしてください。');
  if (conditions.outputBom && conditions.outputEncoding !== 'UTF-8') diagnosticNg_('変換条件の入力不備: Shift_JISを選んだ場合、BOMは「なし」にしてください。');
  return conditions;
}

/** Run once after sync. Re-running reuses questions and preserves responses. */
function setupIndependentDataForm() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const form = FormApp.openById(CONFIG.formId);
    const fields = formConditionFields_();
    // Preflight every existing item before adding or editing any question.
    fields.forEach(field => {
      const matches = form.getItems().filter(item => item.getTitle() === field.title);
      if (matches.length > 1 || (matches.length && String(matches[0].getType()) !== (field.choices ? 'LIST' : 'TEXT'))) throw new Error('変換条件の質問が重複、または形式が異なります: ' + field.title);
    });
    fields.forEach(field => {
      const existing = form.getItems().find(item => item.getTitle() === field.title);
      const item = existing ? (field.choices ? existing.asListItem() : existing.asTextItem()) : (field.choices ? form.addListItem() : form.addTextItem());
      item.setTitle(field.title).setHelpText(field.help).setRequired(true);
      if (field.choices) item.setChoiceValues(field.choices);
      else item.setValidation(FormApp.createTextValidation().requireTextMatchesPattern('^[1-9][0-9]*$').setHelpText('1以上の整数を入力してください。').build());
      // Keep conversion conditions before the final consent questions.
      const consent = form.getItems().find(entry => entry.getTitle() === '受付対象外の条件を確認しました');
      if (consent && item.getIndex() > consent.getIndex()) form.moveItem(item.getIndex(), consent.getIndex());
    });
    setupFormIntake();
    console.log('INDEPENDENT_DATA_FORM_OK: 任意の提出データを、フォームで指定された条件で受け付けます。条件不足・非対応・不一致は通知して停止します。');
  } finally { lock.releaseLock(); }
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
