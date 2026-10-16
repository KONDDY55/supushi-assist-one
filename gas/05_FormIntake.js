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
  return [
    {key: 'beforeHasHeader', title: '前CSVの1行目にヘッダーはありますか？', choices: yesNo, help: '1行目に列名がある場合は「あり」。1行目を除いたデータだけを貼り付けます。'},
    {key: 'expectedHasHeader', title: '正解の後CSVの1行目にヘッダーはありますか？', choices: yesNo, help: '1行目に列名がある場合は「あり」。ヘッダーを除いたデータだけを正解として照合します。'},
    {key: 'inputHasHeader', title: '現Excelの貼り付け先シートの1行目にヘッダーはありますか？', choices: yesNo, help: '「あり」は1行目を残します。既存データがあればその先頭行から、空ならヘッダーの次の行から貼り付けます。'},
    {key: 'outputHasHeader', title: '現Excelの出力元シートの1行目にヘッダーはありますか？', choices: yesNo, help: '「あり」は1行目を除きます。変換後CSVはヘッダーを含まないデータのみで出力します。'}
  ];
}

function parseFormConditions_(answers) {
  // Ignore legacy format answers too: the public form has one output format.
  const conditions = {rowMode: 'AUTO', beforeEncoding: 'AUTO', expectedEncoding: 'AUTO', outputEncoding: 'UTF-8', outputNewline: 'CRLF', outputBom: true};
  formConditionFields_().forEach(field => {
    const text = String(answers[field.title] === undefined ? '' : answers[field.title]).trim();
    if (!text) diagnosticNg_('変換条件の入力不備: 「' + field.title + '」が未回答です。フォームを開き直し、入力して再提出してください。');
    if (!field.choices) {
      if (!/^[1-9][0-9]*$/.test(text) || !Number.isSafeInteger(Number(text))) diagnosticNg_('変換条件の入力不備: 「' + field.title + '」は1以上の整数で指定してください。');
      conditions[field.key] = Number(text);
    } else {
      if (!field.choices.includes(text)) diagnosticNg_('変換条件の入力不備: 「' + field.title + '」は ' + field.choices.join(' / ') + ' から選択してください。');
      conditions[field.key] = ['beforeHasHeader', 'expectedHasHeader', 'inputHasHeader', 'outputHasHeader'].includes(field.key) ? text === 'あり' : text;
    }
  });
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
    simplifyCsvFormatQuestions_(form);
    setupFormIntake();
    console.log('INDEPENDENT_DATA_FORM_OK: 任意の提出データを、フォームで指定された条件で受け付けます。条件不足・非対応・不一致は通知して停止します。');
  } finally { lock.releaseLock(); }
}

/** Remove retired technical questions, backing up their historical answers first. */
function simplifyCsvFormatQuestions_(form) {
  const retired = ['前CSVからコピーする開始行', '現Excelに貼り付ける開始行', '現Excelから後CSVへ出力する開始行', '現Excelから後CSVへ出力する終了行', '前CSVの文字コード', '正解の後CSVの文字コード', '変換後CSVの文字コード', '変換後CSVの改行コード', '変換後CSVにBOMを付けますか？'];
  const items = form.getItems().filter(item => retired.includes(item.getTitle()));
  if (items.length) {
    const ids = items.map(item => String(item.getId()));
    const responses = form.getResponses().map(response => ({
      responseId: response.getId(), submittedAt: response.getTimestamp().toISOString(),
      answers: response.getItemResponses().filter(answer => ids.includes(String(answer.getItem().getId())))
        .map(answer => ({title: answer.getItem().getTitle(), value: answer.getResponse()}))
    })).filter(response => response.answers.length);
    DriveApp.getFolderById(CONFIG.logFolderId).createFile('retired-csv-format-answers-' + Date.now() + '.json', JSON.stringify(responses, null, 2), MimeType.PLAIN_TEXT);
    items.sort((a, b) => b.getIndex() - a.getIndex()).forEach(item => form.deleteItem(item));
  }
  const noticeTitle = 'CSVの文字コード・出力形式について';
  const matches = form.getItems().filter(item => item.getTitle() === noticeTitle);
  if (matches.length > 1 || (matches.length && String(matches[0].getType()) !== 'SECTION_HEADER')) throw new Error('CSV形式の注意書きが重複、または形式が異なります。');
  const notice = matches.length ? matches[0].asSectionHeaderItem() : form.addSectionHeaderItem();
  notice.setTitle(noticeTitle).setHelpText('行番号・文字コード・改行コード・BOMの指定は不要です。ヘッダーは1行目として扱い、「あり」の場合は除外します。指定列の最初のデータ行から、対象列がすべて空白の最初の行の直前までを自動で処理します。途中の空白行より後にデータがある場合は、取りこぼし防止のためNGを通知します。貼り付け先が空の場合はヘッダーの次の行（ヘッダーなしは1行目）から貼り付けます。変換後CSVはヘッダーなしのデータのみです。提出CSVはシステムがUTF-8／Shift_JISを自動判定して読み込みます。判定できない場合は理由と対処方法をメールでお知らせします。変換後CSVはUTF-8（Excelで開きやすいBOM付き）・Windows向け改行でメール添付します。正解CSVとは文字コードや改行形式ではなく、行数・列数・各セルの内容を照合します。');
  if (notice.getIndex() !== 0) form.moveItem(notice.getIndex(), 0);
  console.log('CSV_FORMAT_FORM_OK: 行番号・文字コード・改行・BOMの質問を廃止し、自動判定の注意書きを設定しました。');
}

/** Existing installations can run this one function after syncing. */
function setupAutomaticCsvFormatForm() {
  setupIndependentDataForm();
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

/** One-time migration for existing public form. */
function setupAutomaticRowsForm() {
  setupIndependentDataForm();
}
