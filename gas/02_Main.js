/** Single shared workspace: serialized runs, backed up before replacement. */
function runConversion(rawJob) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('別案件の変換中です。完了後に実行してください。');
  try { return runConversionLocked_(rawJob); }
  finally { lock.releaseLock(); }
}

/** Called only while the shared Script Lock is held. */
function runConversionLocked_(rawJob) {
  const started = new Date();
  const job = JSON.parse(JSON.stringify(rawJob));
  if (!/^[A-Za-z0-9_-]{1,80}$/.test(job.id || '')) throw new Error('安全な案件IDが必要です。');
  const props = PropertiesService.getScriptProperties();
  const key = 'RUN_' + job.id;
  const previous = props.getProperty(key);
  if (previous) return JSON.parse(previous);
  const result = {jobId: job.id, status: 'RUNNING', step: '設定検査', startedAt: started.toISOString()};
  props.setProperty(key, JSON.stringify(result));
  try {
    validateJob_(job);
    result.step = '最新入力ファイル選択';
    const inputs = selectLatestInputs_(job);
    const originals = inputs.map(input => input.file);
    result.inputFiles = inputs.map(input => ({id: input.id, folderId: input.folderId, updatedAt: input.updatedAt}));
    result.step = '原本検査';
    inspectExcelFile_(originals[0]);
    result.inputFormats = {beforeCsv: {}, expectedCsv: {}};
    const before = readCsvFile_(originals[1], job.beforeEncoding, job.beforeHasHeader, result.inputFormats.beforeCsv);
    const expected = readCsvFile_(originals[2], job.expectedEncoding, job.expectedHasHeader, result.inputFormats.expectedCsv);
    result.outputFormat = {encoding: job.outputEncoding, newline: job.outputNewline, bom: job.outputBom};
    const selected = selectInput_(before, job);
    const cellCount = (job.outputEndRow - job.outputStartRow + 1) * (job.outputEnd - job.outputStart + 1);
    if (cellCount > CONFIG.maxCells) diagnosticNg_('出力範囲が試作のセル数上限を超えています。');
    result.step = 'バックアップ・Excel変換';
    result.backupId = DriveApp.getFileById(CONFIG.excelSpreadsheetId).makeCopy('backup-' + job.id, DriveApp.getFolderById(CONFIG.sourceFolderId)).getId();
    props.setProperty(key, JSON.stringify(result));
    Drive.Files.update({mimeType: MimeType.GOOGLE_SHEETS}, CONFIG.excelSpreadsheetId, originals[0].getBlob(), {fields: 'id'});
    const book = SpreadsheetApp.openById(CONFIG.excelSpreadsheetId);
    book.setSpreadsheetLocale('ja_JP');
    book.setSpreadsheetTimeZone('Asia/Tokyo');
    const input = book.getSheetByName(job.inputSheet);
    const output = book.getSheetByName(job.outputSheet);
    if (!input || !output) diagnosticNg_('指定された入力／出力シート名がExcelに存在しません。');
    if (book.getSheets().reduce((sum, s) => sum + s.getLastRow() * s.getLastColumn(), 0) > CONFIG.maxCells) diagnosticNg_('Excelが試作のセル数上限を超えています。');
    result.step = '前CSV貼り付け';
    // RAW prevents CSV values such as =1+1, 00123 and dates from becoming formulas/numbers.
    writeInput_(book.getId(), input, selected, job);
    SpreadsheetApp.flush();
    if (job.outputEndRow > output.getMaxRows() || job.outputEnd > output.getMaxColumns()) diagnosticNg_('出力範囲がシートの外です。');
    result.step = '再計算・互換性診断・正解照合';
    const outputRange = output.getRange(job.outputStartRow, job.outputStart, job.outputEndRow - job.outputStartRow + 1, job.outputEnd - job.outputStart + 1);
    let actual;
    let matched = false;
    let reason = '';
    // Two successive exact matches. Flush alone is not a calculation completion guarantee.
    let consecutive = 0;
    for (let attempt = 0; attempt < 6; attempt++) {
      SpreadsheetApp.flush();
      actual = outputRange.getDisplayValues();
      const errors = errorCells_(book.getId());
      const compared = compareCsv_(actual, expected);
      reason = errors.length ? '数式エラー: ' + errors.slice(0, 10).join(', ') : compared.reason;
      consecutive = !errors.length && compared.ok ? consecutive + 1 : 0;
      if (consecutive >= 2) { matched = true; break; }
      Utilities.sleep(1000);
    }
    if (!matched) diagnosticNg_(reason || '再計算の完了を確認できません。');
    result.step = 'CSV保存';
    verifySelectedInputs_(inputs);
    let text = serializeCsv_(actual, job.outputNewline);
    if (job.outputBom) text = '\uFEFF' + text;
    const blob = Utilities.newBlob('').setDataFromString(text, job.outputEncoding).setContentType('text/csv').setName('converted-' + job.id + '.csv');
    if (blob.getDataAsString(job.outputEncoding) !== text) diagnosticNg_('指定文字コードでは出力文字を保持できません。個別開発で文字コードを確認してください。');
    const csv = DriveApp.getFolderById(CONFIG.outputFolderId).createFile(blob);
    result.outputFileId = csv.getId();
    result.outputUrl = csv.getUrl();
    result.step = '入力3ファイルを05_logへ移管';
    archiveInputs_(inputs, result);
    result.status = 'OK';
    result.step = '完了';
  } catch (error) {
    result.status = result.outputFileId ? 'ARCHIVE_ERROR' : (error.name === 'DiagnosticNG' ? 'NG' : 'ERROR');
    result.reason = error.message;
    if (result.status === 'NG') result.nextAction = 'Gスプシでの自動変換を停止しました。個別開発をご相談ください。';
  } finally {
    result.finishedAt = new Date().toISOString();
    try {
      props.setProperty(key, JSON.stringify(result));
      notifyCompletion_(job, result);
      // Delivery failure must never rerun a successful conversion / archive.
      props.setProperty(key, JSON.stringify(result));
      DriveApp.getFolderById(result.archiveFolderId || CONFIG.logFolderId).createFile('diagnosis-' + job.id + '.json', JSON.stringify(result, null, 2), MimeType.PLAIN_TEXT);
    } finally { /* Lock belongs to the caller, including queue result handling. */ }
  }
  return result;
}
