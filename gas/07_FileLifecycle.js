/** Latest means last modified time; ties stop rather than choosing arbitrarily. */
function latestInput_(folderId, extension) {
  const files = DriveApp.getFolderById(folderId).getFiles();
  let latest = null;
  let latestTime = -1;
  let tied = false;
  while (files.hasNext()) {
    const file = files.next();
    if (file.isTrashed() || !file.getName().toLowerCase().endsWith(extension)) continue;
    const time = file.getLastUpdated().getTime();
    if (time > latestTime) { latest = file; latestTime = time; tied = false; }
    else if (time === latestTime) tied = true;
  }
  if (!latest) diagnosticNg_('入力フォルダに対象ファイルがありません: ' + extension);
  if (tied) diagnosticNg_('最新の更新日時が同じファイルが複数あります。入力を整理してください。');
  return {file: latest, folderId: folderId, id: latest.getId(), updatedAt: latestTime};
}

function selectLatestInputs_(job) {
  if (job.responseId && !job.responseId.startsWith('manual-demo-')) {
    // Queued customers have immutable per-response IDs; another customer's
    // newer files must never replace the current customer's inputs.
    const folders = [CONFIG.inputFolderIds.excel, CONFIG.inputFolderIds.beforeCsv, CONFIG.inputFolderIds.afterCsv];
    const extensions = ['.xlsx', '.csv', '.csv'];
    const inputs = [job.excelFileId, job.beforeCsvFileId, job.expectedCsvFileId].map((id, i) => {
      const file = DriveApp.getFileById(id);
      if (!file.getName().toLowerCase().endsWith(extensions[i])) diagnosticNg_('案件の入力ファイル形式が不正です。');
      return {file: file, id: id, folderId: folders[i], updatedAt: file.getLastUpdated().getTime()};
    });
    verifySelectedInputs_(inputs);
    return inputs;
  }
  const inputs = [
    latestInput_(CONFIG.inputFolderIds.excel, '.xlsx'),
    latestInput_(CONFIG.inputFolderIds.beforeCsv, '.csv'),
    latestInput_(CONFIG.inputFolderIds.afterCsv, '.csv')
  ];
  // A form job must not silently consume another customer's newer submission.
  if (job.responseId) {
    const expected = [job.excelFileId, job.beforeCsvFileId, job.expectedCsvFileId];
    if (inputs.some((input, i) => input.id !== expected[i])) diagnosticNg_('フォーム案件と最新ファイルが一致しません。別案件の混在を確認してください。');
  }
  return inputs;
}

function verifySelectedInputs_(inputs) {
  inputs.forEach(input => {
    const file = DriveApp.getFileById(input.id);
    let inFolder = false;
    const parents = file.getParents();
    while (parents.hasNext()) if (parents.next().getId() === input.folderId) inFolder = true;
    if (!inFolder || file.isTrashed() || file.getLastUpdated().getTime() !== input.updatedAt) diagnosticNg_('処理中に入力ファイルが更新・移動されました。再確認してください。');
  });
}

function archiveInputs_(inputs, result) {
  // Preflight all before moving any; partial failures remain explicit in the result.
  verifySelectedInputs_(inputs);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  const folder = DriveApp.getFolderById(CONFIG.logFolderId).createFolder(stamp + '-' + result.jobId);
  result.archiveFolderId = folder.getId();
  result.archivedFileIds = [];
  inputs.forEach(input => {
    input.file.moveTo(folder);
    result.archivedFileIds.push(input.id);
  });
}
