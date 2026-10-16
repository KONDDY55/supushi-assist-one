function readCsvFile_(file, encoding, header) {
  if (!/\.csv$/i.test(file.getName())) diagnosticNg_('CSV拡張子ではありません。');
  if (file.getSize() > CONFIG.maxBytes) diagnosticNg_('CSVが10 MBを超えています。');
  let text = file.getBlob().getDataAsString(encoding).replace(/^\uFEFF/, '');
  if (text.includes('\uFFFD') || text.includes('\u0000')) diagnosticNg_('文字コードを確認してください。');
  // Remove a terminal record separator only; embedded newlines are retained.
  text = text.replace(/(?:\r\n|\n|\r)$/, '');
  let rows;
  try { rows = Utilities.parseCsv(text); } catch (e) { diagnosticNg_('CSVの構文が不正です。'); }
  validateCsvRows_(rows, header);
  return rows;
}

function inspectExcelFile_(file) {
  if (!/\.xlsx$/i.test(file.getName()) || file.getSize() > CONFIG.maxBytes) diagnosticNg_('10 MB以内の.xlsxを提出してください。');
  let parts;
  try { parts = Utilities.unzip(file.getBlob().setContentType('application/zip')); }
  catch (e) { diagnosticNg_('Excelを読み取れません。パスワード保護または不正な形式を確認してください。'); }
  if (!parts.some(p => p.getName() === 'xl/workbook.xml')) diagnosticNg_('Excelブックの構造がありません。');
  let total = 0;
  const payload = parts.map(p => {
    total += p.getBytes().length;
    if (total > 50 * 1024 * 1024) diagnosticNg_('展開後のExcelが試作のサイズ上限を超えています。');
    return {name: p.getName(), text: /\.xml$/.test(p.getName()) ? p.getDataAsString('UTF-8') : ''};
  });
  inspectExcelParts_(payload);
}

function errorCells_(spreadsheetId) {
  const book = Sheets.Spreadsheets.get(spreadsheetId, {
    fields: 'sheets(properties(title),data(startRow,startColumn,rowData(values(effectiveValue))))',
    includeGridData: true
  });
  const errors = [];
  (book.sheets || []).forEach(sheet => (sheet.data || []).forEach(grid => (grid.rowData || []).forEach((row, r) => (row.values || []).forEach((cell, c) => {
    if (cell.effectiveValue && cell.effectiveValue.errorValue) errors.push(sheet.properties.title + '!' + columnLetters_((grid.startColumn || 0) + c + 1) + ((grid.startRow || 0) + r + 1));
  }))));
  return errors;
}

function inputRange_(sheet, data, job) {
  const endRow = job.pasteStartRow + data.length - 1;
  if (endRow > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), endRow - sheet.getMaxRows());
  if (job.targetEnd > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), job.targetEnd - sheet.getMaxColumns());
  const last = Math.max(sheet.getLastRow(), endRow);
  const range = sheet.getRange(job.pasteStartRow, job.targetStart, last - job.pasteStartRow + 1, job.targetEnd - job.targetStart + 1);
  if (range.getFormulas().some(row => row.some(Boolean))) diagnosticNg_('貼り付け・クリア範囲に数式があります。個別開発で確認してください。');
  range.clearContent();
  return "'" + sheet.getName().replace(/'/g, "''") + "'!" + columnLetters_(job.targetStart) + job.pasteStartRow;
}

function writeInput_(spreadsheetId, sheet, selected, job) {
  const range = inputRange_(sheet, selected, job);
  // Commit Spreadsheet service clears BEFORE the separate Sheets API write.
  SpreadsheetApp.flush();
  Sheets.Spreadsheets.Values.update({values: selected}, spreadsheetId, range, {valueInputOption: 'RAW'});
  const fullRange = range + ':' + columnLetters_(job.targetEnd) + (job.pasteStartRow + selected.length - 1);
  const received = Sheets.Spreadsheets.Values.get(spreadsheetId, fullRange, {valueRenderOption: 'UNFORMATTED_VALUE'}).values || [];
  const normalized = selected.map((row, r) => row.map((_, c) => received[r] && received[r][c] !== undefined ? String(received[r][c]) : ''));
  if (!compareCsv_(normalized, selected).ok) throw new Error('前CSVの書き込み検証に失敗しました。関数互換性の診断は未実施です。');
}
