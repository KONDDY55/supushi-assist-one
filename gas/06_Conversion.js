/** Strict UTF-8 validation rejects overlong encodings, surrogates and truncation. */
function validUtf8Bytes_(bytes) {
  for (let i = 0; i < bytes.length;) {
    const a = bytes[i++];
    if (a <= 0x7f) continue;
    let n, low = 0x80, high = 0xbf;
    if (a >= 0xc2 && a <= 0xdf) n = 1;
    else if (a >= 0xe0 && a <= 0xef) { n = 2; if (a === 0xe0) low = 0xa0; if (a === 0xed) high = 0x9f; }
    else if (a >= 0xf0 && a <= 0xf4) { n = 3; if (a === 0xf0) low = 0x90; if (a === 0xf4) high = 0x8f; }
    else return false;
    if (i + n > bytes.length || bytes[i] < low || bytes[i] > high) return false;
    i++;
    for (let j = 1; j < n; j++) if (bytes[i] < 0x80 || bytes[i++] > 0xbf) return false;
  }
  return true;
}

function csvNewline_(text) {
  const types = new Set((text.match(/\r\n|\r|\n/g) || []));
  return types.size > 1 ? 'MIXED' : types.has('\r\n') ? 'CRLF' : types.has('\n') ? 'LF' : types.has('\r') ? 'CR' : 'NONE';
}

function decodeCsvBlob_(blob, encoding) {
  if (encoding !== 'AUTO') return {text: blob.getDataAsString(encoding), encoding: encoding};
  const bytes = blob.getBytes().map(value => value & 255);
  const invalid = 'CSVの文字コードを自動判定できません。Excel等で「CSV UTF-8」形式に保存し直して再提出してください。';
  if (bytes.includes(0) || bytes.includes(0x1b) || (bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) diagnosticNg_(invalid);
  const bom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf;
  const content = bom ? bytes.slice(3) : bytes;
  const payload = Utilities.newBlob(content.map(value => value > 127 ? value - 256 : value));
  const utf8 = validUtf8Bytes_(content) ? payload.getDataAsString('UTF-8') : null;
  if (bom) {
    if (utf8 === null || utf8.includes('\uFFFD')) diagnosticNg_(invalid);
    return {text: utf8, encoding: 'UTF-8'};
  }
  let sjis = null;
  try {
    const text = payload.getDataAsString('Shift_JIS');
    const roundtrip = Utilities.newBlob('').setDataFromString(text, 'Shift_JIS').getBytes().map(value => value & 255);
    if (!text.includes('\uFFFD') && roundtrip.length === content.length && roundtrip.every((value, i) => value === content[i])) sjis = text;
  } catch (error) { /* Unsupported or invalid Shift_JIS is not guessed. */ }
  if (utf8 !== null && sjis !== null && utf8 !== sjis) diagnosticNg_('CSVの文字コードを自動判定できません。UTF-8とShift_JISの両方で読めるため判定が曖昧です。Excel等で「CSV UTF-8」形式に保存し直して再提出してください。');
  if (utf8 !== null && !utf8.includes('\uFFFD')) return {text: utf8, encoding: 'UTF-8'};
  if (sjis !== null) return {text: sjis, encoding: 'Shift_JIS'};
  diagnosticNg_(invalid);
}

function readCsvFile_(file, encoding, header, formatInfo) {
  if (!/\.csv$/i.test(file.getName())) diagnosticNg_('CSV拡張子ではありません。');
  if (file.getSize() > CONFIG.maxBytes) diagnosticNg_('CSVが10 MBを超えています。');
  const decoded = decodeCsvBlob_(file.getBlob(), encoding);
  let text = decoded.text.replace(/^\uFEFF/, '');
  if (formatInfo) Object.assign(formatInfo, {encoding: decoded.encoding, newline: csvNewline_(text)});
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

function autoPasteStartRow_(sheet, job) {
  const start = job.inputHasHeader ? 2 : 1;
  const last = sheet.getLastRow();
  if (last < start || job.targetStart > sheet.getMaxColumns()) return start;
  const width = Math.min(job.targetEnd, sheet.getMaxColumns()) - job.targetStart + 1;
  const rows = sheet.getRange(start, job.targetStart, last - start + 1, width).getDisplayValues();
  const offset = rows.findIndex(row => row.some(value => value !== ''));
  return offset < 0 ? start : start + offset;
}

function autoOutputBlock_(sheet, job) {
  if (job.outputEnd > sheet.getMaxColumns()) diagnosticNg_('出力範囲がシートの外です。');
  const last = sheet.getLastRow();
  const width = job.outputEnd - job.outputStart + 1;
  if (last * width > CONFIG.maxCells) diagnosticNg_('出力範囲が試作のセル数上限を超えています。');
  const rows = last ? sheet.getRange(1, job.outputStart, last, width).getDisplayValues() : [];
  return dataBlock_(rows, job.outputHasHeader, 'Excelの出力元シート');
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
