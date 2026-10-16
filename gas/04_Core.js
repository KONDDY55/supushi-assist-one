/** Pure logic, shared by GAS and Node tests. */
function diagnosticNg_(message) {
  const error = new Error(message);
  error.name = 'DiagnosticNG';
  throw error;
}

function columnNumber_(text) {
  const value = String(text || '').trim().toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(value)) diagnosticNg_('列はA、H、AAのような列記号1つで指定してください。');
  const result = value.split('').reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  if (result > 16384) diagnosticNg_('Excelの列上限を超えています。');
  return result;
}

function columnLetters_(number) {
  let out = '';
  for (let n = number; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + (n - 1) % 26) + out;
  return out;
}

function validateJob_(job) {
  ['inputSheet', 'outputSheet'].forEach(key => {
    if (!String(job[key] || '').trim()) diagnosticNg_('未設定: ' + key);
  });
  const columns = ['sourceStart', 'sourceEnd', 'targetStart', 'targetEnd', 'outputStart', 'outputEnd'];
  columns.forEach(key => { job[key] = columnNumber_(job[key]); });
  if (job.sourceStart > job.sourceEnd || job.targetStart > job.targetEnd || job.outputStart > job.outputEnd) diagnosticNg_('開始列が終了列より後です。');
  if (job.sourceEnd - job.sourceStart !== job.targetEnd - job.targetStart) diagnosticNg_('コピー元と貼り付け先の列数が一致しません。');
  ['csvStartRow', 'pasteStartRow', 'outputStartRow', 'outputEndRow'].forEach(key => {
    if (!Number.isInteger(job[key]) || job[key] < 1) diagnosticNg_('正の整数の行番号を設定してください: ' + key);
  });
  if (job.outputEndRow < job.outputStartRow) diagnosticNg_('出力行範囲が逆です。');
  if (typeof job.beforeHasHeader !== 'boolean' || typeof job.expectedHasHeader !== 'boolean') diagnosticNg_('CSVのヘッダー有無を指定してください。');
  ['beforeEncoding', 'expectedEncoding', 'outputEncoding'].forEach(key => {
    if (!['UTF-8', 'Shift_JIS'].includes(job[key])) diagnosticNg_('文字コードを指定してください: ' + key);
  });
  if (!['LF', 'CRLF'].includes(job.outputNewline) || typeof job.outputBom !== 'boolean') diagnosticNg_('CSV改行・BOMの設定が必要です。');
  if (job.outputBom && job.outputEncoding !== 'UTF-8') diagnosticNg_('BOMはUTF-8のみ使用できます。');
  return job;
}

function validateCsvRows_(rows, hasHeader) {
  if (!rows.length || rows.every(r => r.every(v => v === ''))) diagnosticNg_('CSVが空です。');
  if (rows.length - (hasHeader ? 1 : 0) > CONFIG.maxDataRows) diagnosticNg_('CSVデータが10,000行以上です。');
  const width = rows[0].length;
  if (rows.some(row => row.length !== width)) diagnosticNg_('CSVの列数が行ごとに異なります。');
  if (rows.length * width > CONFIG.maxCells) diagnosticNg_('試作のセル数上限を超えています。個別開発が必要です。');
}

function selectInput_(rows, job) {
  if (job.sourceEnd > rows[0].length || job.csvStartRow > rows.length) diagnosticNg_('コピー元範囲がCSVの外です。');
  return rows.slice(job.csvStartRow - 1).map(row => row.slice(job.sourceStart - 1, job.sourceEnd));
}

function compareCsv_(actual, expected) {
  if (actual.length !== expected.length) return {ok: false, reason: '出力行数が正解CSVと異なります。'};
  for (let r = 0; r < actual.length; r++) {
    if (actual[r].length !== expected[r].length) return {ok: false, reason: '出力列数が正解CSVと異なります。行: ' + (r + 1)};
    for (let c = 0; c < actual[r].length; c++) {
      if (String(actual[r][c]) !== String(expected[r][c])) return {ok: false, reason: '正解CSVと不一致。行: ' + (r + 1) + '、列: ' + (c + 1)};
    }
  }
  return {ok: true};
}

function serializeCsv_(rows, newline) {
  return rows.map(row => row.map(value => {
    const s = String(value);
    return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }).join(',')).join(newline === 'CRLF' ? '\r\n' : '\n');
}

function inspectExcelParts_(parts) {
  for (const part of parts) {
    if (/vbaProject\.bin|^xl\/(externalLinks|queryTables)\/|^xl\/connections\.xml$/i.test(part.name)) diagnosticNg_('マクロ・外部接続・クエリを含むExcelは個別開発へ進めてください。');
    if (/\.xml$/i.test(part.name) && /<(?:\w+:)?f(?:\s|>)/.test(part.text)) {
      // Conservative preflight only; runtime errors and golden CSV comparison remain mandatory.
      const formulas = part.text.match(/<(?:\w+:)?f(?:\s[^>]*)?>[\s\S]*?<\/(?:\w+:)?f>/g) || [];
      for (const f of formulas) {
        if (/\b(?:CUBE\w*|RTD|WEBSERVICE|FILTERXML|TEXTSPLIT)\s*\(|_xlfn\._xlws\.|_xludf\.|\[[^\]]+\.xlsx?\]/i.test(f)) diagnosticNg_('非対応候補の数式があります: ' + part.name + '。個別開発へ進めてください。');
      }
    }
  }
}
