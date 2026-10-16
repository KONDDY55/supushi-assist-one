const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ctx = vm.createContext({});
['00_Config.example.js','04_Core.js'].forEach(f => vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),ctx));
const sample = () => JSON.parse(fs.readFileSync(path.join(__dirname,'../examples/job.example.json'),'utf8'));

test('列記号を検証し、コピー元と貼付先の列数を照合する',()=>{
  assert.equal(ctx.columnNumber_(' aa '),27);
  assert.equal(ctx.columnLetters_(16384),'XFD');
  assert.throws(()=>ctx.columnNumber_('A:H'));
  assert.throws(()=>ctx.columnNumber_('XFE'));
  const job=sample(); job.targetEnd='I';
  assert.throws(()=>ctx.validateJob_(job),/列数/);
});
test('不足する行番号と文字コードを推測して処理しない',()=>{
  const job=sample(); delete job.csvStartRow;
  assert.throws(()=>ctx.validateJob_(job),/行番号/);
  const bad=sample(); bad.outputEncoding='auto';
  assert.throws(()=>ctx.validateJob_(bad),/文字コード/);
});
test('開始行と列範囲を正しく切り出す',()=>{
  const job=sample(); job.sourceStart='B'; job.sourceEnd='C'; job.targetStart='C'; job.targetEnd='D';
  ctx.validateJob_(job);
  assert.equal(JSON.stringify(ctx.selectInput_([['a','b','c'],['1','001','x']],job)),JSON.stringify([['001','x']]));
});
test('10,000データ行以上と不揃いの列数をNGにする',()=>{
  ctx.validateCsvRows_(Array.from({length:10000},()=>['a']),true);
  assert.throws(()=>ctx.validateCsvRows_(Array.from({length:10001},()=>['a']),true),/10,000/);
  assert.throws(()=>ctx.validateCsvRows_([['a'],['a','b']],false),/列数/);
});
test('正解CSV比較で先頭ゼロ・空白・行数・列数を無視しない',()=>{
  assert.equal(ctx.compareCsv_([['001']],[['001']]).ok,true);
  assert.equal(ctx.compareCsv_([['1']],[['001']]).ok,false);
  assert.equal(ctx.compareCsv_([[' a']],[['a']]).ok,false);
  assert.equal(ctx.compareCsv_([['a']],[['a'],['']]).ok,false);
  assert.equal(ctx.compareCsv_([['a']],[['a','']]).ok,false);
});
test('CSVのカンマ・引用符・改行をエスケープする',()=>{
  assert.equal(ctx.serializeCsv_([['a,b','a"b','a\nb'],['001','','x']],'CRLF'),'"a,b","a""b","a\nb"\r\n001,,x');
});
test('非対応数式と外部接続をNGにし、通常数式は通す',()=>{
  ctx.inspectExcelParts_([{name:'xl/worksheets/sheet1.xml',text:'<f>SUM(A1:A3)</f>'}]);
  assert.throws(()=>ctx.inspectExcelParts_([{name:'xl/worksheets/sheet1.xml',text:'<f>_xlfn.FILTERXML(A1,B1)</f>'}]),/非対応/);
  assert.throws(()=>ctx.inspectExcelParts_([{name:'xl/connections.xml',text:''}]),/外部接続/);
});

test('診断NGでは変換API・CSV生成に進まず、NG結果と個別開発案内を保存する',()=>{
  const saved = new Map();
  const written = [];
  let released = false;
  const pipeline = vm.createContext({
    LockService: {getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{released=true;}})},
    PropertiesService: {getScriptProperties:()=>({getProperty:k=>saved.get(k),setProperty:(k,v)=>saved.set(k,v)})},
    DriveApp: {getFolderById:()=>({createFile:(...args)=>written.push(args)}),getFileById:()=>{throw Error('Unexpected file access');}},
    MimeType: {PLAIN_TEXT:'text/plain'}
    ,notifyCompletion_:()=>{}
  });
  ['00_Config.example.js','04_Core.js','02_Main.js'].forEach(f=>vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),pipeline));
  const bad=sample(); bad.sourceEnd='I';
  const result=pipeline.runConversion(bad);
  assert.equal(result.status,'NG');
  assert.match(result.nextAction,/個別開発/);
  assert.equal(result.outputFileId,undefined);
  assert.equal(written.length,1);
  assert.match(written[0][0],/^diagnosis-/);
  assert.equal(released,true);
  assert.equal(pipeline.runConversion(bad).status,'NG');
  assert.equal(written.length,1,'再実行で診断ファイルを重複生成しない');
});
