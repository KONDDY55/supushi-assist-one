const test=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');const path=require('node:path');
function env(extra={}){const ctx=vm.createContext({CONFIG:{maxDataRows:9999,maxCells:500000},...extra});for(const f of ['04_Core.js','06_Conversion.js'])vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),ctx);return ctx;}
test('ヘッダーと先頭空白を除き、最初の全列空白の手前まで取得する',()=>{
 const ctx=env();const b=ctx.dataBlock_([['ID','値'],['',''],['001',''],['','0'],[' ','x'],['',''],['','']],true,'前CSV');assert.equal(b.startRow,3);assert.equal(b.endRow,5);assert.equal(b.rows.length,3);assert.equal(b.rows[0][0],'001');
});
test('ヘッダーなしは1行目から取得し、途中の空白以降のデータは捨てずNGにする',()=>{
 const ctx=env();assert.equal(ctx.dataBlock_([['0'],['']],false,'CSV').startRow,1);assert.throws(()=>ctx.dataBlock_([['header'],['']],true,'CSV'),/データがありません/);assert.throws(()=>ctx.dataBlock_([['a'],[''],['b']],false,'CSV'),/空白行の後にデータ/);
});
test('コピー元の指定列だけで入力範囲を判定しヘッダーをコピーしない',()=>{
 const ctx=env();const job={rowMode:'AUTO',sourceStart:2,sourceEnd:3,beforeHasHeader:true};const result=ctx.selectInput_([['ignored','ID','値'],['x','',''],['x','001','v'],['x','','']],job);assert.equal(JSON.stringify(result),JSON.stringify([['001','v']]));assert.equal(job.csvStartRow,3);assert.equal(job.csvEndRow,3);
});
test('空の物理CSV行を列幅に揃えるが非空行の列数不一致は拒否する',()=>{
 const ctx=env();const rows=[[''],['ID','値'],['001','x'],['']];ctx.validateCsvRows_(rows,true);assert.equal(rows[0].length,2);assert.throws(()=>ctx.validateCsvRows_([['a','b'],['x']],false),/列数/);
});
test('貼付先の既存データ先頭行を使い、空ならヘッダー直後へ貼る',()=>{
 const ctx=env();const job={inputHasHeader:true,targetStart:1,targetEnd:2};let rows=[['',''],['old','v']];const sheet={getLastRow:()=>rows.length+1,getMaxColumns:()=>2,getRange:(start,col,count,width)=>({getDisplayValues:()=>rows})};assert.equal(ctx.autoPasteStartRow_(sheet,job),3);rows=[];assert.equal(ctx.autoPasteStartRow_(sheet,job),2);job.inputHasHeader=false;assert.equal(ctx.autoPasteStartRow_(sheet,job),1);
});
test('出力範囲は毎回の再計算結果から検出し、ヘッダーを除く',()=>{
 const ctx=env();let rows=[['見出し'],['001'],['']];const sheet={getLastRow:()=>rows.length,getMaxColumns:()=>1,getRange:()=>({getDisplayValues:()=>rows})};const job={outputStart:1,outputEnd:1,outputHasHeader:true};assert.equal(ctx.autoOutputBlock_(sheet,job).rows.length,1);rows=[['見出し'],['001'],['002'],['']];assert.equal(ctx.autoOutputBlock_(sheet,job).endRow,3);
});
function pipeline(mismatch=false){
 const records=new Map();let selected,archived=false,notified=false;
 const input={getLastRow:()=>2,getLastColumn:()=>1,getMaxColumns:()=>1,getRange:()=>({getDisplayValues:()=>[['old']]})};
 const output={getLastRow:()=>3,getLastColumn:()=>1,getMaxColumns:()=>1,getRange:()=>({getDisplayValues:()=>[['出力見出し'],...selected.map(row=>[row[0]+'!']),['']]})};
 const book={getId:()=> 'work',setSpreadsheetLocale:()=>{},setSpreadsheetTimeZone:()=>{},getSheetByName:n=>n==='入力'?input:output,getSheets:()=>[input,output]};
 const mocks={PropertiesService:{getScriptProperties:()=>({getProperty:k=>records.get(k),setProperty:(k,v)=>records.set(k,v)})},MimeType:{GOOGLE_SHEETS:'sheet',PLAIN_TEXT:'text'},
 SpreadsheetApp:{openById:()=>book,flush:()=>{}},Sheets:{},Drive:{Files:{update:()=>{}}},
 DriveApp:{getFileById:()=>({makeCopy:()=>({getId:()=> 'backup'})}),getFolderById:()=>({createFile:()=>({getId:()=> 'csv',getUrl:()=> 'url'})})},
 Utilities:{sleep:()=>{},newBlob:()=>{let value='';const b={setDataFromString:t=>{value=t;return b;},setContentType:()=>b,setName:()=>b,getDataAsString:()=>value};return b;}},
 inspectExcelFile_:()=>{},selectLatestInputs_:()=>[0,1,2].map(id=>({id,file:{id,getBlob:()=>({})}})),readCsvFile_:file=>file.id===1?[['元見出し'],['001'],['']]:[['正解見出し'],[mismatch?'wrong':'001!'],['']],
 writeInput_:(id,sheet,data,job)=>{selected=data;assert.equal(job.pasteStartRow,2);},errorCells_:()=>[],verifySelectedInputs_:()=>{},archiveInputs_:()=>{archived=true;},notifyCompletion_:()=>{notified=true;}};
 const ctx=env(mocks);Object.assign(ctx,mocks);
 vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/02_Main.js'),'utf8'),ctx);
 const job={id:'auto',rowMode:'AUTO',inputSheet:'入力',outputSheet:'出力',sourceStart:'A',sourceEnd:'A',targetStart:'A',targetEnd:'A',outputStart:'A',outputEnd:'A',beforeHasHeader:true,expectedHasHeader:true,inputHasHeader:true,outputHasHeader:true,beforeEncoding:'AUTO',expectedEncoding:'AUTO',outputEncoding:'UTF-8',outputNewline:'CRLF',outputBom:true};
 return {result:ctx.runConversionLocked_(job),get selected(){return selected;},get archived(){return archived;},get notified(){return notified;}};
}
test('自動行モードの変換は双方のヘッダーを除外して照合し、成功時のみ出力・移管する',()=>{
 const p=pipeline();assert.equal(p.result.status,'OK');assert.equal(p.selected[0][0],'001');assert.equal(p.result.resolvedRows.outputStartRow,2);assert.equal(p.result.resolvedRows.outputEndRow,2);assert.equal(p.archived,true);assert.equal(p.notified,true);
});
test('自動行モードでも正解データ不一致ならCSVを作らず通知する',()=>{
 const p=pipeline(true);assert.equal(p.result.status,'NG');assert.match(p.result.reason,/不一致/);assert.equal(p.result.outputFileId,undefined);assert.equal(p.archived,false);assert.equal(p.notified,true);
});
