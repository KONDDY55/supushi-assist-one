const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const iterator = items => { let i=0; return {hasNext:()=>i<items.length,next:()=>items[i++]}; };
function environment() {
  const groups = new Map();
  const byId = new Map();
  const moved = [];
  const context = vm.createContext({
    Utilities: {formatDate:()=> '20261015-120000'},
    DriveApp: {
      getFileById:id=>byId.get(id),
      getFolderById:id=>({getFiles:()=>iterator(groups.get(id)||[]),createFolder:()=>({getId:()=> 'archive'})})
    }
  });
  ['00_Config.example.js','04_Core.js','07_FileLifecycle.js'].forEach(f=>vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),context));
  function file(id,name,time,folder,trashed=false) {
    const f={getId:()=>id,getName:()=>name,getLastUpdated:()=>new Date(time),isTrashed:()=>trashed,getParents:()=>iterator([{getId:()=>folder}]),moveTo:()=>moved.push(id)};
    groups.set(folder,[...(groups.get(folder)||[]),f]); byId.set(id,f); return f;
  }
  return {context,groups,byId,moved,file};
}
test('更新日時が最新の対象拡張子を選び、他形式・古いファイル・ゴミ箱を除外する',()=>{
  const e=environment();
  e.file('old','old.csv',100,'input'); e.file('new','new.CSV',200,'input');
  e.file('other','x.xlsx',300,'input'); e.file('trash','trash.csv',400,'input',true);
  assert.equal(e.context.latestInput_('input','.csv').id,'new');
});
test('最新日時の同率と空フォルダでは推測せずNGにする',()=>{
  const e=environment(); e.file('a','a.csv',100,'input'); e.file('b','b.csv',100,'input');
  assert.throws(()=>e.context.latestInput_('input','.csv'),/同じ/);
  assert.throws(()=>e.context.latestInput_('empty','.csv'),/ありません/);
});
test('フォーム案件は別顧客の最新ファイルではなく、自分の3ファイルIDを使用する',()=>{
  const e=environment();
  const folders=['REPLACE_EXCEL','REPLACE_BEFORECSV','REPLACE_AFTERCSV'];
  folders.forEach((folder,i)=>{e.file('own'+i,'own'+(i===0?'.xlsx':'.csv'),100,folder);e.file('other'+i,'other'+(i===0?'.xlsx':'.csv'),200,folder);});
  const result=e.context.selectLatestInputs_({responseId:'real-response',excelFileId:'own0',beforeCsvFileId:'own1',expectedCsvFileId:'own2'});
  assert.equal(JSON.stringify(result.map(f=>f.id)),JSON.stringify(['own0','own1','own2']));
});
test('選択済み3ファイルだけを移管し、他ファイルは残す',()=>{
  const e=environment();
  e.file('old','old.csv',10,'input');
  const inputs=['excel','before','after'].map((id,i)=>{e.file(id,id+'.csv',100+i,'folder'+i);return e.context.latestInput_('folder'+i,'.csv');});
  const result={jobId:'case1'};
  e.context.archiveInputs_(inputs,result);
  assert.deepEqual(e.moved,['excel','before','after']);
  assert.equal(result.archiveFolderId,'archive');
});
test('処理中の更新を検知すると、1ファイルも移動しない',()=>{
  const e=environment(); e.file('a','a.csv',100,'input');
  const input=e.context.latestInput_('input','.csv');
  e.byId.get('a').getLastUpdated=()=>new Date(200);
  assert.throws(()=>e.context.archiveInputs_([input],{jobId:'case1'}),/更新/);
  assert.equal(e.moved.length,0);
});
test('移動途中の失敗では移管済みIDを保持する',()=>{
  const e=environment(); e.file('a','a.csv',100,'a'); e.file('b','b.csv',100,'b');
  e.byId.get('b').moveTo=()=>{throw Error('permission');};
  const result={jobId:'case1'};
  assert.throws(()=>e.context.archiveInputs_([e.context.latestInput_('a','.csv'),e.context.latestInput_('b','.csv')],result),/permission/);
  assert.equal(JSON.stringify(result.archivedFileIds),JSON.stringify(['a']));
});
