const test=require('node:test'); const assert=require('node:assert/strict');
const fs=require('node:fs');const path=require('node:path');const vm=require('node:vm');
function env(extra={}) {
 const ctx=vm.createContext({CONFIG:{maxBytes:10485760},...extra});
 ['04_Core.js','05_FormIntake.js'].forEach(f=>vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),ctx));return ctx;
}
function answers(ctx, overrides={}) {
 const values={csvStartRow:'2',pasteStartRow:'3',outputStartRow:'1',outputEndRow:'8',beforeHasHeader:'あり',expectedHasHeader:'あり',beforeEncoding:'Shift_JIS',expectedEncoding:'UTF-8',outputEncoding:'UTF-8',outputNewline:'CRLF',outputBom:'あり',...overrides};
 return Object.fromEntries(ctx.formConditionFields_().map(f=>[f.title,values[f.key]]));
}
test('独自データの行・ヘッダー・文字コードをフォーム回答から取得する',()=>{
 const ctx=env();const job=ctx.parseFormConditions_(answers(ctx));
 assert.equal(job.csvStartRow,2);assert.equal(job.pasteStartRow,3);assert.equal(job.outputEndRow,8);assert.equal(job.beforeEncoding,'Shift_JIS');assert.equal(job.outputBom,true);
});
test('条件不足、少数・不正整数、選択肢外の値、逆順、Shift_JISのBOMを拒否する',()=>{
 const ctx=env();
 for(const override of [{csvStartRow:''},{pasteStartRow:'1.5'},{outputEndRow:'0'},{csvStartRow:'2e3'},{csvStartRow:'9007199254740993'},{beforeEncoding:'auto'},{beforeHasHeader:'yes'},{outputStartRow:'9',outputEndRow:'8'},{outputEncoding:'Shift_JIS',outputBom:'あり'}]) assert.throws(()=>ctx.parseFormConditions_(answers(ctx,override)),/変換条件の入力不備/);
 const job=ctx.parseFormConditions_(answers(ctx,{outputEncoding:'Shift_JIS',outputBom:'なし',expectedHasHeader:'なし'}));assert.equal(job.outputBom,false);assert.equal(job.expectedHasHeader,false);
});
test('デモのハッシュに依存せず独自の3ファイルを受付・変換へ渡す',()=>{
 const records=new Map([['FORM_ITEMS','[]']]);const copied=[];let received;
 const folder={createFile:()=>({getId:()=> 'job-file'})};
 const ctx=env({console,CONFIG:{sourceFolderId:'root',inputFolderIds:{excel:'excel',beforeCsv:'before',afterCsv:'after'}},
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>records.get(k),setProperty:(k,v)=>records.set(k,v)})},Utilities:{getUuid:()=> 'unique-job'},MimeType:{PLAIN_TEXT:'text/plain'},
 DriveApp:{getFileById:id=>({getName:()=>id,makeCopy:(name,folder)=>{copied.push(name);return {getId:()=> 'copy-'+id};}}),getFolderById:()=>folder},
 runConversionLocked_:job=>{received=job;return {jobId:job.id,status:'OK'};}});
 const fields={...answers(ctx),'現Excelをアップロードしてください（.xlsx）':['custom.xlsx'],'前CSVをアップロードしてください（.csv）':['custom-before.csv'],'後CSVをアップロードしてください（.csv）':['custom-after.csv'],'前CSVの貼り付け先となる現Excelのシート名':'原票','後CSVの出力元となる現Excelのシート名':'結果','前CSVからコピーする開始列':'B','前CSVからコピーする終了列':'D','前CSVを現Excelに貼り付ける開始列':'C','前CSVを現Excelに貼り付ける終了列':'E','現Excelから後CSVへ出力する開始列':'A','現Excelから後CSVへ出力する終了列':'F'};
 records.set('FORM_ITEMS',JSON.stringify(Object.keys(fields).map((title,i)=>({id:String(i),title}))));
 const response={getId:()=> 'response',getRespondentEmail:()=> 'customer@example.com',getItemResponses:()=>Object.entries(fields).map(([title,value],i)=>({getItem:()=>({getId:()=>i,getTitle:()=>title}),getResponse:()=>value}))};
 assert.equal(ctx.processResponseLocked_(response).status,'OK');assert.equal(copied.length,3);assert.equal(received.excelFileId,'copy-custom.xlsx');assert.equal(received.email,'customer@example.com');assert.equal(received.sourceStart,'B');assert.equal(received.inputSheet,'原票');assert.equal(received.csvStartRow,2);assert.equal(received.pasteStartRow,3);
});
test('質問追加は冪等で、必須・選択肢・整数検証と対応表を設定する',()=>{
 const items=[];let setupCount=0;const records=new Map();
 function add(type){const obj={type,title:'',getType:()=>type,getTitle:()=>obj.title,getIndex:()=>items.indexOf(obj),setTitle:t=>{obj.title=t;return obj;},setHelpText:()=>obj,setRequired:v=>{obj.required=v;return obj;},setChoiceValues:v=>{obj.choices=v;return obj;},setValidation:v=>{obj.validation=v;return obj;},asListItem:()=>obj,asTextItem:()=>obj};items.push(obj);return obj;}
 const form={getItems:()=>items.slice(),addListItem:()=>add('LIST'),addTextItem:()=>add('TEXT')};
 const validation={requireTextMatchesPattern:()=>validation,setHelpText:()=>validation,build:()=> 'validation'};
 const ctx=env({console,LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})},FormApp:{openById:()=>form,createTextValidation:()=>validation}});
 ctx.setupFormIntake=()=>{setupCount++;records.set('FORM_ITEMS','updated');};ctx.setupIndependentDataForm();ctx.setupIndependentDataForm();
 assert.equal(items.length,11);assert.equal(setupCount,2);assert.equal(records.get('FORM_ITEMS'),'updated');assert.ok(items.every(i=>i.required));assert.ok(items.filter(i=>i.type==='TEXT').every(i=>i.validation==='validation'));assert.equal(items.filter(i=>i.type==='LIST').length,7);
});
