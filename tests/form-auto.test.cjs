const test=require('node:test'); const assert=require('node:assert/strict');
const fs=require('node:fs');const path=require('node:path');const vm=require('node:vm');
function env(extra={}) {
 const ctx=vm.createContext({CONFIG:{maxBytes:10485760},...extra});
 ['04_Core.js','05_FormIntake.js'].forEach(f=>vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),ctx));return ctx;
}
function answers(ctx, overrides={}) {
 const values={csvStartRow:'2',pasteStartRow:'3',outputStartRow:'1',outputEndRow:'8',beforeHasHeader:'あり',expectedHasHeader:'あり',inputHasHeader:'あり',outputHasHeader:'あり',beforeEncoding:'Shift_JIS',expectedEncoding:'UTF-8',outputEncoding:'UTF-8',outputNewline:'CRLF',outputBom:'あり',...overrides};
 return Object.fromEntries(ctx.formConditionFields_().map(f=>[f.title,values[f.key]]));
}
test('独自データの行・ヘッダーを取得し、入出力形式はシステム設定する',()=>{
 const ctx=env();const job=ctx.parseFormConditions_(answers(ctx));
 assert.equal(job.rowMode,'AUTO');assert.equal(job.csvStartRow,undefined);assert.equal(job.inputHasHeader,true);assert.equal(job.outputHasHeader,true);assert.equal(job.beforeEncoding,'AUTO');assert.equal(job.expectedEncoding,'AUTO');assert.equal(job.outputEncoding,'UTF-8');assert.equal(job.outputNewline,'CRLF');assert.equal(job.outputBom,true);
});
test('ヘッダーの条件不足・選択肢外を拒否し、旧形式回答は無視する',()=>{
 const ctx=env();
 for(const override of [{inputHasHeader:''},{outputHasHeader:'yes'},{beforeHasHeader:'yes'},{expectedHasHeader:''}]) assert.throws(()=>ctx.parseFormConditions_(answers(ctx,override)),/変換条件の入力不備/);
 const input=answers(ctx,{expectedHasHeader:'なし'});input['変換後CSVの文字コード']='Shift_JIS';input['変換後CSVにBOMを付けますか？']='なし';
 const job=ctx.parseFormConditions_(input);assert.equal(job.outputBom,true);assert.equal(job.outputEncoding,'UTF-8');assert.equal(job.expectedHasHeader,false);
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
 assert.equal(ctx.processResponseLocked_(response).status,'OK');assert.equal(copied.length,3);assert.equal(received.excelFileId,'copy-custom.xlsx');assert.equal(received.email,'customer@example.com');assert.equal(received.sourceStart,'B');assert.equal(received.inputSheet,'原票');assert.equal(received.rowMode,'AUTO');assert.equal(received.csvStartRow,undefined);assert.equal(received.inputHasHeader,true);
});
test('質問追加は冪等で、必須・選択肢・整数検証と対応表を設定する',()=>{
 const items=[];let setupCount=0;const records=new Map();const backups=[];
 function add(type){const obj={type,title:'',getId:()=>obj.title,getType:()=>type,getTitle:()=>obj.title,getIndex:()=>items.indexOf(obj),setTitle:t=>{obj.title=t;return obj;},setHelpText:()=>obj,setRequired:v=>{obj.required=v;return obj;},setChoiceValues:v=>{obj.choices=v;return obj;},setValidation:v=>{obj.validation=v;return obj;},asListItem:()=>obj,asTextItem:()=>obj,asSectionHeaderItem:()=>obj};items.push(obj);return obj;}
 const form={getItems:()=>items.slice(),addListItem:()=>add('LIST'),addTextItem:()=>add('TEXT'),addSectionHeaderItem:()=>add('SECTION_HEADER'),getResponses:()=>[],deleteItem:item=>items.splice(items.indexOf(item),1),moveItem:(from,to)=>items.splice(to,0,items.splice(from,1)[0])};
 ['前CSVの文字コード','正解の後CSVの文字コード','変換後CSVの文字コード','変換後CSVの改行コード','変換後CSVにBOMを付けますか？','前CSVからコピーする開始行','現Excelに貼り付ける開始行','現Excelから後CSVへ出力する開始行','現Excelから後CSVへ出力する終了行'].forEach(title=>add('LIST').setTitle(title));
 const validation={requireTextMatchesPattern:()=>validation,setHelpText:()=>validation,build:()=> 'validation'};
 const ctx=env({console,MimeType:{PLAIN_TEXT:'text/plain'},DriveApp:{getFolderById:()=>({createFile:(...args)=>backups.push(args)})},LockService:{getScriptLock:()=>({waitLock:()=>{},releaseLock:()=>{}})},FormApp:{openById:()=>form,createTextValidation:()=>validation}});
 ctx.setupFormIntake=()=>{setupCount++;records.set('FORM_ITEMS','updated');};ctx.setupIndependentDataForm();ctx.setupIndependentDataForm();
 assert.equal(items.length,5);assert.equal(backups.length,1);assert.equal(items[0].title,'CSVの文字コード・出力形式について');assert.equal(setupCount,2);assert.equal(records.get('FORM_ITEMS'),'updated');assert.ok(items.filter(i=>i.type!=='SECTION_HEADER').every(i=>i.required));assert.ok(items.filter(i=>i.type==='TEXT').every(i=>i.validation==='validation'));assert.equal(items.filter(i=>i.type==='LIST').length,4);
});
test('旧質問の回答を保存できない場合は質問を削除しない',()=>{
 let deleted=false;
 const item={getTitle:()=> '前CSVの文字コード',getId:()=>42};
 const form={getItems:()=>[item],getResponses:()=>[],deleteItem:()=>{deleted=true;}};
 const ctx=env({MimeType:{PLAIN_TEXT:'text/plain'},DriveApp:{getFolderById:()=>({createFile:()=>{throw Error('backup failed');}})}});
 assert.throws(()=>ctx.simplifyCsvFormatQuestions_(form),/backup failed/);assert.equal(deleted,false);
});
