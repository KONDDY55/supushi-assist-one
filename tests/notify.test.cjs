const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
function env() {
  const records=new Map(); const messages=[];
  const ctx=vm.createContext({
    CONFIG:{formId:'form',scriptId:'script',formResponseUrl:'https://example.com/viewform',ownerEmail:'owner@example.com',logFolderId:'log'},console,
    FormApp:{openById:()=>({getResponse:()=>({getRespondentEmail:()=> 'demo@example.com'})})},
    PropertiesService:{getScriptProperties:()=>({getProperty:k=>records.get(k),setProperty:(k,v)=>records.set(k,v)})},
    DriveApp:{getFileById:()=>({getSize:()=>100,getName:()=> 'converted.csv',getBlob:()=>({setName:()=> 'CSV_BLOB'})})},
    MailApp:{getRemainingDailyQuota:()=>10,sendEmail:message=>messages.push(message)}
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/09_Notify.js'),'utf8'),ctx);
  return {ctx,records,messages};
}
const job={id:'test',responseId:'response',email:'demo@example.com'};
test('成功時は変換CSVを添付し、同じ案件を二重送信しない',()=>{
  const e=env(); const result={status:'OK',jobId:'test',outputFileId:'csv'};
  e.ctx.notifyResult_(job,result);e.ctx.notifyResult_(job,result);
  assert.equal(e.messages.length,1); assert.equal(e.messages[0].attachments[0],'CSV_BLOB');assert.equal(result.notification.status,'SENT');
});
test('NGは原因とオンライン会議案内を記載し、CSVを添付しない',()=>{
  const e=env();e.ctx.notifyResult_(job,{status:'NG',jobId:'test',reason:'数式エラー',step:'診断'});
  assert.match(e.messages[0].body,/数式エラー/);assert.match(e.messages[0].body,/オンラインミーティング/);assert.equal(e.messages[0].attachments,undefined);
});
test('システムエラーを関数非対応と断定しない',()=>{
  const e=env();const m=e.ctx.buildResultMail_({status:'ERROR',jobId:'x',reason:'通信失敗'});
  assert.match(m.body,/非対応とは判断していません/);assert.equal(m.attach,false);
});
test('編集された宛先と未完了結果では送信しない',()=>{
  const e=env();assert.throws(()=>e.ctx.notifyResult_({...job,email:'other@example.com'},{status:'OK'}),/一致しません/);
  assert.throws(()=>e.ctx.notifyResult_(job,{status:'RUNNING'}),/未完了/);assert.equal(e.messages.length,0);
});
test('送信結果不明と上限到達で送信せず、変換も再実行しない',()=>{
  const e=env();e.records.set('MAIL_test',JSON.stringify({status:'SENDING'}));
  assert.throws(()=>e.ctx.notifyResult_(job,{status:'NG',jobId:'test'}),/不明/);assert.equal(e.messages.length,0);
  e.records.clear();e.ctx.MailApp.getRemainingDailyQuota=()=>0;
  assert.throws(()=>e.ctx.notifyResult_(job,{status:'NG',jobId:'test'}),/上限/);assert.equal(e.messages.length,0);
});
test('失敗時はユーザーとオーナーの両方へ別メールを送り、オーナーへ対処法を記載',()=>{
 const e=env();const result={jobId:'test',status:'NG',step:'診断',reason:'数式エラー'};
 e.ctx.notifyCompletion_(job,result);e.ctx.notifyCompletion_(job,result);
 assert.equal(e.messages.length,2);assert.equal(e.messages[1].to,'owner@example.com');assert.match(e.messages[1].body,/対処方法/);assert.match(e.messages[1].body,/オンラインミーティング/);assert.equal(e.messages[1].attachments,undefined);
});
test('ユーザー送信が失敗しても、オーナー通知は独立して試行する',()=>{
 const e=env();e.ctx.MailApp.sendEmail=m=>{if(m.to==='demo@example.com')throw Error('delivery failed');e.messages.push(m);};
 const result={jobId:'test',status:'ERROR',step:'貼り付け',reason:'書き込み検証失敗'};e.ctx.notifyCompletion_(job,result);
 assert.equal(result.notification.status,'FAILED');assert.equal(e.messages.length,1);assert.equal(result.ownerNotification.status,'SENT');assert.match(e.messages[0].body,/flush/);
});
test('変換成功・添付通知成功ではオーナーへエラーメールしない',()=>{
 const e=env();e.ctx.notifyCompletion_(job,{jobId:'test',status:'OK',outputFileId:'csv'});assert.equal(e.messages.length,1);
});
test('送信枠不足は双方の通知失敗を記録して、元の変換結果を保持する',()=>{
 const e=env();e.ctx.MailApp.getRemainingDailyQuota=()=>0;const result={jobId:'test',status:'ERROR',reason:'test'};e.ctx.notifyCompletion_(job,result);
 assert.equal(result.status,'ERROR');assert.equal(result.notification.status,'FAILED');assert.equal(result.ownerNotification.status,'FAILED');assert.equal(e.messages.length,0);
});
test('入力不備は具体的な修正方法とフォーム再提出を案内する',()=>{
 const e=env();
 for(const reason of ['コピー元と貼り付け先の列数が一致しません。','指定された入力／出力シート名がExcelに存在しません。','CSVの列数が行ごとに異なります。','コピー元範囲がCSVの外です。']) {
  const m=e.ctx.buildResultMail_({jobId:'x',status:'NG',reason});
  assert.match(m.subject,/再提出/);assert.match(m.body,/修正後/);assert.match(m.body,/viewform/);assert.equal(m.attach,false);
 }
});
test('内部設定不足・非対応関数・照合不一致・未知の原因は当方で対応する',()=>{
 const e=env();
 for(const reason of ['開始行・出力行・文字コードが未設定です。現在は検証済みデモ3ファイルにのみ自動設定できます。','非対応関数があります。','正解CSVと不一致。行: 2、列: 1','原因未確定']) {
  const m=e.ctx.buildResultMail_({jobId:'x',status:'NG',reason});
  assert.match(m.body,/スプシアシストOneにて/);assert.match(m.body,/追ってご連絡/);assert.doesNotMatch(m.body,/再度提出|候補日時|ご返信ください/);
 }
});
test('システム・保管エラーは入力に似た原因でもユーザーへ再提出させない',()=>{
 const e=env();
 for(const status of ['ERROR','ARCHIVE_ERROR']) {
  const m=e.ctx.buildResultMail_({jobId:'x',status,reason:'CSVが空です。'});
  assert.match(m.body,/追ってご連絡/);assert.doesNotMatch(m.body,/再度提出/);
 }
});
test('入力不備でもオーナーへ区分と具体的な修正方法を別送する',()=>{
 const e=env();e.ctx.notifyCompletion_(job,{jobId:'test',status:'NG',reason:'コピー元範囲がCSVの外です。'});
 assert.equal(e.messages.length,2);assert.match(e.messages[1].body,/ユーザーによる修正・再提出/);assert.match(e.messages[1].body,/実際の列数/);
});
test('独自データの条件不足・文字コード不備はユーザーへ修正と再提出を案内する',()=>{
 const e=env();
 for(const reason of ['変換条件の入力不備: 「前CSVからコピーする開始行」が未回答です。','変換条件の入力不備: Shift_JISを選んだ場合、BOMは「なし」にしてください。','文字コードを確認してください。']) {
  const mail=e.ctx.buildResultMail_({jobId:'x',status:'NG',reason});assert.match(mail.subject,/再提出/);assert.match(mail.body,/再度提出/);assert.match(mail.body,new RegExp(reason.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')));
 }
});
