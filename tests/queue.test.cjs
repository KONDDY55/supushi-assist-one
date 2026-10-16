const test=require('node:test');const assert=require('node:assert/strict');const vm=require('node:vm');const fs=require('node:fs');const path=require('node:path');
const response=(id,time)=>({getId:()=>id,getTimestamp:()=>new Date(time)});
function environment(responses, busy=false) {
 const records=new Map([['QUEUE_SINCE','2026-10-16T00:00:00Z']]);const calls=[];let released=0;let locked=false;
 const ctx=vm.createContext({CONFIG:{formId:'form'},console:{log:()=>{},error:()=>{}},
  LockService:{getScriptLock:()=>({tryLock:()=>{if(busy||locked)return false;locked=true;return true;},releaseLock:()=>{released++;locked=false;}})},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>records.get(k),setProperty:(k,v)=>records.set(k,v)})},
  FormApp:{openById:()=>({getResponses:()=>responses})},
  reportQueueIssue_:()=>{},
  processResponseLocked_:r=>{assert.equal(locked,true);calls.push(r.getId());return {jobId:'job-'+r.getId(),status:'OK',notification:{status:'SENT'}};}
 });
 vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas/10_Queue.js'),'utf8'),ctx);
 return {ctx,records,calls,released:()=>released};
}
test('提出時刻順で1件ずつ処理し、前のメール送信結果まで保存してから次を処理',()=>{
 const e=environment([response('later',200),response('first',100)]);
 assert.equal(e.ctx.processFormQueue().jobId,'job-first');assert.deepEqual(e.calls,['first']);
 assert.equal(JSON.parse(e.records.get('QUEUE_first')).notification.status,'SENT');
 assert.equal(e.ctx.processFormQueue().jobId,'job-later');assert.deepEqual(e.calls,['first','later']);
 assert.equal(e.ctx.processFormQueue().status,'IDLE');assert.equal(e.released(),3);
});
test('別の実行が動いている間は入力もメールも触らず待機する',()=>{
 const e=environment([response('first',100)],true);assert.equal(e.ctx.processFormQueue().status,'BUSY');assert.equal(e.calls.length,0);assert.equal(e.records.has('QUEUE_first'),false);
});
test('処理中に別の送信トリガーが起動しても同時処理しない',()=>{
 const e=environment([response('first',100),response('later',200)]);
 e.ctx.processResponseLocked_=r=>{e.calls.push(r.getId());assert.equal(e.ctx.processFormQueue().status,'BUSY');return {status:'OK',jobId:r.getId()};};
 e.ctx.processFormQueue();assert.deepEqual(e.calls,['first']);
});
test('実行時間切れ等で前の完了が不明なら後続を止める',()=>{
 const e=environment([response('first',100),response('later',200)]);e.records.set('QUEUE_first',JSON.stringify({status:'RUNNING'}));
 assert.equal(e.ctx.processFormQueue().status,'BLOCKED');assert.equal(e.calls.length,0);assert.equal(e.records.get('QUEUE_BLOCKED'),'first');
});
test('同一時刻は回答ID順で順序を固定し、完了済みを重複実行しない',()=>{
 const e=environment([response('b',100),response('a',100)]);e.ctx.processFormQueue();e.ctx.processFormQueue();e.ctx.processFormQueue();assert.deepEqual(e.calls,['a','b']);
});
