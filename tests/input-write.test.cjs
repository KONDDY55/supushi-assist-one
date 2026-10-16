const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
function makeContext(bad=false) {
  const calls=[];
  const context=vm.createContext({
    SpreadsheetApp:{flush:()=>calls.push('flush')},
    Sheets:{Spreadsheets:{Values:{
      update:(body,id,range,options)=>{calls.push('write');assert.equal(options.valueInputOption,'RAW');},
      get:(id,range)=>{calls.push('read');assert.equal(range,"'入力CSV'!A1:B2");return {values:bad?[['wrong']]:[['注文ID','備考'],['00123']]};}
    }}}
  });
  ['04_Core.js','06_Conversion.js'].forEach(f=>vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),context));
  context.inputRange_=()=>{calls.push('clear');return "'入力CSV'!A1";};
  return {context,calls};
}
test('クリアをflushで確定してからRAW書込・全セル読戻し検証を行う',()=>{
  const {context,calls}=makeContext();
  context.writeInput_('id',{},[['注文ID','備考'],['00123','']],{targetEnd:2,pasteStartRow:1});
  assert.deepEqual(calls,['clear','flush','write','read']);
});
test('書込失敗は互換性NGと区別して例外にする',()=>{
  const {context}=makeContext(true);
  assert.throws(()=>context.writeInput_('id',{},[['注文ID','備考'],['00123','']],{targetEnd:2,pasteStartRow:1}),e=>e.name!=='DiagnosticNG' && /書き込み検証/.test(e.message));
});
