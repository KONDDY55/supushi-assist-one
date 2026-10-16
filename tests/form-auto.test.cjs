const test=require('node:test'); const assert=require('node:assert/strict');
const fs=require('node:fs');const path=require('node:path');const vm=require('node:vm');
function env(bad=false) {
  const files={excel:'d12af3b1eac81d4711740f248eb46b90',before:'9b5ecd1335ebd94494b60dbabf508c78',after:'498b20893f432a4c8192dddb8ea7a97e'};
  const ctx=vm.createContext({CONFIG:{maxBytes:10485760},DriveApp:{getFileById:id=>({getBlob:()=>({getBytes:()=>[id]})})},
    Utilities:{DigestAlgorithm:{MD5:'md5'},computeDigest:(algo,bytes)=> (bad?'00000000000000000000000000000000':files[bytes[0]]).match(/../g).map(v=>parseInt(v,16))}});
  ['04_Core.js','05_FormIntake.js'].forEach(f=>vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),ctx));return ctx;
}
test('同一デモ3ファイルの場合だけ行・文字コードを補完する',()=>{
 const ctx=env();const job={excelFileId:'excel',beforeCsvFileId:'before',expectedCsvFileId:'after',sourceStart:'B'};
 ctx.applyDemoProfile_(job);assert.equal(job.outputEndRow,11);assert.equal(job.pasteStartRow,1);assert.equal(job.beforeEncoding,'UTF-8');assert.equal(job.sourceStart,'B','フォーム回答は上書きしない');
});
test('別ファイルの行・文字コードを推測せず診断NGへ送る',()=>{
 const ctx=env(true);assert.throws(()=>ctx.applyDemoProfile_({excelFileId:'excel',beforeCsvFileId:'before',expectedCsvFileId:'after'}),/未設定/);
});
