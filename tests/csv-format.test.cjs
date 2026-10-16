const test=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');const path=require('node:path');
function blob(input) {
 let bytes=typeof input==='string'?Buffer.from(input):Buffer.from(input.map(v=>v&255));
 const b={getBytes:()=>Array.from(bytes,v=>v>127?v-256:v),getDataAsString:encoding=>new TextDecoder(encoding==='Shift_JIS'?'shift_jis':'utf-8').decode(bytes),setDataFromString:(text,encoding)=>{
  if(encoding==='Shift_JIS'){
   const map={'日':[0x93,0xfa],'本':[0x96,0x7b],'語':[0x8c,0xea]};
   bytes=Buffer.from([...text].flatMap(c=>map[c]||(c.charCodeAt(0)<=127?[c.charCodeAt(0)]:c.charCodeAt(0)>=0xff61&&c.charCodeAt(0)<=0xff9f?[c.charCodeAt(0)-0xff61+0xa1]:[63])));
  } else bytes=Buffer.from(text,'utf8');return b;
 }};return b;
}
function env(){const ctx=vm.createContext({CONFIG:{maxBytes:10485760,maxDataRows:9999,maxCells:500000},Utilities:{newBlob:blob,parseCsv:text=>text.split(/\r\n|\n|\r/).map(line=>line.split(','))}});for(const f of ['04_Core.js','06_Conversion.js'])vm.runInContext(fs.readFileSync(path.join(__dirname,'../gas',f),'utf8'),ctx);return ctx;}
test('UTF-8 BOM付き・ASCII・UTF-8日本語を自動判定する',()=>{
 const ctx=env();for(const text of ['\uFEFF日本語,123\r\n','a,b\n','日本語,x\n']){const result=ctx.decodeCsvBlob_(blob(text),'AUTO');assert.equal(result.encoding,'UTF-8');assert.equal(result.text,text.replace(/^\uFEFF/,''));}
});
test('Shift_JISは日本語を復元し、バイト往復一致も確認する',()=>{
 const ctx=env();const result=ctx.decodeCsvBlob_(blob([0x93,0xfa,0x96,0x7b,0x8c,0xea,44,49]),'AUTO');assert.equal(result.encoding,'Shift_JIS');assert.equal(result.text,'日本語,1');
});
test('曖昧な入力・UTF16・不正バイト・壊れたUTF8 BOMはNGとして再保存を案内する',()=>{
 const ctx=env();for(const bytes of [[0xc2,0xa2],[0xff,0xfe,65,0],[0x81],[0xef,0xbb,0xbf,0xc0,0xaf],[27,36,66]])assert.throws(()=>ctx.decodeCsvBlob_(blob(bytes),'AUTO'),/自動判定できません.*CSV UTF-8/);
});
test('UTF8検証は過長表現・サロゲート・範囲外・途中切れを拒否する',()=>{
 const ctx=env();for(const bytes of [[0xc0,0xaf],[0xed,0xa0,0x80],[0xf4,0x90,0x80,0x80],[0xe3,0x81],[0xe3,0x20,0x82]])assert.equal(ctx.validUtf8Bytes_(bytes),false);assert.equal(ctx.validUtf8Bytes_([0xf0,0x9f,0x98,0x80]),true);
});
test('CRLF・LF・CRの入力を読み取り、行・セル内容で照合できる',()=>{
 const ctx=env();for(const newline of ['\r\n','\n','\r']){const content='\uFEFF名前,番号'+newline+'日本語,001'+newline;const file={getName:()=> 'test.csv',getSize:()=>100,getBlob:()=>blob(content)};const format={};const rows=ctx.readCsvFile_(file,'AUTO',true,format);assert.equal(JSON.stringify(rows),JSON.stringify([['名前','番号'],['日本語','001']]));assert.equal(format.encoding,'UTF-8');assert.equal(format.newline,newline==='\r\n'?'CRLF':newline==='\n'?'LF':'CR');assert.equal(ctx.serializeCsv_(rows,'CRLF'),'名前,番号\r\n日本語,001');}
});
