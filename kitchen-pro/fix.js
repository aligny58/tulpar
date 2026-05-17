const fs = require('fs');
let c = fs.readFileSync('src/App.jsx', 'utf8');
const old = ' <label style={{...bSt("s",t),fontSize:11,cursor:"pointer"}}>\n          📥 {lang==="tr"?"İçe":"Import"}\n          <input type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={async e=>{';
const neu = '<button onClick={()=>importFileRef.current&&importFileRef.current.click()} style={{...bSt("s",t),fontSize:11}}>📥 {lang==="tr"?"İçe":"Import"}</button>\n        <input ref={importFileRef} type="file" accept=".csv,.xlsx,.xls" style={{display:"none"}} onChange={async e=>{';
if(c.includes(old)){
  c = c.replace(old, neu);
  fs.writeFileSync('src/App.jsx', c);
  console.log('OK');
} else {
  console.log('BULUNAMADI');
}