const fs = require('fs');
let c = fs.readFileSync('src/App.jsx', 'utf8');
// Kalan </label> tag'ini kaldır
const bad = '          }}/>\n        </label>';
const good = '          }}/>'; 
if(c.includes(bad)){
  c = c.replace(bad, good);
  fs.writeFileSync('src/App.jsx', c);
  console.log('OK');
} else {
  console.log('BULUNAMADI');
  const idx = c.indexOf('</label>');
  console.log('</label> konumu:', idx);
  console.log(JSON.stringify(c.substring(idx-100, idx+20)));
}