const fs = require('fs');
let c = fs.readFileSync('c:/Users/yhhoj/ReceivablePayableWebApp/app.js', 'utf8');
let lines = c.split('\n');
lines.splice(0, 10,
  'const partners = [',
  '  { code: "P001", name: "미래산업", manager: "김과장" },',
  '  { code: "P002", name: "부광물류", manager: "이대리" },',
  '  { code: "P003", name: "광주제조", manager: "박팀장" },',
  '];',
  '',
  'let receivables = [];',
  '',
  'let payables = [',
  '  { code: "P001", name: "미래산업", year: 2026, month: 4, purchase: 2800000, paid: 1000000, payDate: "2026-04-20", memo: "장비비", selected: false, decisionAmount: 1800000 },'
);
fs.writeFileSync('c:/Users/yhhoj/ReceivablePayableWebApp/app.js', lines.join('\n'));
