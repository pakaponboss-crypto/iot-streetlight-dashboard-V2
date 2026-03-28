const XLSX = require('./node_modules/xlsx');
const path = require('path');
const filePath = path.join('C:', 'Users', 'pakap', 'Downloads', 'ข้อมูลการควบคุมเสาไฟIOT.xlsx');

const wb = XLSX.readFile(filePath);
const ws = wb.Sheets[wb.SheetNames[0]];
const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });

console.log('Total rows:', rows.length);
console.log('All headers:', JSON.stringify(Object.keys(rows[0])));

// Clean column names
const headers = Object.keys(rows[0]);
headers.forEach(h => {
  console.log('  Col:', JSON.stringify(h), '-> cleaned:', JSON.stringify(h.trim().replace(/^"|"$/g,'').trim()));
});

// Count unique values
const iotCol = headers.find(h => h.includes('ควบคุมiot'));
const apiCol = headers.find(h => h.includes('สถานะข้อมูลจากAPI'));
const contractCol = headers.find(h => h.includes('เลขที่สัญญา'));
const contractorCol = headers.find(h => h.includes('ผู้รับเหมา'));
const lampCol = headers.find(h => h.includes('จำนวนโคม'));

console.log('\nIOT col:', iotCol);
console.log('API col:', apiCol);

const iotVals = [...new Set(rows.map(r => r[iotCol]))];
const apiVals = [...new Set(rows.map(r => r[apiCol]))].slice(0, 10);
console.log('IOT unique values:', JSON.stringify(iotVals));
console.log('API unique values:', JSON.stringify(apiVals));

const totalIot = rows.filter(r => r[iotCol] === 'ได้').length;
const totalApi = rows.filter(r => r[apiCol] === 1 || r[apiCol] === '1').length;
const contracts = [...new Set(rows.map(r => r[contractCol]))];

console.log('\nTotal contracts:', contracts.length);
console.log('Total IOT ได้:', totalIot);
console.log('Total API connected:', totalApi);

// Sample lamp values
const lampVals = [...new Set(rows.map(r => r[lampCol]))].slice(0, 10);
console.log('จำนวนโคม sample values:', JSON.stringify(lampVals));
console.log('Sum จำนวนโคม:', rows.reduce((s,r) => s + (Number(r[lampCol]) || 0), 0));

// Sample by contract
const byContract = {};
for (const r of rows) {
  const cn = r[contractCol];
  if (!byContract[cn]) byContract[cn] = { contractor: r[contractorCol], poles: 0, iot: 0, api: 0, lamps: 0 };
  byContract[cn].poles++;
  if (r[iotCol] === 'ได้') byContract[cn].iot++;
  if (r[apiCol] === 1 || r[apiCol] === '1') byContract[cn].api++;
  byContract[cn].lamps += Number(r[lampCol]) || 0;
}
console.log('\nSample contracts (first 5):');
Object.entries(byContract).slice(0, 5).forEach(([cn, v]) => {
  console.log(`  ${cn}: poles=${v.poles} iot=${v.iot} api=${v.api} lamps=${v.lamps} contractor=${v.contractor}`);
});
