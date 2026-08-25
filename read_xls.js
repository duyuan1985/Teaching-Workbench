const fs = require('fs');
let XLSX;
try {
  XLSX = require('xlsx');
} catch (error) {
  console.error('XLSX_MODULE_MISSING');
  process.exit(2);
}
const file = 'D:/Downloads/教学安排表20260816115742.xls';
const workbook = XLSX.read(fs.readFileSync(file), { type: 'buffer', cellDates: true });
for (const name of workbook.SheetNames) {
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1, defval: '' });
  console.log(`SHEET ${name} ${rows.length}`);
  for (const row of rows.slice(0, 100)) {
    const values = row.map(v => String(v).trim()).filter(Boolean);
    if (values.length) console.log(values.join(' | '));
  }
}
