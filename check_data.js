const XLSX = require('xlsx');
const fs = require('fs');

const workbook = XLSX.readFile('VARIAÇÃO_IPCN.xlsx');
const sheet = workbook.Sheets['VAR_DOS PRODUTOS'];
if (!sheet) {
    console.log('Sheet not found');
    process.exit(1);
}

const data = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: null});
console.log('First 10 rows:');
data.slice(0, 10).forEach((row, i) => {
    console.log(`Row ${i}:`, row);
});

console.log('\nRows 50-55:');
data.slice(50, 56).forEach((row, i) => {
    console.log(`Row ${50+i}:`, row);
});
