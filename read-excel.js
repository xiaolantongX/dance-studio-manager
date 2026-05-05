const XLSX = require('xlsx');
const path = require('path');

// 使用 process.cwd 获取当前目录
const basePath = process.argv[2] || '../业绩';
const fileName = process.argv[3] || '2025 年业绩.xlsx';

console.log('Reading from:', basePath, fileName);

try {
  const filePath = path.join(basePath, fileName);
  console.log('Full path:', filePath);

  const wb = XLSX.readFile(filePath);
  console.log('\n=== Sheet 名称 ===');
  console.log(wb.SheetNames);

  console.log('\n=== 每个 Sheet 的数据量 ===');
  wb.SheetNames.forEach(sheet => {
    const ws = wb.Sheets[sheet];
    const data = XLSX.utils.sheet_to_json(ws);
    console.log(`Sheet: ${sheet} - ${data.length} 条记录`);
    if (data.length > 0 && sheet === wb.SheetNames[0]) {
      console.log('字段:', Object.keys(data[0]));
      console.log('前 2 条数据:');
      data.slice(0, 2).forEach((row, i) => console.log(`  ${i + 1}:`, row));
    }
  });
} catch (error) {
  console.error('Error:', error.message);
}
