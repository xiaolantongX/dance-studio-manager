const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// 使用相对路径
const excelPath = path.join(__dirname, '2025nianjiye.xlsx');

// 先查找文件
const files = fs.readdirSync(__dirname);
console.log('当前目录文件:', files.filter(f => f.endsWith('.xlsx')));

// 尝试多个可能的文件名
const possibleNames = [
  '2025nianjiye.xlsx',
  '2025nian-jiye.xlsx',
  '2025.xlsx',
  '业绩.xlsx',
  'jiye.xlsx'
];

for (const name of possibleNames) {
  const testPath = path.join(__dirname, name);
  if (fs.existsSync(testPath)) {
    console.log('找到文件:', name);
    try {
      const wb = XLSX.readFile(testPath);
      console.log('\n=== Sheet 名称 ===');
      console.log(wb.SheetNames);

      console.log('\n=== 第一个 Sheet 数据结构 ===');
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws);

      if (data.length > 0) {
        console.log('字段:', Object.keys(data[0]));
        console.log('\n前 3 条数据:');
        data.slice(0, 3).forEach((row, i) => {
          console.log(`${i + 1}:`, JSON.stringify(row, null, 2));
        });
      }
    } catch (error) {
      console.error('读取错误:', error.message);
    }
    break;
  }
}
