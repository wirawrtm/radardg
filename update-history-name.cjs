const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

code = code.replace(
  '        monthLabel: item.label,',
  '        monthLabel: item.label,\n        name: item.label,'
);

code = code.replace(
  '        monthLabel: string;',
  '        monthLabel: string;\n        name: string;'
);

fs.writeFileSync('src/App.tsx', code);
console.log("Updated name in overviewHistoryData successfully");
