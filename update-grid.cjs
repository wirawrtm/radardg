const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

code = code.replace(/lg:grid-cols-8/g, 'lg:grid-cols-11');
code = code.replace(/lg:col-span-5 bg-white p-4/g, 'lg:col-span-7 bg-white p-4');
code = code.replace(/lg:col-span-3 flex flex-col h-full min-h-0/g, 'lg:col-span-4 flex flex-col h-full min-h-0');

fs.writeFileSync('src/App.tsx', code);
console.log("Updated grid successfully");
