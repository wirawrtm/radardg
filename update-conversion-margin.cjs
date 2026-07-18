const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

code = code.replace(
  '{/* Left: Conversion Sales Rate by Activity */}\n            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between">',
  '{/* Left: Conversion Sales Rate by Activity */}\n            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between mt-6 lg:mt-8">'
);

fs.writeFileSync('src/App.tsx', code);
console.log("Updated margin top successfully");
