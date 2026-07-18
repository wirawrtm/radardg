const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

// Update activity card styling
code = code.replace(
  'className={`rounded-[22px] flex flex-row items-center justify-between relative overflow-hidden group transition-all duration-300 w-full text-left cursor-pointer border-0 ${',
  'className={`rounded-[32px] lg:w-[94%] lg:mx-auto flex flex-row items-center justify-between relative overflow-hidden group transition-all duration-300 w-full text-left cursor-pointer border-0 ${'
);

code = code.replace(
  'act.name === "TOTAL"\n                          ? "min-h-[116px] py-3 px-5 lg:py-4 lg:px-6"\n                          : "min-h-[96px] py-2 px-4 lg:py-2.5 lg:px-4.5"',
  'act.name === "TOTAL"\n                          ? "min-h-[116px] py-4 px-5 lg:py-6 lg:px-6"\n                          : "min-h-[96px] py-3 px-4 lg:py-4 lg:px-4.5"'
);

code = code.replace(
  'act.name === "TOTAL"\n                          ? "min-h-[116px] py-4 px-5 lg:py-6 lg:px-6"\n                          : "min-h-[96px] py-3 px-4 lg:py-4 lg:px-4.5"',
  'act.name === "TOTAL"\n                          ? "min-h-[116px] py-5 px-5 lg:py-6 lg:px-6"\n                          : "min-h-[96px] py-4 px-4 lg:py-5 lg:px-4.5"'
);

fs.writeFileSync('src/App.tsx', code);
console.log("Updated activity cards styling successfully");
