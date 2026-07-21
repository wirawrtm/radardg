const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

// 1. Remove the user profile icon
const targetProfile = `<div className="flex items-center gap-3">
              <div className="size-10 rounded-2xl bg-[#154be2]/10 flex items-center justify-center shrink-0 border border-[#154be2]/20">
                <User className="size-5 text-[#154be2]" />
              </div>
              <div>
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">`;

const newProfile = `<div className="flex items-center gap-3">
              <div>
                <h1 className="text-lg font-semibold text-[#181a2c] tracking-tight">`;

code = code.replace(targetProfile, newProfile);

// 2. Make BarChart more rounded
code = code.replace(/radius=\{\[6, 6, 6, 6\]\}/g, 'radius={[12, 12, 12, 12]}');

// 3. Make the OverviewXAxisTick label larger
const targetTick = `          <tspan 
            x={0} 
            dy={index === 0 ? 12 : 12} 
            key={\`word-overview-\${index}\`} 
            fill="#8E94B7" 
            style={{ fontSize: "11px", fontWeight: 800 }}
          >
            {word}
          </tspan>`;

const newTick = `          <tspan 
            x={0} 
            dy={index === 0 ? 14 : 14} 
            key={\`word-overview-\${index}\`} 
            fill="#8E94B7" 
            style={{ fontSize: "13px", fontWeight: 900 }}
          >
            {word}
          </tspan>`;

code = code.replace(targetTick, newTick);

fs.writeFileSync('src/App.tsx', code);
console.log("Fixed!");
