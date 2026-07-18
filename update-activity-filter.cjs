const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

// Insert the new state
const stateStr = 'const [conversionSalesFilter, setConversionSalesFilter] = useState("area");';
if (!code.includes('const [conversionActivityFilter, setConversionActivityFilter] = useState("area");')) {
    code = code.replace(stateStr, stateStr + '\n  const [conversionActivityFilter, setConversionActivityFilter] = useState("area");');
}

const oldBlockStart = '{/* Left: Conversion Sales Rate by Activity */}';
const oldBlockEnd = '{/* Right: Conversion Sales Rate by Filter */}';

const startIndex = code.indexOf(oldBlockStart);
const endIndex = code.indexOf(oldBlockEnd);

if (startIndex === -1 || endIndex === -1) {
  console.log("Could not find start or end index");
  process.exit(1);
}

const replacement = `{/* Left: Conversion Sales Rate by Activity */}
            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between">
              <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="material-symbols-outlined text-primary text-sm font-semibold">
                      swap_horiz
                    </span>
                    <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                      Conversion Sales Rate (Activity)
                    </h3>
                  </div>
                  <p className="text-[10px] text-[#8E94B7] mt-0.5">
                    Pilih aktivitas untuk melihat impact konversi
                  </p>
                </div>

                {/* Filter Selection */}
                <select 
                  className="bg-[#f0effc] text-[#181a2c] text-[10px] font-semibold px-3 py-1.5 rounded-lg border-none outline-none cursor-pointer shrink-0"
                  value={conversionActivityFilter}
                  onChange={(e) => setConversionActivityFilter(e.target.value)}
                >
                  <option value="area">Area</option>
                  <option value="province">Province</option>
                  <option value="Sales Agronomist">Sales Agronomist</option>
                  <option value="Hybrids">Hybrids</option>
                </select>
              </div>
              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: "Farmer Meeting", rate: 68, budget: 150000, sales: 102000 },
                  { name: "Demo Plot", rate: 54, budget: 120000, sales: 64800 },
                  { name: "Field Day", rate: 45, budget: 100000, sales: 45000 },
                  { name: "Kiosk Visit", rate: 32, budget: 200000, sales: 64000 },
                  { name: "Farmer Visit", rate: 21, budget: 180000, sales: 37800 },
                ].map((item, index) => (
                  <div 
                    key={item.name} 
                    onClick={() => setActiveConversionActivity(item.name)}
                    className={\`flex items-center justify-between p-3 rounded-xl border transition-all duration-200 cursor-pointer \${activeConversionActivity === item.name ? "bg-[#154be2]/5 border-[#154be2]/30 shadow-sm" : "bg-[#fbfaff] border-[#f0effc] hover:border-[#154be2]/20"}\`}
                  >
                    <div className="flex items-center gap-3">
                      <div className={\`size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs \${activeConversionActivity === item.name ? "bg-[#154be2] text-white" : "bg-[#154be2]/10 text-primary"}\`}>
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate">{item.name}</h4>
                        <div className="flex items-center gap-2 mt-0.5">
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Sales: <span className="font-semibold text-emerald-600">{item.sales.toLocaleString()}</span>
                          </p>
                          <span className="text-[8px] text-slate-300">|</span>
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Budget: <span className="font-semibold text-slate-600">{item.budget.toLocaleString()}</span>
                          </p>
                        </div>
                      </div>
                    </div>
                    <div className="text-right shrink-0">
                      <p className="text-xs font-bold text-primary font-sans">{item.rate}%</p>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            `;

code = code.substring(0, startIndex) + replacement + code.substring(endIndex);
fs.writeFileSync('src/App.tsx', code);
console.log("Updated activity filter successfully");
