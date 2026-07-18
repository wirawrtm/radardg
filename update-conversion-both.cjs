const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

// Change initial state to "activity" if it is "area"
code = code.replace('const [conversionSalesFilter, setConversionSalesFilter] = useState("area");', 'const [conversionSalesFilter, setConversionSalesFilter] = useState("activity");');
code = code.replace('const [conversionActivityFilter, setConversionActivityFilter] = useState("area");', 'const [conversionActivityFilter, setConversionActivityFilter] = useState("activity");');


const oldBlockStart = '{/* Section: Conversion Sales Rate */}';
const oldBlockEnd = '{/* Section: History Bulanan */}';
const startIndex = code.indexOf(oldBlockStart);
const endIndex = code.indexOf(oldBlockEnd);

if (startIndex === -1 || endIndex === -1) {
  console.log("Could not find start or end index");
  process.exit(1);
}

const replacement = `          {/* Section: Conversion Sales Rate */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
            {/* Left: Conversion Sales Rate by Activity */}
            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between">
              <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="material-symbols-outlined text-primary text-sm font-semibold">
                      swap_horiz
                    </span>
                    <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                      Conversion Sales Rate (Left)
                    </h3>
                  </div>
                  <p className="text-[10px] text-[#8E94B7] mt-0.5">
                    Pilih filter untuk melihat impact konversi
                  </p>
                </div>

                {/* Filter Selection */}
                <select 
                  className="bg-[#f0effc] text-[#181a2c] text-[10px] font-semibold px-3 py-1.5 rounded-lg border-none outline-none cursor-pointer shrink-0"
                  value={conversionActivityFilter}
                  onChange={(e) => setConversionActivityFilter(e.target.value)}
                >
                  <option value="activity">Activity</option>
                  <option value="area">Area</option>
                  <option value="province">Province</option>
                  <option value="Sales Agronomist">Sales Agronomist</option>
                  <option value="Hybrids">Hybrids</option>
                </select>
              </div>
              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: conversionActivityFilter === "area" ? "Area 1" : conversionActivityFilter === "province" ? "Jawa Timur" : conversionActivityFilter === "Sales Agronomist" ? "Budi" : conversionActivityFilter === "Hybrids" ? "NK 212" : "Farmer Meeting", rate: 68, budget: 150000, sales: 102000 },
                  { name: conversionActivityFilter === "area" ? "Area 2" : conversionActivityFilter === "province" ? "Jawa Tengah" : conversionActivityFilter === "Sales Agronomist" ? "Agus" : conversionActivityFilter === "Hybrids" ? "NK 6172" : "Demo Plot", rate: 54, budget: 120000, sales: 64800 },
                  { name: conversionActivityFilter === "area" ? "Area 3" : conversionActivityFilter === "province" ? "Jawa Barat" : conversionActivityFilter === "Sales Agronomist" ? "Joko" : conversionActivityFilter === "Hybrids" ? "NK 7328" : "Field Day", rate: 45, budget: 100000, sales: 45000 },
                  { name: conversionActivityFilter === "area" ? "Area 4" : conversionActivityFilter === "province" ? "Sumatera Utara" : conversionActivityFilter === "Sales Agronomist" ? "Rudi" : conversionActivityFilter === "Hybrids" ? "NK 33" : "Kiosk Visit", rate: 32, budget: 200000, sales: 64000 },
                  { name: conversionActivityFilter === "area" ? "Area 5" : conversionActivityFilter === "province" ? "Sulawesi Selatan" : conversionActivityFilter === "Sales Agronomist" ? "Andi" : conversionActivityFilter === "Hybrids" ? "NK 99" : "Farmer Visit", rate: 21, budget: 180000, sales: 37800 },
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

            {/* Right: Conversion Sales Rate by Filter */}
            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between">
              <div className="flex flex-col sm:flex-row sm:items-start justify-between gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <span className="material-symbols-outlined text-primary text-sm font-semibold">
                      filter_alt
                    </span>
                    <h3 className="text-xs font-bold text-[#181a2c] tracking-tight">
                      Impact: {activeConversionActivity}
                    </h3>
                  </div>
                  <p className="text-[10px] text-[#8E94B7] mt-0.5">
                    Konversi penjualan berdasarkan filter
                  </p>
                </div>
                
                {/* Filter Selection */}
                <select 
                  className="bg-[#f0effc] text-[#181a2c] text-[10px] font-semibold px-3 py-1.5 rounded-lg border-none outline-none cursor-pointer"
                  value={conversionSalesFilter}
                  onChange={(e) => setConversionSalesFilter(e.target.value)}
                >
                  <option value="activity">Activity</option>
                  <option value="area">Area</option>
                  <option value="province">Province</option>
                  <option value="Sales Agronomist">Sales Agronomist</option>
                  <option value="Hybrids">Hybrids</option>
                </select>
              </div>

              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: conversionSalesFilter === "activity" ? "Farmer Meeting" : conversionSalesFilter === "area" ? "Area 1" : conversionSalesFilter === "province" ? "Jawa Timur" : conversionSalesFilter === "Sales Agronomist" ? "Budi" : "NK 212", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 75 : 45, budget: 85000, sales: 63750 },
                  { name: conversionSalesFilter === "activity" ? "Demo Plot" : conversionSalesFilter === "area" ? "Area 2" : conversionSalesFilter === "province" ? "Jawa Tengah" : conversionSalesFilter === "Sales Agronomist" ? "Agus" : "NK 6172", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 62 : 40, budget: 70000, sales: 43400 },
                  { name: conversionSalesFilter === "activity" ? "Field Day" : conversionSalesFilter === "area" ? "Area 3" : conversionSalesFilter === "province" ? "Jawa Barat" : conversionSalesFilter === "Sales Agronomist" ? "Joko" : "NK 7328", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 48 : 35, budget: 65000, sales: 31200 },
                  { name: conversionSalesFilter === "activity" ? "Kiosk Visit" : conversionSalesFilter === "area" ? "Area 4" : conversionSalesFilter === "province" ? "Sumatera Utara" : conversionSalesFilter === "Sales Agronomist" ? "Rudi" : "NK 33", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 35 : 25, budget: 50000, sales: 17500 },
                  { name: conversionSalesFilter === "activity" ? "Farmer Visit" : conversionSalesFilter === "area" ? "Area 5" : conversionSalesFilter === "province" ? "Sulawesi Selatan" : conversionSalesFilter === "Sales Agronomist" ? "Andi" : "NK 99", rate: (activeConversionActivity || "").includes("1") || (activeConversionActivity || "").includes("Jawa Timur") || (activeConversionActivity || "").includes("Budi") || (activeConversionActivity || "").includes("212") || (activeConversionActivity || "").includes("Farmer Meeting") ? 20 : 15, budget: 45000, sales: 9000 },
                ].map((item, index) => (
                  <div key={item.name} className="flex items-center justify-between p-3 rounded-xl bg-[#fbfaff] border border-[#f0effc] hover:border-[#154be2]/20 transition-all duration-200">
                    <div className="flex items-center gap-3">
                      <div className="size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs bg-[#154be2]/10 text-primary">
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate">{item.name}</h4>
                        <div className="flex items-center gap-2 mt-0.5">
                          <p className="text-[9px] text-[#8E94B7] truncate">
                            Sales: <span className="font-semibold text-emerald-600">{Math.round(item.budget * (item.rate / 100)).toLocaleString()}</span>
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
          </div>
          
          `;

code = code.substring(0, startIndex) + replacement + code.substring(endIndex);
fs.writeFileSync('src/App.tsx', code);
console.log("Updated both conversion sides successfully");
