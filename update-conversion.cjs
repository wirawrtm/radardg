const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const startStr = "{/* Section: Employee Sales Ranking (Highest & Lowest) */}";
const endStr = "{/* Section: History Bulanan */}";

const startIndex = code.indexOf(startStr);
const endIndex = code.indexOf(endStr);

if (startIndex === -1 || endIndex === -1) {
  console.log("Could not find start or end index");
  process.exit(1);
}

const replacement = `          {/* Section: Conversion Sales Rate */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-6">
            {/* Left: Conversion Sales Rate by Activity */}
            <div className="bg-white p-6 rounded-[40px] shadow-[0_12px_32px_rgba(21,75,226,0.18)] border border-[#154be2]/8 flex flex-col justify-between">
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
                  Tingkat konversi penjualan berdasarkan aktivitas
                </p>
              </div>
              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: "Farmer Meeting", rate: 68 },
                  { name: "Demo Plot", rate: 54 },
                  { name: "Field Day", rate: 45 },
                  { name: "Kiosk Visit", rate: 32 },
                  { name: "Farmer Visit", rate: 21 },
                ].map((item, index) => (
                  <div key={item.name} className="flex items-center justify-between p-3 rounded-xl bg-[#fbfaff] border border-[#f0effc] hover:border-[#154be2]/20 transition-all duration-200">
                    <div className="flex items-center gap-3">
                      <div className="size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs bg-[#154be2]/10 text-primary">
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate">{item.name}</h4>
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
                      Conversion Sales Rate (Impact)
                    </h3>
                  </div>
                  <p className="text-[10px] text-[#8E94B7] mt-0.5">
                    Impact konversi dari aktivitas
                  </p>
                </div>
                
                {/* Filter Selection */}
                <select 
                  className="bg-[#f0effc] text-[#181a2c] text-[10px] font-semibold px-3 py-1.5 rounded-lg border-none outline-none cursor-pointer"
                  value={conversionSalesFilter}
                  onChange={(e) => setConversionSalesFilter(e.target.value)}
                >
                  <option value="area">Area</option>
                  <option value="province">Province</option>
                  <option value="Sales Agronomist">Sales Agronomist</option>
                  <option value="Hybrids">Hybrids</option>
                </select>
              </div>

              <div className="flex flex-col gap-3 mt-4">
                {[
                  { name: conversionSalesFilter === "area" ? "Area 1" : conversionSalesFilter === "province" ? "Jawa Timur" : conversionSalesFilter === "Sales Agronomist" ? "Budi" : "NK 212", rate: 75 },
                  { name: conversionSalesFilter === "area" ? "Area 2" : conversionSalesFilter === "province" ? "Jawa Tengah" : conversionSalesFilter === "Sales Agronomist" ? "Agus" : "NK 6172", rate: 62 },
                  { name: conversionSalesFilter === "area" ? "Area 3" : conversionSalesFilter === "province" ? "Jawa Barat" : conversionSalesFilter === "Sales Agronomist" ? "Joko" : "NK 7328", rate: 48 },
                  { name: conversionSalesFilter === "area" ? "Area 4" : conversionSalesFilter === "province" ? "Sumatera Utara" : conversionSalesFilter === "Sales Agronomist" ? "Rudi" : "NK 33", rate: 35 },
                  { name: conversionSalesFilter === "area" ? "Area 5" : conversionSalesFilter === "province" ? "Sulawesi Selatan" : conversionSalesFilter === "Sales Agronomist" ? "Andi" : "NK 99", rate: 20 },
                ].map((item, index) => (
                  <div key={item.name} className="flex items-center justify-between p-3 rounded-xl bg-[#fbfaff] border border-[#f0effc] hover:border-[#154be2]/20 transition-all duration-200">
                    <div className="flex items-center gap-3">
                      <div className="size-8 shrink-0 rounded-full flex items-center justify-center font-bold text-xs bg-[#154be2]/10 text-primary">
                        #{index + 1}
                      </div>
                      <div className="min-w-0">
                        <h4 className="text-[11px] font-bold text-[#181a2c] leading-tight truncate">{item.name}</h4>
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
console.log("Updated successfully");
