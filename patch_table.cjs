const fs = require('fs');
const code = fs.readFileSync('src/App.tsx', 'utf-8');

const regex = /<table className="w-full text-left border-collapse min-w-\[500px\]">[\s\S]*?(?=<\/table>)<\/table>/;

const replacement = `{(() => {
                      const rawData = [
                        { id: "L1", district: "Malang", subDistrict: "Waru", budget: 1500000, month: "Agustus" },
                        { id: "L2", district: "Malang", subDistrict: "Singosari", budget: 2000000, month: "September" },
                        { id: "L3", district: "Pasuruan", subDistrict: "Bangil", budget: 1200000, month: "Oktober" },
                        { id: "L4", district: "Pasuruan", subDistrict: "Pandaan", budget: 1800000, month: "Agustus" },
                        { id: "L5", district: "Batu", subDistrict: "Bumiaji", budget: 1500000, month: "November" },
                      ];
                      
                      const uniqueDistricts = Array.from(new Set(rawData.map(r => r.district))).sort();
                      const uniqueSubDistricts = Array.from(new Set(rawData.map(r => r.subDistrict))).sort();
                      const uniqueMonths = Array.from(new Set(rawData.map(r => r.month))).sort();
                      
                      const filteredData = rawData.filter(row => 
                        (planningFilterDistrict === "" || row.district === planningFilterDistrict) &&
                        (planningFilterSubDistrict === "" || row.subDistrict === planningFilterSubDistrict) &&
                        (planningFilterMonth === "" || row.month === planningFilterMonth)
                      );

                      return (
                        <table className="w-full text-left border-collapse min-w-[500px]">
                          <thead>
                            <tr className="bg-slate-50">
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                                District
                                <select value={planningFilterDistrict} onChange={e => setPlanningFilterDistrict(e.target.value)} className="mt-1 block w-full px-1.5 py-1 text-[9px] font-normal border border-slate-200 rounded text-slate-800 bg-white focus:outline-none focus:border-[#154be2]">
                                  <option value="">All</option>
                                  {uniqueDistricts.map(d => <option key={d} value={d}>{d}</option>)}
                                </select>
                              </th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                                Sub District
                                <select value={planningFilterSubDistrict} onChange={e => setPlanningFilterSubDistrict(e.target.value)} className="mt-1 block w-full px-1.5 py-1 text-[9px] font-normal border border-slate-200 rounded text-slate-800 bg-white focus:outline-none focus:border-[#154be2]">
                                  <option value="">All</option>
                                  {uniqueSubDistricts.map(d => <option key={d} value={d}>{d}</option>)}
                                </select>
                              </th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200 align-top">Budget</th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200">
                                Month
                                <select value={planningFilterMonth} onChange={e => setPlanningFilterMonth(e.target.value)} className="mt-1 block w-full px-1.5 py-1 text-[9px] font-normal border border-slate-200 rounded text-slate-800 bg-white focus:outline-none focus:border-[#154be2]">
                                  <option value="">All</option>
                                  {uniqueMonths.map(m => <option key={m} value={m}>{m}</option>)}
                                </select>
                              </th>
                              <th className="px-3 py-2 text-[9px] font-extrabold text-slate-500 uppercase tracking-wider border-b border-slate-200 text-right align-top">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100">
                            {filteredData.map((row) => (
                              <tr key={row.id} className="hover:bg-slate-50/50 transition-colors">
                                <td className="px-3 py-2 text-[10px] font-bold text-slate-700">{row.district}</td>
                                <td className="px-3 py-2 text-[10px] font-semibold text-slate-600">{row.subDistrict}</td>
                                <td className="px-3 py-2 text-[10px] font-bold text-slate-800">Rp {row.budget.toLocaleString()}</td>
                                <td className="px-3 py-2 text-[10px] font-medium text-slate-500">{row.month}</td>
                                <td className="px-3 py-2 text-right">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const bsMap = { "Lionel Messi": "lm", "Ronaldo": "ro", "Mbappe": "mb", "Yamal": "ya" };
                                      const actMap = { "Farmer meeting": "fm", "Farmer field day": "ffd", "One day promo": "odp", "Special field trip": "sft", "AIC": "aic", "Expo": "exp", "Caravan": "crv", "Retailer Meeting": "rm" };
                                      const bsCode = bsMap[proposeBs as keyof typeof bsMap] || proposeBs.substring(0, 2).toLowerCase();
                                      const actCode = actMap[proposeActivity as keyof typeof actMap] || "act";
                                      const monthMap = { "Agustus": "08", "September": "09", "Oktober": "10", "November": "11" };
                                      const monthCode = monthMap[row.month as keyof typeof monthMap] || "08";
                                      const yearCode = "26";
                                      const uniqueSuffix = String(Math.floor(Math.random() * 900) + 100);
                                      const projectNo = \`\${bsCode}/\${actCode}/\${monthCode}/\${yearCode}-\${uniqueSuffix}\`;
                                      
                                      const newProj = {
                                        id: (Date.now() + Math.random()).toString(),
                                        projectNo,
                                        bs: proposeBs,
                                        category: proposeCategory,
                                        activity: proposeActivity,
                                        district: row.district,
                                        subDistrict: row.subDistrict,
                                        budget: row.budget,
                                        month: row.month,
                                        farmerReach: "",
                                        hybrids: "",
                                      };
                                      setGeneratedProjects((prev) => [...prev, newProj]);
                                    }}
                                    className="bg-[#154be2] hover:bg-[#154be2]/90 text-white px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all shadow-sm active:scale-95"
                                  >
                                    Propose
                                  </button>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      );
                    })()}`;

if (regex.test(code)) {
  fs.writeFileSync('src/App.tsx', code.replace(regex, replacement));
  console.log("Success");
} else {
  console.log("Failed to match");
}
