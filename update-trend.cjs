const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const startIdx = code.indexOf('{/* Toggle switch for history display options');
const endIdx = code.indexOf('</ResponsiveContainer>\n            </div>', startIdx) + '</ResponsiveContainer>\n            </div>'.length;

if (startIdx === -1 || endIdx === -1) {
  console.log("Could not find boundaries");
  process.exit(1);
}

const replacement = `{/* Metric Toggle for Trend Chart */}
              <div className="flex items-center gap-2 self-start md:self-auto bg-[#fbfaff] px-3.5 py-1.5 rounded-xl border border-[#e2e8f0]/40 shrink-0">
                <div className="flex items-center gap-4 select-none">
                  <button
                    type="button"
                    onClick={() => setShowBudgetBar(prev => !prev)}
                    className={\`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer \${!showBudgetBar ? "opacity-35 line-through" : ""}\`}
                    title="Klik untuk menyembunyikan/menampilkan Budget"
                  >
                    <div className="w-3.5 h-3.5 rounded-sm bg-gradient-to-b from-[#ea580c] to-[#c2410c] shadow-sm"></div>
                    <span className="text-[10.5px] font-bold text-slate-700 uppercase tracking-wide">Budget</span>
                  </button>
                  <button
                    type="button"
                    onClick={() => setShowActualBar(prev => !prev)}
                    className={\`flex items-center gap-2 hover:opacity-85 transition-all cursor-pointer \${!showActualBar ? "opacity-35 line-through" : ""}\`}
                    title="Klik untuk menyembunyikan/menampilkan Actual"
                  >
                    <div className="w-3.5 h-3.5 rounded-sm bg-gradient-to-b from-[#f97316] to-[#ea580c] shadow-sm"></div>
                    <span className="text-[10.5px] font-bold text-slate-700 uppercase tracking-wide">Actual</span>
                  </button>
                </div>
              </div>
            </div>

            <div className="h-64 w-full font-sans mt-4">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart
                  data={overviewHistoryData}
                  margin={{ top: 38, right: 10, left: -10, bottom: 0 }}
                  barGap={currentBarGap}
                  barCategoryGap={currentBarCategoryGap}
                  onMouseMove={(state) => {
                    if (state && state.activeLabel) {
                      setHoveredLabel(state.activeLabel);
                      if (state.activeLabel !== dismissedTooltipLabel) {
                        setDismissedTooltipLabel(null);
                      }
                    } else {
                      setHoveredLabel(null);
                    }
                  }}
                  onMouseLeave={() => {
                    setHoveredLabel(null);
                  }}
                >
                  <defs>
                    <linearGradient
                      id="colorTrendBudget"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="0%"
                        stopColor="#154be2"
                        stopOpacity={0.95}
                      />
                      <stop
                        offset="100%"
                        stopColor="#3b82f6"
                        stopOpacity={0.7}
                      />
                    </linearGradient>
                    <linearGradient
                      id="colorTrendActual"
                      x1="0"
                      y1="0"
                      x2="0"
                      y2="1"
                    >
                      <stop
                        offset="0%"
                        stopColor="#06b6d4"
                        stopOpacity={1.0}
                      />
                      <stop
                        offset="100%"
                        stopColor="#22d3ee"
                        stopOpacity={1.0}
                      />
                    </linearGradient>
                  </defs>
                  <CartesianGrid
                    strokeDasharray="4 4"
                    vertical={false}
                    stroke="#e2e8f0"
                  />
                  <XAxis
                    dataKey="name"
                    tick={<CustomXAxisTick chartData={overviewHistoryData} metricType={overviewMetricFilter} />}
                    axisLine={false}
                    tickLine={false}
                    interval={0}
                    height={65}
                  />
                  <YAxis
                    hide={true}
                    domain={[0, (dataMax) => (dataMax === 0 ? 100 : Math.round(dataMax * 1.25))]}
                    tick={{ fill: "#8E94B7", fontSize: 9, fontWeight: 500 }}
                    axisLine={false}
                    tickLine={false}
                  />
                  <Tooltip
                    cursor={{ fill: "rgba(21, 75, 226, 0.03)" }}
                    content={<CustomChartTooltip metricType={overviewMetricFilter} dismissedLabel={dismissedTooltipLabel} />}
                  />
                  <Bar
                    hide={!showBudgetBar}
                    dataKey={
                      overviewMetricFilter === "activity"
                        ? "budgetActivity"
                        : overviewMetricFilter === "reach"
                        ? "budgetReach"
                        : "budgetNominal"
                    }
                    name="Budget"
                    fill="url(#colorTrendBudget)"
                    radius={[6, 6, 0, 0]}
                    maxBarSize={currentMaxBarSize}
                    background={<CustomBarBackground data={overviewHistoryData} activeKey={activeMainBarKey} />}
                  >
                    {overviewHistoryData.map((entry, index) => {
                      const isActive = activeMainBarKey === entry.name;
                      return (
                        <Cell
                          key={\`cell-trend-budget-\${index}\`}
                          cursor="pointer"
                          fill={isActive ? "#ea580c" : "url(#colorTrendBudget)"}
                          fillOpacity={1.0}
                          onClick={() => {
                            clickedBarRef.current = true;
                            if (activeMainBarKey === entry.name) {
                              setDismissedTooltipLabel(entry.name);
                            } else {
                              setDismissedTooltipLabel(null);
                            }
                            setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                          }}
                        />
                      );
                    })}
                    <LabelList
                      dataKey={
                        overviewMetricFilter === "activity"
                          ? "budgetActivity"
                          : overviewMetricFilter === "reach"
                          ? "budgetReach"
                          : "budgetNominal"
                      }
                      content={<CustomBudgetLabel metricType={overviewMetricFilter} />}
                    />
                  </Bar>
                  <Bar
                    hide={!showActualBar}
                    dataKey={
                      overviewMetricFilter === "activity"
                        ? "actualActivity"
                        : overviewMetricFilter === "reach"
                        ? "actualReach"
                        : "actualNominal"
                    }
                    name="Actual"
                    fill="url(#colorTrendActual)"
                    radius={[6, 6, 0, 0]}
                    maxBarSize={currentMaxBarSize}
                    background={<CustomBarBackground data={overviewHistoryData} activeKey={activeMainBarKey} />}
                  >
                    {overviewHistoryData.map((entry, index) => {
                      const isActive = activeMainBarKey === entry.name;
                      return (
                        <Cell
                          key={\`cell-trend-actual-\${index}\`}
                          cursor="pointer"
                          fill={isActive ? "#f97316" : "url(#colorTrendActual)"}
                          fillOpacity={1.0}
                          onClick={() => {
                            clickedBarRef.current = true;
                            if (activeMainBarKey === entry.name) {
                              setDismissedTooltipLabel(entry.name);
                            } else {
                              setDismissedTooltipLabel(null);
                            }
                            setActiveMainBarKey(prev => prev === entry.name ? null : entry.name);
                          }}
                        />
                      );
                    })}
                    <LabelList
                      dataKey={
                        overviewMetricFilter === "activity"
                          ? "actualActivity"
                          : overviewMetricFilter === "reach"
                          ? "actualReach"
                          : "actualNominal"
                      }
                      content={<CustomActualLabel metricType={overviewMetricFilter} />}
                    />
                  </Bar>
                </BarChart>
              </ResponsiveContainer>
            </div>`;

code = code.substring(0, startIdx) + replacement + code.substring(endIdx);
fs.writeFileSync('src/App.tsx', code);
console.log("Updated trend chart successfully");
