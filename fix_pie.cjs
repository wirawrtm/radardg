const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const targetLabel = `                                    const isClicked = index === clickedPieIndex;

                                    // Shift label coordinates slightly for the active/exploded slice
                                    let lx = x;
                                    let ly = y;
                                    if (index === activeIndex) {
                                      const RADIAN = Math.PI / 180;
                                      const sin = Math.sin(-midAngle * RADIAN);
                                      const cos = Math.cos(-midAngle * RADIAN);
                                      lx += 15 * cos;
                                      ly += 15 * sin;
                                    }

                                    return (
                                      <text
                                        x={lx}
                                        y={ly}
                                        fill={isClicked ? "#154be2" : "#1e293b"}
                                        textAnchor={lx > cx ? "start" : "end"}
                                        dominantBaseline="central"
                                        className={
                                          isClicked
                                            ? "text-[13px] sm:text-[14px] lg:text-[15px] font-black drop-shadow-sm transition-all duration-300"
                                            : "text-[10px] sm:text-[11px] lg:text-[12px] font-extrabold"
                                        }
                                      >
                                        {\`\${name}: \${formattedValue}\`}
                                      </text>
                                    );
                                  }}`;

const newLabel = `                                    const isActive = index === activeIndex;

                                    // Shift label coordinates slightly for the active/exploded slice
                                    let lx = x;
                                    let ly = y;
                                    if (isActive) {
                                      const RADIAN = Math.PI / 180;
                                      const sin = Math.sin(-midAngle * RADIAN);
                                      const cos = Math.cos(-midAngle * RADIAN);
                                      lx += 15 * cos;
                                      ly += 15 * sin;
                                    }
                                    
                                    // Split name if too long to make it wrap
                                    let nameLines = [name];
                                    if (name.length > 10 && name.includes(" ")) {
                                      const words = name.split(" ");
                                      const mid = Math.floor(words.length / 2);
                                      nameLines = [words.slice(0, mid).join(" "), words.slice(mid).join(" ")];
                                    }

                                    return (
                                      <text
                                        x={lx}
                                        y={ly}
                                        fill={isActive ? "#154be2" : "#1e293b"}
                                        textAnchor={lx > cx ? "start" : "end"}
                                        dominantBaseline="central"
                                        className={
                                          isActive
                                            ? "text-[12px] sm:text-[13px] lg:text-[14px] font-black drop-shadow-sm transition-all duration-300"
                                            : "text-[9px] sm:text-[10px] lg:text-[11px] font-extrabold"
                                        }
                                      >
                                        {nameLines.map((line, i) => (
                                          <tspan key={i} x={lx} dy={i === 0 ? (nameLines.length > 1 ? "-1em" : "-0.5em") : "1.2em"}>
                                            {line}
                                          </tspan>
                                        ))}
                                        <tspan x={lx} dy="1.2em" fill={isActive ? "#1d4ed8" : "#64748b"} className="font-semibold">
                                          {formattedValue}
                                        </tspan>
                                      </text>
                                    );
                                  }}`;

code = code.replace(targetLabel, newLabel);
fs.writeFileSync('src/App.tsx', code);
console.log("Pie chart label fixed!");
