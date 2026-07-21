const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

// 1. Replace the parent div className
code = code.replace(
  /\(overviewMetricFilter === "overview" \|\| overviewMetricFilter === "monitoring"\)\n\s*\? "grid grid-cols-2 lg:grid-cols-4 gap-4 lg:gap-6 w-full pb-2"/g,
  '(overviewMetricFilter === "overview" || overviewMetricFilter === "monitoring")\n                  ? "flex flex-col md:flex-row items-center md:items-stretch w-full pb-2"\n'
);

// 2. Change how the cards are rendered
const targetCode = `                    return (
                      <>
                        {activeTab === "overview_v2" && (
                          <div className="rounded-[32px] flex flex-row items-center justify-center relative overflow-hidden group transition-all duration-300 w-full text-center py-4 xs:py-5 sm:py-6 px-4 min-h-[110px] col-span-1 bg-gradient-to-r from-violet-600 to-indigo-600 border border-violet-400/30 shadow-[0_12px_32px_rgba(99,102,241,0.15)] hover:scale-[1.01] hover:shadow-[0_16px_40px_rgba(99,102,241,0.25)] select-none">
                            <div className="flex items-center justify-center z-10 w-full h-full">
                              <img
                                src="/jagoan.png"
                                className="h-24 w-auto object-contain drop-shadow-md"
                                alt="Jagoan Advanta"
                              />
                            </div>
                          </div>
                        )}
                        {overviewCardsLocal.map((card, idx) => {`;

const newCode = `                    const renderedCards = overviewCardsLocal.map((card, idx) => {`;

code = code.replace(targetCode, newCode);

const targetCode2 = `                        })}
                      </>
                    );
                  })()`;

const newCode2 = `                        });

                    return (
                      <>
                        {activeTab === "overview_v2" && (
                          <div className="shrink-0 flex items-center justify-center pt-2 pb-4 md:py-0 md:pr-4 lg:pr-6">
                            <img
                              src="/jagoan.png"
                              className="h-28 sm:h-32 md:h-40 w-auto object-contain drop-shadow-md"
                              alt="Jagoan Advanta"
                            />
                          </div>
                        )}
                        {activeTab === "overview_v2" ? (
                          <div className="flex-1 grid grid-cols-2 lg:grid-cols-3 gap-4 lg:gap-6 w-full">
                            {renderedCards}
                          </div>
                        ) : (
                          renderedCards
                        )}
                      </>
                    );
                  })()`;

code = code.replace(targetCode2, newCode2);

fs.writeFileSync('src/App.tsx', code);
console.log("Done");
