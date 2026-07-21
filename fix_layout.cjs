const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const targetCode = `                        {activeTab === "overview_v2" ? (
                          <div className="w-full grid grid-cols-2 lg:grid-cols-[auto_1fr_1fr_1fr] gap-4 lg:gap-6 items-stretch">
                            <div className="flex items-center justify-center lg:pr-2 col-span-1">
                              <img
                                src="/jagoan.png"
                                className="h-full max-h-[100px] sm:max-h-[110px] md:max-h-[130px] lg:max-h-[140px] w-auto object-contain drop-shadow-md"
                                alt="Jagoan Advanta"
                              />
                            </div>
                            {renderedCards}
                          </div>
                        ) : (
                          renderedCards
                        )}`;

const newCode = `                        {activeTab === "overview_v2" ? (
                          <>
                            {/* Desktop Layout */}
                            <div className="hidden lg:grid w-full lg:grid-cols-[auto_1fr_1fr_1fr] gap-6 items-stretch">
                              <div className="flex items-center justify-start pr-2 w-[160px] xl:w-[200px]">
                                <img
                                  src="/jagoan.png"
                                  className="w-full h-auto max-h-[160px] xl:max-h-[180px] object-contain object-left drop-shadow-md"
                                  alt="Jagoan Advanta"
                                />
                              </div>
                              {renderedCards}
                            </div>
                            {/* Mobile / Tablet Layout */}
                            <div className="w-full flex flex-col gap-3 sm:gap-4 lg:hidden">
                              <div className="w-full flex flex-row gap-3 sm:gap-4 items-stretch">
                                <div className="flex items-center justify-start shrink-0 w-[110px] sm:w-[140px] md:w-[170px]">
                                  <img
                                    src="/jagoan.png"
                                    className="w-full h-auto max-h-[120px] sm:max-h-[140px] md:max-h-[160px] object-contain object-left drop-shadow-md"
                                    alt="Jagoan Advanta"
                                  />
                                </div>
                                <div className="flex-1 min-w-0 flex flex-col">
                                  {renderedCards[0]}
                                </div>
                              </div>
                              <div className="w-full flex flex-row gap-3 sm:gap-4 items-stretch">
                                <div className="flex-1 min-w-0 flex flex-col">
                                  {renderedCards[1]}
                                </div>
                                <div className="flex-1 min-w-0 flex flex-col">
                                  {renderedCards[2]}
                                </div>
                              </div>
                            </div>
                          </>
                        ) : (
                          renderedCards
                        )}`;

code = code.replace(targetCode, newCode);

fs.writeFileSync('src/App.tsx', code);
console.log("Done");
