const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const targetLayout = `                            {/* Desktop Layout */}
                            <div className="hidden lg:grid w-full lg:grid-cols-[auto_1fr_1fr_1fr] gap-3 lg:gap-4 items-stretch">
                              <div className="flex items-center justify-center pr-0 w-[140px] xl:w-[180px]">
                                <img
                                  src="/jagoan.png"
                                  className="w-full h-auto max-h-[160px] xl:max-h-[180px] object-contain object-center drop-shadow-md"
                                  alt="Jagoan Advanta"
                                />
                              </div>
                              {renderedCards}
                            </div>
                            {/* Mobile / Tablet Layout */}
                            <div className="w-full flex flex-col gap-3 sm:gap-4 lg:hidden">
                              <div className="w-full flex flex-row gap-2 sm:gap-3 items-stretch">
                                <div className="flex items-center justify-center shrink-0 w-[95px] sm:w-[125px] md:w-[155px]">
                                  <img
                                    src="/jagoan.png"
                                    className="w-full h-auto max-h-[120px] sm:max-h-[140px] md:max-h-[160px] object-contain object-center drop-shadow-md"
                                    alt="Jagoan Advanta"
                                  />
                                </div>`;

const newLayout = `                            {/* Desktop Layout */}
                            <div className="hidden lg:grid w-full lg:grid-cols-[auto_1fr_1fr_1fr] gap-2 lg:gap-3 items-stretch">
                              <div className="flex items-center justify-end pr-0 w-[130px] xl:w-[160px]">
                                <img
                                  src="/jagoan.png"
                                  className="w-full h-auto max-h-[160px] xl:max-h-[180px] object-contain object-right drop-shadow-md"
                                  alt="Jagoan Advanta"
                                />
                              </div>
                              {renderedCards}
                            </div>
                            {/* Mobile / Tablet Layout */}
                            <div className="w-full flex flex-col gap-3 sm:gap-4 lg:hidden">
                              <div className="w-full flex flex-row gap-1.5 sm:gap-2 items-stretch">
                                <div className="flex items-center justify-end shrink-0 w-[85px] sm:w-[110px] md:w-[140px]">
                                  <img
                                    src="/jagoan.png"
                                    className="w-full h-auto max-h-[120px] sm:max-h-[140px] md:max-h-[160px] object-contain object-right drop-shadow-md"
                                    alt="Jagoan Advanta"
                                  />
                                </div>`;

code = code.replace(targetLayout, newLayout);
fs.writeFileSync('src/App.tsx', code);
console.log("Icon layout updated!");
