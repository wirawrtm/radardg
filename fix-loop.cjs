const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const oldLoop = `        monthlyMap[cfg.key].opening += currentOpening;
        monthlyMap[cfg.key].ending += endingVal;
        monthlyMap[cfg.key].stockIn += stockInVal;
        monthlyMap[cfg.key].idle += idleVal;
        monthlyMap[cfg.key].pog += pogVal;

        currentOpening = endingVal;
      });
    });

    return monthsSequence.map((cfg) => monthlyMap[cfg.key]);`;

const newLoop = `        monthlyMap[cfg.key].opening += currentOpening;
        monthlyMap[cfg.key].ending += endingVal;
        monthlyMap[cfg.key].stockIn += stockInVal;
        monthlyMap[cfg.key].idle += idleVal;
        monthlyMap[cfg.key].pog += pogVal;

        monthlyMap[cfg.key].actualActivity += pogVal;
        monthlyMap[cfg.key].budgetActivity += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2));
        monthlyMap[cfg.key].actualNominal += Math.round(pogVal * 1.12) * 150000;
        monthlyMap[cfg.key].budgetNominal += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2)) * 150000;
        monthlyMap[cfg.key].actualReach += Math.round(pogVal * 0.8 * 1.05);
        monthlyMap[cfg.key].budgetReach += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2)) * 0.8;

        currentOpening = endingVal;
      });
    });

    return monthsSequence.map((cfg) => monthlyMap[cfg.key]);`;

code = code.replace(oldLoop, newLoop);
fs.writeFileSync('src/App.tsx', code);
