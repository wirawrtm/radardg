const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

// We will replace the initial structure of monthlyMap to include budget and actual metrics
const initMapOld = `        opening: 0,
        ending: 0,
        stockIn: 0,
        idle: 0,
        pog: 0,
      };`;
const initMapNew = `        opening: 0,
        ending: 0,
        stockIn: 0,
        idle: 0,
        pog: 0,
        budgetActivity: 0,
        actualActivity: 0,
        budgetNominal: 0,
        actualNominal: 0,
        budgetReach: 0,
        actualReach: 0,
      };`;
code = code.replace(initMapOld, initMapNew);

const mapDefOld = `        idle: number;
        pog: number;
      }
    > = {};`;
const mapDefNew = `        idle: number;
        pog: number;
        budgetActivity: number;
        actualActivity: number;
        budgetNominal: number;
        actualNominal: number;
        budgetReach: number;
        actualReach: number;
      }
    > = {};`;
code = code.replace(mapDefOld, mapDefNew);

const monthlyLoopOld = `        monthlyMap[cfg.key].opening += currentOpening;
        monthlyMap[cfg.key].ending += endingVal;
        monthlyMap[cfg.key].stockIn += stockInVal;
        monthlyMap[cfg.key].idle += idleVal;
        monthlyMap[cfg.key].pog += pogVal;
        currentOpening = endingVal;
      });`;
const monthlyLoopNew = `        monthlyMap[cfg.key].opening += currentOpening;
        monthlyMap[cfg.key].ending += endingVal;
        monthlyMap[cfg.key].stockIn += stockInVal;
        monthlyMap[cfg.key].idle += idleVal;
        monthlyMap[cfg.key].pog += pogVal;
        
        monthlyMap[cfg.key].actualActivity += pogVal;
        monthlyMap[cfg.key].budgetActivity += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2));
        monthlyMap[cfg.key].actualNominal += pogVal * 150000;
        monthlyMap[cfg.key].budgetNominal += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2)) * 150000;
        monthlyMap[cfg.key].actualReach += Math.round(pogVal * 0.8);
        monthlyMap[cfg.key].budgetReach += Math.round(pogVal * 1.15 + (pogVal > 0 ? 5 : 2)) * 0.8;
        
        currentOpening = endingVal;
      });`;
code = code.replace(monthlyLoopOld, monthlyLoopNew);

fs.writeFileSync('src/App.tsx', code);
console.log("Updated overviewHistoryData structure successfully");
