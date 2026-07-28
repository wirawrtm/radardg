const fs = require('fs');
const content = fs.readFileSync('src/App.tsx', 'utf8');
const match = content.match(/const overviewStats = useMemo\(\(\) => \{([\s\S]*?)\}, \[\n?\s*kiosks/);
if (match) console.log(match[1].length);
else console.log("Not found");
