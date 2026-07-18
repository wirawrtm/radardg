const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

const oldTrendStart = '            <div className="h-64 w-full font-sans">';
const oldTrendEnd = '              {/* End of Activity Performance Trend Chart */}'; // Wait, let's just find where the area chart ends.
