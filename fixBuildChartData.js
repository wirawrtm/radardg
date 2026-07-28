const fs = require('fs');
const content = fs.readFileSync('src/App.tsx', 'utf8');

const regex = /const buildChartData = \(dimension: string, filterByMainKey\?: string \| null\) => \{([\s\S]*?)const activeMetric = \(overviewMetricFilter === "overview" \|\| overviewMetricFilter === "monitoring"\) \? overviewSubFilter : overviewMetricFilter;/;
const match = content.match(regex);
if (match) {
    console.log("Found buildChartData!");
} else {
    console.log("Not found.");
}
