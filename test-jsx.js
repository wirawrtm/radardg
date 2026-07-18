const esbuild = require('esbuild');
esbuild.transformSync(require('fs').readFileSync('src/App.tsx', 'utf-8'), { loader: 'tsx' });
