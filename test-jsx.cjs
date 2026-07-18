const esbuild = require('esbuild');
try {
  esbuild.transformSync(require('fs').readFileSync('src/App.tsx', 'utf-8'), { loader: 'tsx' });
} catch (e) {
  console.log(e.message);
}
