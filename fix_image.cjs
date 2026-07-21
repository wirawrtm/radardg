const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

code = code.replace(
    /src="https:\/\/drive\.usercontent\.google\.com\/download\?id=1A0MkFXGsBDmXt67z5uED_jpVQ2QdXUdl&export=download"/g,
    'src="https://lh3.googleusercontent.com/d/1A0MkFXGsBDmXt67z5uED_jpVQ2QdXUdl=w1000"'
);

fs.writeFileSync('src/App.tsx', code);
console.log("Updated image link!");
