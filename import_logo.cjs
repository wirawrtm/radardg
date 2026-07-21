const fs = require('fs');
let code = fs.readFileSync('src/App.tsx', 'utf-8');

code = code.replace(
  'import { AdvantaLogo } from "./AdvantaLogo";',
  'import { AdvantaLogo } from "./AdvantaLogo";\nimport jagoanLogo from "./assets/jagoan.png";'
);

code = code.replace(/src="\.\/jagoan\.png"/g, 'src={jagoanLogo}');

fs.writeFileSync('src/App.tsx', code);
console.log("Logo import added!");
