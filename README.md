npm install
npm install sortablejs
npm install --save-dev @types/xrm

 Add below in "compilerOptions" in tsconfig.json

  "target": "ES5",
    "lib": ["ES5", "ES2015", "DOM"],
    "module": "commonjs"