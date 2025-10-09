import { execSync } from 'child_process';
import fs from 'fs';
import path from 'path';

// Step 1: Build the package
console.log('Building package...');
execSync('npm run build', { stdio: 'inherit' });

// Step 2: Pack the package
console.log('Packing package...');
const packOutput = execSync('npm pack').toString().trim();
const tarball = packOutput.split('\n').pop();

console.log(`Tarball created: ${tarball}`);

// Step 3: Create a temp test project
const tempDir = path.join(process.cwd(), 'npm-test-temp-ts');
if (fs.existsSync(tempDir)) fs.rmSync(tempDir, { recursive: true, force: true });
fs.mkdirSync(tempDir);

fs.writeFileSync(
  path.join(tempDir, 'package.json'),
  JSON.stringify(
    {
      name: 'npm-test-temp-ts',
      version: '1.0.0',
      type: 'module',
      devDependencies: {
        typescript: '^5.0.0',
        '@types/node': '^20.0.0'
      }
    },
    null,
    2
  )
);

// Step 4: Install dependencies and the tarball
console.log('Installing dependencies and tarball in temp project...');
execSync('npm install', { cwd: tempDir, stdio: 'inherit' });
execSync(`npm install ../${tarball}`, { cwd: tempDir, stdio: 'inherit' });

// Step 5: Create TypeScript config
const tsConfig = {
  compilerOptions: {
    target: 'ES2020',
    module: 'ESNext',
    moduleResolution: 'node',
    strict: true,
    esModuleInterop: true,
    skipLibCheck: true
  }
};

fs.writeFileSync(path.join(tempDir, 'tsconfig.json'), JSON.stringify(tsConfig, null, 2));

// Step 6: Create and compile TypeScript test
const tsTestCode = `
import { parseVisioFile, VisioFile, VisioShape, VisioPage } from 'vsdx-js';

console.log('TypeScript compilation test:');
console.log('- parseVisioFile function imported: OK');
console.log('- VisioFile type imported: OK');
console.log('- VisioShape type imported: OK');
console.log('- VisioPage type imported: OK');

// Test that types are properly defined
const testTypes = () => {
  // This should compile without errors if types are properly exported
  const dummyVisioFile: VisioFile = {
    Masters: [],
    Pages: [],
    Relationships: [],
    Stylesheets: []
  };
  console.log('- Type definitions work correctly: OK');
};

testTypes();
console.log('All TypeScript type imports successful!');
`;

fs.writeFileSync(path.join(tempDir, 'test.ts'), tsTestCode);

console.log('Compiling TypeScript test...');
execSync('npx tsc test.ts --outDir dist --module ESNext --target ES2020 --moduleResolution node', { cwd: tempDir, stdio: 'inherit' });

console.log('Running compiled TypeScript test...');
execSync('node dist/test.js', { cwd: tempDir, stdio: 'inherit' });

console.log('TypeScript test complete. Cleaning up...');
fs.rmSync(tempDir, { recursive: true, force: true });
