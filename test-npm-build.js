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
const tempDir = path.join(process.cwd(), 'npm-test-temp');
if (fs.existsSync(tempDir)) fs.rmSync(tempDir, { recursive: true, force: true });
fs.mkdirSync(tempDir);

fs.writeFileSync(
  path.join(tempDir, 'package.json'),
  JSON.stringify(
    {
      name: 'npm-test-temp',
      version: '1.0.0',
      type: 'module'
    },
    null,
    2
  )
);

// Step 4: Install the tarball
console.log('Installing tarball in temp project...');
execSync(`npm install ../${tarball}`, { cwd: tempDir, stdio: 'inherit' });

// Step 5: Create and run import test
const testCode = `
import { parseVisioFile } from 'vsdx-js';

console.log('Testing imports:');
console.log('- parseVisioFile:', typeof parseVisioFile === 'function' ? 'OK' : 'FAIL');
console.log('All imports successful!');
`;

fs.writeFileSync(path.join(tempDir, 'test.mjs'), testCode);

console.log('Running import test...');
execSync('node test.mjs', { cwd: tempDir, stdio: 'inherit' });

console.log('Test complete. Cleaning up...');
fs.rmSync(tempDir, { recursive: true, force: true });
