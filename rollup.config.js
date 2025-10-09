import typescript from '@rollup/plugin-typescript';
import { dts } from 'rollup-plugin-dts';

const config = [
  // JavaScript build
  {
    input: 'src/index.ts',
    output: {
      file: 'dist/index.js',
      format: 'es',
      sourcemap: true
    },
    plugins: [
      typescript({
        tsconfig: './tsconfig.json',
        declaration: false,
        declarationMap: false
      })
    ],
    external: ['adm-zip', 'fs', 'xml2js']
  },
  // TypeScript declarations build
  {
    input: 'src/index.ts',
    output: {
      file: 'dist/index.d.ts',
      format: 'es'
    },
    plugins: [dts()],
    external: ['adm-zip', 'fs', 'xml2js']
  }
];

export default config;
