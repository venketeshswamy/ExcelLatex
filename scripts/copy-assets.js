import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');

function copyDirRecursive(src, dest) {
  if (!fs.existsSync(src)) {
    console.warn(`Source directory not found: ${src}`);
    return;
  }

  if (!fs.existsSync(dest)) {
    fs.mkdirSync(dest, { recursive: true });
  }

  const entries = fs.readdirSync(src, { withFileTypes: true });

  for (const entry of entries) {
    const srcPath = path.join(src, entry.name);
    const destPath = path.join(dest, entry.name);

    if (entry.isDirectory()) {
      copyDirRecursive(srcPath, destPath);
    } else {
      fs.copyFileSync(srcPath, destPath);
    }
  }
}

function copyFile(src, dest) {
  if (!fs.existsSync(src)) {
    console.warn(`Source file not found: ${src}`);
    return;
  }
  const destDir = path.dirname(dest);
  if (!fs.existsSync(destDir)) {
    fs.mkdirSync(destDir, { recursive: true });
  }
  fs.copyFileSync(src, dest);
}

console.log('--- Copying KaTeX & MathLive Offline Assets ---');

// 1. KaTeX Fonts
const katexFontsSrc = path.join(rootDir, 'node_modules', 'katex', 'dist', 'fonts');
const katexFontsDest = path.join(rootDir, 'public', 'assets', 'katex-fonts');
copyDirRecursive(katexFontsSrc, katexFontsDest);
console.log(`[KaTeX Fonts] -> ${katexFontsDest}`);

const katexFontsRelativeDest = path.join(rootDir, 'public', 'assets', 'fonts');
copyDirRecursive(katexFontsSrc, katexFontsRelativeDest);

const katexFontsRootDest = path.join(rootDir, 'public', 'fonts');
copyDirRecursive(katexFontsSrc, katexFontsRootDest);

// 2. KaTeX CSS
const katexCssSrc = path.join(rootDir, 'node_modules', 'katex', 'dist', 'katex.min.css');
const katexCssDest = path.join(rootDir, 'public', 'assets', 'katex.min.css');
copyFile(katexCssSrc, katexCssDest);
console.log(`[KaTeX CSS] -> ${katexCssDest}`);

// 3. MathLive Fonts
const mathliveFontsSrc = path.join(rootDir, 'node_modules', 'mathlive', 'dist', 'fonts');
const mathliveFontsDest = path.join(rootDir, 'public', 'assets', 'mathlive-fonts');
copyDirRecursive(mathliveFontsSrc, mathliveFontsDest);
console.log(`[MathLive Fonts] -> ${mathliveFontsDest}`);

// 4. MathLive Sounds / other assets if present
const mathliveSoundsSrc = path.join(rootDir, 'node_modules', 'mathlive', 'dist', 'sounds');
const mathliveSoundsDest = path.join(rootDir, 'public', 'assets', 'mathlive-sounds');
if (fs.existsSync(mathliveSoundsSrc)) {
  copyDirRecursive(mathliveSoundsSrc, mathliveSoundsDest);
  console.log(`[MathLive Sounds] -> ${mathliveSoundsDest}`);
}

console.log('--- Asset Copy Complete ---');
