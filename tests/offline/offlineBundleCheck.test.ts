import { describe, it, expect } from 'vitest';
import fs from 'fs';
import path from 'path';

describe('Offline & Zero-CDN Bundle Verification Suite', () => {
  const projectRoot = path.resolve(__dirname, '../../');
  const forbiddenDomains = [
    'unpkg.com',
    'cdn.jsdelivr.net',
    'cdnjs.cloudflare.com',
    'fonts.googleapis.com',
    'fonts.gstatic.com',
    'ajax.googleapis.com',
    'cdn.mathjax.org'
  ];

  function scanFileForForbiddenUrls(filePath: string): string[] {
    if (!fs.existsSync(filePath)) return [];
    const content = fs.readFileSync(filePath, 'utf-8');
    const violations: string[] = [];

    for (const domain of forbiddenDomains) {
      if (content.includes(domain)) {
        violations.push(`${filePath} contains reference to forbidden CDN: ${domain}`);
      }
    }
    return violations;
  }

  it('verifies public/manifest.xml contains zero external CDN references', () => {
    const manifestPath = path.join(projectRoot, 'public/manifest.xml');
    const violations = scanFileForForbiddenUrls(manifestPath);
    expect(violations).toEqual([]);
  });

  it('verifies public/taskpane.html contains zero external script CDNs', () => {
    const htmlPath = path.join(projectRoot, 'public/taskpane.html');
    const violations = scanFileForForbiddenUrls(htmlPath);
    expect(violations).toEqual([]);
  });

  it('verifies public/functions.json contains zero external endpoints', () => {
    const functionsPath = path.join(projectRoot, 'public/functions.json');
    const violations = scanFileForForbiddenUrls(functionsPath);
    expect(violations).toEqual([]);
  });

  it('verifies src/ code files contain zero remote CDN imports', () => {
    const srcDir = path.join(projectRoot, 'src');
    const allSrcFiles: string[] = [];

    function walkDir(dir: string) {
      if (!fs.existsSync(dir)) return;
      const entries = fs.readdirSync(dir, { withFileTypes: true });
      for (const entry of entries) {
        const fullPath = path.join(dir, entry.name);
        if (entry.isDirectory()) {
          walkDir(fullPath);
        } else if (entry.isFile() && (entry.name.endsWith('.ts') || entry.name.endsWith('.tsx') || entry.name.endsWith('.js'))) {
          allSrcFiles.push(fullPath);
        }
      }
    }

    walkDir(srcDir);
    expect(allSrcFiles.length).toBeGreaterThan(0);

    const violations: string[] = [];
    for (const file of allSrcFiles) {
      violations.push(...scanFileForForbiddenUrls(file));
    }

    expect(violations).toEqual([]);
  });

  it('verifies asset copy script exists for KaTeX and MathLive fonts packaging', () => {
    const copyScript = path.join(projectRoot, 'scripts/copy-assets.js');
    expect(fs.existsSync(copyScript)).toBe(true);

    const scriptContent = fs.readFileSync(copyScript, 'utf-8');
    expect(scriptContent).toContain('katex');
    expect(scriptContent).toContain('mathlive');
  });

  it('verifies MathfieldElement fontsDirectory points to local assets', () => {
    const taskpaneSrc = fs.readFileSync(path.join(projectRoot, 'src/taskpane/index.tsx'), 'utf-8');
    expect(taskpaneSrc).toContain('fontsDirectory');
    expect(taskpaneSrc).toContain('./assets/mathlive-fonts/');
  });
});
