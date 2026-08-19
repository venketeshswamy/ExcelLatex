import { performance } from 'perf_hooks';
import katex from 'katex';
import { LRUCache } from '../src/core/lruCache.ts';
import { validateLatex, calculateComplexity, compileLatexToHtml } from '../src/core/katexEngine.ts';
import { rasterizeLatex, clearEquationCache, getEquationCacheStats } from '../src/core/imageRasterizer.ts';
import { parseMathKatexParams } from '../src/customfunctions/parameterParser.ts';
import { buildKatexEntityCellValue } from '../src/customfunctions/entityCellBuilder.ts';
import { DEFAULT_MACROS } from '../src/core/macros.ts';

async function runEmpiricalBenchmarks() {
  console.log('================================================================');
  console.log('ExcelKaTeX Challenger 1: Empirical Stress & Benchmark Suite');
  console.log('================================================================\n');

  // -------------------------------------------------------------
  // Benchmark 1: Formula Validation & Complexity Scoring Throughput
  // -------------------------------------------------------------
  const formulas = [
    'x + 1',
    'E = mc^2',
    '\\frac{a+b}{c+d}',
    '\\int_{-\\infty}^\\infty e^{-x^2} dx = \\sqrt{\\pi}',
    '\\sum_{k=1}^n k^3 = \\left(\\frac{n(n+1)}{2}\\right)^2',
    '\\begin{pmatrix} \\cos\\theta & -\\sin\\theta \\\\ \\sin\\theta & \\cos\\theta \\end{pmatrix}',
    '\\nabla \\times \\mathbf{E} = -\\frac{\\partial \\mathbf{B}}{\\partial t}',
    '\\lim_{x \\to 0} \\frac{\\sin x}{x} = 1',
    'f(x) = \\begin{cases} x^2 & x \\ge 0 \\\\ -x & x < 0 \\end{cases}',
    '\\mathcal{H} |\\psi\\rangle = i\\hbar \\frac{\\partial}{\\partial t} |\\psi\\rangle'
  ];

  const N_VALIDATE = 5000;
  const t0_val = performance.now();
  for (let i = 0; i < N_VALIDATE; i++) {
    const f = formulas[i % formulas.length];
    validateLatex(f);
    calculateComplexity(f);
  }
  const t1_val = performance.now();
  const val_duration = t1_val - t0_val;
  const val_ops_per_sec = Math.round((N_VALIDATE / val_duration) * 1000);
  console.log(`[1] Validation & Complexity Throughput:`);
  console.log(`    - Executed ${N_VALIDATE} operations in ${val_duration.toFixed(2)} ms`);
  console.log(`    - Throughput: ${val_ops_per_sec.toLocaleString()} ops/sec (Average: ${(val_duration / N_VALIDATE).toFixed(4)} ms/op)\n`);

  // -------------------------------------------------------------
  // Benchmark 2: Deep Nesting Stress Test (50, 100, 150 levels)
  // -------------------------------------------------------------
  console.log(`[2] Deep Nesting Stress Tests:`);
  const nestingLevels = [25, 50, 100, 150];
  for (const depth of nestingLevels) {
    let nested = 'x';
    for (let d = 0; d < depth; d++) {
      nested = `\\frac{1}{${nested} + 1}`;
    }
    const t0_nest = performance.now();
    const val = validateLatex(nested);
    const html = compileLatexToHtml(nested, { throwOnError: false });
    const t1_nest = performance.now();
    console.log(`    - Nested Fraction (Depth ${depth}): valid=${val.isValid}, renderTime=${(t1_nest - t0_nest).toFixed(2)} ms, HTML length=${html.length} chars`);
  }
  console.log('');

  // -------------------------------------------------------------
  // Benchmark 3: Large Matrix Scaling (5x5, 10x10, 15x15, 20x20)
  // -------------------------------------------------------------
  console.log(`[3] Giant Matrix Scaling:`);
  const matrixSizes = [5, 10, 15, 20];
  for (const size of matrixSizes) {
    const rows = [];
    for (let r = 0; r < size; r++) {
      const cols = [];
      for (let c = 0; c < size; c++) {
        cols.push(`a_{${r + 1},${c + 1}}`);
      }
      rows.push(cols.join(' & '));
    }
    const matLatex = `\\begin{pmatrix} ${rows.join(' \\\\ ')} \\end{pmatrix}`;
    const t0_mat = performance.now();
    const val = validateLatex(matLatex);
    const html = compileLatexToHtml(matLatex);
    const t1_mat = performance.now();
    console.log(`    - Matrix ${size}x${size} (${size * size} cells): valid=${val.isValid}, compileTime=${(t1_mat - t0_mat).toFixed(2)} ms, HTML length=${html.length}`);
  }
  console.log('');

  // -------------------------------------------------------------
  // Benchmark 4: LRU Cache High-Throughput & Eviction Integrity
  // -------------------------------------------------------------
  console.log(`[4] LRU Cache Performance & Eviction Benchmark:`);
  const cache = new LRUCache({ maxSize: 1000 });
  const N_CACHE_OPS = 100000;
  const t0_cache = performance.now();
  for (let i = 0; i < N_CACHE_OPS; i++) {
    const key = `key_${i % 5000}`; // 5000 unique keys into 1000 capacity
    if (i % 2 === 0) {
      cache.get(key);
    } else {
      cache.set(key, { result: `data_${i}` });
    }
  }
  const t1_cache = performance.now();
  const cache_duration = t1_cache - t0_cache;
  const cacheStats = cache.getStats();
  console.log(`    - Executed ${N_CACHE_OPS.toLocaleString()} ops in ${cache_duration.toFixed(2)} ms`);
  console.log(`    - Throughput: ${Math.round((N_CACHE_OPS / cache_duration) * 1000).toLocaleString()} ops/sec`);
  console.log(`    - Final Size: ${cacheStats.size}/${cacheStats.maxSize}, Evictions: ${cacheStats.evictions}, Hits: ${cacheStats.hits}, Misses: ${cacheStats.misses}, Hit Ratio: ${(cacheStats.hitRatio * 100).toFixed(1)}%\n`);

  // -------------------------------------------------------------
  // Benchmark 5: Custom Function Parameter Parsing & Error Robustness
  // -------------------------------------------------------------
  console.log(`[5] Parameter Parsing & Adversarial Fuzzing:`);
  const adversarialInputs = [
    '',
    '   ',
    '\\frac{1}{',
    '\\begin{matrix} 1 & 2 \\end{aligned}',
    '\\unknownCmd{123}',
    '<script>alert(1)</script>',
    '\\text{<img src=x onerror=alert(1)>}',
    null,
    undefined,
    12345,
    'x + \\alpha\0 + \\beta\u0001'
  ];

  let errorsHandled = 0;
  const t0_fuzz = performance.now();
  for (let i = 0; i < 2000; i++) {
    const input = adversarialInputs[i % adversarialInputs.length];
    const outcome = parseMathKatexParams(input);
    if (!outcome.isValid) {
      errorsHandled++;
    }
  }
  const t1_fuzz = performance.now();
  console.log(`    - Tested 2,000 adversarial inputs in ${(t1_fuzz - t0_fuzz).toFixed(2)} ms`);
  console.log(`    - Handled ${errorsHandled}/2,000 invalid inputs with zero uncaught exceptions (100% safe)\n`);

  // -------------------------------------------------------------
  // Benchmark 6: End-to-End Rasterization & Entity Builder Pipeline
  // -------------------------------------------------------------
  console.log(`[6] End-to-End Image Rasterization & Entity Builder:`);
  clearEquationCache();
  const testEquations = [
    'E = mc^2',
    '\\int_0^1 \\sqrt{x} dx = \\frac{2}{3}',
    '\\sum_{n=1}^\\infty \\frac{1}{n^2} = \\frac{\\pi^2}{6}',
    '\\begin{bmatrix} a & b \\\\ c & d \\end{bmatrix}',
    '\\nabla \\cdot \\mathbf{E} = \\frac{\\rho}{\\varepsilon_0}'
  ];

  for (const eq of testEquations) {
    const res = await rasterizeLatex(eq, { scale: 3, fontSize: 16 });
    const entity = buildKatexEntityCellValue(eq, res, { fontSize: 16, displayMode: true });
    console.log(`    - Formula: "${eq}" -> Dimensions: ${res.width}x${res.height}, AspectRatio: ${res.aspectRatio.toFixed(2)}, EntityType: ${entity.type}, Text: "${entity.text}"`);
  }

  const eqStats = getEquationCacheStats();
  console.log(`    - Equation Cache Stats: size=${eqStats.size}, hits=${eqStats.hits}, misses=${eqStats.misses}\n`);

  console.log('================================================================');
  console.log('All Empirical Benchmarks Completed Successfully!');
  console.log('================================================================');
}

runEmpiricalBenchmarks().catch(console.error);
