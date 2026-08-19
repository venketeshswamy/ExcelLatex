/**
 * Standard LaTeX Math Macros for ExcelKaTeX.
 * Provides pre-configured, domain-specific math shortcuts for calculus, linear algebra,
 * statistics, physics, and logic.
 */

export type MacroDictionary = Record<string, string>;

export const DEFAULT_MACROS: MacroDictionary = {
  // --- Number Systems & Sets ---
  '\\R': '\\mathbb{R}',
  '\\N': '\\mathbb{N}',
  '\\Z': '\\mathbb{Z}',
  '\\Q': '\\mathbb{Q}',
  '\\C': '\\mathbb{C}',
  '\\Complex': '\\mathbb{C}',
  '\\Reals': '\\mathbb{R}',
  '\\Naturals': '\\mathbb{N}',
  '\\Integers': '\\mathbb{Z}',
  '\\Rationals': '\\mathbb{Q}',
  '\\set': '\\left\\{ #1 \\right\\}',
  '\\emptyset': '\\varnothing',

  // --- Linear Algebra & Vectors ---
  '\\vec': '\\mathbf{#1}',
  '\\bm': '\\boldsymbol{#1}',
  '\\mat': '\\begin{pmatrix} #1 \\end{pmatrix}',
  '\\bmat': '\\begin{bmatrix} #1 \\end{bmatrix}',
  '\\vmat': '\\begin{vmatrix} #1 \\end{vmatrix}',
  '\\norm': '\\left\\| #1 \\right\\|',
  '\\abs': '\\left| #1 \\right|',
  '\\inner': '\\left\\langle #1, #2 \\right\\rangle',
  '\\det': '\\operatorname{det}\\left(#1\\right)',
  '\\tr': '\\operatorname{tr}\\left(#1\\right)',
  '\\rank': '\\operatorname{rank}\\left(#1\\right)',
  '\\diag': '\\operatorname{diag}\\left(#1\\right)',
  '\\span': '\\operatorname{span}\\left(#1\\right)',
  '\\proj': '\\operatorname{proj}_{#1}\\left(#2\\right)',

  // --- Calculus & Analysis ---
  '\\d': '\\mathrm{d}#1',
  '\\diff': '\\frac{\\mathrm{d}#1}{\\mathrm{d}#2}',
  '\\ddiff': '\\frac{\\mathrm{d}^2 #1}{\\mathrm{d}#2^2}',
  '\\pd': '\\frac{\\partial #1}{\\partial #2}',
  '\\ppd': '\\frac{\\partial^2 #1}{\\partial #2^2}',
  '\\intinf': '\\int_{-\\infty}^{+\\infty}',
  '\\suminf': '\\sum_{#1=1}^{\\infty}',
  '\\limto': '\\lim_{#1 \\to #2}',
  '\\grad': '\\nabla #1',
  '\\div': '\\nabla \\cdot #1',
  '\\curl': '\\nabla \\times #1',
  '\\laplacian': '\\nabla^2 #1',

  // --- Probability & Statistics ---
  '\\E': '\\mathbb{E}\\left[ #1 \\right]',
  '\\Var': '\\operatorname{Var}\\left( #1 \\right)',
  '\\Cov': '\\operatorname{Cov}\\left( #1, #2 \\right)',
  '\\Prob': '\\mathbb{P}\\left( #1 \\right)',
  '\\std': '\\operatorname{std}\\left( #1 \\right)',
  '\\corr': '\\operatorname{Corr}\\left( #1, #2 \\right)',
  '\\MSE': '\\operatorname{MSE}\\left( #1 \\right)',
  '\\bias': '\\operatorname{Bias}\\left( #1 \\right)',
  '\\Gaussian': '\\mathcal{N}\\left( #1, #2 \\right)',

  // --- Physics & Engineering ---
  '\\hbar': '\\hslash',
  '\\ket': '\\left| #1 \\right\\rangle',
  '\\bra': '\\left\\langle #1 \\right|',
  '\\braket': '\\left\\langle #1 \\middle| #2 \\right\\rangle',
  '\\expectation': '\\left\\langle #1 \\right\\rangle',
  '\\ii': '\\mathrm{i}',
  '\\ee': '\\mathrm{e}',
  '\\dag': '\\dagger',
  '\\const': '\\text{const}',

  // --- Logic & Relations ---
  '\\defeq': '\\coloneqq',
  '\\coloneqq': ':=',
  '\\approxeq': '\\cong'
};

class MacroRegistry {
  private macros: MacroDictionary = { ...DEFAULT_MACROS };

  public getAll(): MacroDictionary {
    return { ...this.macros };
  }

  public get(name: string): string | undefined {
    const formattedName = name.startsWith('\\') ? name : `\\${name}`;
    return this.macros[formattedName];
  }

  public has(name: string): boolean {
    const formattedName = name.startsWith('\\') ? name : `\\${name}`;
    return formattedName in this.macros;
  }

  public register(name: string, expansion: string): void {
    const formattedName = name.startsWith('\\') ? name : `\\${name}`;
    this.macros[formattedName] = expansion;
  }

  public registerBatch(newMacros: MacroDictionary): void {
    for (const [name, expansion] of Object.entries(newMacros)) {
      this.register(name, expansion);
    }
  }

  public unregister(name: string): void {
    const formattedName = name.startsWith('\\') ? name : `\\${name}`;
    delete this.macros[formattedName];
  }

  public reset(): void {
    this.macros = { ...DEFAULT_MACROS };
  }
}

export const macroRegistry = new MacroRegistry();
