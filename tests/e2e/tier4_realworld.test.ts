import { describe, it, expect, beforeEach } from 'vitest';
import { compileLatex, compileLatexToHtml, validateLatex } from '../../src/core/katexEngine';
import { mathKatexFunction } from '../helpers/testHelpers';
import { resetOfficeMock, ExcelMock } from '../../src/mocks/officeMock';
import { resetCustomFunctionsMock, CustomFunctionsMock } from '../../src/mocks/customFunctionsMock';

describe('Tier 4: Real-World Enterprise Mathematical Scenarios Suite', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
  });

  describe('4.1 Advanced Calculus & Differential Equations', () => {
    it('compiles Navier-Stokes Incompressible Momentum Equation', async () => {
      const navierStokes = '\\rho \\left( \\frac{\\partial \\mathbf{u}}{\\partial t} + \\mathbf{u} \\cdot \\nabla \\mathbf{u} \\right) = -\\nabla p + \\mu \\nabla^2 \\mathbf{u} + \\mathbf{f}';
      
      const validation = validateLatex(navierStokes);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');

      const render = await compileLatex(navierStokes, { fontSize: 16 });
      expect(render.html).toContain('katex');
      expect(render.pngDataUrl.startsWith('data:image/')).toBe(true);
    });

    it('compiles Time-Dependent Schrödinger Equation in 3D', async () => {
      const schrodinger = 'i\\hbar \\frac{\\partial}{\\partial t}\\Psi(\\mathbf{r},t) = \\left[ -\\frac{\\hbar^2}{2m}\\nabla^2 + V(\\mathbf{r},t) \\right]\\Psi(\\mathbf{r},t)';
      
      const validation = validateLatex(schrodinger);
      expect(validation.isValid).toBe(true);

      const html = compileLatexToHtml(schrodinger);
      expect(html).toContain('katex');
    });

    it('compiles Infinite Taylor Series Expansion with Big-O notation', async () => {
      const taylor = 'f(x) = \\sum_{n=0}^{\\infty} \\frac{f^{(n)}(a)}{n!} (x - a)^n = f(a) + f\'(a)(x-a) + \\frac{f\'\'(a)}{2!}(x-a)^2 + \\mathcal{O}((x-a)^3)';
      
      const validation = validateLatex(taylor);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.tokens).toBeGreaterThan(30);
    });

    it('compiles Forward and Inverse Continuous Fourier Transforms', async () => {
      const fourierPair = '\\hat{f}(\\xi) = \\int_{-\\infty}^{\\infty} f(x) e^{-2\\pi i x \\xi} dx, \\quad f(x) = \\int_{-\\infty}^{\\infty} \\hat{f}(\\xi) e^{2\\pi i x \\xi} d\\xi';
      
      const validation = validateLatex(fourierPair);
      expect(validation.isValid).toBe(true);
    });
  });

  describe('4.2 Modern Physics & Field Theory', () => {
    it('compiles Differential Maxwell Equations Set', async () => {
      const maxwell = `\\begin{aligned}
        \\nabla \\cdot \\mathbf{E} &= \\frac{\\rho}{\\varepsilon_0} \\\\
        \\nabla \\cdot \\mathbf{B} &= 0 \\\\
        \\nabla \\times \\mathbf{E} &= -\\frac{\\partial \\mathbf{B}}{\\partial t} \\\\
        \\nabla \\times \\mathbf{B} &= \\mu_0 \\mathbf{J} + \\mu_0 \\varepsilon_0 \\frac{\\partial \\mathbf{E}}{\\partial t}
      \\end{aligned}`;

      const validation = validateLatex(maxwell);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');
    });

    it('compiles Einstein Field Equations with Cosmological Constant', async () => {
      const einsteinField = 'G_{\\mu\\nu} + \\Lambda g_{\\mu\\nu} = \\frac{8\\pi G}{c^4} T_{\\mu\\nu}';
      
      const validation = validateLatex(einsteinField);
      expect(validation.isValid).toBe(true);
    });

    it('compiles Relativistic Dirac Equation', async () => {
      const dirac = '\\left( i\\gamma^\\mu \\partial_\\mu - m \\right) \\psi = 0';
      
      const validation = validateLatex(dirac);
      expect(validation.isValid).toBe(true);
    });

    it('compiles Quantum Electrodynamics (QED) Lagrangian Density', async () => {
      const qed = '\\mathcal{L}_{\\text{QED}} = \\bar{\\psi}(i\\gamma^\\mu D_\\mu - m)\\psi - \\frac{1}{4}F_{\\mu\\nu}F^{\\mu\\nu}';
      
      const validation = validateLatex(qed);
      expect(validation.isValid).toBe(true);
    });
  });

  describe('4.3 Quantitative Finance & Actuarial Equations', () => {
    it('compiles Black-Scholes Analytical Option Pricing Model & Greeks', async () => {
      const blackScholes = `\\begin{aligned}
        C(S_t, t) &= S_t N(d_1) - K e^{-r(T - t)} N(d_2) \\\\
        d_1 &= \\frac{\\ln(S_t/K) + \\left(r + \\frac{\\sigma^2}{2}\\right)(T - t)}{\\sigma\\sqrt{T - t}} \\\\
        d_2 &= d_1 - \\sigma\\sqrt{T - t}
      \\end{aligned}`;

      const validation = validateLatex(blackScholes);
      expect(validation.isValid).toBe(true);

      const inv = CustomFunctionsMock.createInvocation('Sheet1!B2');
      const entity = (await mathKatexFunction(blackScholes, undefined, undefined, 14, true, 'cell', inv)) as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });

    it('compiles Capital Asset Pricing Model (CAPM) with Expected Return', async () => {
      const capm = '\\mathbb{E}[R_i] = R_f + \\beta_i \\left( \\mathbb{E}[R_m] - R_f \\right) + \\epsilon_i';
      
      const validation = validateLatex(capm);
      expect(validation.isValid).toBe(true);
    });

    it('compiles Heston Stochastic Volatility Couple SDEs', async () => {
      const heston = `\\begin{cases}
        dS_t = \\mu S_t dt + \\sqrt{v_t} S_t dW_t^S \\\\
        dv_t = \\kappa(\\theta - v_t) dt + \\xi \\sqrt{v_t} dW_t^v \\\\
        d\\langle W^S, W^v \\rangle_t = \\rho dt
      \\end{cases}`;

      const validation = validateLatex(heston);
      expect(validation.isValid).toBe(true);
      expect(validation.complexity?.score).toBe('Advanced');
    });
  });

  describe('4.4 Advanced Statistics & Machine Learning', () => {
    it('compiles Multivariate Gaussian Probability Density Function', async () => {
      const multiGaussian = 'f_{\\mathbf{X}}(x_1,\\dots,x_k) = \\frac{\\exp\\left(-\\frac{1}{2}(\\mathbf{x}-\\boldsymbol{\\mu})^T\\boldsymbol{\\Sigma}^{-1}(\\mathbf{x}-\\boldsymbol{\\mu})\\right)}{\\sqrt{(2\\pi)^k |\\boldsymbol{\\Sigma}|}}';
      
      const validation = validateLatex(multiGaussian);
      expect(validation.isValid).toBe(true);
    });

    it('compiles Bayes Rule with Continuous Integral Normalization', async () => {
      const bayes = 'P(\\theta \\mid \\mathcal{D}) = \\frac{P(\\mathcal{D} \\mid \\theta) P(\\theta)}{\\int_{\\Omega} P(\\mathcal{D} \\mid \\theta\') P(\\theta\') d\\theta\'}';
      
      const validation = validateLatex(bayes);
      expect(validation.isValid).toBe(true);
    });

    it('compiles Kullback-Leibler Divergence Definition', async () => {
      const klDiv = 'D_{\\text{KL}}(P \\parallel Q) = \\int_{-\\infty}^{\\infty} p(x) \\log\\left(\\frac{p(x)}{q(x)}\\right) dx';
      
      const validation = validateLatex(klDiv);
      expect(validation.isValid).toBe(true);
    });

    it('compiles Scaled Dot-Product Transformer Attention Mechanism', async () => {
      const attention = '\\operatorname{Attention}(\\mathbf{Q}, \\mathbf{K}, \\mathbf{V}) = \\operatorname{softmax}\\left( \\frac{\\mathbf{Q}\\mathbf{K}^T}{\\sqrt{d_k}} \\right) \\mathbf{V}';
      
      const validation = validateLatex(attention);
      expect(validation.isValid).toBe(true);
    });
  });
});
