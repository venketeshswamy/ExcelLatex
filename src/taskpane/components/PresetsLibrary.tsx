import React, { useState, useMemo } from 'react';
import {
  makeStyles,
  tokens,
  shorthands,
  Input,
  TabList,
  Tab,
  Card,
  Caption1,
  Subtitle2,
  Badge,
  Button
} from '@fluentui/react-components';
import {
  Search20Regular,
  Add20Regular,
  ArrowRight20Regular
} from '@fluentui/react-icons';
import { compileLatexToHtml } from '../../core/katexEngine';

export interface PresetFormula {
  id: string;
  name: string;
  category: 'Calculus' | 'Linear Algebra' | 'Statistics' | 'Physics' | 'Financial Math' | 'Algebra';
  latex: string;
  description?: string;
}

export const PRESET_CATEGORIES = [
  'All',
  'Calculus',
  'Linear Algebra',
  'Statistics',
  'Physics',
  'Financial Math'
] as const;

export const PRESET_FORMULAS: PresetFormula[] = [
  // Calculus
  {
    id: 'euler_identity',
    name: "Euler's Identity",
    category: 'Calculus',
    latex: 'e^{i\\pi} + 1 = 0',
    description: 'Fundamental relationship between five fundamental mathematical constants.'
  },
  {
    id: 'ftc',
    name: 'Fundamental Theorem of Calculus',
    category: 'Calculus',
    latex: '\\int_a^b f\'(x) \\, dx = f(b) - f(a)',
    description: 'Relates differentiation with integration.'
  },
  {
    id: 'gaussian_integral',
    name: 'Gaussian Integral',
    category: 'Calculus',
    latex: '\\int_{-\\infty}^{+\\infty} e^{-x^2} \\, dx = \\sqrt{\\pi}',
    description: 'Integral of the Gaussian function over the entire real line.'
  },
  {
    id: 'taylor_series',
    name: 'Taylor Series Expansion',
    category: 'Calculus',
    latex: 'f(x) = \\sum_{n=0}^{\\infty} \\frac{f^{(n)}(a)}{n!} (x - a)^n',
    description: 'Representation of a function as an infinite sum of terms.'
  },
  {
    id: 'fourier_transform',
    name: 'Fourier Transform',
    category: 'Calculus',
    latex: '\\hat{f}(\\xi) = \\int_{-\\infty}^{\\infty} f(x) e^{-2\\pi i x \\xi} \\, dx',
    description: 'Transforms a function from time domain to frequency domain.'
  },
  {
    id: 'navier_stokes',
    name: 'Navier-Stokes Momentum',
    category: 'Calculus',
    latex: '\\rho \\left( \\frac{\\partial \\mathbf{u}}{\\partial t} + \\mathbf{u} \\cdot \\nabla \\mathbf{u} \\right) = -\\nabla p + \\mu \\nabla^2 \\mathbf{u} + \\mathbf{f}',
    description: 'Describes the motion of viscous fluid substances.'
  },

  // Linear Algebra
  {
    id: 'eigenvalue',
    name: 'Eigenvalue Equation',
    category: 'Linear Algebra',
    latex: 'A \\mathbf{v} = \\lambda \\mathbf{v}',
    description: 'Characteristic equation for linear transformations.'
  },
  {
    id: 'matrix_inverse_2x2',
    name: '2x2 Matrix Inverse',
    category: 'Linear Algebra',
    latex: 'A^{-1} = \\frac{1}{ad - bc} \\begin{pmatrix} d & -b \\\\ -c & a \\end{pmatrix}',
    description: 'Analytical inverse formula for a 2x2 invertible matrix.'
  },
  {
    id: 'determinant_2x2',
    name: '2x2 Matrix Determinant',
    category: 'Linear Algebra',
    latex: '\\det(A) = \\begin{vmatrix} a & b \\\\ c & d \\end{vmatrix} = ad - bc',
    description: 'Determinant calculation for a 2x2 matrix.'
  },
  {
    id: 'matrix_mult',
    name: 'Matrix Multiplication',
    category: 'Linear Algebra',
    latex: 'C_{ij} = \\sum_{k=1}^n A_{ik} B_{kj}',
    description: 'General formula for the element of a product matrix.'
  },
  {
    id: 'dot_product',
    name: 'Vector Dot Product',
    category: 'Linear Algebra',
    latex: '\\mathbf{u} \\cdot \\mathbf{v} = \\|\\mathbf{u}\\| \\|\\mathbf{v}\\| \\cos(\\theta)',
    description: 'Geometric definition of scalar product between vectors.'
  },

  // Statistics
  {
    id: 'normal_pdf',
    name: 'Normal Distribution PDF',
    category: 'Statistics',
    latex: 'f(x) = \\frac{1}{\\sigma\\sqrt{2\\pi}} e^{-\\frac{1}{2}\\left(\\frac{x-\\mu}{\\sigma}\\right)^2}',
    description: 'Probability density function of Gaussian normal distribution.'
  },
  {
    id: 'bayes_theorem',
    name: "Bayes' Theorem",
    category: 'Statistics',
    latex: 'P(A|B) = \\frac{P(B|A) P(A)}{P(B)}',
    description: 'Calculates conditional probability of an event based on prior knowledge.'
  },
  {
    id: 'standard_deviation',
    name: 'Standard Deviation (Population)',
    category: 'Statistics',
    latex: '\\sigma = \\sqrt{\\frac{1}{N} \\sum_{i=1}^N (x_i - \\mu)^2}',
    description: 'Measure of the amount of variation or dispersion of a set of values.'
  },
  {
    id: 'binomial_dist',
    name: 'Binomial Distribution PMF',
    category: 'Statistics',
    latex: 'P(X=k) = \\binom{n}{k} p^k (1-p)^{n-k}',
    description: 'Discrete probability distribution of successes in n Bernoulli trials.'
  },
  {
    id: 'ols_slope',
    name: 'OLS Regression Slope',
    category: 'Statistics',
    latex: '\\hat{\\beta}_1 = \\frac{\\sum (x_i - \\bar{x})(y_i - \\bar{y})}{\\sum (x_i - \\bar{x})^2}',
    description: 'Ordinary least squares estimate for the linear slope coefficient.'
  },

  // Physics
  {
    id: 'schrodinger',
    name: 'Schrödinger Wave Equation',
    category: 'Physics',
    latex: 'i\\hbar \\frac{\\partial}{\\partial t} \\Psi(\\mathbf{r}, t) = \\hat{H} \\Psi(\\mathbf{r}, t)',
    description: 'Quantum state evolution governed by the Hamiltonian operator.'
  },
  {
    id: 'mass_energy',
    name: 'Mass-Energy Equivalence',
    category: 'Physics',
    latex: 'E = mc^2',
    description: "Einstein's principle of equivalence of mass and energy."
  },
  {
    id: 'maxwell_faraday',
    name: "Maxwell's Induction Law",
    category: 'Physics',
    latex: '\\nabla \\times \\mathbf{E} = -\\frac{\\partial \\mathbf{B}}{\\partial t}',
    description: 'Maxwell-Faraday equation of electromagnetic induction.'
  },
  {
    id: 'gravitation',
    name: "Newton's Law of Gravitation",
    category: 'Physics',
    latex: 'F = G \\frac{m_1 m_2}{r^2}',
    description: 'Attractive gravitational force between two point masses.'
  },
  {
    id: 'lorentz_factor',
    name: 'Lorentz Factor',
    category: 'Physics',
    latex: '\\gamma = \\frac{1}{\\sqrt{1 - \\frac{v^2}{c^2}}}',
    description: 'Relativistic factor in special relativity time dilation and length contraction.'
  },

  // Financial Math
  {
    id: 'black_scholes',
    name: 'Black-Scholes Call Option',
    category: 'Financial Math',
    latex: 'C = S_t N(d_1) - K e^{-rt} N(d_2)',
    description: 'Theoretical pricing model for European-style call options.'
  },
  {
    id: 'compound_interest',
    name: 'Compound Interest Formula',
    category: 'Financial Math',
    latex: 'A = P \\left( 1 + \\frac{r}{n} \\right)^{nt}',
    description: 'Future value calculation with periodic interest compounding.'
  },
  {
    id: 'annuity_pv',
    name: 'Present Value of Ordinary Annuity',
    category: 'Financial Math',
    latex: 'PV = PMT \\times \\left[ \\frac{1 - (1 + r)^{-n}}{r} \\right]',
    description: 'Present discounted value of a series of equal future cash flows.'
  },
  {
    id: 'capm',
    name: 'Capital Asset Pricing Model (CAPM)',
    category: 'Financial Math',
    latex: 'E(R_i) = R_f + \\beta_i \\left( E(R_m) - R_f \\right)',
    description: 'Expected return on an asset based on systematic market risk.'
  },
  {
    id: 'sharpe_ratio',
    name: 'Sharpe Ratio',
    category: 'Financial Math',
    latex: 'S = \\frac{R_p - R_f}{\\sigma_p}',
    description: 'Measure of risk-adjusted return relative to risk-free benchmark.'
  }
];

const useStyles = makeStyles({
  container: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('12px'),
    width: '100%'
  },
  header: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('6px')
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1
  },
  tabList: {
    overflowX: 'auto',
    ...shorthands.borderBottom('1px', 'solid', tokens.colorNeutralStroke2)
  },
  grid: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap('10px'),
    maxHeight: '400px',
    overflowY: 'auto',
    ...shorthands.padding('2px', '4px')
  },
  presetCard: {
    flexShrink: 0,
    ...shorthands.padding('10px', '12px'),
    backgroundColor: tokens.colorNeutralBackground1,
    ...shorthands.border('1px', 'solid', tokens.colorNeutralStroke2),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    transitionProperty: 'box-shadow, border-color',
    transitionDuration: '0.15s',
    ':hover': {
      boxShadow: tokens.shadow4,
      ...shorthands.borderColor(tokens.colorCompoundBrandStroke)
    }
  },
  cardTop: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    marginBottom: '6px'
  },
  formulaDisplay: {
    ...shorthands.padding('8px', '10px'),
    backgroundColor: tokens.colorNeutralBackground2,
    ...shorthands.borderRadius(tokens.borderRadiusSmall),
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: '40px',
    overflowX: 'auto',
    margin: '6px 0'
  },
  actionsRow: {
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'flex-end',
    ...shorthands.gap('6px'),
    marginTop: '6px'
  }
});

export interface PresetsLibraryProps {
  onSelectPreset: (latex: string) => void;
  onAppendPreset?: (latex: string) => void;
}

export const PresetsLibrary: React.FC<PresetsLibraryProps> = ({
  onSelectPreset,
  onAppendPreset
}) => {
  const styles = useStyles();
  const [selectedCategory, setSelectedCategory] = useState<string>('All');
  const [searchQuery, setSearchQuery] = useState<string>('');

  const filteredPresets = useMemo(() => {
    return PRESET_FORMULAS.filter((item) => {
      const matchCategory =
        selectedCategory === 'All' || item.category === selectedCategory;
      const q = searchQuery.toLowerCase().trim();
      const matchSearch =
        !q ||
        item.name.toLowerCase().includes(q) ||
        item.category.toLowerCase().includes(q) ||
        item.latex.toLowerCase().includes(q) ||
        (item.description && item.description.toLowerCase().includes(q));
      return matchCategory && matchSearch;
    });
  }, [selectedCategory, searchQuery]);

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Subtitle2 className={styles.title}>Formula Preset Catalog</Subtitle2>
        <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
          Explore standard academic, scientific, and financial equations
        </Caption1>

        <Input
          contentBefore={<Search20Regular />}
          placeholder="Search formulas (e.g. Euler, Black-Scholes, integral)..."
          value={searchQuery}
          onChange={(_e, data) => setSearchQuery(data.value)}
          aria-label="Search formula presets"
        />

        <TabList
          selectedValue={selectedCategory}
          onTabSelect={(_e, data) => setSelectedCategory(data.value as string)}
          className={styles.tabList}
          size="small"
        >
          {PRESET_CATEGORIES.map((cat) => (
            <Tab key={cat} value={cat}>
              {cat}
            </Tab>
          ))}
        </TabList>
      </div>

      <div className={styles.grid} role="list" aria-label="Preset Equations">
        {filteredPresets.length === 0 ? (
          <div style={{ textAlign: 'center', padding: '20px', color: tokens.colorNeutralForeground4 }}>
            No presets found matching "{searchQuery}".
          </div>
        ) : (
          filteredPresets.map((item) => {
            const html = compileLatexToHtml(item.latex, { displayMode: true });
            return (
              <Card key={item.id} className={styles.presetCard} role="listitem">
                <div className={styles.cardTop}>
                  <Subtitle2 style={{ fontWeight: tokens.fontWeightSemibold }}>
                    {item.name}
                  </Subtitle2>
                  <Badge appearance="tint" color="brand" size="small">
                    {item.category}
                  </Badge>
                </div>

                {item.description && (
                  <Caption1 style={{ color: tokens.colorNeutralForeground3 }}>
                    {item.description}
                  </Caption1>
                )}

                <div
                  className={styles.formulaDisplay}
                  dangerouslySetInnerHTML={{ __html: html }}
                />

                <div className={styles.actionsRow}>
                  {onAppendPreset && (
                    <Button
                      size="small"
                      appearance="subtle"
                      icon={<Add20Regular />}
                      onClick={() => onAppendPreset(item.latex)}
                      title="Append to active equation"
                    >
                      Append
                    </Button>
                  )}
                  <Button
                    size="small"
                    appearance="primary"
                    icon={<ArrowRight20Regular />}
                    onClick={() => onSelectPreset(item.latex)}
                    title="Load as current equation"
                  >
                    Load Equation
                  </Button>
                </div>
              </Card>
            );
          })
        )}
      </div>
    </div>
  );
};
