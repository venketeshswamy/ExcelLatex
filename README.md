# LaTeX Math — Excel Add-in

A fast, 100% offline LaTeX mathematical equation editor and custom function engine for Microsoft Excel.

## Features

- **Visual & Code Math Editor**: Interactive visual MathLive formula editor with bidirectional synchronization to raw LaTeX code.
- **`=MATH.KATEX()` Custom Function**: Insert mathematical equations directly using formulas with customizable backgrounds (`0: Transparent`, `1: White`, `2: Black`), text colors, font sizes, and display modes.
- **100% Offline & Local**: Built-in local KaTeX engine and font bundles with zero external CDN dependencies.
- **High-DPI In-Cell Rendering & Floating Shapes**: Insert clean vector/raster math graphics into cells or floating shapes.
- **Equation Catalog**: Pre-built library of common calculus, linear algebra, statistics, and physics equations.

## Getting Started

### Prerequisites

- Node.js (v18 or higher)
- Microsoft Excel (Desktop or Web)

### Installation

```bash
# Clone the repository
git clone https://github.com/venketeshswamy/ExcelLatex.git
cd ExcelLatex

# Install dependencies
npm install
```

### Development

```bash
# Start the local development server
npm run dev

# Sideload the add-in into Excel
npx office-addin-dev-settings sideload public/manifest.xml -a excel
```

### Testing & Verification

```bash
# Run unit, integration, and adversarial test suites
npm test
```

### Production Build

```bash
# Build the production bundle into dist/
npm run build
```

## License

MIT License
