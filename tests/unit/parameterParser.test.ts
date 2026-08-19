import { describe, it, expect } from 'vitest';
import {
  parseMathKatexParams,
  sanitizeBackground,
  sanitizeTextColor,
  sanitizeColor,
  sanitizeFontSize,
  sanitizeDisplayMode,
  sanitizeOutputMethod,
  extractCellAddress,
  isInvocationObject,
  isDarkColor,
  getDefaultTextColor,
  stripLatexDelimiters
} from '../../src/customfunctions/parameterParser';

describe('Parameter Parser Unit Tests', () => {
  describe('sanitizeBackground', () => {
    it('maps numeric codes 0, 1, 2 to transparent, white, and black', () => {
      expect(sanitizeBackground(0)).toBe('transparent');
      expect(sanitizeBackground(1)).toBe('#ffffff');
      expect(sanitizeBackground(2)).toBe('#000000');
    });

    it('maps string digit codes "0", "1", "2" to transparent, white, and black', () => {
      expect(sanitizeBackground('0')).toBe('transparent');
      expect(sanitizeBackground('1')).toBe('#ffffff');
      expect(sanitizeBackground('2')).toBe('#000000');
      expect(sanitizeBackground(' 0 ')).toBe('transparent');
      expect(sanitizeBackground(' 1 ')).toBe('#ffffff');
      expect(sanitizeBackground(' 2 ')).toBe('#000000');
    });

    it('maps case-insensitive named codes transparent, white, black', () => {
      expect(sanitizeBackground('transparent')).toBe('transparent');
      expect(sanitizeBackground('TRANSPARENT')).toBe('transparent');
      expect(sanitizeBackground('white')).toBe('#ffffff');
      expect(sanitizeBackground('WHITE')).toBe('#ffffff');
      expect(sanitizeBackground('White')).toBe('#ffffff');
      expect(sanitizeBackground('black')).toBe('#000000');
      expect(sanitizeBackground('BLACK')).toBe('#000000');
      expect(sanitizeBackground('Black')).toBe('#000000');
    });

    it('supports custom hex colors and standard CSS color strings', () => {
      expect(sanitizeBackground('#ffffff')).toBe('#ffffff');
      expect(sanitizeBackground('#fff')).toBe('#fff');
      expect(sanitizeBackground('#107c41')).toBe('#107c41');
      expect(sanitizeBackground('rgba(0, 0, 0, 0.5)')).toBe('rgba(0, 0, 0, 0.5)');
      expect(sanitizeBackground('rgb(10, 20, 30)')).toBe('rgb(10, 20, 30)');
      expect(sanitizeBackground('red')).toBe('red');
      expect(sanitizeBackground('blue')).toBe('blue');
    });

    it('defaults to transparent for empty, null, undefined, invalid numbers, and invalid strings', () => {
      expect(sanitizeBackground(undefined)).toBe('transparent');
      expect(sanitizeBackground(null)).toBe('transparent');
      expect(sanitizeBackground('')).toBe('transparent');
      expect(sanitizeBackground('   ')).toBe('transparent');
      expect(sanitizeBackground(99)).toBe('transparent');
      expect(sanitizeBackground(-1)).toBe('transparent');
      expect(sanitizeBackground('not;a-color')).toBe('transparent');
    });
  });

  describe('isDarkColor & getDefaultTextColor (Adaptive Color Detection)', () => {
    it('identifies black and dark hex colors as dark', () => {
      expect(isDarkColor('#000000')).toBe(true);
      expect(isDarkColor('#000')).toBe(true);
      expect(isDarkColor('black')).toBe(true);
      expect(isDarkColor('#111111')).toBe(true);
      expect(isDarkColor('#1a1a1a')).toBe(true);
      expect(isDarkColor('rgb(0, 0, 0)')).toBe(true);
      expect(isDarkColor('darkblue')).toBe(true);
      expect(isDarkColor('navy')).toBe(true);
    });

    it('identifies light and transparent colors as non-dark', () => {
      expect(isDarkColor('transparent')).toBe(false);
      expect(isDarkColor('#ffffff')).toBe(false);
      expect(isDarkColor('#fff')).toBe(false);
      expect(isDarkColor('white')).toBe(false);
      expect(isDarkColor('#f0f0f0')).toBe(false);
      expect(isDarkColor('#ffff00')).toBe(false);
      expect(isDarkColor('rgb(255, 255, 255)')).toBe(false);
    });

    it('returns white default text color for dark backgrounds and black for light/transparent', () => {
      expect(getDefaultTextColor('#000000')).toBe('#ffffff');
      expect(getDefaultTextColor('#1a1a1a')).toBe('#ffffff');
      expect(getDefaultTextColor('transparent')).toBe('#000000');
      expect(getDefaultTextColor('#ffffff')).toBe('#000000');
    });
  });

  describe('sanitizeTextColor', () => {
    it('returns adaptive white text when background is black and color is unspecified', () => {
      expect(sanitizeTextColor(undefined, '#000000')).toBe('#ffffff');
      expect(sanitizeTextColor(null, '#000000')).toBe('#ffffff');
      expect(sanitizeTextColor('', '#000000')).toBe('#ffffff');
      expect(sanitizeTextColor('   ', '#000000')).toBe('#ffffff');
    });

    it('returns adaptive black text when background is white/transparent and color is unspecified', () => {
      expect(sanitizeTextColor(undefined, 'transparent')).toBe('#000000');
      expect(sanitizeTextColor(undefined, '#ffffff')).toBe('#000000');
      expect(sanitizeTextColor(null, '#ffffff')).toBe('#000000');
      expect(sanitizeTextColor('', 'transparent')).toBe('#000000');
    });

    it('preserves user explicit color override on black background', () => {
      expect(sanitizeTextColor('#00ff00', '#000000')).toBe('#00ff00');
      expect(sanitizeTextColor('#ffff00', '#000000')).toBe('#ffff00');
      expect(sanitizeTextColor('yellow', '#000000')).toBe('yellow');
      expect(sanitizeTextColor('#000000', '#000000')).toBe('#000000');
      expect(sanitizeTextColor('#ffffff', '#000000')).toBe('#ffffff');
    });

    it('preserves user explicit color override on white/transparent background', () => {
      expect(sanitizeTextColor('#107c41', '#ffffff')).toBe('#107c41');
      expect(sanitizeTextColor('#ff0000', 'transparent')).toBe('#ff0000');
      expect(sanitizeTextColor('blue', '#ffffff')).toBe('blue');
    });

    it('falls back to adaptive default if invalid color string is supplied', () => {
      expect(sanitizeTextColor('invalid;color', '#000000')).toBe('#ffffff');
      expect(sanitizeTextColor('invalid;color', 'transparent')).toBe('#000000');
    });
  });

  describe('sanitizeColor (Backward Compatibility)', () => {
    it('returns default color on undefined, null, or empty string', () => {
      expect(sanitizeColor(undefined, '#000000')).toBe('#000000');
      expect(sanitizeColor('', 'transparent')).toBe('transparent');
      expect(sanitizeColor('   ', 'red')).toBe('red');
    });

    it('accepts valid hex, rgb, rgba, hsl, hsla, transparent, and named colors', () => {
      expect(sanitizeColor('#fff', '#000')).toBe('#fff');
      expect(sanitizeColor('#107c41', '#000')).toBe('#107c41');
      expect(sanitizeColor('rgba(255, 0, 0, 0.5)', '#000')).toBe('rgba(255, 0, 0, 0.5)');
      expect(sanitizeColor('rgb(10, 20, 30)', '#000')).toBe('rgb(10, 20, 30)');
      expect(sanitizeColor('transparent', '#000')).toBe('transparent');
      expect(sanitizeColor('currentColor', '#000')).toBe('currentColor');
      expect(sanitizeColor('blue', '#000')).toBe('blue');
    });

    it('falls back to default for invalid colors', () => {
      expect(sanitizeColor('not;a-color', '#000000')).toBe('#000000');
    });
  });

  describe('sanitizeFontSize', () => {
    it('returns default 16 for undefined, null, NaN, or non-positive numbers', () => {
      expect(sanitizeFontSize(undefined)).toBe(16);
      expect(sanitizeFontSize(null)).toBe(16);
      expect(sanitizeFontSize('')).toBe(16);
      expect(sanitizeFontSize(NaN)).toBe(16);
      expect(sanitizeFontSize(0)).toBe(16);
      expect(sanitizeFontSize(-5)).toBe(16);
    });

    it('parses numeric strings and clamps between min (4) and max (1000)', () => {
      expect(sanitizeFontSize(2)).toBe(4);
      expect(sanitizeFontSize('14')).toBe(14);
      expect(sanitizeFontSize(14)).toBe(14);
      expect(sanitizeFontSize(48)).toBe(48);
      expect(sanitizeFontSize(2000)).toBe(1000);
      expect(sanitizeFontSize('invalid')).toBe(16);
    });
  });

  describe('sanitizeDisplayMode', () => {
    it('defaults to true when undefined, null, or empty string', () => {
      expect(sanitizeDisplayMode(undefined)).toBe(true);
      expect(sanitizeDisplayMode(null)).toBe(true);
      expect(sanitizeDisplayMode('')).toBe(true);
    });

    it('handles booleans, numbers, and strings', () => {
      expect(sanitizeDisplayMode(true)).toBe(true);
      expect(sanitizeDisplayMode(false)).toBe(false);
      expect(sanitizeDisplayMode(1)).toBe(true);
      expect(sanitizeDisplayMode(0)).toBe(false);
      expect(sanitizeDisplayMode('true')).toBe(true);
      expect(sanitizeDisplayMode('false')).toBe(false);
      expect(sanitizeDisplayMode('TRUE')).toBe(true);
      expect(sanitizeDisplayMode('FALSE')).toBe(false);
      expect(sanitizeDisplayMode('1')).toBe(true);
      expect(sanitizeDisplayMode('0')).toBe(false);
    });
  });

  describe('sanitizeOutputMethod', () => {
    it('returns default "auto" when undefined or invalid', () => {
      expect(sanitizeOutputMethod(undefined)).toBe('auto');
      expect(sanitizeOutputMethod('')).toBe('auto');
      expect(sanitizeOutputMethod('invalidMethod')).toBe('auto');
    });

    it('normalizes valid case-insensitive output methods', () => {
      expect(sanitizeOutputMethod('cell')).toBe('cell');
      expect(sanitizeOutputMethod('CELL')).toBe('cell');
      expect(sanitizeOutputMethod('shape')).toBe('shape');
      expect(sanitizeOutputMethod('SHAPE')).toBe('shape');
      expect(sanitizeOutputMethod('auto')).toBe('auto');
    });
  });

  describe('extractCellAddress & isInvocationObject', () => {
    it('detects invocation objects accurately', () => {
      expect(isInvocationObject(undefined)).toBe(false);
      expect(isInvocationObject('Sheet1!A1')).toBe(false);
      expect(isInvocationObject(1)).toBe(false);
      expect(isInvocationObject({ address: 'Sheet1!B2' })).toBe(true);
      expect(isInvocationObject({ parameterAddresses: [] })).toBe(true);
      expect(isInvocationObject({ functionName: 'KATEX' })).toBe(true);
      expect(isInvocationObject({})).toBe(true);
    });

    it('defaults to Sheet1!A1 when invocation is absent', () => {
      expect(extractCellAddress(undefined)).toBe('Sheet1!A1');
      expect(extractCellAddress(null)).toBe('Sheet1!A1');
    });

    it('extracts address from string or invocation object', () => {
      expect(extractCellAddress('Sheet2!C5')).toBe('Sheet2!C5');
      expect(extractCellAddress({ address: 'Sheet3!D10' })).toBe('Sheet3!D10');
      expect(extractCellAddress({ address: "'Quarterly Data'!$B$12" })).toBe("'Quarterly Data'!$B$12");
    });
  });

  describe('stripLatexDelimiters', () => {
    it('strips $$, \\[\\] \\(\\), and $ math delimiters', () => {
      expect(stripLatexDelimiters('$$x^2 + y^2$$')).toBe('x^2 + y^2');
      expect(stripLatexDelimiters('\\[ \\alpha + \\beta \\]')).toBe('\\alpha + \\beta');
      expect(stripLatexDelimiters('\\( E = mc^2 \\)')).toBe('E = mc^2');
      expect(stripLatexDelimiters('$a^2 + b^2 = c^2$')).toBe('a^2 + b^2 = c^2');
      expect(stripLatexDelimiters('  $$ \\frac{1}{2} $$  ')).toBe('\\frac{1}{2}');
    });

    it('leaves plain equations untouched', () => {
      expect(stripLatexDelimiters('x = y')).toBe('x = y');
    });
  });

  describe('parseMathKatexParams', () => {
    it('returns error when latexString is missing or empty', () => {
      const res1 = parseMathKatexParams(undefined as any);
      expect(res1.isValid).toBe(false);
      expect(res1.error).toBeDefined();
      expect(res1.errorCode).toBe('#VALUE!');

      const res2 = parseMathKatexParams('');
      expect(res2.isValid).toBe(false);
      expect(res2.error).toContain('empty');

      const res3 = parseMathKatexParams('   \t\n  ');
      expect(res3.isValid).toBe(false);
    });

    it('returns error when LaTeX syntax is invalid', () => {
      const res = parseMathKatexParams('\\frac{1}');
      expect(res.isValid).toBe(false);
      expect(res.error).toContain('KaTeX parse error');
    });

    it('parses valid parameters with defaults applied', () => {
      const res = parseMathKatexParams('E = mc^2');
      expect(res.isValid).toBe(true);
      expect(res.params).toBeDefined();
      expect(res.params?.latexString).toBe('E = mc^2');
      expect(res.params?.background).toBe('transparent');
      expect(res.params?.color).toBe('#000000');
      expect(res.params?.fontSize).toBe(16);
      expect(res.params?.displayMode).toBe(true);
      expect(res.params?.outputMethod).toBe('auto');
      expect(res.params?.cellAddress).toBe('Sheet1!A1');
    });

    it('parses numeric background codes 0, 1, 2 with adaptive text coloring', () => {
      // 0: Transparent bg -> Black text
      const res0 = parseMathKatexParams('x^2', 0);
      expect(res0.isValid).toBe(true);
      expect(res0.params?.background).toBe('transparent');
      expect(res0.params?.color).toBe('#000000');

      // 1: White bg -> Black text
      const res1 = parseMathKatexParams('x^2', 1);
      expect(res1.isValid).toBe(true);
      expect(res1.params?.background).toBe('#ffffff');
      expect(res1.params?.color).toBe('#000000');

      // 2: Black bg -> White text (Adaptive)
      const res2 = parseMathKatexParams('x^2', 2);
      expect(res2.isValid).toBe(true);
      expect(res2.params?.background).toBe('#000000');
      expect(res2.params?.color).toBe('#ffffff');
    });

    it('parses string background codes "0", "1", "2", "transparent", "white", "black"', () => {
      const res0 = parseMathKatexParams('x^2', '0');
      expect(res0.params?.background).toBe('transparent');
      expect(res0.params?.color).toBe('#000000');

      const res1 = parseMathKatexParams('x^2', '1');
      expect(res1.params?.background).toBe('#ffffff');
      expect(res1.params?.color).toBe('#000000');

      const res2 = parseMathKatexParams('x^2', '2');
      expect(res2.params?.background).toBe('#000000');
      expect(res2.params?.color).toBe('#ffffff');

      const resBlack = parseMathKatexParams('x^2', 'black');
      expect(resBlack.params?.background).toBe('#000000');
      expect(resBlack.params?.color).toBe('#ffffff');

      const resWhite = parseMathKatexParams('x^2', 'white');
      expect(resWhite.params?.background).toBe('#ffffff');
      expect(resWhite.params?.color).toBe('#000000');
    });

    it('honors explicit text color override on black background', () => {
      const res = parseMathKatexParams('x^2', 2, '#00ff00');
      expect(res.isValid).toBe(true);
      expect(res.params?.background).toBe('#000000');
      expect(res.params?.color).toBe('#00ff00');

      const resHex = parseMathKatexParams('x^2', 2, '#ffffff');
      expect(resHex.params?.background).toBe('#000000');
      expect(resHex.params?.color).toBe('#ffffff');
    });

    it('dynamically extracts invocation object from any parameter slot', () => {
      // Invocation in slot 2 (background position) -> =MATH.KATEX("x^2")
      const inv2 = { address: 'Sheet1!B2' };
      const resSlot2 = parseMathKatexParams('x^2', inv2);
      expect(resSlot2.isValid).toBe(true);
      expect(resSlot2.params?.cellAddress).toBe('Sheet1!B2');
      expect(resSlot2.params?.background).toBe('transparent');
      expect(resSlot2.params?.color).toBe('#000000');

      // Invocation in slot 3 (color position) -> =MATH.KATEX("x^2", 2)
      const inv3 = { address: 'Sheet1!C3' };
      const resSlot3 = parseMathKatexParams('x^2', 2, inv3);
      expect(resSlot3.isValid).toBe(true);
      expect(resSlot3.params?.cellAddress).toBe('Sheet1!C3');
      expect(resSlot3.params?.background).toBe('#000000');
      expect(resSlot3.params?.color).toBe('#ffffff');

      // Invocation in slot 4 (fontSize position) -> =MATH.KATEX("x^2", 1, "#107c41")
      const inv4 = { address: 'Sheet1!D4' };
      const resSlot4 = parseMathKatexParams('x^2', 1, '#107c41', inv4);
      expect(resSlot4.isValid).toBe(true);
      expect(resSlot4.params?.cellAddress).toBe('Sheet1!D4');
      expect(resSlot4.params?.background).toBe('#ffffff');
      expect(resSlot4.params?.color).toBe('#107c41');
      expect(resSlot4.params?.fontSize).toBe(16);

      // Invocation in slot 5 (displayMode position)
      const inv5 = { address: 'Sheet1!E5' };
      const resSlot5 = parseMathKatexParams('x^2', 0, undefined, 20, inv5);
      expect(resSlot5.isValid).toBe(true);
      expect(resSlot5.params?.cellAddress).toBe('Sheet1!E5');
      expect(resSlot5.params?.fontSize).toBe(20);
      expect(resSlot5.params?.displayMode).toBe(true);

      // Invocation in slot 6 (outputMethod position)
      const inv6 = { address: 'Sheet1!F6' };
      const resSlot6 = parseMathKatexParams('x^2', 1, undefined, 18, false, inv6);
      expect(resSlot6.isValid).toBe(true);
      expect(resSlot6.params?.cellAddress).toBe('Sheet1!F6');
      expect(resSlot6.params?.displayMode).toBe(false);
      expect(resSlot6.params?.outputMethod).toBe('auto');

      // Invocation in slot 7 (invocation position)
      const inv7 = { address: 'Sheet1!G7' };
      const resSlot7 = parseMathKatexParams('x^2', 1, '#107c41', 20, false, 'cell', inv7);
      expect(resSlot7.isValid).toBe(true);
      expect(resSlot7.params?.cellAddress).toBe('Sheet1!G7');
      expect(resSlot7.params?.outputMethod).toBe('cell');
    });

    it('resiliently handles omitted arguments and empty string commas', () => {
      // Simulated: =MATH.KATEX("x^2", , , 24)
      const resOmitted = parseMathKatexParams('x^2', undefined, undefined, 24);
      expect(resOmitted.isValid).toBe(true);
      expect(resOmitted.params?.background).toBe('transparent');
      expect(resOmitted.params?.color).toBe('#000000');
      expect(resOmitted.params?.fontSize).toBe(24);

      // Simulated: =MATH.KATEX("x^2", , , , , "shape")
      const resShapeOmitted = parseMathKatexParams('x^2', null, null, null, null, 'shape');
      expect(resShapeOmitted.isValid).toBe(true);
      expect(resShapeOmitted.params?.outputMethod).toBe('shape');

      // Trailing commas with empty strings
      const resEmptyStrings = parseMathKatexParams('x^2', '', '', '', '', '');
      expect(resEmptyStrings.isValid).toBe(true);
      expect(resEmptyStrings.params?.background).toBe('transparent');
      expect(resEmptyStrings.params?.color).toBe('#000000');
      expect(resEmptyStrings.params?.fontSize).toBe(16);
      expect(resEmptyStrings.params?.displayMode).toBe(true);
      expect(resEmptyStrings.params?.outputMethod).toBe('auto');
    });

    it('parses fully specified custom parameters', () => {
      const res = parseMathKatexParams(
        '\\int_0^1 x dx = \\frac{1}{2}',
        '#ffffff',
        '#107c41',
        24,
        false,
        'shape',
        { address: 'Calculus!B15' }
      );
      expect(res.isValid).toBe(true);
      expect(res.params?.latexString).toBe('\\int_0^1 x dx = \\frac{1}{2}');
      expect(res.params?.background).toBe('#ffffff');
      expect(res.params?.color).toBe('#107c41');
      expect(res.params?.fontSize).toBe(24);
      expect(res.params?.displayMode).toBe(false);
      expect(res.params?.outputMethod).toBe('shape');
      expect(res.params?.cellAddress).toBe('Calculus!B15');
    });
  });
});
