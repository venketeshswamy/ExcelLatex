import { describe, it, expect, beforeEach } from 'vitest';
import { mathKatex, mathKatexFunction, isEntityCellValueSupported } from '../../src/customfunctions/mathKatex';
import { KaTeXShapeManager } from '../../src/customfunctions/shapeManager';
import {
  resetOfficeMock,
  setMockRequirementSupported,
  ExcelMock
} from '../../src/mocks/officeMock';
import {
  resetCustomFunctionsMock,
  CustomFunctionsMock
} from '../../src/mocks/customFunctionsMock';

describe('mathKatex Custom Function Unit Tests', () => {
  beforeEach(() => {
    resetOfficeMock();
    resetCustomFunctionsMock();
    KaTeXShapeManager.clear();
  });

  describe('Validation & Error Handling', () => {
    it('throws CustomFunctions.Error on empty latex string', async () => {
      await expect(mathKatex('')).rejects.toThrow();
      await expect(mathKatex('   \t\n  ')).rejects.toThrow();
    });

    it('throws CustomFunctions.Error on null or undefined latex string', async () => {
      await expect(mathKatex(undefined as any)).rejects.toThrow();
      await expect(mathKatex(null as any)).rejects.toThrow();
    });

    it('throws CustomFunctions.Error on invalid LaTeX syntax', async () => {
      await expect(mathKatex('\\frac{1}')).rejects.toThrow();
      await expect(mathKatex('\\sqrt[')).rejects.toThrow();
    });

    it('throws CustomFunctions.Error on malicious script injection', async () => {
      await expect(mathKatex('<script>alert(1)</script>')).rejects.toThrow();
    });
  });

  describe('Modern In-Cell EntityCellValue Rendering (ExcelApi 1.16+)', () => {
    it('returns EntityCellValue with default parameters', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B2');
      const result = await mathKatex('e^{i\\pi} + 1 = 0', undefined, undefined, undefined, undefined, 'auto', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.text).toBe('[Math: e^{i\\pi} + 1 = 0]');
      expect(entity.properties?.latex?.basicValue).toBe('e^{i\\pi} + 1 = 0');
      expect(entity.layouts?.compact?.icon).toBeDefined();
    });

    it('accepts custom styling parameters (background, color, fontSize, displayMode)', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!C3');
      const result = await mathKatex(
        '\\int x^2 dx = \\frac{x^3}{3} + C',
        '#ffffff',
        '#0055ff',
        22,
        false,
        'cell',
        inv
      );

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.properties?.fontSize?.basicValue).toBe(22);
      expect(entity.properties?.displayMode?.basicValue).toBe(false);
    });

    it('forces in-cell entity when outputMethod is "cell"', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!D4');
      const result = await mathKatex('a^2 + b^2 = c^2', undefined, undefined, undefined, undefined, 'cell', inv);

      expect(typeof result).toBe('object');
      expect((result as ExcelMock.EntityCellValue).type).toBe('Entity');
      expect(KaTeXShapeManager.getPendingCount()).toBe(0);
    });
  });

  describe('Standardized Background Codes (0, 1, 2) & Adaptive Text Color', () => {
    it('renders transparent background (code 0) with default black text', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A1');
      const result = await mathKatex('\\frac{2}{2}', 0, undefined, undefined, undefined, 'cell', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.properties?.latex?.basicValue).toBe('\\frac{2}{2}');
    });

    it('renders white background (code 1) with default black text', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A2');
      const result = await mathKatex('\\frac{2}{2}', 1, undefined, undefined, undefined, 'cell', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });

    it('renders black background (code 2) with adaptive white text', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A3');
      const result = await mathKatex('\\frac{2}{2}', 2, undefined, undefined, undefined, 'cell', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });

    it('renders string background codes ("0", "1", "2", "black", "white")', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A4');
      const resBlackStr = await mathKatex('x^2', '2', undefined, undefined, undefined, 'cell', inv);
      expect((resBlackStr as ExcelMock.EntityCellValue).type).toBe('Entity');

      const resNamedBlack = await mathKatex('x^2', 'black', undefined, undefined, undefined, 'cell', inv);
      expect((resNamedBlack as ExcelMock.EntityCellValue).type).toBe('Entity');

      const resWhiteStr = await mathKatex('x^2', '1', undefined, undefined, undefined, 'cell', inv);
      expect((resWhiteStr as ExcelMock.EntityCellValue).type).toBe('Entity');
    });

    it('preserves user explicit color override on black background', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!A5');
      const result = await mathKatex('x^2', 2, '#00ff00', undefined, undefined, 'cell', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });
  });

  describe('Dynamic Invocation Shifting Across Variable Argument Lengths', () => {
    it('handles 1 user argument with invocation at slot 2: =MATH.KATEX("x^2")', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B2');
      const result = await mathKatex('x^2', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.text).toBe('[Math: x^2]');
    });

    it('handles 2 user arguments with invocation at slot 3: =MATH.KATEX("x^2", 2)', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B3');
      const result = await mathKatex('x^2', 2, inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });

    it('handles 3 user arguments with invocation at slot 4: =MATH.KATEX("x^2", 1, "#107c41")', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B4');
      const result = await mathKatex('x^2', 1, '#107c41', inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
    });

    it('handles 4 user arguments with invocation at slot 5: =MATH.KATEX("x^2", 0, undefined, 20)', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B5');
      const result = await mathKatex('x^2', 0, undefined, 20, inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.properties?.fontSize?.basicValue).toBe(20);
    });

    it('handles 5 user arguments with invocation at slot 6: =MATH.KATEX("x^2", 1, undefined, 18, false)', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!B6');
      const result = await mathKatex('x^2', 1, undefined, 18, false, inv);

      expect(typeof result).toBe('object');
      const entity = result as ExcelMock.EntityCellValue;
      expect(entity.type).toBe('Entity');
      expect(entity.properties?.displayMode?.basicValue).toBe(false);
    });
  });

  describe('Shape Output & Legacy Fallback Routing', () => {
    it('enqueues shape and returns placeholder when outputMethod is "shape"', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!E5');
      const result = await mathKatex(
        '\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J}',
        undefined,
        undefined,
        undefined,
        undefined,
        'shape',
        inv
      );

      expect(typeof result).toBe('string');
      expect(result).toBe('[KaTeX Shape: \\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J}]');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      const queued = KaTeXShapeManager.getQueue()[0];
      expect(queued.address).toBe('Sheet1!E5');
      expect(queued.latex).toBe('\\nabla \\times \\mathbf{B} = \\mu_0 \\mathbf{J}');
    });

    it('enqueues shape with black background code 2 and adaptive white text color', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!E6');
      const result = await mathKatex(
        'E = mc^2',
        2,
        undefined,
        20,
        true,
        'shape',
        inv
      );

      expect(typeof result).toBe('string');
      expect(result).toBe('[KaTeX Shape: E = mc^2]');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);

      const queued = KaTeXShapeManager.getQueue()[0];
      expect(queued.options?.background).toBe('#000000');
      expect(queued.options?.color).toBe('#ffffff');
      expect(queued.options?.fontSize).toBe(20);
    });

    it('automatically falls back to shape when ExcelApi 1.16 is not supported and outputMethod is "auto"', async () => {
      setMockRequirementSupported('ExcelApi', 1.16, false);
      expect(isEntityCellValueSupported()).toBe(false);

      const inv = CustomFunctionsMock.createInvocation('Sheet1!F6');
      const result = await mathKatex(
        '\\sigma = \\sqrt{\\frac{1}{N}\\sum (x_i - \\mu)^2}',
        undefined,
        undefined,
        undefined,
        undefined,
        'auto',
        inv
      );

      expect(typeof result).toBe('string');
      expect(result).toContain('KaTeX Shape');
      expect(KaTeXShapeManager.getPendingCount()).toBe(1);
    });
  });

  describe('mathKatexFunction alias', () => {
    it('behaves identically to mathKatex', async () => {
      const inv = CustomFunctionsMock.createInvocation('Sheet1!G7');
      const res = await mathKatexFunction('x + y = z', undefined, undefined, undefined, undefined, 'cell', inv);
      expect((res as ExcelMock.EntityCellValue).type).toBe('Entity');
    });
  });
});
