import { describe, it, expect } from 'vitest';
import { buildKatexEntityCellValue } from '../../src/customfunctions/entityCellBuilder';

describe('Entity Cell Builder Unit Tests', () => {
  it('builds standard Excel.EntityCellValue with correct type and text fallback', () => {
    const entity = buildKatexEntityCellValue('f(x) = x^2', {
      pngDataUrl: 'data:image/png;base64,mockPng',
      svg: '<svg>mock</svg>',
      width: 100,
      height: 40
    });

    expect(entity.type).toBe('Entity');
    expect(entity.text).toBe('[Math: f(x) = x^2]');
  });

  it('populates metadata properties for latex, dimensions, svg, fontSize, displayMode, and webImage', () => {
    const entity = buildKatexEntityCellValue(
      '\\sum_{k=1}^n k = \\frac{n(n+1)}{2}',
      {
        pngDataUrl: 'data:image/png;base64,data123',
        svg: '<svg>content</svg>',
        width: 180,
        height: 50
      },
      {
        fontSize: 20,
        displayMode: true
      }
    );

    expect(entity.properties?.latex?.type).toBe('String');
    expect(entity.properties?.latex?.basicValue).toBe('\\sum_{k=1}^n k = \\frac{n(n+1)}{2}');

    expect(entity.properties?.dimensions?.type).toBe('String');
    expect(entity.properties?.dimensions?.basicValue).toBe('180x50');

    expect(entity.properties?.image?.type).toBe('WebImage');
    expect(entity.properties?.image?.address).toBe('data:image/png;base64,data123');
    expect(entity.properties?.image?.altText).toContain('\\sum_{k=1}^n k');

    expect(entity.properties?.svg?.type).toBe('String');
    expect(entity.properties?.svg?.basicValue).toBe('<svg>content</svg>');

    expect(entity.properties?.fontSize?.basicValue).toBe(20);
    expect(entity.properties?.displayMode?.basicValue).toBe(true);
  });

  it('configures compact icon layout with high-DPI PNG data URL', () => {
    const entity = buildKatexEntityCellValue('a^2 + b^2 = c^2', {
      pngDataUrl: 'data:image/png;base64,compactIconTest',
      width: 120,
      height: 35
    });

    expect(entity.layouts?.compact?.icon).toBe('data:image/png;base64,compactIconTest');
  });

  it('configures rich card preview layout with section properties', () => {
    const entity = buildKatexEntityCellValue('\\int_0^\\infty e^{-x} dx = 1', {
      pngDataUrl: 'data:image/png;base64,cardIcon',
      width: 150,
      height: 45
    });

    expect(entity.layouts?.card?.title).toBe('LaTeX Equation: \\int_0^\\infty e^{-x} dx = 1');
    expect(entity.layouts?.card?.sections).toBeDefined();
    expect(entity.layouts?.card?.sections?.length).toBeGreaterThan(0);
    expect(entity.layouts?.card?.sections?.[0].properties).toContain('latex');
    expect(entity.layouts?.card?.sections?.[0].properties).toContain('dimensions');
  });

  it('attaches provider info for Excel rich data source attribution', () => {
    const entity = buildKatexEntityCellValue('\\lambda = 42', {
      pngDataUrl: 'data:image/png;base64,logo',
      width: 80,
      height: 30
    });

    expect(entity.provider?.description).toBe('ExcelKaTeX Local Math Engine');
    expect(entity.provider?.logoSourceAddress).toBe('data:image/png;base64,logo');
  });

  it('handles empty or missing optional properties gracefully', () => {
    const entity = buildKatexEntityCellValue('', {});
    expect(entity.type).toBe('Entity');
    expect(entity.text).toBe('[Math: ]');
    expect(entity.properties?.dimensions?.basicValue).toBe('120x35');
  });
});
