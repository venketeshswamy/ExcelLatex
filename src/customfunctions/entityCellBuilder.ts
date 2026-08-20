/**
 * Entity Cell Value Builder for ExcelKaTeX.
 * Constructs modern Excel.EntityCellValue rich data types with embedded WebImageCellValue
 * and detailed formula metadata properties for in-cell mathematical rendering.
 */

import { RenderResult, RenderOptions, serializeEquationMetadata } from '../core/imageRasterizer';

/**
 * Checks whether the current Excel build supports ExcelApi 1.16 EntityCellValue.
 */
export function isEntityCellValueSupported(): boolean {
  try {
    const office = (globalThis as any).Office;
    if (typeof office !== 'undefined' && office?.context?.requirements?.isSetSupported) {
      return Boolean(
        office.context.requirements.isSetSupported('ExcelApi', 1.16) ||
        office.context.requirements.isSetSupported('ExcelApi', '1.16')
      );
    }
    return false;
  } catch {
    return false;
  }
}

export interface EntityProperty {
  type: string;
  basicValue: string | number | boolean;
}

export interface EntityCellValueLayouts {
  compact?: { icon?: string };
  card?: {
    title?: string;
    sections?: Array<{
      layout?: string;
      title?: string;
      properties?: string[];
    }>;
  };
}

export interface EntityCellValueProvider {
  description?: string;
  logoSourceAddress?: string;
}

export interface KatexWebImageCellValue {
  type: 'WebImage';
  address: string;
  altText: string;
}

export interface KatexEntityCellValue {
  type: 'Entity';
  text: string;
  properties?: Record<string, EntityProperty>;
  layouts?: EntityCellValueLayouts;
  provider?: EntityCellValueProvider;
}

/**
 * Builds an Excel.WebImageCellValue object representing an in-cell equation image.
 */
export function buildKatexWebImageCellValue(
  latex: string,
  renderResult: Partial<RenderResult> & { pngDataUrl?: string; width?: number; height?: number; svg?: string },
  options?: RenderOptions
): KatexWebImageCellValue {
  const cleanLatex = latex || '';
  const pngDataUrl = renderResult.pngDataUrl || '';
  const metadata = serializeEquationMetadata(cleanLatex, options);
  return {
    type: 'WebImage',
    address: pngDataUrl,
    altText: metadata
  };
}

/**
 * Builds an Excel.EntityCellValue object representing a rendered KaTeX equation.
 *
 * @param latex The LaTeX mathematical formula string.
 * @param renderResult The compiled rendering result containing SVG, PNG data URL, and dimensions.
 * @param options Optional rendering parameters (font size, display mode, background, color).
 */
export function buildKatexEntityCellValue(
  latex: string,
  renderResult: Partial<RenderResult> & { pngDataUrl?: string; width?: number; height?: number; svg?: string },
  options?: RenderOptions
): KatexEntityCellValue {
  const cleanLatex = latex || '';
  const pngDataUrl = renderResult.pngDataUrl || '';
  const width = renderResult.width ?? 120;
  const height = renderResult.height ?? 35;
  const svg = renderResult.svg || '';
  const fontSize = options?.fontSize ?? 16;
  const displayMode = options?.displayMode ?? true;

  const entity: KatexEntityCellValue = {
    type: 'Entity',
    text: `[Math: ${cleanLatex}]`,
    properties: {
      latex: {
        type: 'String',
        basicValue: cleanLatex
      },
      dimensions: {
        type: 'String',
        basicValue: `${width}x${height}`
      },
      image: {
        type: 'WebImage',
        address: pngDataUrl,
        altText: `LaTeX: ${cleanLatex}`
      },
      svg: {
        type: 'String',
        basicValue: svg
      },
      fontSize: {
        type: 'Double',
        basicValue: fontSize
      },
      displayMode: {
        type: 'Boolean',
        basicValue: displayMode
      }
    },
    layouts: {
      compact: {
        icon: pngDataUrl
      },
      card: {
        title: `LaTeX Equation: ${cleanLatex}`,
        sections: [
          {
            layout: 'List',
            title: 'Equation Metadata',
            properties: ['latex', 'dimensions', 'fontSize']
          }
        ]
      }
    },
    provider: {
      description: 'ExcelKaTeX Local Math Engine',
      logoSourceAddress: pngDataUrl
    }
  };

  return entity;
}
