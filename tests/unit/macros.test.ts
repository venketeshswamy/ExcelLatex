import { describe, it, expect, beforeEach } from 'vitest';
import { macroRegistry, DEFAULT_MACROS } from '../../src/core/macros';

describe('macros.ts - MacroRegistry and DEFAULT_MACROS', () => {
  beforeEach(() => {
    macroRegistry.reset();
  });

  it('should contain standard number system macros', () => {
    expect(DEFAULT_MACROS['\\R']).toBe('\\mathbb{R}');
    expect(DEFAULT_MACROS['\\N']).toBe('\\mathbb{N}');
    expect(DEFAULT_MACROS['\\C']).toBe('\\mathbb{C}');
    expect(DEFAULT_MACROS['\\Z']).toBe('\\mathbb{Z}');
  });

  it('should contain linear algebra macros', () => {
    expect(DEFAULT_MACROS['\\vec']).toBe('\\mathbf{#1}');
    expect(DEFAULT_MACROS['\\mat']).toBe('\\begin{pmatrix} #1 \\end{pmatrix}');
    expect(DEFAULT_MACROS['\\norm']).toBe('\\left\\| #1 \\right\\|');
  });

  it('should allow registering and retrieving custom macros', () => {
    macroRegistry.register('myOp', '\\operatorname{myOp}\\left(#1\\right)');
    expect(macroRegistry.get('\\myOp')).toBe('\\operatorname{myOp}\\left(#1\\right)');

    macroRegistry.register('\\anotherOp', '123');
    expect(macroRegistry.get('\\anotherOp')).toBe('123');
  });

  it('should allow batch macro registration', () => {
    macroRegistry.registerBatch({
      '\\foo': 'FOO',
      '\\bar': 'BAR'
    });
    expect(macroRegistry.get('\\foo')).toBe('FOO');
    expect(macroRegistry.get('\\bar')).toBe('BAR');
  });

  it('should allow unregistering a macro', () => {
    macroRegistry.register('temp', 'TEMP');
    expect(macroRegistry.get('\\temp')).toBe('TEMP');
    macroRegistry.unregister('temp');
    expect(macroRegistry.get('\\temp')).toBeUndefined();
  });

  it('should reset back to DEFAULT_MACROS', () => {
    macroRegistry.register('custom', 'VALUE');
    expect(macroRegistry.get('\\custom')).toBe('VALUE');
    macroRegistry.reset();
    expect(macroRegistry.get('\\custom')).toBeUndefined();
    expect(macroRegistry.get('\\R')).toBe('\\mathbb{R}');
  });
});
