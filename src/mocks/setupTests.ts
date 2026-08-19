/**
 * Global Test Environment Setup for Vitest with JSDOM
 */
import { beforeEach, afterEach, vi } from 'vitest';
import { setupOfficeMock, resetOfficeMock } from './officeMock';
import { setupCustomFunctionsMock, resetCustomFunctionsMock } from './customFunctionsMock';

// 1. Mock HTMLCanvasElement.getContext('2d')
if (typeof window !== 'undefined') {
  const canvasStateStack: any[] = [];
  const mockCanvasContext2D = {
    fillStyle: '#000000',
    strokeStyle: '#000000',
    font: '16px KaTeX_Main',
    lineWidth: 1,
    textBaseline: 'top',
    fillText: vi.fn(),
    strokeText: vi.fn(),
    fillRect: vi.fn(),
    clearRect: vi.fn(),
    getImageData: vi.fn(() => ({
      data: new Uint8ClampedArray(4 * 100 * 100),
      width: 100,
      height: 100
    })),
    putImageData: vi.fn(),
    createImageData: vi.fn(),
    setTransform: vi.fn(),
    drawImage: vi.fn(),
    save: vi.fn(function (this: any) {
      canvasStateStack.push({
        fillStyle: this.fillStyle,
        strokeStyle: this.strokeStyle,
        font: this.font,
        lineWidth: this.lineWidth,
        textBaseline: this.textBaseline
      });
    }),
    restore: vi.fn(function (this: any) {
      const state = canvasStateStack.pop();
      if (state) {
        Object.assign(this, state);
      }
    }),
    beginPath: vi.fn(),
    moveTo: vi.fn(),
    lineTo: vi.fn(),
    closePath: vi.fn(),
    stroke: vi.fn(),
    fill: vi.fn(),
    scale: vi.fn(),
    translate: vi.fn(),
    rotate: vi.fn(),
    measureText: vi.fn((text: string) => ({
      width: text.length * 9,
      actualBoundingBoxAscent: 12,
      actualBoundingBoxDescent: 4,
      actualBoundingBoxLeft: 0,
      actualBoundingBoxRight: text.length * 9,
      fontBoundingBoxAscent: 14,
      fontBoundingBoxDescent: 4
    }))
  };

  HTMLCanvasElement.prototype.getContext = function (contextType: string, _options?: any): any {
    if (contextType === '2d') {
      return mockCanvasContext2D;
    }
    return null;
  };

  HTMLCanvasElement.prototype.toDataURL = function (type = 'image/png', _quality?: any): string {
    return `data:${type};base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==`;
  };

  // Mock Range.prototype.getClientRects and getBoundingClientRect for JSDOM
  if (typeof Range !== 'undefined') {
    if (!Range.prototype.getClientRects) {
      Range.prototype.getClientRects = function (): DOMRectList {
        const rect = {
          x: 10,
          y: 10,
          width: 20,
          height: 15,
          top: 10,
          right: 30,
          bottom: 25,
          left: 10,
          toJSON: () => ({})
        } as DOMRect;
        return [rect] as unknown as DOMRectList;
      };
    }
    if (!Range.prototype.getBoundingClientRect) {
      Range.prototype.getBoundingClientRect = function (): DOMRect {
        return {
          x: 10,
          y: 10,
          width: 20,
          height: 15,
          top: 10,
          right: 30,
          bottom: 25,
          left: 10,
          toJSON: () => ({})
        } as DOMRect;
      };
    }
  }

  // 2. Mock Image constructor with synchronous onload handling for data URIs
  class MockImage {
    public width = 100;
    public height = 40;
    public naturalWidth = 100;
    public naturalHeight = 40;
    public complete = true;
    public onload: (() => void) | null = null;
    public onerror: ((error: any) => void) | null = null;
    private _src = '';

    get src(): string {
      return this._src;
    }

    set src(value: string) {
      this._src = value;
      // Auto-trigger onload asynchronously in next tick or synchronously
      setTimeout(() => {
        if (this.onload) {
          this.onload();
        }
      }, 0);
    }

    addEventListener(event: string, callback: () => void): void {
      if (event === 'load') {
        this.onload = callback;
      } else if (event === 'error') {
        this.onerror = callback;
      }
    }

    removeEventListener(_event: string, _callback: () => void): void {
      // no-op
    }
  }

  (window as any).Image = MockImage;

  // 3. Mock matchMedia
  window.matchMedia =
    window.matchMedia ||
    function (query: string) {
      return {
        matches: query.includes('dark') ? false : false,
        media: query,
        onchange: null,
        addListener: vi.fn(),
        removeListener: vi.fn(),
        addEventListener: vi.fn(),
        removeEventListener: vi.fn(),
        dispatchEvent: vi.fn()
      } as any;
    };

  // 4. Mock ResizeObserver
  class MockResizeObserver {
    public observe = vi.fn();
    public unobserve = vi.fn();
    public disconnect = vi.fn();
  }
  (window as any).ResizeObserver = MockResizeObserver;
  (globalThis as any).ResizeObserver = MockResizeObserver;

  // 5. Mock URL.createObjectURL / revokeObjectURL
  if (!window.URL.createObjectURL) {
    window.URL.createObjectURL = vi.fn((_blob: any) => `blob:mock-url-${Math.random()}`);
  }
  if (!window.URL.revokeObjectURL) {
    window.URL.revokeObjectURL = vi.fn();
  }
}

// 6. Global Mock Initialization
setupOfficeMock();
setupCustomFunctionsMock();

beforeEach(() => {
  resetOfficeMock();
  resetCustomFunctionsMock();
});

afterEach(() => {
  vi.clearAllMocks();
});
