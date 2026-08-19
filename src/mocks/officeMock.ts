/**
 * Headless Office.js & Excel Mock Framework for ExcelKaTeX Testing
 */

export namespace OfficeMock {
  export enum HostType {
    Excel = 'Excel',
    Word = 'Word',
    PowerPoint = 'PowerPoint',
    Outlook = 'Outlook'
  }

  export enum PlatformType {
    PC = 'PC',
    Mac = 'Mac',
    OfficeOnline = 'OfficeOnline',
    iOS = 'iOS',
    Android = 'Android'
  }

  export interface OfficeTheme {
    bodyBackgroundColor: string;
    bodyForegroundColor: string;
    controlBackgroundColor: string;
    controlForegroundColor: string;
    accentColor: string;
    isDark?: boolean;
  }
}

export namespace ExcelMock {
  export enum CellValueType {
    empty = 'Empty',
    string = 'String',
    number = 'Number',
    boolean = 'Boolean',
    error = 'Error',
    array = 'Array',
    entity = 'Entity',
    webImage = 'WebImage',
    formattedNumber = 'FormattedNumber'
  }

  export interface WebImageCellValue {
    type: CellValueType.webImage | 'WebImage';
    address?: string;
    altText?: string;
  }

  export interface EntityProperty {
    type: string;
    basicValue?: string | number | boolean;
    basicType?: string;
  }

  export interface EntityCellValue {
    type: CellValueType.entity | 'Entity';
    text: string;
    properties?: Record<string, EntityProperty | any>;
    layouts?: {
      compact?: {
        icon?: string;
      };
      card?: {
        title?: string;
        sections?: Array<{
          layout: string;
          properties: string[];
        }>;
      };
    };
    provider?: {
      description?: string;
      logoSourceAddress?: string;
    };
  }

  export class MockShape {
    public id: string;
    public name: string;
    public type: string;
    public left: number;
    public top: number;
    public width: number;
    public height: number;
    public altTextTitle: string = '';
    public altTextDescription: string = '';
    public visible: boolean = true;
    public lockAspectRatio: boolean = true;
    public imageContent?: string; // base64 or svg
    public isSvg: boolean = false;

    constructor(id: string, name: string, type = 'Image', left = 0, top = 0, width = 120, height = 40) {
      this.id = id;
      this.name = name;
      this.type = type;
      this.left = left;
      this.top = top;
      this.width = width;
      this.height = height;
    }

    public delete(): void {
      this.visible = false;
    }

    public scaleWidth(factor: number): void {
      this.width *= factor;
    }

    public scaleHeight(factor: number): void {
      this.height *= factor;
    }
  }

  export class MockShapeCollection {
    public items: MockShape[] = [];
    private shapeCounter = 1;

    public addImage(base64ImageString: string): MockShape {
      const id = `shape_${this.shapeCounter++}`;
      const shape = new MockShape(id, `Equation_${id}`, 'Image');
      shape.imageContent = base64ImageString;
      shape.isSvg = false;
      this.items.push(shape);
      return shape;
    }

    public addSvg(svgString: string): MockShape {
      const id = `shape_${this.shapeCounter++}`;
      const shape = new MockShape(id, `EquationSvg_${id}`, 'Svg');
      shape.imageContent = svgString;
      shape.isSvg = true;
      this.items.push(shape);
      return shape;
    }

    public getCount(): { value: number } {
      return { value: this.items.filter((s) => s.visible).length };
    }

    public getItem(key: string): MockShape | undefined {
      return this.items.find((s) => s.id === key || s.name === key);
    }

    public load(_props?: string | string[]): MockShapeCollection {
      return this;
    }
  }

  export class MockRange {
    public address: string;
    public values: any[][];
    public formulas: string[][];
    public numberFormat: string[][];
    public left: number = 0;
    public top: number = 0;
    public width: number = 80;
    public height: number = 22;
    public rowCount: number = 1;
    public columnCount: number = 1;
    public format: { font: { name: string; size: number; color: string } } = {
      font: { name: 'Calibri', size: 11, color: '#000000' }
    };
    private _loadedProperties: Set<string> = new Set();
    private _pendingLoads: Set<string> = new Set();

    constructor(address = 'A1', values: any[][] = [['']]) {
      this.address = address;
      this.values = values;
      this.formulas = values.map((row) => row.map((cell) => (typeof cell === 'string' && cell.startsWith('=') ? cell : '')));
      this.numberFormat = values.map((row) => row.map(() => 'General'));
      this._computeDefaultCoordinates(address);
    }

    private _computeDefaultCoordinates(address: string): void {
      try {
        const exclamationIdx = address.lastIndexOf('!');
        const cellPart = exclamationIdx !== -1 ? address.substring(exclamationIdx + 1) : address;
        const clean = cellPart.replace(/\$/g, '').toUpperCase();
        const match = clean.match(/^([A-Z]+)(\d+)$/);
        if (match) {
          const colStr = match[1];
          const rowNum = parseInt(match[2], 10);
          let colNum = 0;
          for (let i = 0; i < colStr.length; i++) {
            colNum = colNum * 26 + (colStr.charCodeAt(i) - 64);
          }
          this.left = (colNum - 1) * 80;
          this.top = (rowNum - 1) * 22;
        }
      } catch {
        this.left = 0;
        this.top = 0;
      }
    }

    public isPropertyLoaded(prop: string): boolean {
      return this._loadedProperties.has(prop) || this._loadedProperties.has('*');
    }

    public getLoadedProperties(): string[] {
      return Array.from(this._loadedProperties);
    }

    public load(props?: string | string[]): MockRange {
      if (!props) {
        this._pendingLoads.add('*');
        return this;
      }
      const propList = Array.isArray(props) ? props : props.split(',').map((p) => p.trim());
      for (const p of propList) {
        this._pendingLoads.add(p);
      }
      return this;
    }

    public _syncLoadedProperties(): void {
      for (const p of this._pendingLoads) {
        this._loadedProperties.add(p);
      }
      this._pendingLoads.clear();
    }

    public clear(): void {
      this.values = [['']];
      this.formulas = [['']];
    }

    public select(): void {
      // no-op mock
    }

    public getCell(_row: number, _column: number): MockRange {
      return this;
    }

    public getOffsetRange(_rowOffset: number, _columnOffset: number): MockRange {
      return this;
    }
  }

  export class MockWorksheet {
    public id: string;
    public name: string;
    public shapes: MockShapeCollection;
    public ranges: Map<string, MockRange> = new Map();
    public activeCellAddress: string = 'A1';

    constructor(id = 'sheet_1', name = 'Sheet1') {
      this.id = id;
      this.name = name;
      this.shapes = new MockShapeCollection();
    }

    public getRange(address = 'A1'): MockRange {
      const exclamationIdx = address.lastIndexOf('!');
      const cellPart = exclamationIdx !== -1 ? address.substring(exclamationIdx + 1) : address;
      const normalized = cellPart.replace(/\$/g, '').toUpperCase();
      if (!this.ranges.has(normalized)) {
        const range = new MockRange(normalized);
        this.ranges.set(normalized, range);
      }
      return this.ranges.get(normalized)!;
    }

    public getUsedRange(): MockRange {
      return this.getRange('A1:Z100');
    }

    public activate(): void {
      // active
    }

    public load(_props?: string | string[]): MockWorksheet {
      return this;
    }
  }

  export class MockWorksheetCollection {
    public items: MockWorksheet[] = [];
    public activeIndex = 0;

    constructor() {
      this.items.push(new MockWorksheet('sheet_1', 'Sheet1'));
    }

    public getActiveWorksheet(): MockWorksheet {
      return this.items[this.activeIndex] || this.items[0];
    }

    public getItem(name: string): MockWorksheet | undefined {
      return this.items.find((s) => s.name === name || s.id === name);
    }

    public add(name?: string): MockWorksheet {
      const sheet = new MockWorksheet(`sheet_${this.items.length + 1}`, name || `Sheet${this.items.length + 1}`);
      this.items.push(sheet);
      return sheet;
    }

    public load(_props?: string | string[]): MockWorksheetCollection {
      return this;
    }
  }

  export class MockWorkbook {
    public worksheets: MockWorksheetCollection;
    public activeCellAddress: string = 'A1';

    constructor() {
      this.worksheets = new MockWorksheetCollection();
    }

    public getSelectedRange(): MockRange {
      const activeSheet = this.worksheets.getActiveWorksheet();
      return activeSheet.getRange(this.activeCellAddress);
    }

    public load(_props?: string | string[]): MockWorkbook {
      return this;
    }
  }

  export class MockRequestContext {
    public workbook: MockWorkbook;
    public trackedObjects: { add: (obj: any) => void; remove: (obj: any) => void } = {
      add: () => {},
      remove: () => {}
    };

    constructor(workbook: MockWorkbook) {
      this.workbook = workbook;
    }

    public async sync(): Promise<void> {
      // Simulate asynchronous batch sync completion and property loading lifecycle
      for (const sheet of this.workbook.worksheets.items) {
        for (const range of sheet.ranges.values()) {
          range._syncLoadedProperties();
        }
      }
      return Promise.resolve();
    }

    public load(_obj: any, _props?: string | string[]): void {
      // no-op mock
    }
  }
}

export interface MockState {
  workbook: ExcelMock.MockWorkbook;
  theme: OfficeMock.OfficeTheme;
  supportedRequirementSets: Map<string, string | number>;
}

const mockState: MockState = {
  workbook: new ExcelMock.MockWorkbook(),
  theme: {
    bodyBackgroundColor: '#ffffff',
    bodyForegroundColor: '#000000',
    controlBackgroundColor: '#f3f2f1',
    controlForegroundColor: '#201f1e',
    accentColor: '#107c41',
    isDark: false
  },
  supportedRequirementSets: new Map([
    ['ExcelApi', 1.16],
    ['CustomFunctionsRuntime', 1.3],
    ['SharedRuntime', 1.1]
  ])
};

export function getMockExcelState(): MockState {
  return mockState;
}

export function resetOfficeMock(): void {
  mockState.workbook = new ExcelMock.MockWorkbook();
  mockState.theme = {
    bodyBackgroundColor: '#ffffff',
    bodyForegroundColor: '#000000',
    controlBackgroundColor: '#f3f2f1',
    controlForegroundColor: '#201f1e',
    accentColor: '#107c41',
    isDark: false
  };
  mockState.supportedRequirementSets = new Map([
    ['ExcelApi', 1.16],
    ['CustomFunctionsRuntime', 1.3],
    ['SharedRuntime', 1.1]
  ]);
  setupOfficeMock();
}

export function setMockTheme(theme: Partial<OfficeMock.OfficeTheme>): void {
  Object.assign(mockState.theme, theme);
}

export function setMockRequirementSupported(set: string, version: string | number, supported: boolean): void {
  if (supported) {
    mockState.supportedRequirementSets.set(set, version);
  } else {
    mockState.supportedRequirementSets.delete(set);
  }
}

export function setupOfficeMock(): void {
  const OfficeGlobal = {
    HostType: OfficeMock.HostType,
    PlatformType: OfficeMock.PlatformType,
    context: {
      host: OfficeMock.HostType.Excel,
      platform: OfficeMock.PlatformType.PC,
      officeTheme: mockState.theme,
      requirements: {
        isSetSupported: (name: string, minVersion: string | number) => {
          const supportedVer = mockState.supportedRequirementSets.get(name);
          if (supportedVer === undefined) return false;
          return Number(supportedVer) >= Number(minVersion);
        }
      },
      ui: {
        displayDialogAsync: (_url: string, _options: any, callback?: any) => {
          if (callback) callback({ status: 'succeeded' });
        }
      }
    },
    onReady: (callback?: (info: { host: string; platform: string }) => void) => {
      const info = { host: OfficeMock.HostType.Excel, platform: OfficeMock.PlatformType.PC };
      if (callback) callback(info);
      return Promise.resolve(info);
    }
  };

  const ExcelGlobal = {
    CellValueType: ExcelMock.CellValueType,
    run: async <T>(callback: (context: ExcelMock.MockRequestContext) => Promise<T>): Promise<T> => {
      const context = new ExcelMock.MockRequestContext(mockState.workbook);
      return await callback(context);
    }
  };

  const OfficeRuntimeGlobal = {
    displayWebDialog: async () => ({ close: () => {} }),
    storage: {
      getItem: async (_key: string) => null,
      setItem: async (_key: string, _value: string) => {},
      removeItem: async (_key: string) => {}
    }
  };

  (globalThis as any).Office = OfficeGlobal;
  (globalThis as any).Excel = ExcelGlobal;
  (globalThis as any).OfficeRuntime = OfficeRuntimeGlobal;
}
