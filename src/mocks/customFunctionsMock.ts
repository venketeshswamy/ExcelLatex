/**
 * Headless CustomFunctions Mock Framework for ExcelKaTeX Testing
 */

export namespace CustomFunctionsMock {
  export enum ErrorCode {
    invalidValue = '#VALUE!',
    notAvailable = '#N/A',
    divisionByZero = '#DIV/0!',
    invalidNumber = '#NUM!',
    nullValue = '#NULL!',
    invalidReference = '#REF!',
    invalidName = '#NAME?'
  }

  export class CustomFunctionsError extends Error {
    public code: ErrorCode;
    public override message: string;

    constructor(code: ErrorCode, message?: string) {
      super(message || code);
      this.name = 'CustomFunctions.Error';
      this.code = code;
      this.message = message || code;
      Object.setPrototypeOf(this, CustomFunctionsError.prototype);
    }
  }

  export interface Invocation {
    address?: string;
    parameterAddresses?: string[];
    functionName?: string;
  }

  export function createInvocation(address = 'Sheet1!A1', parameterAddresses: string[] = []): Invocation {
    return {
      address,
      parameterAddresses,
      functionName: 'MATH.KATEX'
    };
  }
}

// Global runtime attachment helper
export function setupCustomFunctionsMock(): void {
  const customFunctionsGlobal = {
    Error: CustomFunctionsMock.CustomFunctionsError,
    ErrorCode: CustomFunctionsMock.ErrorCode
  };

  (globalThis as any).CustomFunctions = customFunctionsGlobal;
}

export function resetCustomFunctionsMock(): void {
  setupCustomFunctionsMock();
}
