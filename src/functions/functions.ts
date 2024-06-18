/* global clearInterval, console, CustomFunctions, setInterval */

/// <reference path="../office-experiment52.d.ts" />

import path from "path";

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Echo the input. If input is null, return "null".
 * @customfunction
 * @param {any} input
 * @returns {any} The input value.
 */
export function echo(input: any): any {
  if (input == null) {
    return "null";
  }
  return input;
}

/**
 * Creates a formatted number.
 * @customfunction
 * @param {number} input
 * @param {string} format
 * @returns {any} The formatted number.
 */
function createFormattedNumber(value, format) {
  return {
    type: "FormattedNumber",
    basicValue: value,
    numberFormat: format
  }
}

/**
 * Creates a PwM for number.
 * @customfunction
 * @param {number} value
 * @returns {any}
 */
function createPWMForNumber(value: number) {
  return {
    type: Excel.CellValueType.double,
    basicValue: value,
    basicType: Excel.RangeValueType.double,
    properties:
    {
      Name:
      {
        type: Excel.CellValueType.string,
        basicValue: "Metadata for the number"
      }
    },
    layouts:
    {
      compact:
      {
        icon: Excel.EntityCompactLayoutIcons.airplane,
      }
    }
  }
}

/**
 * Returns a result of input + 1 for type 'any'.
 * @customfunction
 * @param input
 * @returns
 */
function plusOneForAny(input: any): any {
  if (typeof input === "number") {
    return input + 1;
  } else if (typeof input === "object") {
    if (input.type === Excel.CellValueType.double) {
      input.basicValue = input.basicValue + 1;
    } else if (input.type === Excel.CellValueType.formattedNumber) {
      input.basicValue = input.basicValue + 1;
    }
    return input;
  }

  return input;
}

/**
 * Returns a result of input + 1 for number.
 * @customfunction
 * @param {number} input
 * @returns {number}
 */
function plusOneForNumber(input: number): number {
  let result = input + 1;
  return result;
}

/**
 * [NOTE] No customfunction tag as the new types will lead to error.
 * Returns a result of input + 1 for Excel.DoubleCellValue.
 * @param {Excel.DoubleCellValue} input
 * @returns {Excel.DoubleCellValue}
 */
function plusOneForDoubleCellValue(input: Excel.DoubleCellValue): Excel.DoubleCellValue {
  input.basicValue = input.basicValue + 1;
  return input;
}

/**
 * [NOTE] No customfunction tag as the new types will lead to error.
 * Returns a result of input + 1 for Excel.FormattedNumberCellValue.
 * @param {Excel.FormattedNumberCellValue} input
 * @returns {Excel.FormattedNumberCellValue}
 */
function plusOneForFormattedNumberCellValue(input: Excel.FormattedNumberCellValue): Excel.FormattedNumberCellValue {
  input.basicValue = input.basicValue + 1;
  return input;
}

/**
 * Streaming function that returns an entity every interval seconds.
 * @customfunction
 * @param {any} dependency
 * @param {number} interval
 * @param {CustomFunctions.StreamingInvocation<any>} invocation
 */
function testStreaming(dependency: any, interval: number, invocation: CustomFunctions.StreamingInvocation<any>): void {
  let result = 0;
  let resEntity = {
    type: "Entity",
    text: "Entity " + result,
    properties: {
      propNumber: {
        type: "Double",
        basicValue: 123,
      },
    }
  };

  const timer = setInterval(() => {
    result += 1;
    resEntity.text = "Entity " + result;
    invocation.setResult(resEntity);
  }, interval * 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * @customfunction
 * @param {any[]} input
 * @param {CustomFunctions.Invocation} invocation
 * @returns {Promise<string>} Concate the input array.
 * @requiresParameterAddresses
 */
async function testRepeatingParameter(input: any[], invocation: CustomFunctions.Invocation): Promise<string> {
  let result = "";
  const context = new Excel.RequestContext();
  let sheet = context.workbook.worksheets.getActiveWorksheet();
  for (let index = 0; index < input.length; index++) {
    const element = input[index];
    if (element === 0 && invocation.parameterAddresses[index] != undefined) {
      let range = sheet.getRange(invocation.parameterAddresses[index]).load("text");
      await context.sync();
      if (range.text[0][0] == "") {
        result += "[]"; // '0' comes from the empty cell.
      } else {
        result += range.text[0][0]; // '0' is the real value.
      }
    } else {
      result += element;
    }
  }

  return result;
}


/**
 * Simulate latency and return the number in millisecond.
 * @customfunction
 * @param {number} latency Average latency in millisecond
 * @param {any} dependency Only for triggering chained calc.
 * @returns {Promise<number>}
 */
function returnAfterAsyncLatency(latency: number, dependency?: any) {
  let simulateLatency = (Math.random() * 2 - 1) * 1000 + latency;
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(Math.floor(simulateLatency));
    }, simulateLatency);
  });
}


/**
 * Return latency in millisecond after sleep.
 * @customfunction
 * @param {number} latency Time to sleep in millisecond before return.
 * @param {any} dependency Only for triggering chained calc.
 * @returns {Promise<number>}
 */
function returnAfterSleep(latency: number, dependency?: any) {
  let date = new Date().getTime();
  let curDate = null;
  do { curDate = new Date().getTime(); }
  while (curDate - date < latency);
  return latency;
}

/**
 * Returns a simple entity.
 * @customfunction
 * @returns {any} A simple entity.
 */
function getSimpleEntity() {
  console.log(`Start getSimpleEntity`);
  let randomValue = Math.floor(Math.random() * 100);
  const entity = {
    type: Excel.CellValueType.entity,
    text: "Random Entity " + randomValue,
    properties: {
      randomNumber: {
        type: Excel.CellValueType.double,
        basicValue: randomValue,
      },
    },
  };

  return entity;
}

/**
 * Returns a simple entity.
 * @customfunction
 * @param {number} latency Latency in millisecond.
 * @param {any} dependency Only for triggering chained calc.
 * @returns {any} A simple entity.
 */
function getRandomEntityAfterAsyncLatentcy(latency?: number, dependency?: any) {
  console.log(`Start getSimpleEntityAfterAsyncLatentcy`);
  let randomValue = Math.floor(Math.random() * 100);
  const entity = {
    type: Excel.CellValueType.entity,
    text: "Random Entity " + randomValue,
    properties: {
      randomNumber: {
        type: Excel.CellValueType.double,
        basicValue: randomValue,
      },
    },
  };
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve(entity);
    }, latency);
  });
}

/**
 * Returns a rich error. Error type: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-data-types-concepts#improved-error-support
 * @customfunction
 * @param {string} errorType The type of error to return.
 * @returns {any} A rich error.
 */
function getRichError(errorTypeString?: string) {
  console.log(`Start getRichError`);
  let errorType = Excel.ErrorCellValueType.value;
  let errorSubType = null;
  switch(errorTypeString.toLowerCase()) {
    case "blocked":
      errorType = Excel.ErrorCellValueType.blocked;
      errorSubType = Excel.BlockedErrorCellValueSubType.dataTypeUnsupportedApp;
      break;

    case "busy":
      errorType = Excel.ErrorCellValueType.busy;
      errorSubType = Excel.BusyErrorCellValueSubType.loadingImage;
      break;

    case "calc":
      errorType = Excel.ErrorCellValueType.calc;
      errorSubType = Excel.CalcErrorCellValueSubType.tooDeeplyNested;
      break;

    case "connect":
      errorType = Excel.ErrorCellValueType.connect;
      errorSubType = Excel.ConnectErrorCellValueSubType.externalLinksAccessFailed;
      break;

    case "div0":
      errorType = Excel.ErrorCellValueType.div0;
      // div0 does not have subType
      break;

    case "external": // Not in the documentation
      errorType = Excel.ErrorCellValueType.external;
      errorSubType = Excel.ExternalErrorCellValueSubType.unknown;
      break;

    case "field":
      errorType = Excel.ErrorCellValueType.field;
      errorSubType = Excel.FieldErrorCellValueSubType.webImageMissingFilePart;
      break;

    case "gettingdata":
      errorType = Excel.ErrorCellValueType.gettingData;
      break;

    case "notavailable":
      errorType = Excel.ErrorCellValueType.notAvailable;
      break;

    case "name":
      errorType = Excel.ErrorCellValueType.name;
      // "#NAME!" does not have subType
      break;

    case "null":
      errorType = Excel.ErrorCellValueType.null;
      // null does not have subType
      break;

    case "num":
      errorType = Excel.ErrorCellValueType.num;
      errorSubType = Excel.NumErrorCellValueSubType.arrayTooLarge;
      break;

    case "ref":
      errorType = Excel.ErrorCellValueType.ref;
      errorSubType = Excel.RefErrorCellValueSubType.externalLinksCalculatedRef;
      break;

    case "spill":
      errorType = Excel.ErrorCellValueType.spill;
      errorSubType = Excel.SpillErrorCellValueSubType.collision;
      break;

    case "timeout": // Not in the documentation
      errorType = Excel.ErrorCellValueType.timeout;
      errorSubType = Excel.TimeoutErrorCellValueSubType.pythonTimeoutLimitReached;
      break;

    case "value":
      errorType = Excel.ErrorCellValueType.value;
      errorSubType = Excel.ValueErrorCellValueSubType.coerceStringToNumberInvalid;
      break;

    default:
      errorType = Excel.ErrorCellValueType.name;
      // "#NAME!" does not have subType
      break;
  }

  let error = {};
  if (errorSubType) {
    error = {
      type: Excel.CellValueType.error,
      basicType: Excel.RangeValueType.error,
      errorType: errorType,
      errorSubType: errorSubType,
    };
  } else {
    error = {
      type: Excel.CellValueType.error,
      basicType: Excel.RangeValueType.error,
      errorType: errorType,
    };
  }

  return error;
}