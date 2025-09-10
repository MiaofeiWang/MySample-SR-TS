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
 * Returns a result of input + 1 for Excel.DoubleCellValue.
 * @customfunction
 * @param {Excel.DoubleCellValue} input
 * @returns {Excel.DoubleCellValue}
 */
function plusOneForDoubleCellValue(input: Excel.DoubleCellValue): Excel.DoubleCellValue {
  input.basicValue = input.basicValue + 1;
  return input;
}

/**
 * Returns a result of input + 1 for Excel.FormattedNumberCellValue.
 * @customfunction
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
 * This function will call the write API to write "Hello" to A1.
 * @customfunction
 * @returns {string} 
 */
async function testCallWriteAPI() {
  Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getActiveWorksheet();
    let range = sheet.getRange("A1");
    range.values = [["Hello"]];
    await context.sync();
  });

  return "Write API called";
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

/**
 * @customfunction
 * @param errorTypeString Error type
 * @param noMessage Whether to include message
 * @returns A custom function error.
 */
function getCFError(errorTypeString?: string, noMessage?: boolean) {
  console.log(`Start getCFError`);
  let errorType = CustomFunctions.ErrorCode.notAvailable;
  switch(errorTypeString.toLowerCase()) {
    case "divisionbyzero":
      errorType = CustomFunctions.ErrorCode.divisionByZero;
      break;
    case "invalidvalue":
      errorType = CustomFunctions.ErrorCode.invalidValue;
      break;
    case "notavailable":
      errorType = CustomFunctions.ErrorCode.notAvailable;
      break;
    default:
      // default NA error
      break;
  }

  if (noMessage) {
    return new CustomFunctions.Error(errorType);
  } else {
    let message = "Customized CF error message";
    return new CustomFunctions.Error(errorType, message);
  }
}

/**
 * @customfunction
 * @param {any} input Input value
 * @returns {string} Error message
 */
function getCFErrorMessage(input: any) {
  if (input.type == CustomFunctions.Error) {
    return input.message;
  }
  else if (input.type == Excel.CellValueType.error) {
    return "Not CF error but Excel error";
  }
  else {
    return "Not a CF error";
  }
}

export function getRandom0to99() {
  return Math.floor(Math.random() * 100);
}

/**
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<any>} invocation 
 */
function testFormattedNumberStreaming(invocation: CustomFunctions.StreamingInvocation<any>): void {
  let format = "0.0";
  let value = 0;
  const result = {
    basicValue: getRandom0to99(),
    numberFormat: `${format}`,
    type: Excel.CellValueType.formattedNumber,
  };
  invocation.setResult(result);

  const timeoutId = setInterval(async () => {
    value++;
    const now = {
      basicValue: getRandom0to99(),
      numberFormat: `${format}`,
      type: Excel.CellValueType.formattedNumber,
    };
    invocation.setResult(now);
  }, 2000);
}

/**
 * Filters data based on date range. Accepts various input types:
 * - String dates (yyyy-mm-dd format)
 * - Excel date serial numbers
 * - Relative day numbers (negative for past, positive for future)
 * @customfunction
 * @param {any} startInput - Start date (string, Excel date, or relative days)
 * @param {any} endInput - End date (string, Excel date, or relative days)
 * @returns {string} Message showing the filtered date range
 */
export function filterDataByDateRange(startInput: any, endInput: any): string {
  try {
    // Convert inputs to date strings
    const startDate = convertToDateString(startInput);
    const endDate = convertToDateString(endInput);
    
    // Call mock method to get data for this time period
    return mockGetDataInTimeRange(startDate, endDate);
  } catch (error) {
    return `Error: ${error.message}`;
  }
}

/**
 * Converts various input types to yyyy-mm-dd date string
 * @param {any} input - String, number (Excel date or relative days), or Excel date object
 * @returns {string} Date in yyyy-mm-dd format
 */
function convertToDateString(input: any): string {
  if (input == null) {
    throw new Error("Input cannot be null or undefined");
  }

  // Handle string input - try to parse various date formats
  if (typeof input === "string") {
    const parsedDate = parseFlexibleDateString(input);
    if (parsedDate) {
      return formatDateToYYYYMMDD(parsedDate);
    } else {
      throw new Error("Unable to parse date string. Supported formats: yyyy-mm-dd, yyyy/mm/dd, mm/dd/yyyy, dd/mm/yyyy");
    }
  }

  // Handle number input
  if (typeof input === "number") {
    // Calculate Excel serial date range for past 20 years to future 20 years
    // Today (September 10, 2025) is approximately serial 45903
    // 20 years ago (2005) ≈ serial 38718
    // 20 years from now (2045) ≈ serial 53088
    const minExcelDate = 38718; // Approximately January 1, 2005
    const maxExcelDate = 53088; // Approximately December 31, 2045
    
    // Check if it's a relative day number (typically small numbers, positive or negative)
    if (input >= -365 && input <= 365) {
      // Treat as relative days from today
      return getRelativeDate(input);
    } else if (input >= minExcelDate && input <= maxExcelDate) {
      // Treat as Excel date serial number within valid range
      return convertExcelDateToYYYYMMDD(input);
    } else {
      throw new Error(`Number input must be either relative days (-365 to 365) or Excel date serial (${minExcelDate} to ${maxExcelDate} for years 2005-2045)`);
    }
  }

  // Handle Excel date objects
  if (typeof input === "object" && input.type === Excel.CellValueType.double) {
    const dateValue = input.basicValue;
    const minExcelDate = 38718; // Approximately January 1, 2005
    const maxExcelDate = 53088; // Approximately December 31, 2045
    
    if (typeof dateValue === "number" && dateValue >= minExcelDate && dateValue <= maxExcelDate) {
      return convertExcelDateToYYYYMMDD(dateValue);
    }
  }

  throw new Error("Unsupported input type");
}

/**
 * Parses flexible date string formats
 * @param {string} dateStr - Date string in various formats
 * @returns {Date | null} Parsed Date object or null if parsing fails
 */
function parseFlexibleDateString(dateStr: string): Date | null {
  // Remove extra whitespace
  const cleanStr = dateStr.trim();
  
  // Try different date patterns
  const patterns = [
    // yyyy-mm-dd or yyyy/mm/dd
    /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/,
    // mm/dd/yyyy or mm-dd-yyyy
    /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/,
    // dd/mm/yyyy or dd-mm-yyyy (European format)
    /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/
  ];

  // Pattern 1: yyyy-mm-dd or yyyy/mm/dd
  let match = cleanStr.match(patterns[0]);
  if (match) {
    const year = parseInt(match[1]);
    const month = parseInt(match[2]) - 1; // JavaScript months are 0-indexed
    const day = parseInt(match[3]);
    const date = new Date(year, month, day);
    if (isValidDate(date, year, month + 1, day)) {
      return date;
    }
  }

  // Pattern 2: mm/dd/yyyy (US format)
  match = cleanStr.match(patterns[1]);
  if (match) {
    const month = parseInt(match[1]) - 1; // JavaScript months are 0-indexed
    const day = parseInt(match[2]);
    const year = parseInt(match[3]);
    const date = new Date(year, month, day);
    if (isValidDate(date, year, month + 1, day)) {
      return date;
    }
  }

  // Try JavaScript's native Date parsing as fallback
  const nativeDate = new Date(cleanStr);
  if (!isNaN(nativeDate.getTime())) {
    return nativeDate;
  }

  return null;
}

/**
 * Validates if the created date matches the input values
 * @param {Date} date - The created Date object
 * @param {number} year - Expected year
 * @param {number} month - Expected month (1-12)
 * @param {number} day - Expected day
 * @returns {boolean} True if date is valid and matches input
 */
function isValidDate(date: Date, year: number, month: number, day: number): boolean {
  return date.getFullYear() === year && 
         date.getMonth() === month - 1 && 
         date.getDate() === day &&
         !isNaN(date.getTime());
}

/**
 * Formats Date object to yyyy-mm-dd string
 * @param {Date} date - Date object to format
 * @returns {string} Date in yyyy-mm-dd format
 */
function formatDateToYYYYMMDD(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Converts Excel serial date number to yyyy-mm-dd format
 * @param {number} excelDate - Excel serial date number
 * @returns {string} Date in yyyy-mm-dd format
 */
function convertExcelDateToYYYYMMDD(excelDate: number): string {
  // Excel's epoch starts at 1900-01-01, but there's a leap year bug
  // Excel incorrectly treats 1900 as a leap year
  const excelEpoch = new Date(1899, 11, 30); // December 30, 1899
  
  // Convert Excel serial number to milliseconds and add to epoch
  const jsDate = new Date(excelEpoch.getTime() + (excelDate * 24 * 60 * 60 * 1000));
  
  // Format as yyyy-mm-dd
  const year = jsDate.getFullYear();
  const month = String(jsDate.getMonth() + 1).padStart(2, '0');
  const day = String(jsDate.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Gets date relative to today
 * @param {number} relativeDays - Number of days relative to today (negative for past, positive for future)
 * @returns {string} Date in yyyy-mm-dd format
 */
function getRelativeDate(relativeDays: number): string {
  const today = new Date();
  const targetDate = new Date(today.getTime() + (relativeDays * 24 * 60 * 60 * 1000));
  
  const year = targetDate.getFullYear();
  const month = String(targetDate.getMonth() + 1).padStart(2, '0');
  const day = String(targetDate.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}

/**
 * Mock method to simulate getting data in a time range
 * @param {string} startDate - Start date in yyyy-mm-dd format
 * @param {string} endDate - End date in yyyy-mm-dd format
 * @returns {string} Simple string showing the date range
 */
function mockGetDataInTimeRange(startDate: string, endDate: string): string {
  return `Filtering data from ${startDate} to ${endDate}`;
}
