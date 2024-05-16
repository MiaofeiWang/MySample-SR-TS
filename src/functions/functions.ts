﻿/* global clearInterval, console, CustomFunctions, setInterval */

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
    text: "Sample Entity " + randomValue,
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
 * Returns a rich error.
 * @customfunction
 * @param {string} errorType The type of error to return.
 * @returns {any} A rich error.
 */
function getRichError(errorTypeString?: string) {
  console.log(`Start getRichError`);
  let errorType = Excel.ErrorCellValueType.value;
  let errorSubType = null;
  switch(errorTypeString) {
    case "timeout":
      errorType = Excel.ErrorCellValueType.timeout;
      errorSubType = Excel.TimeoutErrorCellValueSubType.pythonTimeoutLimitReached;
      break;
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
    case "div0":
      errorType = Excel.ErrorCellValueType.div0;
      // div0 does not have subType
      break;
    case "external":
      errorType = Excel.ErrorCellValueType.external;
      errorSubType = Excel.ExternalErrorCellValueSubType.unknown;
      break;
    case "value":
    default:
      errorType = Excel.ErrorCellValueType.value;
      errorSubType = Excel.ValueErrorCellValueSubType.coerceStringToNumberInvalid;
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