/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/// <reference path="../office-experiment52.d.ts" />

// The initialize function must be run each time a new page is loaded
(async () => {
  await Office.onReady();
  console.log("Office is ready");

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = registerLinkedEntityDomains;

  // Add event listeners for the function management buttons
  document.getElementById("showFunctionsBtn").onclick = showFunctions;
  document.getElementById("hideFunctionsBtn").onclick = hideFunctions;
  
  // Add event listener for the range reader button
  document.getElementById("getRangeBtn").onclick = getRangeValues;
  
  // Add event listener for the spill error test button
  document.getElementById("testSpillErrorBtn").onclick = testSpillErrorDetection;

  await registerLinkedEntityDomains();
})();

//Office.onReady(async () => {
  //Office.context.ui.displayDialogAsync('https://www.bing.com');
  //console.log("Office is ready from 2nd onReady");
//});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

/// Linked entity samples below


// Linked entity data domain constants
const domainDataProvider = "MYTSSAMPLE";
const domainLoadFunctionId = "PRODUCTLINKEDENTITYSERVICE"; // IMPORTANT: update the function namespace to match your own

// Linked entity cell value constants
const addinDomainServiceId = 268436224;
const defaultCulture = "en-US";

// // Linked entity data domains represent a specific category or field of information that shares common
// // characteristics or attributes.
// const productsDomain: Excel.LinkedEntityDataDomainCreateOptions = {
//   dataProvider: domainDataProvider,
//   id: "products",
//   name: "Products",
//   // Id of the custom function that will be called on demand by Excel to resolve/refresh linked entity
//   // cell values of this data domain.
//   loadFunctionId: domainLoadFunctionId,
//   // periodicRefreshInterval is only required when supportedRefreshModes contains "Periodic".
//   periodicRefreshInterval: 300,
//   // Manual refresh mode is always supported, even if unspecified.
//   supportedRefreshModes: [
//     "Periodic",
//     "OnLoad",
//   ]
// };

// Linked entity data domains can use unique load functions or the same load function can be used for
// multiple data domains.
const categoriesDomain: Excel.LinkedEntityDataDomainCreateOptions = {
  dataProvider: domainDataProvider,
  id: "categories",
  name: "Cateogories",
  loadFunctionId: domainLoadFunctionId
};

const suppliersDomain: Excel.LinkedEntityDataDomainCreateOptions = {
  dataProvider: domainDataProvider,
  id: "suppliers",
  name: "Suppliers",
  loadFunctionId: domainLoadFunctionId
};

export async function registerLinkedEntityDomains() {
  // Linked entity data domains represent a specific category or field of information that shares common
  // characteristics or attributes.
  const productsDomain: Excel.LinkedEntityDataDomainCreateOptions = {
    dataProvider: domainDataProvider,
    id: "products",
    name: "Products",
    // Id of the custom function that will be called on demand by Excel to resolve/refresh linked entity
    // cell values of this data domain.
    loadFunctionId: domainLoadFunctionId,
    // periodicRefreshInterval is only required when supportedRefreshModes contains "Periodic".
    periodicRefreshInterval: 300,
    // Manual refresh mode is always supported, even if unspecified.
    supportedRefreshModes: [
      "Periodic",
      "OnLoad",
    ]
  };

  await Excel.run(async (context) => {
    // Before we can create linked entity cell values, we need to register the linked entity data domains
    // with Excel. A linked entity data domain can only be registered once per workbook.
    const linkedEntityDataDomains = context.workbook.linkedEntityDataDomains;
    linkedEntityDataDomains.add(productsDomain);
    linkedEntityDataDomains.add(categoriesDomain);
    linkedEntityDataDomains.add(suppliersDomain);

    await context.sync();
    console.log("Linked entity data domains registered.");
  });
}

/// Linked entity samples above

/**
 * Function to handle showing functions based on comma-separated input
 */
export async function showFunctions() {
  try {
    const input = (document.getElementById("showFunctionsInput") as HTMLInputElement).value;
    const functionNames = parseFunctionNames(input);
    
    console.log("Functions to show:", functionNames);

    await Excel.CustomFunctionManager.setVisibility({show: functionNames, hide: []});
    
  } catch (error) {
    console.error("Error in showFunctions:", error);
  }
}

/**
 * Function to handle hiding functions based on comma-separated input
 */
export async function hideFunctions() {
  try {
    const input = (document.getElementById("hideFunctionsInput") as HTMLInputElement).value;
    const functionNames = parseFunctionNames(input);
    
    console.log("Functions to hide:", functionNames);

    await Excel.CustomFunctionManager.setVisibility({show: [], hide: functionNames});
  } catch (error) {
    console.error("Error in hideFunctions:", error);
  }
}

/**
 * Helper function to parse comma-separated function names
 * @param input - Comma-separated string of function names
 * @returns Array of trimmed function names
 */
function parseFunctionNames(input: string): string[] {
  if (!input || input.trim() === "") {
    return [];
  }
  
  return input
    .split(",")
    .map(name => name.trim())
    .filter(name => name.length > 0);
}

/**
 * Function to get range values from Sheet1!A1:B2, wait 10 seconds, and display the values
 */
export async function getRangeValues() {
  const outputElement = document.getElementById("rangeOutput");
  const button = document.getElementById("getRangeBtn") as HTMLButtonElement;
  
  try {
    // Disable button and show loading state
    button.disabled = true;
    button.textContent = "Getting range...";
    
    if (outputElement) {
      outputElement.textContent = "Getting range object...";
    }

    await Excel.run(async (context) => {
      // Get the range Sheet1!A1:B2
      const range = context.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");
      
      // Load the values property of the range
      range.load("values, address");

      await context.sync();

      if (outputElement) {
        outputElement.textContent = `Range ${range.address} retrieved. Waiting 10 seconds...`;
      }

      // Wait for 10 seconds asynchronously
      await new Promise(resolve => setTimeout(resolve, 10000));

      // Display the range values
      const values = range.values;
      let output = `Range: ${range.address}\nValues:\n`;
      
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values[i].length; j++) {
          output += `[${i},${j}]: ${values[i][j]}\n`;
        }
      }

      if (outputElement) {
        outputElement.textContent = output;
      }

      console.log("Range values:", values);
    });

  } catch (error) {
    console.error("Error getting range values:", error);
    if (outputElement) {
      outputElement.textContent = `Error: ${error.message || error}`;
    }
  } finally {
    // Re-enable button
    button.disabled = false;
    button.textContent = "Get Range Values (Sheet1!A1:B2)";
  }
}

/**
 * Helper function to get worksheet - mocks the getWorksheet function from your code
 */
async function getWorksheet(worksheetName: string, { context }: { context: Excel.RequestContext }): Promise<Excel.Worksheet> {
  let worksheet: Excel.Worksheet;
  
  try {
    // Try to get existing worksheet
    worksheet = context.workbook.worksheets.getItem(worksheetName);
  } catch {
    // If worksheet doesn't exist, create it
    worksheet = context.workbook.worksheets.add(worksheetName);
  }
  
  // Sync to ensure worksheet is available
  await context.sync();
  return worksheet;
}

/**
 * Function to mock the spill error detection behavior that sometimes fails
 * This reproduces the code snippet you provided
 */
export async function testSpillErrorDetection() {
  const outputElement = document.getElementById("spillErrorOutput");
  const button = document.getElementById("testSpillErrorBtn") as HTMLButtonElement;
  const addressInput = document.getElementById("testCellInput") as HTMLInputElement;
  const waitDurationInput = document.getElementById("waitDurationInput") as HTMLInputElement;
  
  const address = addressInput.value || "A1";
  const worksheetName = "Sheet1";
  const waitDuration = parseInt(waitDurationInput.value) || 0;
  
  try {
    // Disable button and show loading state
    button.disabled = true;
    button.textContent = "Testing...";
    
    if (outputElement) {
      outputElement.textContent = "Starting spill error detection test...";
    }

    // Mock the invokeExcelRun behavior
    await Excel.run(async (context) => {
      if (outputElement) {
        outputElement.textContent = `Getting worksheet: ${worksheetName}...`;
      }

      const worksheet = await getWorksheet(worksheetName, { context }); // syncs inside

      if (outputElement) {
        outputElement.textContent = `Getting range: ${address}...`;
      }

      const cellRange = worksheet.getRange(address);

      if (outputElement) {
        outputElement.textContent = `Checking for spill parent range...`;
      }

      const cellSpillParentRange = cellRange.getSpillParentOrNullObject();

      await context.sync();

      // Wait asynchronously after the spill parent sync if duration is specified
      if (waitDuration > 0) {
        if (outputElement) {
          outputElement.textContent = `Waiting ${waitDuration} seconds after spill parent sync...`;
        }
        await new Promise(resolve => setTimeout(resolve, waitDuration * 1000));
      }

      /**
       * If the cell belongs to a spill range and contains an error, attempting to get its error formula range
       * will return a null object, and we won't check whether it still contains an error later.
       * Therefore, we need to use the spill parent cell in such cases.
       */
      const cellOrSpillParentRange = !cellSpillParentRange.isNullObject
          ? cellSpillParentRange
          : cellRange;

      if (outputElement) {
        const rangeType = !cellSpillParentRange.isNullObject ? "spill parent" : "original cell";
        outputElement.textContent = `Using ${rangeType} range. Getting special cells with errors...`;
      }

      const erroredCell = cellOrSpillParentRange.getSpecialCellsOrNullObject(
          Excel.SpecialCellType.formulas,
          Excel.SpecialCellValueType.errors
      );

      erroredCell.load({ address: true });

      if (outputElement) {
        outputElement.textContent = `Loading errored cell data... (this is where the sync might fail)`;
      }

      // This is the sync that sometimes throws in your original code
      await context.sync(); // <--- this sync threw

      // If we get here, the sync succeeded
      let output = `✅ Test completed successfully!\n`;
      output += `Original address: ${address}\n`;
      output += `Wait duration: ${waitDuration} seconds\n`;
      output += `Has spill parent: ${!cellSpillParentRange.isNullObject}\n`;
      
      if (!erroredCell.isNullObject) {
        output += `Errored cell address: ${erroredCell.address}\n`;
      } else {
        output += `No errored cells found\n`;
      }

      if (outputElement) {
        outputElement.textContent = output;
      }

      console.log("Spill error detection test completed successfully");

    });

  } catch (error) {
    console.error("Error in spill error detection test:", error);
    if (outputElement) {
      let errorOutput = `❌ Error occurred during test:\n`;
      errorOutput += `Error: ${error.message || error}\n`;
      errorOutput += `Address tested: ${address}\n`;
      errorOutput += `Worksheet: ${worksheetName}\n`;
      errorOutput += `Wait duration: ${waitDuration} seconds\n`;
      errorOutput += `\nThis might be the error you're trying to reproduce!`;
      outputElement.textContent = errorOutput;
    }
  } finally {
    // Re-enable button
    button.disabled = false;
    button.textContent = "Test Spill Error Detection";
  }
}