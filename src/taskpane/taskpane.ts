/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/// <reference path="../office-experiment52.d.ts" />

// The initialize function must be run each time a new page is loaded
(async () => {
  await Office.onReady();

  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = registerLinkedEntityDomains;

  await registerLinkedEntityDomains();
  // async () => { await registerLinkedEntityDomains(); };
})();

// Office.onReady(async () => {
//   document.getElementById("sideload-msg").style.display = "none";
//   document.getElementById("app-body").style.display = "flex";
//   document.getElementById("run").onclick = registerLinkedEntityDomains;

//   //async () => { await registerLinkedEntityDomains(); };
// });

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