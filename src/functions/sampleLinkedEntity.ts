/// Linked entity samples below

// Linked entity data domain constants
const productsDomainId = "products";
const categoriesDomainId = "categories";
const suppliersDomainId = "suppliers";

// Linked entity cell value constants
const addinDomainServiceId = 268436224;
const defaultCulture = "en-US";

/**
 * Custom function which demonstrates insertion of a third-party `LinkedEntityCellValue`.
 * @customfunction
 * @param {string} productId Unique id of the product.
 * @return {Promise<any>} `LinkedEntityCellValue` for the requested product, if found.
 */
async function getProductById(productId: string): Promise<any> {
    console.log(`Start getProductById: Fetching product with id ${productId} ...`);
    const linkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: "Chai",
        id: {
            entityId: productId,
            domainId: productsDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        }
    };

    return linkedEntity;
}


/**
 * Custom function which acts as the "service" or the data provider for a `LinkedEntityDataDomain`, that is
 * called on demand by Excel to resolve/refresh `LinkedEntityCellValue`s of that `LinkedEntityDataDomain`.
 * @customfunction
 * @linkedEntityDataProvider
 * @param {any} linkedEntityId Unique `LinkedEntityId` of the `LinkedEntityCellValue`s which is being
 * requested for resolution/refresh.
 * @return {Promise<any>} Resolved/Updated `LinkedEntityCellValue` that was requested by the passed-in id.
 */
async function productLinkedEntityService(linkedEntityId: any): Promise<any> {
    console.log(`Start productLinkedEntityService: Fetching linked entity with id ${linkedEntityId} ...`);
    const notAvailableError = new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    console.log(`Fetching linked entity with id ${linkedEntityId} ...`);

    try {
        const parsedLinkedEntityId: Excel.LinkedEntityId = JSON.parse(linkedEntityId);
        return makeSimpleLinkedEntity(parsedLinkedEntityId.entityId);
        // // Parse the linkedEntityId that was passed-in by Excel.
        // const parsedLinkedEntityId: Excel.LinkedEntityId = JSON.parse(linkedEntityId);

        // // Identify the domainId of the requested linkedEntityId and call the corresponding function to create
        // // linked entity cell values for that linked entity data domain.
        // var linkedEntityResult = null;
        // switch (parsedLinkedEntityId.domainId) {
        //   case productsDomainId: {
        //     linkedEntityResult = makeProductLinkedEntity(parsedLinkedEntityId.entityId);
        //     break;
        //   }

        //   case categoriesDomainId: {
        //     linkedEntityResult = makeCategoryLinkedEntity(parsedLinkedEntityId.entityId);
        //     break;
        //   }

        //   case suppliersDomainId: {
        //     linkedEntityResult = makeSupplierLinkedEntity(parsedLinkedEntityId.entityId);
        //     break;
        //   }

        //   default:
        //     throw notAvailableError;
        // }

        // if (linkedEntityResult === null) {
        //   // Throw an error to signify to Excel that resolution/refresh of the requested linkedEntityId failed.
        //   throw notAvailableError;
        // }

        // return linkedEntityResult;
    } catch (error) {
        console.error(error);
        throw notAvailableError;
    }
}

function makeSimpleLinkedEntity(productID: string): any {
    const productLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: "Chai",
        id: {
            entityId: productID,
            domainId: productsDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Product ID": {
                type: "String",
                basicValue: productID.toString()
            },
            "Product Name": {
                type: "String",
                basicValue: "Chai"
            },
            "Quantity Per Unit": {
                type: "String",
                basicValue: "10 boxes x 20 bags"
            },
            // Add Unit Price as a formatted number.
            "Unit Price": {
                type: "FormattedNumber",
                basicValue: 18,
                numberFormat: "$* #,##0.00"
            },
            Discontinued: {
                type: "Boolean",
                basicValue: false
            }
        },
        cardLayout: {
            title: { property: "Product Name" },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false,
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsed: true,
                    properties: ["Discontinued"]
                }
            ]
        }
    };
    productLinkedEntity.properties["Image"] = {
        type: "WebImage",
        address: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Masala_Chai.JPG/320px-Masala_Chai.JPG"
    };
    productLinkedEntity.cardLayout.mainImage = { property: "Image" };
    // Add a deferred nested linked entity for the product category.
    productLinkedEntity.properties["Category"] = {
        type: "LinkedEntity",
        text: "Beverages",
        id: {
            entityId: "C1",
            domainId: categoriesDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        }
    };
    // Add nested product category to the card layout.
    productLinkedEntity.cardLayout.sections[0].properties.push("Category");
    // Add a deferred nested linked entity for the supplier.
    productLinkedEntity.properties["Supplier"] = {
        type: "LinkedEntity",
        text: "Exotic Liquids",
        id: {
            entityId: "S1",
            domainId: suppliersDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        }
    };
    // Add nested product supplier to the card layout.
    productLinkedEntity.cardLayout.sections[2].properties.push("Supplier");
    return productLinkedEntity;
}

/** Helper function to create linked entity from product properties. */
function makeProductLinkedEntity(productID: string): any {
    // Search the sample JSON product data for a matching product ID.
    const product = getProduct(productID);
    if (product === null) {
        // Return null if no matching product is found.
        return null;
    }

    const productLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: product.productName,
        id: {
            entityId: product.productID,
            domainId: productsDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Product ID": {
                type: "String",
                basicValue: product.productID
            },
            "Product Name": {
                type: "String",
                basicValue: product.productName
            },
            "Quantity Per Unit": {
                type: "String",
                basicValue: product.quantityPerUnit
            },
            // Add Unit Price as a formatted number.
            "Unit Price": {
                type: "FormattedNumber",
                basicValue: product.unitPrice,
                numberFormat: "$* #,##0.00"
            },
            Discontinued: {
                type: "Boolean",
                basicValue: product.discontinued
            }
        },
        cardLayout: {
            title: { property: "Product Name" },
            sections: [
                {
                    layout: "List",
                    properties: ["Product ID"]
                },
                {
                    layout: "List",
                    title: "Quantity and price",
                    collapsible: true,
                    collapsed: false,
                    properties: ["Quantity Per Unit", "Unit Price"]
                },
                {
                    layout: "List",
                    title: "Additional information",
                    collapsed: true,
                    properties: ["Discontinued"]
                }
            ]
        }
    };

    // Add image property to the linked entity and then add it to the card layout.
    if (product.productImage) {
        productLinkedEntity.properties["Image"] = {
            type: "WebImage",
            address: product.productImage
        };
        productLinkedEntity.cardLayout.mainImage = { property: "Image" };
    }

    // Add a deferred nested linked entity for the product category.
    const category = getCategory(product.categoryID.toString());
    if (category) {
        productLinkedEntity.properties["Category"] = {
            type: "LinkedEntity",
            text: category.categoryName,
            id: {
                entityId: category.categoryID.toString(),
                domainId: categoriesDomainId,
                serviceId: addinDomainServiceId,
                culture: defaultCulture
            }
        };

        // Add nested product category to the card layout.
        productLinkedEntity.cardLayout.sections[0].properties.push("Category");
    }

    // Add a deferred nested linked entity for the supplier.
    const supplier = getSupplier(product.supplierID.toString());
    if (supplier) {
        productLinkedEntity.properties["Supplier"] = {
            type: "LinkedEntity",
            text: supplier.companyName,
            id: {
                entityId: supplier.supplierID.toString(),
                domainId: suppliersDomainId,
                serviceId: addinDomainServiceId,
                culture: defaultCulture
            }
        };

        // Add nested product supplier to the card layout.
        productLinkedEntity.cardLayout.sections[2].properties.push("Supplier");
    }

    return productLinkedEntity;
}

/** Helper function to create linked entity from category properties. */
function makeCategoryLinkedEntity(categoryID: string): any {
    // Search the sample JSON category data for a matching category ID.
    const category = getCategory(categoryID);
    if (category === null) {
        // Return null if no matching category is found.
        return null;
    }

    const categoryLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: category.categoryName,
        id: {
            entityId: category.categoryID,
            domainId: categoriesDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Category ID": {
                type: "String",
                basicValue: category.categoryID,
                propertyMetadata: {
                    // Exclude the category ID property from the card view and auto complete.
                    excludeFrom: {
                        cardView: true,
                        autoComplete: true
                    }
                }
            },
            "Category Name": {
                type: "String",
                basicValue: category.categoryName
            },
            Description: {
                type: "String",
                basicValue: category.description
            }
        }
    };

    return categoryLinkedEntity;
}

/** Helper function to create linked entity from supplier properties. */
function makeSupplierLinkedEntity(supplierID: string): any {
    // Search the sample JSON category data for a matching supplier ID.
    const supplier = getSupplier(supplierID);
    if (supplier === null) {
        // Return null if no matching supplier is found.
        return null;
    }

    const supplierLinkedEntity: Excel.LinkedEntityCellValue = {
        type: "LinkedEntity",
        text: supplier.companyName,
        id: {
            entityId: supplier.supplierID,
            domainId: suppliersDomainId,
            serviceId: addinDomainServiceId,
            culture: defaultCulture
        },
        properties: {
            "Supplier ID": {
                type: "String",
                basicValue: supplier.supplierID
            },
            "Company Name": {
                type: "String",
                basicValue: supplier.companyName
            },
            "Contact Name": {
                type: "String",
                basicValue: supplier.contactName
            },
            "Contact Title": {
                type: "String",
                basicValue: supplier.contactTitle
            }
        },
        cardLayout: {
            title: { property: "Company Name" },
            sections: [
                {
                    layout: "List",
                    properties: ["Supplier ID", "Company Name", "Contact Name", "Contact Title"]
                }
            ]
        }
    };

    return supplierLinkedEntity;
}

/** Get products and product properties. */
function getProduct(productID: string): any {
    return products.find((p) => p.productID === productID);
}

/** Get product categories and category properties. */
function getCategory(categoryID: string): any {
    return categories.find((c) => c.categoryID === categoryID);
}

/** Get product suppliers and supplier properties. */
function getSupplier(supplierID: string): any {
    return suppliers.find((s) => s.supplierID === supplierID);
}

/** Sample JSON product data. */
const products = [
    {
        productID: "P1",
        productName: "Chai",
        supplierID: "S1",
        categoryID: "C1",
        quantityPerUnit: "10 boxes x 20 bags",
        unitPrice: 18,
        discontinued: false,
        productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/04/Masala_Chai.JPG/320px-Masala_Chai.JPG"
    },
    {
        productID: "P2",
        productName: "Chang",
        supplierID: "S1",
        categoryID: "C1",
        quantityPerUnit: "24 - 12 oz bottles",
        unitPrice: 19,
        discontinued: false,
        productImage: ""
    },
    {
        productID: "P3",
        productName: "Aniseed Syrup",
        supplierID: "S1",
        categoryID: "C2",
        quantityPerUnit: "12 - 550 ml bottles",
        unitPrice: 10,
        discontinued: false,
        productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/Maltose_syrup.jpg/185px-Maltose_syrup.jpg"
    },
    {
        productID: "P4",
        productName: "Chef Anton's Cajun Seasoning",
        supplierID: "S2",
        categoryID: "C2",
        quantityPerUnit: "48 - 6 oz jars",
        unitPrice: 22,
        discontinued: false,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/8/82/Kruidenmengeling-spice.jpg/193px-Kruidenmengeling-spice.jpg"
    },
    {
        productID: "P5",
        productName: "Chef Anton's Gumbo Mix",
        supplierID: "S2",
        categoryID: "C2",
        quantityPerUnit: "36 boxes",
        unitPrice: 21.35,
        discontinued: true,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/Okra_in_a_Bowl_%28Unsplash%29.jpg/180px-Okra_in_a_Bowl_%28Unsplash%29.jpg"
    },
    {
        productID: "P6",
        productName: "Grandma's Boysenberry Spread",
        supplierID: "S3",
        categoryID: "C2",
        quantityPerUnit: "12 - 8 oz jars",
        unitPrice: 25,
        discontinued: false,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/1/10/Making_cranberry_sauce_-_in_the_jar.jpg/90px-Making_cranberry_sauce_-_in_the_jar.jpg"
    },
    {
        productID: "P7",
        productName: "Uncle Bob's Organic Dried Pears",
        supplierID: "S3",
        categoryID: "C7",
        quantityPerUnit: "12 - 1 lb pkgs.",
        unitPrice: 30,
        discontinued: false,
        productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fd/DriedPears.JPG/120px-DriedPears.JPG"
    },
    {
        productID: "P8",
        productName: "Northwoods Cranberry Sauce",
        supplierID: "S3",
        categoryID: "C2",
        quantityPerUnit: "12 - 12 oz jars",
        unitPrice: 40,
        discontinued: false,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/0/07/Making_cranberry_sauce_-_stovetop.jpg/90px-Making_cranberry_sauce_-_stovetop.jpg"
    },
    {
        productID: "P9",
        productName: "Mishi Kobe Niku",
        supplierID: "S4",
        categoryID: "C6",
        quantityPerUnit: "18 - 500 g pkgs.",
        unitPrice: 97,
        discontinued: true,
        productImage: ""
    },
    {
        productID: "P10",
        productName: "Ikura",
        supplierID: "S4",
        categoryID: "C8",
        quantityPerUnit: "12 - 200 ml jars",
        unitPrice: 31,
        discontinued: false,
        productImage: ""
    },
    {
        productID: "P11",
        productName: "Queso Cabrales",
        supplierID: "S5",
        categoryID: "C4",
        quantityPerUnit: "1 kg pkg.",
        unitPrice: 21,
        discontinued: false,
        productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/9/96/Tilsit_cheese.jpg/190px-Tilsit_cheese.jpg"
    },
    {
        productID: "P12",
        productName: "Queso Manchego La Pastora",
        supplierID: "S5",
        categoryID: "C4",
        quantityPerUnit: "10 - 500 g pkgs.",
        unitPrice: 38,
        discontinued: false,
        productImage: "https://upload.wikimedia.org/wikipedia/commons/thumb/5/59/Manchego.jpg/177px-Manchego.jpg"
    },
    {
        productID: "P13",
        productName: "Konbu",
        supplierID: "S6",
        categoryID: "C8",
        quantityPerUnit: "2 kg box",
        unitPrice: 6,
        discontinued: false,
        productImage: ""
    },
    {
        productID: "P14",
        productName: "Tofu",
        supplierID: "S6",
        categoryID: "C7",
        quantityPerUnit: "40 - 100 g pkgs.",
        unitPrice: 23.25,
        discontinued: false,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e5/Korean.food-Dubu.gui-01.jpg/120px-Korean.food-Dubu.gui-01.jpg"
    },
    {
        productID: "P15",
        productName: "Genen Shouyu",
        supplierID: "S6",
        categoryID: "C2",
        quantityPerUnit: "24 - 250 ml bottles",
        unitPrice: 15.5,
        discontinued: false,
        productImage: ""
    },
    {
        productID: "P16",
        productName: "Pavlova",
        supplierID: "S7",
        categoryID: "C3",
        quantityPerUnit: "32 - 500 g boxes",
        unitPrice: 17.45,
        discontinued: false,
        productImage: ""
    },
    {
        productID: "P17",
        productName: "Alice Mutton",
        supplierID: "S7",
        categoryID: "C6",
        quantityPerUnit: "20 - 1 kg tins",
        unitPrice: 39,
        discontinued: true,
        productImage: ""
    },
    {
        productID: "P18",
        productName: "Carnarvon Tigers",
        supplierID: "S7",
        categoryID: "C8",
        quantityPerUnit: "16 kg pkg.",
        unitPrice: 62.5,
        discontinued: false,
        productImage: ""
    },
    {
        productID: "P19",
        productName: "Teatime Chocolate Biscuits",
        supplierID: "S8",
        categoryID: "C3",
        quantityPerUnit: "10 boxes x 12 pieces",
        unitPrice: 9.2,
        discontinued: false,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/d/df/Macau_Koi_Kei_Bakery_Almond_Biscuits_2.JPG/120px-Macau_Koi_Kei_Bakery_Almond_Biscuits_2.JPG"
    },
    {
        productID: "P20",
        productName: "Sir Rodney's Marmalade",
        supplierID: "S8",
        categoryID: "C3",
        quantityPerUnit: "30 gift boxes",
        unitPrice: 81,
        discontinued: false,
        productImage:
            "https://upload.wikimedia.org/wikipedia/commons/thumb/3/30/Homemade_marmalade%2C_England.jpg/135px-Homemade_marmalade%2C_England.jpg"
    }
];

const categories = [
    {
        categoryID: "C1",
        categoryName: "Beverages",
        description: "Soft drinks, coffees, teas, beers, and ales"
    },
    {
        categoryID: "C2",
        categoryName: "Condiments",
        description: "Sweet and savory sauces, relishes, spreads, and seasonings"
    },
    {
        categoryID: "C3",
        categoryName: "Confections",
        description: "Desserts, candies, and sweet breads"
    },
    {
        categoryID: "C4",
        categoryName: "Dairy Products",
        description: "Cheeses"
    },
    {
        categoryID: "C5",
        categoryName: "Grains/Cereals",
        description: "Breads, crackers, pasta, and cereal"
    },
    {
        categoryID: "C6",
        categoryName: "Meat/Poultry",
        description: "Prepared meats"
    },
    {
        categoryID: "C7",
        categoryName: "Produce",
        description: "Dried fruit and bean curd"
    },
    {
        categoryID: "C8",
        categoryName: "Seafood",
        description: "Seaweed and fish"
    }
];

const suppliers = [
    {
        supplierID: "S1",
        companyName: "Exotic Liquids",
        contactName: "Charlotte Cooper",
        contactTitle: "Purchasing Manager"
    },
    {
        supplierID: "S2",
        companyName: "New Orleans Cajun Delights",
        contactName: "Shelley Burke",
        contactTitle: "Order Administrator"
    },
    {
        supplierID: "S3",
        companyName: "Grandma Kelly's Homestead",
        contactName: "Regina Murphy",
        contactTitle: "Sales Representative"
    },
    {
        supplierID: "S4",
        companyName: "Tokyo Traders",
        contactName: "Yoshi Nagase",
        contactTitle: "Marketing Manager",
        address: "9-8 Sekimai Musashino-shi"
    },
    {
        supplierID: "S5",
        companyName: "Cooperativa de Quesos 'Las Cabras'",
        contactName: "Antonio del Valle Saavedra",
        contactTitle: "Export Administrator"
    },
    {
        supplierID: "S6",
        companyName: "Mayumi's",
        contactName: "Mayumi Ohno",
        contactTitle: "Marketing Representative"
    },
    {
        supplierID: "S7",
        companyName: "Pavlova, Ltd.",
        contactName: "Ian Devling",
        contactTitle: "Marketing Manager"
    },
    {
        supplierID: "S8",
        companyName: "Specialty Biscuits, Ltd.",
        contactName: "Peter Wilson",
        contactTitle: "Sales Representative"
    }
];


/// Linked entity samples above