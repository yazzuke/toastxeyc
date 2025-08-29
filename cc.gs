/**
 * Creates a menu item in Google Sheets to run the import
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('POS')
    .addItem('Import Products', 'importProducts')
    .addItem('Import Orders', 'importOrders')
    .addItem('Import Orders Detailed', 'importOrdersDetailed')
    .addToUi();
}

/**
 * Main function to import products from the Eyecatch API
 */
function importProducts() {
  try {
    // Get active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Set up or clear sheet
    var productsSheet = setupSheet(ss, "Products");

    // Fetch data from Eyecatch API
    Logger.log("Fetching data from Eyecatch API...");
    var data = fetchEyecatchData();

    if (!data || !data.response || !data.response.products) {
      Logger.log("No product data available");
      return;
    }

    Logger.log("Found " + data.response.products.length + " products");

    // Process products
    var row = 2; // Start after headers
    data.response.products.forEach(function (product) {
      row = processProduct(product, productsSheet, row);
    });

    // Auto-resize columns
    productsSheet.autoResizeColumns(1, productsSheet.getLastColumn());

    Logger.log("Products imported successfully!");
    Logger.log("Total products processed: " + (productsSheet.getLastRow() - 1));
  } catch (e) {
    Logger.log("Error: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
  }
}

/**
 * Set up the products sheet with headers
 */
function setupSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = [
    "ID", "Created", "Updated", "POS ID", "ID Brand",
    "Name", "Description", "Price", "In Stock", "Category",
    "Image URL", "Calories", "Tags", "Status", "Not Found",
    "Modifier Groups", "Custom Fields (JSON)"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Fetches data from the Eyecatch API
 */
function fetchEyecatchData() {
  var apiUrl = "https://pos.eyecatch.co/products/active?brand=ceaa397d-47df-426d-8a9a-92550fe36164&category=&tag=&dataset=standar&stock=false&modifiers=false";

  try {
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log("API Error: " + responseCode);
      return null;
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("Error fetching data: " + e.toString());
    return null;
  }
}

/**
 * Process a single product and add it to the sheet
 */
function processProduct(product, productsSheet, row) {
  // Extract image URL from custom fields
  var imageUrl = "";
  var calories = "";
  var customFieldsJson = "";

  if (product.custom_fields && product.custom_fields.length > 0) {
    // Create a simplified object for custom fields
    var customFieldsObj = {};

    product.custom_fields.forEach(function (field) {
      if (field.custom_field && field.custom_field.fieldKey) {
        var fieldKey = field.custom_field.fieldKey;
        var fieldValue = field.fieldValue || "";

        // Extract specific fields
        if (fieldKey === "image") {
          imageUrl = fieldValue;
        } else if (fieldKey === "calories") {
          calories = fieldValue;
        }

        // Add to custom fields object
        customFieldsObj[fieldKey] = fieldValue;
      }
    });

    // Convert to JSON string for the sheet
    customFieldsJson = JSON.stringify(customFieldsObj);
  }

  // Extract category name
  var category = "";
  if (product.product_category && product.product_category.category) {
    category = product.product_category.category;
  }

  // Format modifier groups
  var modifierGroups = "";
  if (product.modifier_groups) {
    modifierGroups = JSON.stringify(product.modifier_groups);
  }

  // Convert timestamps to readable dates
  var createdDate = product.Created ? new Date(product.Created * 1000) : "";
  var updatedDate = product.Updated ? new Date(product.Updated * 1000) : "";

  // Prepare product data
  var productData = [
    product.id || "",
    createdDate,
    updatedDate,
    product.pos_id || "",
    product.id_brand || "",
    product.name || "",
    product.description || "",
    product.price || "",
    product.in_stock ? "Yes" : "No",
    category,
    imageUrl,
    calories,
    product.tags || "",
    product.status || "",
    product.not_found ? "Yes" : "No",
    modifierGroups,
    customFieldsJson
  ];

  productsSheet.getRange(row, 1, 1, productData.length).setValues([productData]);
  return row + 1;
}

/**
 * Optional: Function to import products with detailed custom fields in separate columns
 */
function importProductsDetailed() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var productsSheet = setupDetailedSheet(ss, "Products Detailed");

    Logger.log("Fetching data from Eyecatch API...");
    var data = fetchEyecatchData();

    if (!data || !data.response || !data.response.products) {
      Logger.log("No product data available");
      return;
    }

    Logger.log("Found " + data.response.products.length + " products");

    var row = 2;
    data.response.products.forEach(function (product) {
      row = processProductDetailed(product, productsSheet, row);
    });

    productsSheet.autoResizeColumns(1, productsSheet.getLastColumn());

    Logger.log("Detailed products imported successfully!");
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

/**
 * Set up detailed sheet with separate columns for each custom field
 */
function setupDetailedSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = [
    "ID", "Created", "Updated", "POS ID", "ID Brand",
    "Name", "Description", "Price", "In Stock", "Category",
    "Image URL", "Calories", "Tags", "Status", "Not Found",
    "Modifier Groups", "Custom Field - Image", "Custom Field - Calories",
    "All Custom Fields (JSON)"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Process product with detailed custom fields breakdown
 */
function processProductDetailed(product, productsSheet, row) {
  var imageUrl = "";
  var calories = "";
  var customFieldsJson = "";

  if (product.custom_fields && product.custom_fields.length > 0) {
    var customFieldsObj = {};

    product.custom_fields.forEach(function (field) {
      if (field.custom_field && field.custom_field.fieldKey) {
        var fieldKey = field.custom_field.fieldKey;
        var fieldValue = field.fieldValue || "";

        if (fieldKey === "image") {
          imageUrl = fieldValue;
        } else if (fieldKey === "calories") {
          calories = fieldValue;
        }

        customFieldsObj[fieldKey] = fieldValue;
      }
    });

    customFieldsJson = JSON.stringify(customFieldsObj);
  }

  var category = "";
  if (product.product_category && product.product_category.category) {
    category = product.product_category.category;
  }

  var modifierGroups = "";
  if (product.modifier_groups) {
    modifierGroups = JSON.stringify(product.modifier_groups);
  }

  var createdDate = product.Created ? new Date(product.Created * 1000) : "";
  var updatedDate = product.Updated ? new Date(product.Updated * 1000) : "";

  var productData = [
    product.id || "",
    createdDate,
    updatedDate,
    product.pos_id || "",
    product.id_brand || "",
    product.name || "",
    product.description || "",
    product.price || "",
    product.in_stock ? "Yes" : "No",
    category,
    imageUrl,
    calories,
    product.tags || "",
    product.status || "",
    product.not_found ? "Yes" : "No",
    modifierGroups,
    imageUrl, // Separate column for image
    calories, // Separate column for calories
    customFieldsJson
  ];

  productsSheet.getRange(row, 1, 1, productData.length).setValues([productData]);
  return row + 1;
}


function importOrders() {
  try {
    // Get active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Set up or clear sheet
    var ordersSheet = setupOrdersSheet(ss, "Orders");

    // Fetch data from Toast API
    Logger.log("Fetching data from Toast API...");
    var data = fetchToastOrdersData();

    if (!data || !Array.isArray(data)) {
      Logger.log("No order data available");
      return;
    }

    Logger.log("Found " + data.length + " orders");

    // Process orders
    var row = 2; // Start after headers
    data.forEach(function (order) {
      row = processOrder(order, ordersSheet, row);
    });

    // Auto-resize columns
    ordersSheet.autoResizeColumns(1, ordersSheet.getLastColumn());

    Logger.log("Orders imported successfully!");
    Logger.log("Total orders processed: " + (ordersSheet.getLastRow() - 1));
  } catch (e) {
    Logger.log("Error: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
  }
}

/**
 * Set up the orders sheet with headers
 */
function setupOrdersSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = [
    "Order GUID", "Display Number", "Source", "Business Date",
    "Opened Date", "Paid Date", "Closed Date", "Duration (min)",
    "Number of Guests", "Voided", "Approval Status",
    "Total Amount", "Tax Amount", "Payment Type", "Payment Status",
    "Server GUID", "Device ID", "Created in Test Mode",
    "Total Items", "Order Summary", "Checks Info"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Fetches data from the Toast Orders API
 */
function fetchToastOrdersData() {
  // Get today's date in yyyyMMdd format
  var today = new Date();
  var businessDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyyMMdd");

  var apiUrl = "https://eycxtoast.zeabur.app/api/toast/orders?businessDate=" + businessDate;

  try {
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    var responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      Logger.log("API Error: " + responseCode);
      Logger.log("Response: " + response.getContentText());
      return null;
    }

    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("Error fetching data: " + e.toString());
    return null;
  }
}

/**
 * Process a single order and add it to the sheet
 */
function processOrder(order, ordersSheet, row) {
  // Format dates
  var openedDate = order.openedDate ? new Date(order.openedDate) : "";
  var paidDate = order.paidDate ? new Date(order.paidDate) : "";
  var closedDate = order.closedDate ? new Date(order.closedDate) : "";

  // Calculate total amount and payment info from checks
  var totalAmount = 0;
  var taxAmount = 0;
  var paymentType = "";
  var paymentStatus = "";
  var totalItems = 0;
  var orderSummary = "";

  if (order.checks && order.checks.length > 0) {
    order.checks.forEach(function (check) {
      totalAmount += check.totalAmount || 0;
      taxAmount += check.taxAmount || 0;

      // Get payment info from first payment
      if (check.payments && check.payments.length > 0) {
        paymentType = check.payments[0].type || "";
        paymentStatus = check.paymentStatus || "";
      }

      // Count items and create summary
      if (check.selections && check.selections.length > 0) {
        check.selections.forEach(function (selection) {
          totalItems += selection.quantity || 0;
          if (orderSummary) orderSummary += ", ";
          orderSummary += (selection.quantity || 1) + "x " + (selection.displayName || "Unknown Item");
        });
      }
    });
  }

  // Prepare order data
  var orderData = [
    order.guid || "",
    order.displayNumber || "",
    order.source || "",
    order.businessDate || "",
    openedDate,
    paidDate,
    closedDate,
    order.duration || "",
    order.numberOfGuests || "",
    order.voided ? "Yes" : "No",
    order.approvalStatus || "",
    totalAmount,
    taxAmount,
    paymentType,
    paymentStatus,
    order.server ? order.server.guid : "",
    order.createdDevice ? order.createdDevice.id : "",
    order.createdInTestMode ? "Yes" : "No",
    totalItems,
    orderSummary,
    order.checks ? order.checks.length + " check(s)" : "0 checks"
  ];

  ordersSheet.getRange(row, 1, 1, orderData.length).setValues([orderData]);
  return row + 1;
}

/**
 * Detailed function to import orders with item-level breakdown
 */
function importOrdersDetailed() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ordersSheet = setupDetailedOrdersSheet(ss, "Orders Detailed");

    Logger.log("Fetching data from Toast API...");
    var data = fetchToastOrdersData();

    if (!data || !Array.isArray(data)) {
      Logger.log("No order data available");
      return;
    }

    Logger.log("Found " + data.length + " orders");

    var row = 2;
    data.forEach(function (order) {
      row = processOrderDetailed(order, ordersSheet, row);
    });

    ordersSheet.autoResizeColumns(1, ordersSheet.getLastColumn());

    Logger.log("Detailed orders imported successfully!");
  } catch (e) {
    Logger.log("Error: " + e.toString());
  }
}

/**
 * Set up detailed sheet with item-level breakdown
 */
function setupDetailedOrdersSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  var headers = [
    "Order GUID", "Order Number", "Business Date", "Order Date",
    "Check GUID", "Check Number", "Item GUID", "Item Name",
    "Item Quantity", "Item Price", "Item Total", "Item Category",
    "Item Group", "Modifiers", "Payment Type", "Payment Amount",
    "Server", "Device", "Order Status", "Item Status"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Process order with detailed item breakdown
 */
function processOrderDetailed(order, ordersSheet, row) {
  var orderDate = order.openedDate ? new Date(order.openedDate) : "";
  var serverGuid = order.server ? order.server.guid : "";
  var deviceId = order.createdDevice ? order.createdDevice.id : "";

  if (order.checks && order.checks.length > 0) {
    order.checks.forEach(function (check) {
      var paymentType = "";
      var paymentAmount = 0;

      if (check.payments && check.payments.length > 0) {
        paymentType = check.payments[0].type || "";
        paymentAmount = check.payments[0].amount || 0;
      }

      if (check.selections && check.selections.length > 0) {
        check.selections.forEach(function (selection) {
          var modifiers = "";
          if (selection.modifiers && selection.modifiers.length > 0) {
            modifiers = selection.modifiers.map(function (mod) {
              return mod.displayName || "Unknown Modifier";
            }).join(", ");
          }

          var itemData = [
            order.guid || "",
            order.displayNumber || "",
            order.businessDate || "",
            orderDate,
            check.guid || "",
            check.displayNumber || "",
            selection.guid || "",
            selection.displayName || "",
            selection.quantity || 0,
            selection.price || 0,
            (selection.quantity || 0) * (selection.price || 0),
            selection.salesCategory ? selection.salesCategory.guid : "",
            selection.itemGroup ? selection.itemGroup.guid : "",
            modifiers,
            paymentType,
            paymentAmount,
            serverGuid,
            deviceId,
            order.approvalStatus || "",
            selection.fulfillmentStatus || ""
          ];

          ordersSheet.getRange(row, 1, 1, itemData.length).setValues([itemData]);
          row++;
        });
      } else {
        // If no selections, still add order info
        var itemData = [
          order.guid || "",
          order.displayNumber || "",
          order.businessDate || "",
          orderDate,
          check.guid || "",
          check.displayNumber || "",
          "",
          "No items",
          0,
          0,
          0,
          "",
          "",
          "",
          paymentType,
          paymentAmount,
          serverGuid,
          deviceId,
          order.approvalStatus || "",
          ""
        ];

        ordersSheet.getRange(row, 1, 1, itemData.length).setValues([itemData]);
        row++;
      }
    });
  }

  return row;
}

/**
 * Helper function to get orders for a specific date
 */
function getOrdersForDate(dateString) {
  var apiUrl = "https://eycxtoast.zeabur.app/api/toast/orders?businessDate=" + dateString;

  try {
    var response = UrlFetchApp.fetch(apiUrl, {
      method: 'get',
      muteHttpExceptions: true
    });

    return JSON.parse(response.getContentText());
  } catch (e) {
    Logger.log("Error fetching orders for date " + dateString + ": " + e.toString());
    return null;
  }
}