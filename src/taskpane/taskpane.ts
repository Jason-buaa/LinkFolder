/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = parse;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      let sheet = context.workbook.worksheets.getActiveWorksheet();

      // Create the headers and format them to stand out.
      let headers = [["Product", "Quantity", "Unit Price", "Totals"]];
      let headerRange = sheet.getRange("B2:E2");
      headerRange.values = headers;
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.color = "white";

      // Create the product data rows.
      let productData = [
        ["Almonds", 6, 7.5],
        ["Coffee", 20, 34.5],
        ["Chocolate", 10, 9.56],
      ];
      let dataRange = sheet.getRange("B3:D5");
      dataRange.values = productData;

      // Create the formulas to total the amounts sold.
      let totalFormulas = [["=C3 * D3"], ["=C4 * D4"], ["=C5 * D5"], ["=SUM(E3:E5)"]];
      let totalRange = sheet.getRange("E3:E6");
      totalRange.formulas = totalFormulas;
      totalRange.format.font.bold = true;

      // Display the totals as US dollar amounts.
      totalRange.numberFormat = [["$0.00"]];

      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}

// Function to parse the snippet
function parseSnippet(snippet: String) {
  // Define a regular expression to match the content between /begin HEADER and /end HEADER
  const headerRegex = /\/begin HEADER.*?\n([\s\S]*?)\/end HEADER/m;
  const match = snippet.match(headerRegex);

  if (match && match[1]) {
    const content = match[1].trim();
    const lines = content.split("\n");
    const parsedData = {};

    lines.forEach((line) => {
      // Split each line by whitespace to get key and value
      const [key, ...valueParts] = line.trim().split(/\s+/);
      const value = valueParts.join(" ");
      parsedData[key] = value.replace(/"/g, ""); // Remove quotes from values if any
    });

    return parsedData;
  } else {
    throw new Error("Header content not found");
  }
}
/*
Type in A1:
/begin PROJECT A211 "MDG1"
  /begin HEADER ""
    VERSION    "C914"
    PROJECT_NO A211
  /end HEADER
  
*/


export async function parse() {
  try {
    await Excel.run(async (context) => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      let headerRange = sheet.getRange("A1");
      headerRange.load("values");
      await context.sync();
      console.log(headerRange.values);
      const parsedData = parseSnippet(headerRange.values[0][0]);
      console.log(parsedData);
      // Start writing from cell B2
      let startRow = 2; // B2 means row 2, column B
      const startColumn = "B";

      // Iterate over the parsedData object
      for (const [key, value] of Object.entries(parsedData)) {
        // Set the key in the current row, column B
        sheet.getRange(`${startColumn}${startRow}`).values = [[key]];
        // Set the value in the current row, column C
        sheet.getRange(`${String.fromCharCode(startColumn.charCodeAt(0) + 1)}${startRow}`).values = [[value]];
        // Move to the next row
        startRow++;
      }
      //await context.sync();
    });
  } catch (error) {
    console.error(error.message);
  }
}
