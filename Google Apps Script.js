// These are the defined donor tiers. Change these values if the donor tiers need to be changed.
// MAKE SURE OF THE FOLLOWING:
// - HEADERS MUST BE ON ROW 1.
// - ALL HEADERS MUST MATCH THESE EXACTLY ON THE SPREADSHEET.
// - ALL HEADERS MUST BE PRESENT.
// - PUT A SECOND COLUMN WITH THE EXACT HEADER NAME + "SORT KEYS" TO THE RIGHT OF EACH COLUMN (eg. "MAESTRO CLUB $2500+ SORT KEYS")
const Donor_Tiers = Object.freeze({
  MAESTRO: "MAESTRO CLUB $2500+",
  BRAVISSIMO: "BRAVISSIMO CLUB $1000 - $2499",
  VIRTUOSO: "VIRTUOSO CLUB $500 - $999",
  DIVERTIMENTO: "DIVERTIMENTO CLUB $250 - $499",
  ALLEGRO: "ALLEGRO CLUB  $50 - $249",
  CORPORATE: "CORPORATE MATCHING PARTNERS",
  COMMUNITY: "COMMUNITY PARTNERS"
});

// These are suffixes to omit for sorting names. Add to this list if a new suffix needs to be accounted for.
const SUFFIXES = Object.freeze(["Sr.", "Jr.", "III", "IV", "MD", "PhD"]);

// These are prefixes to omit for sorting names. Add to this list if a new prefix needs to be accounted for.
const PREFIXES = Object.freeze(["Dr."]);

// These are strategies for generating default sort keys for trimmed names. INDIVIDUAL will generate a sort key with the last name coming first, then first name, then middle names, then prefixes, then suffixes. ORGANIZATION will simply treat the original string as the sort key but remove "The" at the beginning.
const SORT_KEY_GENERATION_STRATEGIES = Object.freeze({
  INDIVIDUAL: (name) => {
    if (typeof name !== "string") {
      throw new TypeError(`name must be type string. Got ${typeof name} instead.`);
    }
    if (!name) {
      return "";
    }

    const cleanedName = name.replace(/[()",&]/g, "");
    if (cleanedName === "") {
      return "";
    }
    const words = cleanedName.split(/\s+/);

    const suffixSet = new Set(SUFFIXES.map(s => s.toLowerCase()));
    const suffixes = []
    while (words.length > 0 && suffixSet.has(words[words.length - 1].toLowerCase())) {
      suffixes.push(words.pop());
    }
    suffixes.reverse();

    const prefixSet = new Set(PREFIXES.map(p => p.toLowerCase()));
    const prefixes = []
    while (words.length > 0 && prefixSet.has(words[0].toLowerCase())) {
      prefixes.push(words.shift());
    }

    let firstName = "";
    let lastName = "";
    const middleNames = [];
    if (words.length > 1) {
      firstName = words.shift();
      lastName = words.pop();
      while (words.length > 0) {
        middleNames.push(words.shift());
      }
    } else {
      lastName = words.pop();
    }
    
    const components = [
      lastName,
      firstName,
      middleNames.join(" "),
      prefixes.join(" "),
      suffixes.join(" ")
    ];

    return components.filter(part => part !== "").join(" ");
  },
  ORGANIZATION: (name) => {
    if (typeof name !== "string") {
      throw new TypeError(`name must be type string. Got ${typeof name} instead.`);
    }
    if (!name) {
      return "";
    }
    if (name.toLowerCase().startsWith("the ")) {
      return name.slice(4);
    }
    return name;
  }
});

// These are the sort strategies that should be used for each donor tier. Update this when the tiers are changed or the sorting strategy should be changed.
const TIER_SORTING_STRATEGIES = Object.freeze({
  [Donor_Tiers.MAESTRO]: SORT_KEY_GENERATION_STRATEGIES.INDIVIDUAL,
  [Donor_Tiers.BRAVISSIMO]: SORT_KEY_GENERATION_STRATEGIES.INDIVIDUAL,
  [Donor_Tiers.VIRTUOSO]: SORT_KEY_GENERATION_STRATEGIES.INDIVIDUAL,
  [Donor_Tiers.DIVERTIMENTO]: SORT_KEY_GENERATION_STRATEGIES.INDIVIDUAL,
  [Donor_Tiers.ALLEGRO]: SORT_KEY_GENERATION_STRATEGIES.INDIVIDUAL,
  [Donor_Tiers.CORPORATE]: SORT_KEY_GENERATION_STRATEGIES.ORGANIZATION,
  [Donor_Tiers.COMMUNITY]: SORT_KEY_GENERATION_STRATEGIES.ORGANIZATION
});

// This is the API key for uploading the JSON to the Cloudflare worker attached to the R2 bin. Generally speaking, don't worry about this.
const API_KEY = PropertiesService.getScriptProperties().getProperty("R2_API_KEY");

// This is the URL of the Cloudflare worker attached to the R2 bin. Don't change this unless necessary.
const R2_URL = "https://sfcma-donors-api.grantshimada.workers.dev/";

/**
 * Automatically runs when the spreadsheet opens to add the upload button to the toolbar.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🚀 Custom Tools')
    .addItem('Process Data', 'processDonorData')
    .addItem('Upload Data', 'updateDonorData')
    .addToUi();
}

/**
 * Main function to clean and sort donor data.
 * 
 * @returns {Map<string, string[]} - The map of the required headers and their sort keys to the cleaned and sorted data columns
 */
function processDonorData(sheet = SpreadsheetApp.getActiveSheet(), ui = SpreadsheetApp.getUi()) {
  // Initialize sheet and data.
  const allData = sheet.getDataRange().getValues();

  const headerRow = allData[0];
  const dataRows = allData.slice(1);

  // Get the required headers and their corresponding indices and column names.
  const headerIndexMap = (() => {
    try {
      return validateAndMapHeaders(headerRow);
    } catch(e) {
      const userMessage = `${e.message}\n\nPlease resolve all sheet layout issues before trying again.`;
      ui.alert(`[${e.name}]`, `Upload Halted\n${userMessage}`, ui.ButtonSet.OK);
      throw e;
    }
  })();

  const headerDataMap = extractColumns(headerIndexMap, dataRows);

  // Deduplicate duplicate names from the extracted columns based on the user strategy.
  let duplicateHandlingStrategy = null;
  
  const promptUserForDuplicateStrategy = (name) => {
    if (duplicateHandlingStrategy !== null) return duplicateHandlingStrategy;

    const response = ui.alert(
      "Duplicate name detected!",
      `A duplicate name '${name}' was found. Would you like to delete duplicates?\n\n
      [Yes] - Delete any and all duplicate names, keeping the first instance in the highest tier\n
      [No] - Keep all duplicates\n
      [Cancel] - Cancel this upload and handle duplicates manually,`,
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (response === ui.Button.YES) {
      duplicateHandlingStrategy = "DELETE";
    } else if (response === ui.Button.NO) {
      duplicateHandlingStrategy = "KEEP";
    }

    return duplicateHandlingStrategy;
  }

  deduplicateColumns(headerDataMap, promptUserForDuplicateStrategy);
  generateSortKeys(headerDataMap);
  sortColumns(headerDataMap);
  writeColumnsToSheet(sheet, headerIndexMap, headerDataMap);
  
  if (arguments.length === 0) {
    ui.alert("✨ Data Processing Complete!");
  }
  return headerDataMap;
}

/**
 * Main function to upload the data to the Cloudflare database. Calls {@link processDonorData} to make sure the data is processed.
 */
function updateDonorData() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const headerDataMap = processDonorData(sheet, ui);
  const sheetName = sheet.getSheetName().trim();
  const uploadConfirmationAlert = ui.alert("Data Processing Complete!", `Are you sure you want to upload the data for ${sheetName}?`, ui.ButtonSet.YES_NO);

  if (uploadConfirmationAlert !== ui.Button.YES) {
    ui.alert("Upload cancelled.");
    return;
  }

  const payload = createJSONPayload(headerDataMap);
  const response = uploadJSONPayload(sheetName.toLowerCase(), payload);
  Logger.log('Status: ' + response.getResponseCode());
  Logger.log('Response: ' + response.getContentText());
  ui.alert("🚀 Upload was successful!");
}

/**
 * Validates that all required tiers and their sort key columns are present as defined in {@link Donor_Tiers} and that there are no duplicates.
 * Then, returns a lookup map of the required headers to their indices and column names.
 *  
 * @param {Array<string>} headersToCheck - The raw first row of headers from the sheet data
 * @returns {Map<string, {index: number, columnName: string}>} The map of the required headers and their sort keys to their indices and column names
 * @throws {TypeError} If headersToCheck is not an array
 * @throws {Error} If a duplicate required header is found
 * @throws {Error} If a required header is missing
 */

function validateAndMapHeaders(headersToCheck) {
  // Validate that headersToCheck is an array.
  if (!Array.isArray(headersToCheck)) {
    throw new TypeError(`headersToCheck must be an array. Got ${typeof headersToCheck} instead.`);
  }

  const seenMap = new Map();
  const requiredHeaders = new Set(Object.values(Donor_Tiers));
  Array.from(requiredHeaders)
    .map(header => header + " SORT KEYS")
    .forEach(sortHeader => requiredHeaders.add(sortHeader));
  const missingHeaders = new Set(requiredHeaders);

  for (let i = 0; i < headersToCheck.length; i++) {
    const headerText = headersToCheck[i];
    
    // Validate that a duplicate required header is not present. seenMap should only contain required headers.
    if (seenMap.has(headerText)) {
      throw new Error(`Duplicate header ${headerText} found at columns ${getColumnName(i)} and ${seenMap.get(headerText).columnName}.`);
    }
    
    // Check off the header from missingHeaders and add it to seenMap.
    if (requiredHeaders.has(headerText)) {
      missingHeaders.delete(headerText);
      seenMap.set(headerText, {
        index: i,
        columnName: getColumnName(i)
      });
    }
  }

  // Validate that there are no missing headers.
  if (missingHeaders.size > 0) {
    const missingList = Array.from(missingHeaders).join(', ');
    throw new Error(`The following required headers are missing from the sheet layout: ${missingList}.`);
  }

  return seenMap;
}

/**
 * Get the column name associated with an index of data from a Google Sheet.
 * 
 * @param {number} index - The 0-based index of the data
 * @returns {string} The column name associated with the index (ie. 0 => A, 25 => Z, 26 => AZ, 701 => ZZ)
 * @throws {TypeError} If index is not an integer
 * @throws {RangeError} If index is not >= 0
 */

function getColumnName(index) {
  // Validate that index is a non-negative integer.
  if (!Number.isInteger(index)) {
    throw new TypeError(`index must be an integer. Got ${typeof index} instead.`);
  }
  if (index < 0) {
    throw new RangeError(`index must be >= 0. Got ${index}.`);
  }

  // Build the column name
  const res = [];
  const asciiBaseA = "A".charCodeAt(0);

  while (index >= 0) {
    let remainder = index % 26;
    res.push(String.fromCharCode(remainder + asciiBaseA));
    index = Math.floor(index / 26) - 1;
  }

  return res.reverse().join('');
}

/**
 * Extract the columns for the specified headers from the raw data (excluding headers), representing empty cells as "".
 * 
 * @param {Map<string, {index: number, columnName: string}>} headerIndexMap - The map of the header strings to their indices and column names (see {@link validateAndMapHeaders})
 * @param {Array<string[]>} rawData - The matrix of data from the sheet excluding the header row
 * @returns {Map<string, string[]>} The map of the headers to the extracted columns
 */

function extractColumns(headerIndexMap, rawData) {
  const headerDataMap = new Map();
  for (const [header, values] of headerIndexMap) {
    headerDataMap.set(header, rawData
    .map(row => String(row[values.index] || "").trim()));
  }

  return headerDataMap;
}

/**
 * Detects and handles duplicates for the {@link Donor_Tiers} columns in headerDataMap based on the duplicateHandlingStrategy from {@link Duplicate_Handling_Strategies} excluding empty cells. It will turn duplicates into empty strings if deletion is chosen to avoid changing the length of the array.
 * ["DELETE"] - Turn all duplicate instances to the empty string ""
 * ["KEEP"] - Keep all duplicates
 * [null] - Terminate the process
 * 
 * @param {Map<string, string[]} headerDataMap - The map of headers to their columns (arrays) of names/sort keys
 * @param {function} duplicateHandlingStrategy - A function to get the user duplicate handling strategy
 * @throws {Error} If the duplicateHandlingStrategy yields null
 * @modifies {headerDataMap}
 */
function deduplicateColumns(headerDataMap, duplicateHandlingStrategy) {
  const seen = new Set();
  for (const header of Object.values(Donor_Tiers)) {
    const column = headerDataMap.get(header);
    for (let i = 0; i < column.length; i++) {
      const name = column[i];
      if (name === "") {
        continue;
      }
      if (seen.has(name)) {
        const duplicateHandling = duplicateHandlingStrategy(name);
        if (duplicateHandling === null) {
          throw new Error("User halted process due to duplicate name.");
        } else if (duplicateHandling === "DELETE") {
          column[i] = "";
          continue;
        } else if (duplicateHandling === "KEEP") {
          continue;
        }
      }
      seen.add(name);
    }
  }
}

/** Generates default sort keys for the names of each column in headerDataMap based on the sorting strategy indicated by {@link TIER_SORTING_STRATEGIES}, defined in {@link SORT_KEY_GENERATION_STRATEGIES}, if a sort key does not exist already.
 * 
 * @param {Map<string, string[]} headerDataMap - The map of headers to their columns (arrays) of names/sort keys
 * @modifies headerDataMap 
 */
function generateSortKeys(headerDataMap) {
  for (const header of Object.values(Donor_Tiers)) {
    const sortKeyHeader = header + " SORT KEYS";
    const sortStrategy = TIER_SORTING_STRATEGIES[header] || SORT_KEY_GENERATION_STRATEGIES.INDIVIDUAL;
    const headerColumn = headerDataMap.get(header);
    const headerSortColumn = headerDataMap.get(sortKeyHeader);

    for (let i = 0; i < headerColumn.length; i++) {
      if (!headerColumn[i]) {
        headerSortColumn[i] = "";
      } else if (headerColumn[i].startsWith("Anonymous")) {
        headerSortColumn[i] = "";
      } else if (!headerSortColumn[i]) {
        headerSortColumn[i] = sortStrategy(headerColumn[i]);
      }
    }
  }
}

/**
 * Sorts the columns of the donor tiers in headerDataMap according to their corresponding sort key columns, putting "Anonymous - X" entries at the beginning and empty strings at the end.
 * 
 * @param {Map<string, string[]} headerDataMap - The map of headers to their columns (arrays) of names/sort keys
 * @modifies {headerDataMap}
 */
function sortColumns(headerDataMap) {
  for (const header of Object.values(Donor_Tiers)) {
    const headerColumn = headerDataMap.get(header); 
    const sortKeyColumn = headerDataMap.get(header + " SORT KEYS")
    const zippedColumns = headerColumn.map((name, index) => ({
      name: name,
      sortKey: sortKeyColumn[index]
    }));

    zippedColumns.sort((a, b) => {
      const nameA = a.name;
      const sortKeyA = a.sortKey;

      const nameB = b.name;
      const sortKeyB = b.sortKey;

      // Handle anonymous count entries
      const isAnonA = nameA.startsWith("Anonymous");
      const isAnonB = nameB.startsWith("Anonymous");
      if (isAnonA && !isAnonB) return -1;
      if (!isAnonA && isAnonB) return 1;
      if (isAnonA && isAnonB) return 0;

      // Take care of edge cases where names are present but have a blank sort key.
      let rankA = 1;
      if (nameA !== "" && sortKeyA === "") rankA = 2;
      if (nameA === "" && sortKeyA === "") rankA = 3;

      let rankB = 1;
      if (nameB !== "" && sortKeyB === "") rankB = 2;
      if (nameB === "" && sortKeyB === "") rankB = 3;
      
      if (rankA !== rankB) {
        return rankA - rankB;
      } else if (rankA === 1) {
        return sortKeyA.localeCompare(sortKeyB, undefined, {sensitivity: 'base', numeric: true});
      } else if (rankA === 2) {
        return nameA.localeCompare(nameB, undefined, {sensitivity: 'base', numeric: true});
      }
      return 0;
    });

    for (let i = 0; i < headerColumn.length; i++) {
      headerColumn[i] = zippedColumns[i].name;
      sortKeyColumn[i] = zippedColumns[i].sortKey;
    }
  }
}

/**
 * Write the final columns to the sheet.
 * 
 * @param {Sheet} sheet - The active Google Sheet object
 * @param {headerIndexMap} - The map of the headers to their indices and column names (see {@link validateAndMapHeaders})
 * @param {headerDataMap} - The map of the headers to the column data
 */
function writeColumnsToSheet(sheet, headerIndexMap, headerDataMap) {
  for (const [header, column] of headerDataMap) {
    const columnIndex = headerIndexMap.get(header).index;
    sheet.getRange(2, columnIndex + 1, column.length, 1)
      .setValues(column.map(name => [name]));
  }
}

/**
 * Create the JSON string to be uploaded.
 * 
 * @param {Map<string, string[]} headerDataMap - The map of headers to the processed (sorted and duplicate handled) columns
 * @returns {string} - A JSON String payload with the headers as the keys and the columns as the values
 */
function createJSONPayload(headerDataMap) {
  for ([header, column] of headerDataMap) {
    const idx = column.indexOf("");
    if (idx !== -1) {
      column.length = idx;
    }
  }
  const payloadObject = Object.fromEntries(
    [...headerDataMap].filter(([header]) => !header.endsWith("SORT KEYS"))
  );
  return JSON.stringify(payloadObject);
}

/**
 * Upload the JSON payload based on the years (ie. '2026-2027') given by the sheet's name to the Cloudflare R2 database.
 * 
 * @param <string> years - The years for the donor information payload, which should match the sheet name (ie. '2026-2027')
 * @param <string> payload - The JSON string payload of donor information
 * @returns {UrlFetchApp.HTTPResponse} - The full Google Apps Script response object
 * @throws {TypeError} - If years or payload are not strings
 * @throws {Exception} - If the HTTP request fails (ie. 404 error)
 */
function uploadJSONPayload(years, payload) {
  if (typeof years !== "string") {
    throw new TypeError(`years must be type string, got ${typeof years} instead.`);
  }
  if (typeof payload !== "string") {
    throw new TypeError(`payload must be type string, got ${typeof payload} instead.`);
  }

  const fullUrl = `${R2_URL}donor-recognition-${years}`;
  const response = UrlFetchApp.fetch(fullUrl, {
    method: 'put',
    contentType: 'application/json',
    headers: {
      'X-API-Key': API_KEY
    },
    payload
  })

  return response
}

