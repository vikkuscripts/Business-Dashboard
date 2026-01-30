function doGet(request) {
  return HtmlService.createTemplateFromFile('main')
      .evaluate()
      .setTitle("Login Page")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Checks whether the logged‐in user’s project is active
function license_Status() {
  var userInfoId = '1xo40DrkDnkEkWjYJ5-SjJ6sZgI26ZdXYNjZuq_rGciE';
  var now = new Date();
  var timeZone = Session.getScriptTimeZone();
  var currentDate = Utilities.formatDate(now, timeZone, "dd/MM/yyyy");
  var usr = Session.getActiveUser().getEmail();
  var domain = usr.substring(usr.lastIndexOf("@") + 1);
  
  var ss = SpreadsheetApp.openById(userInfoId);
  var sheet = ss.getSheetByName('Demo');
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, 4).getValues();
  
  // Find row that matches the domain (skipping header row)
  var matchIndex = data.findIndex((row, idx) => idx > 0 && row[0] === domain);
  if (matchIndex < 0) {
    return HtmlService.createHtmlOutput('Your project is not active. Please contact support.');
  }
  
  var projectStatus = data[matchIndex][2];
  if (projectStatus === 'Active') {
    return HtmlService.createHtmlOutputFromFile('main')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("BAIITBOX-ERP Solutions");
  } else {
    return HtmlService.createHtmlOutput('Your project is not active. Please contact support.');
  }
}

// Validates login credentials and retrieves user's metrics visibility setting
function checkLogin(username, password, role) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("User");
  const data = sheet.getDataRange().getValues();
  // Assuming "Metrics Visibility" is the 7th column (index 6)
  const METRICS_VISIBILITY_COLUMN_INDEX = 6; 

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === password && String(data[i][3]).toLowerCase() === String(role).toLowerCase()) {
      
      // Check the 'Metrics Visibility' column. Treat "Yes" as true, otherwise false.
      const metricsVisibilityValue = data[i][METRICS_VISIBILITY_COLUMN_INDEX] || '';
      const canViewMetrics = String(metricsVisibilityValue).trim().toLowerCase() === 'yes';

      return { 
        success: true, 
        displayName: data[i][2],
        metricsVisible: canViewMetrics,
        profilePictureUrl: data[i][7] || '' // Fetching profile picture from column 8
      };
    }
  }
  return { success: false, displayName: null, metricsVisible: false };
}

// Returns the menu structure (from "Master Sheet" and "Permission" sheets) for the given user
/**
 * Returns the menu structure (from "Master Sheet", "Permission", and "Create Button") for the given user.
 * Also attaches any permitted buttons for each sub-menu.
 */
/**
 * Returns the menu structure (from "Master Sheet", "Permission", and "Create Button")
 * for the given userName. Looks up the user's role from the "User" sheet:
 *   - If role = 'admin', attach ALL buttons for each sub-menu.
 *   - Otherwise, filter by Permission sheet as before.
 */
function getMenuStructure(userName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master Sheet");
  const permissionSheet = ss.getSheetByName("Permission");
  const createButtonSheet = ss.getSheetByName("Create Button");

  // Read the user's role (used later for table data; for menu, we now always filter by Permission)
  const role = getUserRoleByEmail(userName) || 'user';

  // 1) Read the Master Sheet (columns: mainMenu, subMenu, sheetName, dbLink, externalId, rangeSpec, Question, Table)
  const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 8).getValues();
  const menuStructure = {};
  masterData.forEach(row => {
    const mainMenu = String(row[0]).trim();
    const subMenu = String(row[1]).trim();
    const sheetName = String(row[2]).trim();
    const dbLink = String(row[3]).trim();
    const externalId = String(row[4]).trim();
    const rangeSpec = row[5] ? String(row[5]).trim() : "";
    
    // Combine Question and Table into (Q)(T) for frontend compatibility if needed, 
    // or just pass them as requested. For now, let's pass a combined string for backward compatibility.
    const q = row[6] ? String(row[6]).trim() : "";
    const t = row[7] ? String(row[7]).trim() : "";
    const entryViewable = (q ? "(" + q + ")" : "") + (t ? "(" + t + ")" : "");
    
    if (!menuStructure[mainMenu]) {
      menuStructure[mainMenu] = [];
    }
    menuStructure[mainMenu].push({
      subMenu: subMenu,
      sheetName: sheetName,
      dbLink: dbLink,
      externalId: externalId,
      rangeSpec: rangeSpec,
      entryViewable: entryViewable,
      buttons: []
    });
  });

  // 2) Read the "Create Button" sheet into a button map.
    const lastRowBtn = createButtonSheet.getLastRow();
    let buttonData = [];
    if (lastRowBtn >= 2) { 
      buttonData = createButtonSheet.getRange(2, 1, lastRowBtn - 1, 4).getValues();
    }
  const buttonMap = {};
  buttonData.forEach(row => {
    const bMenu = String(row[0]).trim();
    const bSub = String(row[1]).trim();
    const bName = String(row[2]).trim();
    const bUrl = String(row[3]).trim();
    if (bMenu && bSub && bName && bUrl) {
      const key = bMenu + "||" + bSub + "||" + bName;
      buttonMap[key] = bUrl;
    }
  });

  // 3) Read the Permission sheet and build a permission map.
  const lastRowPerm = permissionSheet.getLastRow();
  let permData = [];
  if (lastRowPerm >= 2) {
    permData = permissionSheet.getRange(2, 1, lastRowPerm - 1, 5).getValues();
  }
  const userPermissions = [];
  const accessTypeMap = {}; 
  
  permData.forEach(row => {
    let [permUser, permMenu, permSubMenu, permButtonNames, accessType] = row;
    if (!permUser || !permMenu || !permSubMenu) return;
    
    if (String(permUser).toLowerCase() === String(userName).toLowerCase()) {
      if (permButtonNames && typeof permButtonNames === 'string' && permButtonNames.indexOf(',') >= 0) {
        const splitBtns = permButtonNames.split(',').map(s => s.trim()).filter(s => s !== "");
        splitBtns.forEach(bName => {
          userPermissions.push({
            menu: String(permMenu).trim(),
            subMenu: String(permSubMenu).trim(),
            buttonName: bName
          });
        });
      } else {
        userPermissions.push({
          menu: String(permMenu).trim(),
          subMenu: String(permSubMenu).trim(),
          buttonName: permButtonNames ? String(permButtonNames).trim() : ""
        });
      }
      
      if (!accessTypeMap[permMenu]) {
        accessTypeMap[permMenu] = {};
      }
      accessTypeMap[permMenu][permSubMenu] = String(accessType || "").trim();
    }
  });

  // Group permissions by main menu and sub-menu.
  // Always register a permission record, even if the button name is empty.
  const userPermMap = {};
  userPermissions.forEach(({ menu, subMenu, buttonName }) => {
    if (!userPermMap[menu]) {
      userPermMap[menu] = {};
    }
    if (!userPermMap[menu][subMenu]) {
      userPermMap[menu][subMenu] = [];
    }
    userPermMap[menu][subMenu].push(buttonName);
  });

  // Only include sub-menus for which there is a permission record.
  // 4) Filter the master menu based on permissions (SKIP if no userName provided):
  Object.keys(menuStructure).forEach(mainMenu => {
    // If userName is provided, filter sub-menus for this main menu based on the Permission sheet:
    if (userName && userName.trim() !== "") {
      menuStructure[mainMenu] = menuStructure[mainMenu].filter(item => {
        const subName = item.subMenu;
        // Only include the sub-menu if there's a permission record for it
        if (!userPermMap[mainMenu] || !userPermMap[mainMenu][subName]) {
          return false;
        }
        // For permitted sub-menus, attach only the allowed buttons:
        const permBtnNames = userPermMap[mainMenu][subName];
        const buttonObjects = [];
        permBtnNames.forEach(bName => {
          const key = mainMenu + "||" + subName + "||" + bName;
          if (buttonMap[key]) {
            buttonObjects.push({
              name: bName,
              url: buttonMap[key]
            });
          }
        });
        item.buttons = buttonObjects;
        
        // Add the access type from our separate map
        item.accessType = accessTypeMap[mainMenu] && accessTypeMap[mainMenu][subName]
          ? accessTypeMap[mainMenu][subName] 
          : "";
          
        return true;
      });
    } else {
      // IF NEW USER: Attach ALL buttons from the buttonMap for each sub-menu
      menuStructure[mainMenu].forEach(item => {
        const subName = item.subMenu;
        const buttonObjects = [];
        // Loop through all buttons in buttonMap to find those belonging to this menu/submenu
        for (const key in buttonMap) {
          if (key.indexOf(mainMenu + "||" + subName + "||") === 0) {
            const bName = key.split("||")[2];
            buttonObjects.push({
              name: bName,
              url: buttonMap[key]
            });
          }
        }
        item.buttons = buttonObjects;
        item.accessType = ""; // New user has no access types set yet
      });
    }
    
    // If there are no permitted sub-menus for this main menu, remove it:
    if (menuStructure[mainMenu].length === 0) {
      delete menuStructure[mainMenu];
    }
  });

  return menuStructure;
}

// Formats a date to "yyyy-MM-dd" or returns the value if not a date
function formatDate(date) {
  if (date instanceof Date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return date;
}

// Retrieves data from a specified sheet
function getSheetData(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    return { error: `Sheet "${sheetName}" does not exist.` };
  }
  const data = sheet.getDataRange().getValues();
  return { success: true, data: data };
}

// Retrieves the image URL from the "Setup" sheet (cell B2)
function getImageUrl() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setup");
  return sheet.getRange('B2').getValue();
}

// Retrieves the business title from the "Setup" sheet (cell B1)
function getBusinessTitle() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Setup");
  return sheet.getRange('B1').getValue();
}

// Retrieves and formats the data from the "Dashboard" sheet
function getSheetData2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const data = sheet.getDataRange().getValues();
  return data.map(row => row.map(cell => (cell instanceof Date ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'MM/dd/yyyy') : cell)));
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// Retrieves the login card image URL from the "SetUP" sheet (cell B3)
function getLoginImageUrl() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("SetUP");
  if (!sheet) {
    return "Sheet 'SetUP' not found";
  }
  var value = sheet.getRange("B3").getValue();
  return value;
}

// Retrieves the company logo from the "SetUP" sheet (cell B4) using Drive thumbnail logic
function getCompanyLogo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SetUP");
  const fileUrl = sheet.getRange("B2").getValue();
  if (fileUrl) {
    const fileId = extractFileId(fileUrl);
    if (fileId) {
      return getThumbnailImageUrl(fileId);
    }
  }
  return ""; // Return empty string if no logo is available
}

function extractFileId(fileUrl) {
  const regex = /\/d\/([a-zA-Z0-9_-]+)/;
  const matches = fileUrl.match(regex);
  if (matches && matches[1]) {
    return matches[1];
  }
  return null;
}

function getThumbnailImageUrl(fileId) {
  return `https://drive.google.com/thumbnail?id=${fileId}&sz=200`;
}

/**
 * Given a user email, returns the user's role by looking it up
 * in the "User" sheet. Returns null if the user is not found.
 */
function getUserRoleByEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var userSheet = ss.getSheetByName("User");
  if (!userSheet) {
    // No 'User' sheet found
    return null;
  }

  // Read all rows in the 'User' sheet
  var data = userSheet.getDataRange().getValues();
  // data[0] is the header row: ["User ID", "Password", "Display Name", "Role", "Manager", "Profile Picture", ...]
  // We'll loop from row 1 downward
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var userID = row[0];  // "User ID" (the email, presumably)
    var role   = row[3];  // "Role" is in the 4th column (index 3)

    // Compare case-insensitively if the userID matches the login email
    if (String(userID).toLowerCase() === String(email).toLowerCase()) {
      return role; // Return the role
    }
  }

  // If not found, return null (or "user" if you want a default)
  return null;
}


/**
 * getParamHtml(sheetName, pageTitleArg, externalId, rangeSpec, userName, role, buttons)
 *
 * If role = 'admin', show all data and all buttons.
 * If role != 'admin', filter rows by user’s email and show only the buttons that were passed in.
 */
/**
 * SERVER-SIDE: Return the ParamIndex.html for a given sheet, injecting
 * the final array of data. If a column is "DateRange," keep typed date
 * objects; otherwise store the sheet's displayed text.
 * 
 * Also skip creating filters for "DriveImage" in ParamIndex (see Part B below).
 */
function getParamHtml(sheetName, pageTitleArg, externalId, rangeSpec, userName, role, buttons, accessType) {
  // 1) Decide which spreadsheet to open
  let ss;
  if (externalId && externalId.trim() !== "") {
    ss = SpreadsheetApp.openById(externalId.trim());
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  // 2) Get the sheet
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return "ERROR: Sheet '" + sheetName + "' not found.";
  }

  // 3) Decide which range to read
  let range;
  if (rangeSpec && rangeSpec.trim() !== "") {
    try {
      range = sheet.getRange(rangeSpec.trim());
    } catch (e) {
      return "ERROR: Invalid range '" + rangeSpec + "' for sheet '" + sheetName + "'. " + e;
    }
  } else {
    // fallback: entire sheet
    range = sheet.getDataRange();
  }

  // Read typed values and displayed text
  let typedValues    = range.getValues();
  let displayValues  = range.getDisplayValues();

  // We require at least 3 rows for columnIDs, filterTypes, headers
  if (typedValues.length < 3) {
    return "ERROR: Not enough rows in '" + sheetName + "'.";
  }

  // Pull row0 = columnIdentifiers, row1 = filterTypes, row2 = headers
  var columnIdentifiers = typedValues[0];
  var filterTypes       = typedValues[1];
  var headers           = typedValues[2];

  // NEW: Identify logic columns (columns with headers starting with "Logic")
  var logicColumnIndices = [];
  var dataColumnIndices = [];
  var logicData = {};

  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).toLowerCase().startsWith('logic')) {
      logicColumnIndices.push(i);
    } else {
      dataColumnIndices.push(i);
    }
  }

  // NEW: Extract logic data from rows 3 onwards
  for (var r = 3; r < typedValues.length; r++) {
    var rowIsEmpty = typedValues[r].every(function(cell) {
      return String(cell).trim() === "";
    });
    if (rowIsEmpty) continue;

    logicData[r] = {};
    for (var l = 0; l < logicColumnIndices.length; l++) {
      var logicColIndex = logicColumnIndices[l];
      var logicName = headers[logicColIndex];
      var colorValue = displayValues[r][logicColIndex];
      if (colorValue && colorValue.trim() !== "") {
        logicData[r][logicName] = colorValue.trim();
      }
    }
  }

  // Build a data array from row3 onward, excluding logic columns
  let data = [];
  for (var r = 3; r < typedValues.length; r++) {
    // skip any row that is completely blank
    var rowIsEmpty = typedValues[r].every(function(cell) {
      return String(cell).trim() === "";
    });
    if (rowIsEmpty) continue;

    var typedRow   = typedValues[r];
    var dispRow    = displayValues[r];
    var finalRow   = [];

    // Only include data columns, not logic columns
    for (var c = 0; c < dataColumnIndices.length; c++) {
      var colIndex = dataColumnIndices[c];
      if (filterTypes[colIndex] === 'DateRange') {
        // Keep the typed value so it can remain a Date object if the cell was a real date
        finalRow.push(typedRow[colIndex]);
      } else {
        // Keep the raw displayed text for everything else
        finalRow.push(dispRow[colIndex]);
      }
    }
    
    // Add original row index as the last element of the array
    finalRow.push(r);
    data.push(finalRow);
  }

  // Filter column-related arrays to only include data columns
  var dataColumnIdentifiers = dataColumnIndices.map(i => columnIdentifiers[i]);
  var dataFilterTypes = dataColumnIndices.map(i => filterTypes[i]);
  var dataHeaders = dataColumnIndices.map(i => headers[i]);

  // 4) If role not provided, find it from the User sheet
  if (!role) {
    role = getUserRoleByEmail(userName) || 'user';
  }

  // 5) Filter out data if user is not admin AND accessType is not "Full"
  if (role.toLowerCase() !== 'admin' && accessType !== 'Full') {
    // For example, if col0 is the user's email
    data = data.filter(function(row) {
      return String(row[0]).toLowerCase() === String(userName).toLowerCase();
    });
  }
  // If role is admin OR accessType is "Full", we keep all data

  // 6) Build the ParamIndex template
  var template = HtmlService.createTemplateFromFile('ParamIndex');
  template.columnIdentifiers = dataColumnIdentifiers;
  template.filterTypes       = dataFilterTypes;
  template.headers           = dataHeaders;
  template.filteredData      = data;
  template.pageTitle         = pageTitleArg;
  template.timeZone          = ss.getSpreadsheetTimeZone();

  template.entryViewable     = getEntryViewable(sheetName, externalId);

  // 7) Attach buttons
  template.buttons = buttons || [];

  // 8) Add context parameters to the template 
  template.sheetName = sheetName;
  template.externalId = externalId; 
  template.rangeSpec = rangeSpec;
  template.userName = userName;
  template.role = role;
  template.accessType = accessType;

  // NEW: Add logic data to template
  template.logicData = logicData;
  template.logicColumnNames = logicColumnIndices.map(i => headers[i]);

  // 9) Return the rendered HTML
  return template.evaluate().getContent();
}

// In Code.gs, ensure the refreshParamData() function is fully refreshing the data:
// In Code.gs, replace the entire refreshParamData function with this:
function refreshParamData(params) {
  // Return the complete HTML with fresh data
  // This will ensure that all columns, including new ones, are included
  
  // If accessType was not provided, look it up from Master Sheet
  let accessType = params.accessType || "";
  if (!accessType) {
    // Get accessType from the Master Sheet (similar to how getMenuStructure does it)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const permissionSheet = ss.getSheetByName("Permission");
    const lastRowPerm = permissionSheet.getLastRow();
    let permData = [];
    if (lastRowPerm >= 2) {
      permData = permissionSheet.getRange(2, 1, lastRowPerm - 1, 5).getValues();
    }
    
    // Find matching permission for this user and submenu
    for (let i = 0; i < permData.length; i++) {
      let [permUser, permMenu, permSubMenu, permButtonNames, permAccessType] = permData[i];
      if (String(permUser).toLowerCase() === String(params.userName).toLowerCase() &&
          String(permSubMenu).trim() === String(params.pageTitle).trim()) {
        accessType = String(permAccessType || "").trim();
        break;
      }
    }
  }
  
  // Get the buttons for this submenu
  let buttons = [];
  try {
    // Step 1: Get menu structure to retrieve the buttons
    const menuStructure = getMenuStructure(params.userName);
    
    // Step 2: Find buttons for this specific submenu
    for (const mainMenu in menuStructure) {
      for (const item of menuStructure[mainMenu]) {
        if (item.sheetName === params.sheetName && 
            item.subMenu === params.pageTitle && 
            item.externalId === params.externalId) {
          buttons = item.buttons || [];
          break;
        }
      }
    }
  } catch (e) {
    // If there's an error, we'll just use an empty array for buttons
    // but we'll log the error for debugging
    Logger.log("Error retrieving buttons: " + e.toString());
  }
  
  return getParamHtml(
    params.sheetName, 
    params.pageTitle, 
    params.externalId, 
    params.rangeSpec, 
    params.userName, 
    params.role, 
    buttons, // Pass the retrieved buttons here instead of empty array
    accessType // Use looked-up or provided accessType
  );
}

/**
 * Gets the "Entry Viewable" setting for a sheet from the Master Sheet.
 * The format should be: (A,B,D),(C,E) where:
 * - First parenthesis group (A,B,D) contains columns to display in label-data format
 * - Second parenthesis group (C,E) contains columns to display in tabular format
 * - Column references can be single letters (A,B,C...) or multi-letter (AA,AB,AC...)
 * - If empty, no view card will be available
 * - If only one group is provided like (A,B,C), only label-data format will be used
 */
function getEntryViewable(sheetNameToFind, externalIdToFind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName("Master Sheet");
  const data = masterSheet.getDataRange().getValues();
  
  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Match both sheet name and external ID (or empty external ID)
    if (row[2] === sheetNameToFind && (row[4] === externalIdToFind || (!row[4] && !externalIdToFind))) {
      // Columns G (6) and H (7)
      const q = row[6] ? String(row[6]).trim() : "";
      const t = row[7] ? String(row[7]).trim() : "";
      return (q ? "(" + q + ")" : "") + (t ? "(" + t + ")" : "");
    }
  }
  return ""; 
}

// MIS Score Backend Functions - Add these to Code.gs

/**
 * Get the external sheet ID from SetUp sheet
 */
function getExternalSheetId() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const setupSheet = ss.getSheetByName("SetUP");
    if (!setupSheet) return null;
    
    return setupSheet.getRange("B3").getValue();
  } catch (error) {
    console.error("Error getting external sheet ID:", error);
    return null;
  }
}

/**
 * Get MIS data from external sheet
 */
function getMISData(userName, role) {
  try {
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) {
      return { error: "External sheet ID not found in SetUP sheet cell B3" };
    }
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    const dataSheet = externalSS.getSheetByName("data");
    
    if (!dataSheet) {
      return { error: "Data sheet not found in external spreadsheet" };
    }
    
    const data = dataSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    // Filter data based on user role
    let filteredData = data.slice(1); // Remove header
    
    if (role !== 'admin') {
      filteredData = filteredData.filter(row => 
        String(row[0]).toLowerCase() === String(userName).toLowerCase()
      );
    }
    
    return { success: true, data: filteredData, headers: data[0] };
  } catch (error) {
    console.error("Error getting MIS data:", error);
    return { error: "Failed to retrieve MIS data: " + error.toString() };
  }
}

/**
 * Get week mapping data
 */
function getWeekMappingData() {
  try {
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) return [];
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    const weekMapSheet = externalSS.getSheetByName("WeekMap");
    
    if (!weekMapSheet) {
      return []; // Return empty array if WeekMap doesn't exist
    }
    
    const data = weekMapSheet.getDataRange().getValues();
    return data.length > 1 ? data.slice(1) : []; // Remove header if exists
  } catch (error) {
    console.error("Error getting week mapping:", error);
    return [];
  }
}

/**
 * Get work commitments data
 */
function getWorkCommitments() {
  try {
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) return [];
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    const commitmentSheet = externalSS.getSheetByName("workCommitments");
    
    if (!commitmentSheet) {
      return []; // Return empty array if sheet doesn't exist
    }
    
    const data = commitmentSheet.getDataRange().getValues();
    return data.length > 1 ? data.slice(1) : []; // Remove header if exists
  } catch (error) {
    console.error("Error getting work commitments:", error);
    return [];
  }
}

/**
 * Calculate MIS scores for dashboard
 */
function calculateMISScores(userName, role, startDate, endDate) {
  try {
    const misDataResult = getMISData(userName, role);
    if (misDataResult.error) {
      return misDataResult;
    }
    
    const misData = misDataResult.data;
    const workCommitments = getWorkCommitments();
    
    // Parse dates - ensure we include the full day
    const filterStartDate = new Date(startDate);
    filterStartDate.setHours(0, 0, 0, 0);
    
    const filterEndDate = new Date(endDate);
    filterEndDate.setHours(23, 59, 59, 999);
    
    // Filter data by date range (Planned column - index 2)
    const filteredData = misData.filter(row => {
      if (!row[2]) return false; // Skip empty planned dates
      const plannedDate = new Date(row[2]);
      plannedDate.setHours(0, 0, 0, 0);
      return plannedDate >= filterStartDate && plannedDate <= filterEndDate;
    });
    
    // Group by user
    const userGroups = {};
    filteredData.forEach(row => {
      const email = row[0];
      const doerName = row[1];
      
      if (!userGroups[email]) {
        userGroups[email] = {
          email: email,
          name: doerName,
          tasks: []
        };
      }
      userGroups[email].tasks.push(row);
    });
    
    // Calculate scores for each user
    const results = [];
    
    Object.keys(userGroups).forEach(email => {
      const userGroup = userGroups[email];
      const tasks = userGroup.tasks;
      
      // Calculate metrics
      const totalTasks = tasks.length;
      const completedTasks = tasks.filter(task => task[3] && task[3] !== '').length; // Actual column
      const onTimeTasks = tasks.filter(task => {
        if (!task[3] || task[3] === '') return false;
        const actualDate = new Date(task[3]);
        const plannedDate = new Date(task[2]);
        return actualDate <= plannedDate;
      }).length;
      
      // Get the week number for the SELECTED week (not current calendar week)
      const selectedWeekNumber = getWeekNumber(filterStartDate);
      const previousWeekNumber = selectedWeekNumber - 1;
      
      // Get previous week commitments (for "Current Week Score")
      const previousCommitments = workCommitments.find(commitment => 
        commitment[0] === email && commitment[1] === previousWeekNumber
      );
      
      const currentWeekWND = previousCommitments ? (previousCommitments[4] || 0) : 0;
      const currentWeekWNDOT = previousCommitments ? (previousCommitments[5] || 0) : 0;
      
      // Calculate scores
      const wndScore = totalTasks > 0 ? ((completedTasks / totalTasks) * 100) - 100 : -100;
      const wndotScore = completedTasks > 0 ? ((onTimeTasks / completedTasks) * 100) - 100 : -100;
      
      // Get current week commitments (for "Next Commitment" - the week being viewed)
      const currentCommitments = workCommitments.find(commitment => 
        commitment[0] === email && commitment[1] === selectedWeekNumber
      );
      
      results.push({
        userId: email,
        name: userGroup.name,
        currentWeekScoreWND: currentWeekWND,
        totalTasks: totalTasks,
        totalCompletedTasks: completedTasks,
        wndScore: Math.round(wndScore * 100) / 100,
        nextCommitmentWND: currentCommitments ? (currentCommitments[4] || '') : '',
        currentWeekScoreWNDOT: currentWeekWNDOT,
        totalCompleteTasks: completedTasks,
        onTimeDoneTasks: onTimeTasks,
        wndotScore: Math.round(wndotScore * 100) / 100,
        nextCommitmentWNDOT: currentCommitments ? (currentCommitments[5] || '') : '',
        weekNumber: selectedWeekNumber
      });
    });
    
    return { success: true, data: results };
  } catch (error) {
    console.error("Error calculating MIS scores:", error);
    return { error: "Failed to calculate MIS scores: " + error.toString() };
  }
}

/**
 * Get week number for a date
 */
function getWeekNumber(date) {
  const weekMapData = getWeekMappingData();
  
  if (weekMapData.length > 0) {
    // Use custom week mapping
    for (let i = 0; i < weekMapData.length; i++) {
      const startDate = new Date(weekMapData[i][0]);
      const endDate = new Date(weekMapData[i][1]);
      if (date >= startDate && date <= endDate) {
        return weekMapData[i][2];
      }
    }
  }
  
  // Default week calculation (starting from Jan 1)
  const startOfYear = new Date(date.getFullYear(), 0, 1);
  const diffInTime = date.getTime() - startOfYear.getTime();
  const diffInDays = Math.ceil(diffInTime / (1000 * 3600 * 24));
  return Math.ceil(diffInDays / 7);
}

/**
 * Get process details for a specific user and date range
 */
function getProcessDetails(userEmail, startDate, endDate) {
  try {
    const misDataResult = getMISData(userEmail, 'user'); // Always filter to specific user
    if (misDataResult.error) {
      return misDataResult;
    }
    
    const misData = misDataResult.data;
    
    // Filter by date range - ensure we include the full day
    const filterStartDate = new Date(startDate);
    filterStartDate.setHours(0, 0, 0, 0);
    
    const filterEndDate = new Date(endDate);
    filterEndDate.setHours(23, 59, 59, 999);
    
    const filteredData = misData.filter(row => {
      if (!row[2]) return false; // Skip empty planned dates
      const plannedDate = new Date(row[2]);
      plannedDate.setHours(0, 0, 0, 0);
      return plannedDate >= filterStartDate && plannedDate <= filterEndDate;
    });
    
    // Group by process
    const processGroups = {};
    filteredData.forEach(row => {
      const processName = row[5]; // Process Name column
      
      if (!processGroups[processName]) {
        processGroups[processName] = [];
      }
      processGroups[processName].push(row);
    });
    
    // Calculate process metrics
    const results = [];
    Object.keys(processGroups).forEach(processName => {
      const tasks = processGroups[processName];
      const totalTasks = tasks.length;
      const completedTasks = tasks.filter(task => task[3] && task[3] !== '').length;
      const onTimeTasks = tasks.filter(task => {
        if (!task[3] || task[3] === '') return false;
        const actualDate = new Date(task[3]);
        const plannedDate = new Date(task[2]);
        return actualDate <= plannedDate;
      }).length;
      
      const taskScore = totalTasks > 0 ? ((completedTasks / totalTasks) * 100) - 100 : -100;
      const timeScore = completedTasks > 0 ? ((onTimeTasks / completedTasks) * 100) - 100 : -100;
      
      results.push({
        processName: processName,
        totalTasks: totalTasks,
        totalCompleted: completedTasks,
        taskScore: Math.round(taskScore * 100) / 100,
        totalCompleteTasks: completedTasks,
        onTimeDoneTasks: onTimeTasks,
        timeScore: Math.round(timeScore * 100) / 100
      });
    });
    
    return { success: true, data: results };
  } catch (error) {
    console.error("Error getting process details:", error);
    return { error: "Failed to get process details: " + error.toString() };
  }
}

/**
 * Save work commitments
 */
function saveWorkCommitments(userEmail, weekNumber, startDate, endDate, wndCommitment, wndotCommitment) {
  try {
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) {
      return { error: "External sheet ID not found" };
    }
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    let commitmentSheet = externalSS.getSheetByName("workCommitments");
    
    // Create sheet if it doesn't exist
    if (!commitmentSheet) {
      commitmentSheet = externalSS.insertSheet("workCommitments");
      commitmentSheet.getRange(1, 1, 1, 6).setValues([
        ["UserId", "Week Number", "Start Date", "End Date", "Next Commitment (Work not done)", "Next Commitment (Work Done on Time)"]
      ]);
    }
    
    const data = commitmentSheet.getDataRange().getValues();
    
    // Find existing row for this user and week
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail && data[i][1] == weekNumber) {
        rowIndex = i + 1; // Convert to 1-based index
        break;
      }
    }
    
    const newRow = [userEmail, parseInt(weekNumber), startDate, endDate, wndCommitment, wndotCommitment];
    
    if (rowIndex > 0) {
      // Update existing row
      commitmentSheet.getRange(rowIndex, 1, 1, 6).setValues([newRow]);
    } else {
      // Add new row
      commitmentSheet.appendRow(newRow);
    }
    
    return { success: true, message: "Commitments saved successfully" };
  } catch (error) {
    console.error("Error saving work commitments:", error);
    return { error: "Failed to save commitments: " + error.toString() };
  }
}

/**
 * Get all available weeks for dropdown
 */
function getAvailableWeeks() {
  try {
    const weekMapData = getWeekMappingData();
    
    if (weekMapData.length > 0) {
      // Use custom week mapping
      return weekMapData.map(row => ({
        weekNumber: row[2],
        startDate: formatDateForInput(new Date(row[0])),
        endDate: formatDateForInput(new Date(row[1])),
        label: `Week ${row[2]} (${formatDateForDisplay(new Date(row[0]))} - ${formatDateForDisplay(new Date(row[1]))})`
      }));
    } else {
      // Generate default weeks for current year
      const currentYear = new Date().getFullYear();
      const weeks = [];
      
      // Start from the first Monday of the year
      const jan1 = new Date(currentYear, 0, 1);
      const jan1DayOfWeek = jan1.getDay();
      const firstMonday = new Date(currentYear, 0, 1 + (jan1DayOfWeek === 0 ? 1 : 8 - jan1DayOfWeek));
      
      for (let week = 1; week <= 52; week++) {
        const startDate = new Date(firstMonday);
        startDate.setDate(firstMonday.getDate() + (week - 1) * 7);
        
        const endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 6);
        
        weeks.push({
          weekNumber: week,
          startDate: formatDateForInput(startDate),
          endDate: formatDateForInput(endDate),
          label: `Week ${week} (${formatDateForDisplay(startDate)} - ${formatDateForDisplay(endDate)})`
        });
      }
      
      return weeks;
    }
  } catch (error) {
    console.error("Error getting available weeks:", error);
    // Return at least current week as fallback
    const today = new Date();
    const monday = new Date(today);
    const dayOfWeek = today.getDay();
    const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
    monday.setDate(today.getDate() + mondayOffset);
    
    const sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);
    
    return [{
      weekNumber: 1,
      startDate: formatDateForInput(monday),
      endDate: formatDateForInput(sunday),
      label: `Week 1 (${formatDateForDisplay(monday)} - ${formatDateForDisplay(sunday)})`
    }];
  }
}

/**
 * Helper function to format date for input fields (YYYY-MM-DD)
 */
function formatDateForInput(date) {
  // Use local timezone instead of UTC to avoid date shifts
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Helper function to format date for display (DD/MM/YYYY)
 */
function formatDateForDisplay(date) {
  // Use consistent local date formatting
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  return `${day}/${month}/${year}`;
}

/**
 * Get MIS Score page HTML
 */
function getMISScorePageHtml(userName, role) {
  const template = HtmlService.createTemplateFromFile('MISScorePage');
  template.userName = userName;
  template.role = role;
  
  return template.evaluate().getContent();
}

// ============================================
// USER MANAGEMENT FUNCTIONS
// ============================================

/**
 * Get all users from the User sheet
 * Returns user data excluding sensitive password info for display
 */
// ... (previous functions remain unchanged)

/**
 * Get all users from the User sheet
 * Returns user data excluding sensitive password info for display
 */
function getUsersData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("User");
    
    if (!userSheet) {
      return { error: "User sheet not found" };
    }
    
    const data = userSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: true, users: [] };
    }
    
    // Skip header row, map to user objects
    // Columns: User ID, Password, Display Name, Role, Manager, Number, Metrics Visibility, Profile Picture
    const users = data.slice(1).map(row => ({
      userId: row[0] || '',
      password: row[1] || '',
      displayName: row[2] || '',
      role: row[3] || '',
      manager: row[4] || '',
      mobileNumber: row[5] || '',
      metricsVisible: row[6] || '',
      profilePicture: row[7] || ''
    }));
    
    return { success: true, users: users };
  } catch (error) {
    Logger.log("Error in getUsersData: " + error);
    return { error: "Failed to fetch users: " + error.toString() };
  }
}

/**
 * Get all menus and their buttons from Master Sheet and Create Button
 * This will be used to build the permission checkboxes
 */
function getAllMenusAndButtons() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("Master Sheet");
    const createButtonSheet = ss.getSheetByName("Create Button");
    
    if (!masterSheet) {
      return { error: "Master Sheet not found" };
    }
    
    // Read Master Sheet to get all submenus
    const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 3).getValues();
    
    // Group by mainMenu -> subMenus
    const menuStructure = {};
    masterData.forEach(row => {
      const mainMenu = String(row[0]).trim();
      const subMenu = String(row[1]).trim();
      
      if (!mainMenu || !subMenu) return;
      
      if (!menuStructure[mainMenu]) {
        menuStructure[mainMenu] = new Set();
      }
      menuStructure[mainMenu].add(subMenu);
    });
    
    // Convert Sets to Arrays for JSON serialization
    const menus = {};
    for (let mainMenu in menuStructure) {
      menus[mainMenu] = Array.from(menuStructure[mainMenu]);
    }
    
    // Read Create Button sheet to get all buttons per submenu
    const buttons = {};
    if (createButtonSheet) {
      const lastRowBtn = createButtonSheet.getLastRow();
      if (lastRowBtn >= 2) {
        const buttonData = createButtonSheet.getRange(2, 1, lastRowBtn - 1, 4).getValues();
        
        buttonData.forEach(row => {
          const bMenu = String(row[0]).trim();
          const bSub = String(row[1]).trim();
          const bName = String(row[2]).trim();
          
          if (!bMenu || !bSub || !bName) return;
          
          const key = bMenu + "||" + bSub;
          if (!buttons[key]) {
            buttons[key] = [];
          }
          buttons[key].push(bName);
        });
      }
    }
    
    return { success: true, menus: menus, buttons: buttons };
  } catch (error) {
    Logger.log("Error in getAllMenusAndButtons: " + error);
    return { error: "Failed to fetch menus and buttons: " + error.toString() };
  }
}

/**
 * Create a new user and set up their permissions
 * @param {Object} userData - User information
 * @param {Array} permissions - Array of permission objects {mainMenu, subMenu, buttons: []}
 */
/**
 * Create a new user and set up their permissions
 * @param {Object} userData - User information
 * @param {Array} permissions - Array of permission objects {mainMenu, subMenu, buttons: []}
 */
function createNewUser(userData, permissions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("User");
    const permissionSheet = ss.getSheetByName("Permission");
    
    if (!userSheet || !permissionSheet) {
      return { error: "Required sheets not found" };
    }
    
    // Check if user already exists
    const existingUsers = userSheet.getDataRange().getValues();
    for (let i = 1; i < existingUsers.length; i++) {
      if (existingUsers[i][0] === userData.userId) {
        return { error: "User with this email already exists" };
      }
    }

    // Handle File Upload to Drive if present
    let profilePictureUrl = '';
    if (userData.fileData && userData.fileData.content) {
      try {
        const contentType = userData.fileData.mimeType || 'application/octet-stream';
        const blob = Utilities.newBlob(
          Utilities.base64Decode(userData.fileData.content), 
          contentType, 
          userData.fileData.name || 'ProfilePicture'
        );
        
        // Create file in specific folder for profile pictures
        const PROFILE_PICTURES_FOLDER_ID = '1ZZ4HRkT0VZY6x-YXjlim1-fT4WPRAMsS';
        const folder = DriveApp.getFolderById(PROFILE_PICTURES_FOLDER_ID);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // We use getUrl() which returns the preview link. 
        // The frontend 'getThumbnailImageUrl' function extracts the ID from this to convert to a thumbnail.
        profilePictureUrl = file.getUrl(); 
      } catch (fileError) {
        console.error('Error uploading file:', fileError);
        return { error: "Failed to upload profile picture: " + fileError.toString() };
      }
    }
    
    // Add user to User sheet
    // Columns: User ID, Password, Display Name, Role, Manager, Number, Metrics Visibility, Profile Picture
    userSheet.appendRow([
      userData.userId,
      userData.password,
      userData.displayName,
      userData.role ? String(userData.role).toLowerCase() : 'user',
      userData.manager || '',
      userData.mobileNumber || '',
      userData.metricsVisible || 'No', 
      profilePictureUrl || ''
    ]);
    
    // Add permissions to Permission sheet
    permissions.forEach(perm => {
      const buttonNames = perm.buttons && perm.buttons.length > 0 
        ? perm.buttons.join(', ') 
        : '';
      
      permissionSheet.appendRow([
        userData.userId,
        perm.mainMenu,
        perm.subMenu,
        buttonNames,
        perm.accessType || ''
      ]);
    });
    
    return { success: true, message: "User created successfully" };
  } catch (error) {
    Logger.log("Error in createNewUser: " + error);
    return { error: "Failed to create user: " + error.toString() };
  }
}

/**
 * Update an existing user
 */
function updateUser(userData, permissions) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("User");
    const permissionSheet = ss.getSheetByName("Permission");
    
    if (!userSheet || !permissionSheet) {
      return { error: "Required sheets not found" };
    }
    
    const data = userSheet.getDataRange().getValues();
    let rowIndex = -1;
    let existingRow = [];
    
    // Find user row (skip header)
    for (let i = 1; i < data.length; i++) {
        // Compare User ID (Email)
      if (String(data[i][0]).toLowerCase() === String(userData.userId).toLowerCase()) {
        rowIndex = i + 1; // 1-based index
        existingRow = data[i];
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { error: "User not found" };
    }

    // Handle File Upload: 
    // If new file content is provided, upload and get new URL.
    // If NOT provided, keep the existing URL (index 7).
    let profilePictureUrl = existingRow[7]; 
    
    if (userData.fileData && userData.fileData.content) {
      try {
        const contentType = userData.fileData.mimeType || 'application/octet-stream';
        const blob = Utilities.newBlob(
          Utilities.base64Decode(userData.fileData.content), 
          contentType, 
          userData.fileData.name || 'ProfilePicture'
        );
        // Create file in specific folder for profile pictures
        const PROFILE_PICTURES_FOLDER_ID = '1ZZ4HRkT0VZY6x-YXjlim1-fT4WPRAMsS';
        const folder = DriveApp.getFolderById(PROFILE_PICTURES_FOLDER_ID);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        profilePictureUrl = file.getUrl();
      } catch (fileError) {
        console.error('Error uploading file:', fileError);
        return { error: "Failed to upload new profile picture: " + fileError.toString() };
      }
    }

    // Update User Sheet Row
    // Columns: User ID, Password, Display Name, Role, Manager, Number, Metrics Visibility, Profile Picture
    // We update all fields except User ID (assuming Email is immutable key for now)
    const updatedRow = [
      userData.userId, 
      userData.password || existingRow[1], // Update password if sent, else keep existing? usually we should probably just update it if it's there. 
      userData.displayName,
      userData.role ? String(userData.role).toLowerCase() : existingRow[3],
      userData.manager,
      userData.mobileNumber,
      userData.metricsVisible,
      profilePictureUrl
    ];
    
    // Set values for the row (startRow, startCol, numRows, numCols)
    userSheet.getRange(rowIndex, 1, 1, 8).setValues([updatedRow]);
    
    // Update Permissions: Sync instead of Delete/Add (unless deactivating)
    if (userData.role && userData.role.toLowerCase() !== 'deactivate') {
      const permData = permissionSheet.getDataRange().getValues();
      const userEmail = String(userData.userId).toLowerCase();
      
      // Find row indices for existing permissions of this user
      const existingPermIndices = [];
      for (let i = 1; i < permData.length; i++) {
        if (String(permData[i][0]).toLowerCase() === userEmail) {
          existingPermIndices.push(i + 1); // 1-based index
        }
      }
      
      // Synchronize permissions
      const newPerms = permissions || [];
      const countExisting = existingPermIndices.length;
      const countNew = newPerms.length;
      
      // 1. Update existing rows as much as possible
      const updateCount = Math.min(countExisting, countNew);
      for (let i = 0; i < updateCount; i++) {
        const rowInd = existingPermIndices[i];
        const perm = newPerms[i];
        const buttonNames = perm.buttons && perm.buttons.length > 0 ? perm.buttons.join(', ') : '';
        const rowData = [userData.userId, perm.mainMenu, perm.subMenu, buttonNames, perm.accessType || ''];
        permissionSheet.getRange(rowInd, 1, 1, 5).setValues([rowData]);
      }
      
      // 2. Handle mismatch (Delete extras or Add missing)
      if (countExisting > countNew) {
        // Delete extra rows from BOTTOM to TOP to keep indices valid
        for (let i = countExisting - 1; i >= updateCount; i--) {
          permissionSheet.deleteRow(existingPermIndices[i]);
        }
      } else if (countNew > countExisting) {
        // Append missing rows
        for (let i = updateCount; i < countNew; i++) {
          const perm = newPerms[i];
          const buttonNames = perm.buttons && perm.buttons.length > 0 ? perm.buttons.join(', ') : '';
          const rowData = [userData.userId, perm.mainMenu, perm.subMenu, buttonNames, perm.accessType || ''];
          permissionSheet.appendRow(rowData);
        }
      }
    } else {
      // If deactivating, delete all permissions (as per current behavior)
      const permData = permissionSheet.getDataRange().getValues();
      for (let i = permData.length - 1; i >= 1; i--) {
        if (String(permData[i][0]).toLowerCase() === String(userData.userId).toLowerCase()) {
          permissionSheet.deleteRow(i + 1);
        }
      }
    }
    
    return { success: true, message: "User updated successfully" };
    
  } catch (error) {
    Logger.log("Error in updateUser: " + error);
    return { error: "Failed to update user: " + error.toString() };
  }
}

/**
 * Get permissions for a specific user
 */
function getUserPermissions(userId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const permissionSheet = ss.getSheetByName("Permission");
    
    if (!permissionSheet) return { success: true, permissions: [] };
    
    const data = permissionSheet.getDataRange().getValues();
    const permissions = [];
    
    // Skip header
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === String(userId).toLowerCase()) {
        // Row: User, Main Menu, Sub Menu, Button Names, Access Type
        permissions.push({
          mainMenu: data[i][1],
          subMenu: data[i][2],
          buttons: data[i][3] ? String(data[i][3]).split(',').map(s => s.trim()) : [],
          accessType: data[i][4]
        });
      }
    }
    
    return { success: true, permissions: permissions };
  } catch (error) {
    return { error: "Failed to fetch user permissions: " + error.toString() };
  }
}

/**
 * Delete a user and all their permissions
 * @param {string} userId - The email/user ID to delete
 */
function deleteUser(userId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName("User");
    const permissionSheet = ss.getSheetByName("Permission");
    
    if (!userSheet) {
      return { error: "User sheet not found" };
    }
    
    // 1. Find and delete user from User sheet
    const userData = userSheet.getDataRange().getValues();
    let userRowIndex = -1;
    
    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][0]).toLowerCase() === String(userId).toLowerCase()) {
        userRowIndex = i + 1; // 1-based index
        break;
      }
    }
    
    if (userRowIndex === -1) {
      return { error: "User not found" };
    }
    
    // Delete user row
    userSheet.deleteRow(userRowIndex);
    
    // 2. Delete all permissions for this user
    if (permissionSheet) {
      const permData = permissionSheet.getDataRange().getValues();
      // Iterate backwards to delete rows safely
      for (let i = permData.length - 1; i >= 1; i--) {
        if (String(permData[i][0]).toLowerCase() === String(userId).toLowerCase()) {
          permissionSheet.deleteRow(i + 1);
        }
      }
    }
    
    return { success: true, message: "User deleted successfully" };
  } catch (error) {
    Logger.log("Error in deleteUser: " + error);
    return { error: "Failed to delete user: " + error.toString() };
  }
}
 
// ... (rest of the file)

/**
 * Get User Management page HTML
 */
function getUserManagementPageHtml(userName, role) {
  const template = HtmlService.createTemplateFromFile('UserManagement');
  template.userName = userName;
  template.role = role;
  
  return template.evaluate().getContent();
}

// Add after the existing MIS functions

/**
 * Get delegation data from external sheet
 */
function getDelegationData(userName, role) {
  try {
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) {
      return { error: "External sheet ID not found in SetUP sheet cell B3" };
    }
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    const dataSheet = externalSS.getSheetByName("deligation_data");
    
    if (!dataSheet) {
      return { error: "Delegation data sheet not found in external spreadsheet" };
    }
    
    const data = dataSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    // Filter data based on user role
    let filteredData = data.slice(1); // Remove header
    
    if (role !== 'admin') {
      filteredData = filteredData.filter(row => 
        String(row[1]).toLowerCase() === String(userName).toLowerCase() // Column B is Email
      );
    }
    
    return { success: true, data: filteredData, headers: data[0] };
  } catch (error) {
    console.error("Error getting delegation data:", error);
    return { error: "Failed to retrieve delegation data: " + error.toString() };
  }
}

/**
 * Get delegation work commitments data
 */
function getDelegationWorkCommitments() {
  try {
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) return [];
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    const commitmentSheet = externalSS.getSheetByName("deligation_workCommitments");
    
    if (!commitmentSheet) {
      return []; // Return empty array if sheet doesn't exist
    }
    
    const data = commitmentSheet.getDataRange().getValues();
    return data.length > 1 ? data.slice(1) : []; // Remove header if exists
  } catch (error) {
    console.error("Error getting delegation work commitments:", error);
    return [];
  }
}

/**
 * Calculate delegation scores for dashboard
 */
function calculateDelegationScores(userName, role, startDate, endDate) {
  try {
    const delegationDataResult = getDelegationData(userName, role);
    if (delegationDataResult.error) {
      return delegationDataResult;
    }
    
    const delegationData = delegationDataResult.data;
    const workCommitments = getDelegationWorkCommitments();
    
    // Parse dates
    const filterStartDate = new Date(startDate);
    filterStartDate.setHours(0, 0, 0, 0);
    
    const filterEndDate = new Date(endDate);
    filterEndDate.setHours(23, 59, 59, 999);
    
    // Filter data by date range (Due Date column - index 4)
    const filteredData = delegationData.filter(row => {
      if (!row[4]) return false; // Skip empty due dates
      const dueDate = new Date(row[4]);
      dueDate.setHours(0, 0, 0, 0);
      return dueDate >= filterStartDate && dueDate <= filterEndDate;
    });
    
    // Group by user
    const userGroups = {};
    filteredData.forEach(row => {
      const email = row[1]; // Email column
      const assignTo = row[2]; // Assign to column
      
      if (!userGroups[email]) {
        userGroups[email] = {
          email: email,
          name: assignTo,
          tasks: []
        };
      }
      userGroups[email].tasks.push(row);
    });
    
    // Calculate scores for each user
    const results = [];
    const selectedWeekNumber = getWeekNumber(filterStartDate);
    const previousWeekNumber = selectedWeekNumber - 1;
    
    Object.keys(userGroups).forEach(email => {
      const userGroup = userGroups[email];
      const tasks = userGroup.tasks;
      
      // Calculate task counts
      const totalTasks = tasks.length;
      const completedTasks = tasks.filter(task => task[9] === 'Complete').length; // Status column
      const pendingTasks = tasks.filter(task => task[9] === 'Pending').length;
      const shiftedTasks = tasks.filter(task => task[9] === 'Shifted').length;
      
      // Calculate scores for completed and shifted tasks only
      const completedAndShiftedTasks = tasks.filter(task => 
        task[9] === 'Complete' || task[9] === 'Shifted'
      );
      const completedAndShiftedCount = completedAndShiftedTasks.length;
      
      let redScore = 0, yellowScore = 0, greenScore = 0;
      let redCount = 0, yellowCount = 0, greenCount = 0;  // Declare outside the if block

      if (completedAndShiftedCount > 0) {
        // Red: both revision 1 and 2 are non-empty
        redCount = completedAndShiftedTasks.filter(task => 
          task[5] && task[5] !== '' && task[6] && task[6] !== ''
        ).length;
        
        // Yellow: revision 1 is non-empty but revision 2 is empty
        yellowCount = completedAndShiftedTasks.filter(task => 
          task[5] && task[5] !== '' && (!task[6] || task[6] === '')
        ).length;
        
        // Green: both revision 1 and 2 are empty
        greenCount = completedAndShiftedTasks.filter(task => 
          (!task[5] || task[5] === '') && (!task[6] || task[6] === '')
        ).length;
        
        redScore = Math.round((redCount / completedAndShiftedCount) * 100 * 100) / 100;
        yellowScore = Math.round((yellowCount / completedAndShiftedCount) * 100 * 100) / 100;
        greenScore = Math.round((greenCount / completedAndShiftedCount) * 100 * 100) / 100;
      }
      
      // Get previous week commitments (for "Current Planned")
      const previousCommitments = workCommitments.find(commitment => 
        commitment[0] === email && commitment[1] === previousWeekNumber
      );
      
      const currentPlannedRed = previousCommitments ? (previousCommitments[4] || 0) : 0;
      const currentPlannedYellow = previousCommitments ? (previousCommitments[5] || 0) : 0;
      const currentPlannedGreen = previousCommitments ? (previousCommitments[6] || 0) : 0;
      
      // Get current week commitments (for "Next Committed")
      const currentCommitments = workCommitments.find(commitment => 
        commitment[0] === email && commitment[1] === selectedWeekNumber
      );
      
      results.push({
        userId: email,
        name: userGroup.name,
        currentPlannedRed: currentPlannedRed,
        currentPlannedYellow: currentPlannedYellow,
        currentPlannedGreen: currentPlannedGreen,
        currentScoreRed: redScore,
        currentScoreYellow: yellowScore,
        currentScoreGreen: greenScore,
        currentScoreRedCount: redCount,
        currentScoreYellowCount: yellowCount,
        currentScoreGreenCount: greenCount,
        totalTasks: totalTasks,
        completedTasks: completedTasks,
        pendingTasks: pendingTasks,
        shiftedTasks: shiftedTasks,
        nextCommittedRed: currentCommitments ? (currentCommitments[4] || '') : '',
        nextCommittedYellow: currentCommitments ? (currentCommitments[5] || '') : '',
        nextCommittedGreen: currentCommitments ? (currentCommitments[6] || '') : '',
        weekNumber: selectedWeekNumber
      });
    });
    
    return { success: true, data: results };
  } catch (error) {
    console.error("Error calculating delegation scores:", error);
    return { error: "Failed to calculate delegation scores: " + error.toString() };
  }
}

/**
 * Save delegation work commitments
 */
function saveDelegationWorkCommitments(userEmail, weekNumber, startDate, endDate, redCommitment, yellowCommitment, greenCommitment) {
  try {
    console.log('saveDelegationWorkCommitments called with:', {
      userEmail, weekNumber, startDate, endDate, redCommitment, yellowCommitment, greenCommitment
    });
    
    const externalSheetId = getExternalSheetId();
    if (!externalSheetId) {
      console.error('External sheet ID not found');
      return { error: "External sheet ID not found" };
    }
    
    console.log('External sheet ID:', externalSheetId);
    
    const externalSS = SpreadsheetApp.openById(externalSheetId);
    let commitmentSheet = externalSS.getSheetByName("deligation_workCommitments");
    
    // Create sheet if it doesn't exist
    if (!commitmentSheet) {
      console.log('Creating deligation_workCommitments sheet');
      commitmentSheet = externalSS.insertSheet("deligation_workCommitments");
      commitmentSheet.getRange(1, 1, 1, 7).setValues([
        ["UserId", "Week Number", "Start Date", "End Date", "Red Score", "Yellow Score", "Green Score"]
      ]);
    }
    
    const data = commitmentSheet.getDataRange().getValues();
    console.log('Existing data rows:', data.length);
    
    // Find existing row for this user and week
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === userEmail && data[i][1] == weekNumber) {
        rowIndex = i + 1; // Convert to 1-based index
        console.log('Found existing row at index:', rowIndex);
        break;
      }
    }
    
    const newRow = [userEmail, parseInt(weekNumber), startDate, endDate, redCommitment, yellowCommitment, greenCommitment];
    console.log('New row data:', newRow);
    
    if (rowIndex > 0) {
      // Update existing row
      console.log('Updating existing row');
      commitmentSheet.getRange(rowIndex, 1, 1, 7).setValues([newRow]);
    } else {
      // Add new row
      console.log('Adding new row');
      commitmentSheet.appendRow(newRow);
    }
    
    console.log('Save completed successfully');
    return { success: true, message: "Delegation commitments saved successfully" };
  } catch (error) {
    console.error("Error saving delegation work commitments:", error);
    return { error: "Failed to save delegation commitments: " + error.toString() };
  }
}
/**
 * Master Sheet Management Functions (Create Button)
 */

function getMasterSheetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master Sheet");
    if (!sheet) return { error: "Master Sheet not found" };
    
    if (sheet.getLastRow() < 2) return { success: true, data: [] };
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    return { success: true, data: data };
  } catch (e) {
    return { error: e.toString() };
  }
}

function saveMasterSheetEntry(entryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master Sheet");
    if (!sheet) return { error: "Master Sheet not found" };

    const split = parseEntryViewable(entryData.entryViewable);
    
    const newRow = [
      entryData.mainMenu,
      entryData.subMenu,
      entryData.sheetName,
      entryData.dbLink,
      entryData.externalId,
      entryData.rangeSpec,
      split.question,
      split.table
    ];
    
    sheet.appendRow(newRow);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

function updateMasterSheetEntry(entryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master Sheet");
    if (!sheet) return { error: "Master Sheet not found" };

    const rowIndex = entryData.index + 2; // +1 for 0-indexing, +1 for header
    const split = parseEntryViewable(entryData.entryViewable);
    
    const updatedRow = [
      entryData.mainMenu,
      entryData.subMenu,
      entryData.sheetName,
      entryData.dbLink,
      entryData.externalId,
      entryData.rangeSpec,
      split.question,
      split.table
    ];
    
    sheet.getRange(rowIndex, 1, 1, 8).setValues([updatedRow]);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

function deleteMasterSheetEntry(index) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master Sheet");
    if (!sheet) return { error: "Master Sheet not found" };

    const rowIndex = index + 2;
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Helper to parse (a,b)(c,d) into {question: "a,b", table: "c,d"}
 */
function parseEntryViewable(str) {
  const result = { question: "", table: "" };
  if (!str) return result;
  
  // Look for parenthesis groups
  const matches = str.match(/\(([^)]+)\)/g);
  if (matches) {
    if (matches[0]) result.question = matches[0].replace(/[()]/g, '');
    if (matches[1]) result.table = matches[1].replace(/[()]/g, '');
  } else {
    // Fallback if no parenthesis, just put everything in question
    result.question = str;
  }
  return result;
}
/**
 * Create Button Sheet Management Functions
 */

function getCreateButtonData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Create Button");
    if (!sheet) return { error: "Create Button sheet not found" };
    
    if (sheet.getLastRow() < 2) return { success: true, data: [] };
    
    // Fetching 4 columns as seen in getAllMenusAndButtons
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    return { success: true, data: data };
  } catch (e) {
    return { error: e.toString() };
  }
}

function saveCreateButtonEntry(entryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Create Button");
    if (!sheet) return { error: "Create Button sheet not found" };

    const newRow = [
      entryData.mainMenu,
      entryData.subMenu,
      entryData.buttonName,
      entryData.extra || "" // Column 4 placeholder
    ];
    
    sheet.appendRow(newRow);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

function updateCreateButtonEntry(entryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Create Button");
    if (!sheet) return { error: "Create Button sheet not found" };

    const rowIndex = entryData.index + 2; 
    const updatedRow = [
      entryData.mainMenu,
      entryData.subMenu,
      entryData.buttonName,
      entryData.extra || ""
    ];
    
    sheet.getRange(rowIndex, 1, 1, 4).setValues([updatedRow]);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

function deleteCreateButtonEntry(index) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Create Button");
    if (!sheet) return { error: "Create Button sheet not found" };

    const rowIndex = index + 2;
    sheet.deleteRow(rowIndex);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Bulk Save Master Sheet Entries
 */
function saveBulkMasterSheetEntries(entries) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Master Sheet");
    if (!sheet) return { error: "Master Sheet not found" };

    const rowsToAdd = entries.map(entry => {
      const split = parseEntryViewable(entry.entryViewable);
      return [
        entry.mainMenu,
        entry.subMenu,
        entry.sheetName,
        entry.dbLink,
        entry.externalId,
        entry.rangeSpec,
        split.question,
        split.table
      ];
    });

    if (rowsToAdd.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 8).setValues(rowsToAdd);
    }
    
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Bulk Save Create Button Entries
 */
function saveBulkCreateButtonEntries(entries) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Create Button");
    if (!sheet) return { error: "Create Button sheet not found" };

    const rowsToAdd = entries.map(entry => [
      entry.mainMenu,
      entry.subMenu,
      entry.buttonName,
      entry.extra || ""
    ]);

    if (rowsToAdd.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 4).setValues(rowsToAdd);
    }
    
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}
