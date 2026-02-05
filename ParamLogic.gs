function doGetParam(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Instead of hardcoding "Data", decide which sheet to use.
  //    We can take it from a URL parameter: e.parameter.sheetName
  var sheetName = e.parameter.sheetName || "Sheet1"; 
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return ContentService
      .createTextOutput("Sheet '"+sheetName+"' not found.")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  // 2. If you wish, also fetch a "Base" sheet for the Title/Logo
  //    Or fallback to some defaults if not found
  var baseSheet = ss.getSheetByName("Base");
  var pageTitle = "Fix Title"; // fallback
  if (baseSheet) {
    pageTitle = baseSheet.getRange("B2").getValue();
  }

  // 3. Get the time zone
  var timeZone = ss.getSpreadsheetTimeZone();

  // 4. Read the data range. (We expect row1=colIDs, row2=filterTypes, row3=Headers, row4+ = data)
  var range  = sheet.getDataRange();
  var values = range.getValues();
  if (values.length < 3) {
    return ContentService
      .createTextOutput("Not enough rows in '"+sheetName+"'.")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  var columnIdentifiers = values[0]; // Row1
  var filterTypes       = values[1]; // Row2
  var headers           = values[2]; // Row3
  var data              = values.slice(3);

  // 5. Apply server-side filtering (like your old applyUrlFilters).
  var filteredData = applyUrlFilters(e.parameter, data, columnIdentifiers, filterTypes);

  var userEmail = e.parameter.userName; // or e.parameter.userEmail if you prefer that key
  var role = e.parameter.role || 'user';
  console.log(userEmail);
  console.log(role);
  if (userEmail && role.toLowerCase() !== 'admin') {
    filteredData = filteredData.filter(function(row) {
      // Assumes the first column is the Email column.
      return String(row[0]).toLowerCase() === String(userEmail).toLowerCase();
    });
  }

  // 6. Build the HTML template from the ParamIndex.html file
  var template = HtmlService.createTemplateFromFile('ParamIndex');
  template.columnIdentifiers = columnIdentifiers;
  template.filterTypes       = filterTypes;
  template.headers           = headers;
  template.filteredData      = filteredData;
  template.pageTitle         = pageTitle;
  template.timeZone          = timeZone;

  return template.evaluate()
    .setTitle("Dynamic Filter Demo")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Copy/paste the same applyUrlFilters() from your Param project here,
 * unchanged, or rename it if you like.
 */
function applyUrlFilters(params, data, colIds, filterTypes) {
  var colFilters = [];
  for (var i = 0; i < colIds.length; i++) {
    colFilters.push({
      filterType: filterTypes[i],
      searchValue: null,
      searchFilterValue: null,
      dateStart: null,
      dateEnd: null
    });
  }

  var datePattern = /^\d{4}-\d{2}-\d{2}$/; // basic YYYY-MM-DD

  // Read each URL param
  for (var key in params) {
    var val = params[key];
    var colIndex = colIds.indexOf(key);

    if (colIndex >= 0) {
      var ft = filterTypes[colIndex];
      
      if (ft === 'Search') {
        colFilters[colIndex].searchValue = val;
      } else if (ft === 'SearchFilter') {
        colFilters[colIndex].searchFilterValue = val;
      } else if (ft === 'Hidden') {
        colFilters[colIndex].searchValue = val;
      } else if (ft === 'DateRange') {
        // handled below
      } else if (["Avg", "Sum", "Min", "Max", "Count"].indexOf(ft) !== -1) {
        // Aggregate columns do not apply URL filtering.
      } else if (!ft || ft.trim() === '') {
        // fallback partial substring
        colFilters[colIndex].searchValue = val;
      }
      
    } else {
      // Possibly "ColXStart" or "ColXEnd"
      var matchesStart = key.match(/^(.*)Start$/);
      var matchesEnd   = key.match(/^(.*)End$/);

      if (matchesStart) {
        var baseKey = matchesStart[1];
        var idx2 = colIds.indexOf(baseKey);
        if (idx2 >= 0 && filterTypes[idx2] === 'DateRange' && datePattern.test(val)) {
          colFilters[idx2].dateStart = new Date(val);
        }
      } else if (matchesEnd) {
        var baseKey2 = matchesEnd[1];
        var idx3 = colIds.indexOf(baseKey2);
        if (idx3 >= 0 && filterTypes[idx3] === 'DateRange' && datePattern.test(val)) {
          colFilters[idx3].dateEnd = new Date(val);
        }
      }
    }
  }

  // Filter the data
  var filtered = [];
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var keep = true;
    
    for (var c = 0; c < row.length; c++) {
      var fObj = colFilters[c];
      var cellVal = row[c];
      var ft = filterTypes[c];

      // 1) Partial substring match
      if (fObj.searchValue) {
        var cellStr = String(cellVal || '').toLowerCase();
        if (cellStr.indexOf(fObj.searchValue.toLowerCase()) === -1) {
          keep = false;
          break;
        }
      }
      // 2) Exact match
      if (fObj.searchFilterValue) {
        var cellStr2 = String(cellVal || '').toLowerCase();
        if (cellStr2 !== fObj.searchFilterValue.toLowerCase()) {
          keep = false;
          break;
        }
      }
      // 3) Date range
      if (ft === 'DateRange') {
        if (fObj.dateStart || fObj.dateEnd) {
          var cellDate = new Date(cellVal);
          if (fObj.dateStart && cellDate < fObj.dateStart) {
            keep = false;
            break;
          }
          if (fObj.dateEnd && cellDate > fObj.dateEnd) {
            keep = false;
            break;
          }
        }
      }
    }

    if (keep) {
      filtered.push(row);
    }
  }

  return filtered;
}