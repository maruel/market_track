// Copyright 2016 Marc-Antoine Ruel. All rights reserved.
// Use of this source code is governed under the Apache License, Version 2.0
// that can be found in the LICENSE file.
//
//
// Please see https://github.com/maruel/market_track/ for up to date information!
//
//
// On Google Sheets templates, it is impossible to "sync" a spreadsheet an updated template
// version. Yet, this project is constantly being improved. This is why it is important
// to  replace the content of this file with the updated version at
// https://github.com/maruel/market_track/blob/master/code.gs
//
//
// References:
// https://support.google.com/docs/answer/3093281
// https://developers.google.com/apps-script/reference/base/logger
// https://developers.google.com/apps-script/reference/spreadsheet/range
// https://developers.google.com/apps-script/reference/spreadsheet/sheet
// https://developers.google.com/apps-script/reference/spreadsheet/embedded-chart-builder
//
//
// TODO(maruel):
// - Figure out a way to fill dividend via a third party service.
// - Use the currency when formatting the lines based on B1, e.g. £ for GBP, € for EUR.
// - Have this script warn the user when a new version of this script is live.
//   This needs to include an easy way to turn the check off and explanation
//   about using multiple files for user scripts.
// - Do not get all values in createChart() but only the ranges needed.
// - Update shares # in updateOneSheet().


// Hooks

function onOpen(e) {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu('Stocks');
  menu.addItem('Update all sheets', 'updateAllSheets').addToUi();
  menu.addItem('Update current sheet', 'updateCurrentSheet').addToUi();
  menu.addItem('Add new sheet', 'addNewSheet').addToUi();
  updateAllSheets();
}

// Stock sheets management.

// Returns true if it is a sheet that contains stocks.
function isValidStockSheet(sheet) {
  return (sheet.getRange("A1").getValue() == "Currency");
}

// Update all sheets, silently ignoring invalid ones.
function updateAllSheets() {
  for (var sheet in SpreadsheetApp.getActive().getSheets()) {
    updateOneSheet(sheet);
  }
}

// Update the current sheet, warn the user if it is invalid.
function updateCurrentSheet() {
  if (!updateOneSheet(SpreadsheetApp.getActiveSheet())) {
    var ui = SpreadsheetApp.getUi();
    ui.alert("Invalid sheet!");
  }
}

function updateOneSheet(sheet) {
  if (!isValidStockSheet(sheet)) {
    return false;
  }

  // Get the last line in column "A" to retrieve the date, then add new lines until yesterday.
  var ticker = sheet.getName();
  var lastRow = sheet.getLastRow();
  var startDate = nextDay(toStr(sheet.getRange(lastRow, 1).getValue()));
  Logger.log("", startDate, lastRow);
  var values = getDayValuesUpToYesterday(sheet, ticker, lastRow + 1, startDate, false);
  if (values == null) {
    return false;
  }
  sheet.getRange(lastRow + 1, 1, values.length, values[0].length).setValues(values);
  formatLines(sheet, lastRow+1, values.length, getCurrentFmt(ticker));
  lastRow += values.length;
  
  // Find the first row of data based on the content. This permits to safely add new header lines up to 40.
  var values = sheet.getRange(1, 1, 40, 1).getValues();
  var firstRow = 0;
  for (var y in values) {
    if (values[y][0] == "Date") {
      firstRow = y + 1;
      break
    }
  }
  return updateOneSheetInner(sheet, ticker, firstRow, lastRow);
}

function updateOneSheetInner(sheet, ticker, firstRow, lastRow) {
  // Trim lines.
  var delta = sheet.getMaxRows() - lastRow;
  if (delta > 10) {
    sheet.deleteRows(lastRow+1, delta - 10);
  }

  // Market cap is last share value * current price.
  if (!isCurrency(ticker)) {
    sheet.getRange("B3").setValue("=$B$2*$B$" + (lastRow-1));
  }
  // Min and max.
  sheet.getRange("B4:B5").setValues(
    [
      ["=MIN($E$" + firstRow + ":$E$" + lastRow + ")"],
      ["=MAX($E$" + firstRow + ":$E$" + lastRow + ")"],
  ])
  // Resize the columns with some margin. This is very slow, 1s per column.
  SpreadsheetApp.flush();
  for (var i = 1; i < 8; i++) {
    sheet.autoResizeColumn(i);
    sheet.setColumnWidth(i, sheet.getColumnWidth(i)+10);
  }
  createChart(sheet, ticker, firstRow, lastRow);
  return true;
}

function createChart(sheet, ticker, firstRow, lastRow) {
  //sheet.activate();
  // Recreate the chart everytime. This permits updating the chart style for older sheets.
  var charts = sheet.getCharts();
  for (var i in charts) {
    sheet.removeChart(charts[i]);
  }

  // It is very confusing that EmbeddedChartBuilder is very difference from Charts.newAreaChart().
  // Don't be misled by reading the wrong doc!
  // https://developers.google.com/chart/interactive/docs/reference#options
  // https://developers.google.com/chart/interactive/docs/gallery/linechart#configuration-options
  // http://icu-project.org/apiref/icu4c/classDecimalFormat.html#details
  var currency = sheet.getRange("B1").getValue();
  var width = sheet.getColumnWidth(8);
  var legend = {
  };
  var hAxis = {
    "gridlines": {
      "count": -1,
      "units": {
        "years": {"format": ["yy-MM"]},
        "months": {"format": ["yy-MM", "-MM", ""]},
      },
    },
    "minorGridlines": {
      "count": -1,
      "units": {
        "months": {"format": ["yy-MM", "-MM", ""]},
      },
    },
  };
  var range = sheet.getRange(firstRow, 1, lastRow - firstRow, 6);
  var values = range.getValues();
  if (isCurrency(ticker)) {
    var series = {};
  } else {
    // Cut the volume line at half of the graph.
    var maxVolume = 0;
    for (var i in values) {
      if (values[i][5] > maxVolume) {
        maxVolume = values[i][5];
      }
    }
    var series = {
      0: {"targetAxisIndex": 0},
      1: {"targetAxisIndex": 0},
      2: {"targetAxisIndex": 0},
      3: {"targetAxisIndex": 0},
      4: {"targetAxisIndex": 1},
    }
  }
  var vAxes = {
    0: {
      "format": "#,##0.00 [$$-C0C]",
      "title": currency,
    },
    1: {
      "format": "#,##0",
      "gridlines": {"count": 0},
      "title": "Volume",
      "maxValue": maxVolume * 2,
    },
  };
  var chart = sheet.newChart().addRange(range)
       .asLineChart()
       .setPosition(4, 8, 0, 0)
       .setOption("legend", legend)
       .setOption("hAxis", hAxis)
       .setOption("series", series)
       .setOption("vAxes", vAxes)
       .setOption("title", ticker + " " + toStr(values[1][0]) + " ~ " + toStr(values[values.length-1][0]))
       .setOption("width", width)
       .build();
  sheet.insertChart(chart);
}

// Creates a new sheet to track a new stock, ETF or exchange rate.
function addNewSheet() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Initializing new sheet', 'Please give the ticker symbol.', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() == ui.Button.CANCEL) {
    return false;
  }
  var ticker = response.getResponseText().toUpperCase();
  var sheet = SpreadsheetApp.getActive().getSheetByName(ticker);
  if (sheet != null) {
    ui.alert("Sheet already exist! Delete it first.");
    return false;
  }
  sheet = SpreadsheetApp.getActive().insertSheet(ticker);
  SpreadsheetApp.flush();

  // Bootstrap currency and # of shares.
  if (isCurrency(ticker)) {
    var currency = ticker.substr(9, 3);
    var cell1 = sheet.getRange("B1");
  } else {
    // Interleaving calls may help with latency (?)
    var cell1 = runGoogleFinanceInternal(sheet, "B1", [ticker, "currency"]);
    var cell2 = runGoogleFinanceInternal(sheet, "B2", [ticker, "shares"]);
    var currency = cell1.getValue();
    var shares = cell2.getValue();
    if (shares != "#N/A") {
      shares = shares / 1000000.;
     cell2.setNumberFormat("0.00 \"M\"");
    }
    cell2.setValue(shares);
    cell2.setHorizontalAlignment("right");
    cell2 = null;
  }
  if (currency == "#N/A") {
    ui.alert(ticker + " is not a valid ticker. Delete the sheet and try again.");
    return false;
  }
  cell1.setValue(currency);
  cell1.setHorizontalAlignment("right");
  cell1 = null;

  // Information header.
  var sections = [
    ["Currency"],
    ["Shares"],
    ["Market Cap"],
    ["Close low"],
    ["Close high"],
  ];

  // Get the starting date.
  while (true) {
    var response = ui.prompt('Initializing new sheet', 'When do you want to start? The format is YYYY, YYYY-MM or YYYY-MM-DD. If unspecified, it starts at Jan 1st of the current year', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() == ui.Button.CANCEL) {
      return false;
    }
    var startDate = response.getResponseText();
    if (startDate == "") {
      startDate = (new Date()).getFullYear() + "-01-01";
    }
    if (startDate.match(/^\d\d\d\d$/)) {
      startDate += "-01-01";
    }
    if (startDate.match(/^\d\d\d\d-\d\d$/)) {
      startDate += "-01";
    }
    Logger.log("Chose date", startDate);
    if (!startDate.match(/^\d\d\d\d-\d\d-\d\d$/)) {
      ui.alert("Invalid date.");
      continue;
    }
    if (!getDaysBetween(startDate, previousDay(getToday()))) {
      ui.alert("Date must be at least one day before yesterday.");
      continue;
    }
    // Initialize the data right away before the headers. This saves a few RPCs.
    var values = getDayValuesUpToYesterday(sheet, ticker, sections.length + 1, startDate, true);
    if (values == null) {
      ui.alert("Failed to find a date where this ticker is valid. Try earlier?");
      continue;
    }
    break;
  }

  // Sets the data.
  sheet.getRange(sections.length + 1, 1, values.length, values[0].length).setValues(values);
  var currencyFmt = getCurrentFmt(ticker);
  formatLines(sheet, sections.length + 2, values.length - 1, currencyFmt);

  // Sadly dividends are not supported by GOOGLEFINANCE() so one has to track it manually! :(
  sheet.getRange(sections.length + 1, values[0].length + 1).setValue("Dividend");
  if (isCurrency(ticker)) {
    sections[1][0] = "";
    sections[2][0] = "";
    sheet.getRange(sections.length + 1, 2, 1, 3).setValue("");
    sheet.getRange(sections.length + 1, 6, 1, 2).setValue("");
    sheet.setColumnWidth(3, 15);
    sheet.setColumnWidth(4, 15);
    sheet.setColumnWidth(6, 15);
    sheet.setColumnWidth(7, 15);
  } else {
    var cell = sheet.getRange("B3");
    cell.setNumberFormat("#,##0\\ \"M\"[$$-C0C]");
    cell.setHorizontalAlignment("right");
    cell.setValue("=$B$2*$B$" + (sections.length + 2));
    sheet.getRange("B4:B5").setNumberFormat(currencyFmt);
  }
  var range = sheet.getRange(1, 1, sections.length, 1);
  range.setValues(sections);
  range.setFontWeight("bold");
  var range = sheet.getRange(sections.length + 1, 1, 1, values[0].length + 1);
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");

  // Tidy the sheet.
  var delta = sheet.getMaxColumns() - 8;
  if (delta) {
    sheet.deleteColumns(9, delta);
  }
  sheet.setColumnWidth(8, 1000);
  SpreadsheetApp.flush();
  
  // Format the headers and create the graph.
  return updateOneSheetInner(sheet, ticker, sections.length + 2, sections.length + 1 + values.length);
}

// Format lines of data.
function formatLines(sheet, row, numRows, currencyFmt) {
  var range = sheet.getRange(row, 1, numRows, 1);
  range.setNumberFormat("yyyy-MM-dd");
  range.setHorizontalAlignment("left");

  sheet.getRange(row, 2, numRows, 4).setNumberFormat(currencyFmt);
  sheet.getRange(row, 6, numRows, 1).setNumberFormat("#,##0");
}

// Returns the format to use for the ticker.
function getCurrentFmt(ticker) {
  if (isCurrency(ticker)) {
    return "#,##0.0000\ [$$-C0C]";
  }
  return "#,##0.00\ [$$-C0C]";
}

// Utilities

// Utilities: Date

// Returns today as "YYYY-MM-DD".
function getToday() {
  return toStr(new Date());
}

// Converts a "YYYY-MM-DD" (or a integer) to a javascript Date object.
function toDate(s) {
  if (s instanceof Date) {
    return s;
  }
  if (typeof s === "number") {
    return new Date(s);
  }
  var parts = s.split("-");
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

// Converts a javascript Date object to "YYYY-MM-DD".
// Most importantly, it keeps the "0" to keep the length constant.
// For stocks, d will be "YYYY-MM-DD 16:00:00" but for currencies, it is "YYYY-MM-DD 23:58:00".
function toStr(d) {
  if (typeof d === "number") {
    assert(false);
    d = new Date(d);
  }
  if (d instanceof Date) {
    // Counter intuitively, we must not use d.getUTCFullYear() or d.toISOString() since the date was stored in locale timezone.
    return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate());
  }
  return d.split(" ")[0];
}

function pad(n) {
  if ((n + "").length == 1) {
    return "0" + n;
  }
  return n + "";
}

// Given "YYY-MM-DD", returns "YYYY-MM-DD" that is the day after.
function nextDay(day) {
  return toStr(new Date(toDate(day).getTime() + 24 * 60 * 60000));
}

// Given "YYY-MM-DD", returns "YYYY-MM-DD" that is the day before.
function previousDay(day) {
  return toStr(new Date(toDate(day).getTime() - 24 * 60 * 60000));
}

// Returns the number of days between two dates as Date instances.
function getDaysBetween(start, end) {
  var s = toDate(start).getTime() / 24 / 60 / 60000;
  var e = toDate(end).getTime() / 24 / 60 / 60000;
  if (s >= e) {
    return 0;
  }
  return e - s;
}

// Utilities: Stocks

// Returns true if the ticker is a currency exchange.
// This function doesn't do RPC.
function isCurrency(ticker) {
  return ticker.substr(0, 9) == "CURRENCY:";
}

// Utilities: low level GOOGLEFINANCE functions.
// These could be replaced by something else than GOOGLEFINANCE, e.g. UrlFetchApp.fetch() to another data provider.

// Gets the stock value from the start day until yesterday.
// Returns a list of list, e.g. [[]].
function getDayValuesUpToYesterday(sheet, ticker, startRow, startDate, includeHeader) {
  var endDate = previousDay(getToday());
  var days = getDaysBetween(startDate, endDate);
  Logger.log("Between", startDate, endDate, days);
  if (!days) {
    return null;
  }
  var spare = sheet.getMaxRows() - startRow - 1;
  if (spare < days) {
    sheet.insertRowsAfter(startRow, days - spare);
  }
  var cell = runGoogleFinanceInternal(sheet, "A" + startRow, [ticker, "all", startDate, endDate]);
  var lastRow = sheet.getLastRow();
  Logger.log("", startRow, lastRow);
  var values = null;
  if (lastRow > startRow) {
    var start = startRow + 1;
    if (includeHeader) {
      start = startRow;
    }
    values = sheet.getRange(start, 1, lastRow - start + 1, 6).getValues(); 
    // Change the date to be only YYYY-MM-DD, not 16:00.
    for (var y in values) {
      if (!includeHeader || y) {
        values[y][0] = toStr(values[y][0]);
        for (var x in values[y]) {
          if (values[y][x] == "#N/A") {
            values[y][x] = ""
          }
        }
      }
    }
  }
  cell.clearContent();
  return values;
}

// https://support.google.com/docs/answer/3093281
// Sadly, GOOGLEFINANCE() is not accessible from Google Apps (!).
function runGoogleFinanceInternal(sheet, location, args) {
  var cell = sheet.getRange(location);
  cell.setValue(getGOOGLEFINANCE(args));
  return cell
}

// Returns a GOOGLEFINANCE() call to put in a cell with properly escaped arguments.
// This function doesn't do RPC.
function getGOOGLEFINANCE(args) {
  var asStr = [];
  for (var a in args) {
    asStr.push("\"" + args[a] + "\"");
  }
  return "=GOOGLEFINANCE(" + asStr.join("; ") + ")";
}
