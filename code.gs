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
// - Use the currency when formatting the lines baed on B1.
// - Have this script warn the user when a new version of this script is live.
//   This needs to include an easy way to turn the check off and explanation
//   about using multiple files for user scripts.
// - Discover titleLine value in createChart() based on the sheet content, so a user can
//   add more headers and it just works.
// - Do not get all values in createChart() but only the ranges needed.
// - Update shares # in updateOneSheet().


// Line at which the data starts. Update this value if you insert more headers.
var titleLine = 6;


// Hooks

function onOpen(e) {
  var ui = SpreadsheetApp.getUi()
  var menu = ui.createMenu('Stocks');
  menu.addItem('Update all sheets', 'updateAllSheets').addToUi();
  menu.addItem('Update current sheet', 'updateCurrentSheet').addToUi();
  menu.addItem('Add new sheet', 'addNewSheet').addToUi();
  var scratchSheet = SpreadsheetApp.getActive().getSheetByName("Explanations");
  if (scratchSheet == null) {
    ui.alert("Please recreate the sheet 'Explanations' before doing anything!\nAdd explanations to it. Lines 50 and later will be used as scratch space.");
  } else {
    updateAllSheets();
  }
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
  var date = nextDay(toStr(sheet.getRange(lastRow, 1).getValue()));
  var changed = false;

  // Trim lines.
  var delta = sheet.getMaxRows() - lastRow;
  if (delta) {
    sheet.deleteRows(lastRow+1, delta);
  }

  // Do it in batches. This is more efficient than calling GOOGLEFINANCE() one by one.
  while (true) {
    var values = getDayValuesUpToToday(ticker, date);
    if (values == null) {
      break
    }
    changed = true;
    sheet.insertRowsAfter(lastRow, values.length);
    sheet.getRange(lastRow+1, 1, values.length, values[0].length).setValues(values);
    formatLines(sheet, lastRow+1, values.length);
    lastRow += values.length;
    date = nextDay(toStr(values[values.length - 1][0]));
  }
  // Save some work and skip right away.
  if (!changed) {
    return true;
  }
  // Market cap is last share value * current price.
  if (!isCurrency(ticker)) {
    sheet.getRange("B3").setValue("=$B$2*$B$" + (lastRow-1));
  }
  sheet.getRange("B4:B5").setValues(
    [
      ["=MIN(E7:E" + lastRow + ")"],
      ["=MAX(E7:E" + lastRow + ")"],
  ])
  // Resize the columns with some margin.
  SpreadsheetApp.flush();
  for (var i = 1; i < 8; i++) {
    sheet.autoResizeColumn(i);
    sheet.setColumnWidth(i, sheet.getColumnWidth(i)+10);
  }
  createChart(sheet);
  return true;
}

function createChart(sheet) {
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
  var ticker = sheet.getName();
  var legend = {
  };
  var hAxis = {
    "gridlines": {
      "count": -1,
      "units": {
        "years": {
          "format": ["yy-mm"],
        },
      },
    },
    "minorGridlines": {
      "count": -1,
      "units": {
        "months": {"format": ["yy-mm", "-mm", ""]},
      },
    },
  };
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(titleLine, 1, lastRow - titleLine - 1, 6);
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
  }
  if (currency == "#N/A") {
    ui.alert(ticker + " is not a valid ticker. Delete the sheet and try again.");
    return false;
  }
  cell1.setValue(currency);

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
    return false;
  }

  // Initialize the headers.
  var sections = [
    ["Currency"],
    ["Shares"],
    ["Market Cap"],
    ["Close low"],
    ["Close high"],
  ];
  // Sadly dividends are not supported by GOOGLEFINANCE() so one has to track it manually! :(
  var headers = [
    ["Date", "Open", "High", "Low", "Close", "Volume", "Dividend"],
  ];
  if (isCurrency(ticker)) {
    sections[1][0] = "";
    sections[2][0] = "";
    headers[0][1] = "";
    headers[0][2] = "";
    headers[0][3] = "";
    headers[0][5] = "";
    headers[0][6] = "";
    sheet.setColumnWidth(3, 15);
    sheet.setColumnWidth(4, 15);
    sheet.setColumnWidth(6, 15);
    sheet.setColumnWidth(7, 15);
  } else {
    var cell = sheet.getRange("B3");
    cell.setNumberFormat("#,##0\\ \"M\"[$$-C0C]");
    cell.setValue("=$B$2*$B$" + (sections.length + 2));
    var range = sheet.getRange("B4:B5");
    range.setNumberFormat("#,##0.00\ [$$-C0C]");
    range.setValues([["=MIN(E7:E7)"], ["=MAX(E7:E7)"]]);
  }
  var range = sheet.getRange(1, 1, sections.length, 1);
  range.setValues(sections);
  range.setFontWeight("bold");
  var range = sheet.getRange(sections.length + 1, 1, 1, headers[0].length);
  range.setValues(headers);
  range.setFontWeight("bold");
  range.setHorizontalAlignment("center");

  // Initialize with the start date. Search for the first valid date.
  var values = getDayValues(ticker, startDate);
  if (values == null) {
    ui.alert("Failed to find a date where this ticker is valid.");
    return false;
  }
  sheet.getRange(sections.length + 2, 1, 1, values[0].length).setValues(values);
  formatLines(sheet, sections.length + 2, 1);

  // Tidy the sheet.
  var delta = sheet.getMaxColumns() - 8;
  if (delta) {
    sheet.deleteColumns(9, delta);
  }
  /*
  var delta = sheet.getMaxRows() - titleLine-2;
  if (delta) {
    sheet.deleteRows(titleLine+3, delta);
  }
  */
  sheet.setColumnWidth(8, 1000);
  SpreadsheetApp.flush();
  
  // Add the remaining values and create the graph.
  updateOneSheet(sheet);
}

// Format lines of data.
function formatLines(sheet, row, numRows) {
  var range = sheet.getRange(row, 1, numRows, 1);
  range.setNumberFormat("yyyy-MM-dd");
  range.setHorizontalAlignment("left");

  sheet.getRange(row, 2, numRows, 4).setNumberFormat("#,##0.00\ [$$-C0C]");
  sheet.getRange(row, 6, numRows, 1).setNumberFormat("#,##0");
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
function toStr(d) {
  if (typeof d === "number") {
    d = new Date(d);
  }
  if (d instanceof Date) {
    return d.toISOString().split("T")[0];
  }
  return d;
}

// Given "YYY-MM-DD", returns "YYYY-MM-DD" that is the day after.
function nextDay(day) {
  return toStr(new Date(toDate(day).getTime() + 24 * 60 * 60 * 1000));
}

// Utilities: Stocks

// Returns the stock value for the day or any following that is valid.
function getDayValues(ticker, day) {
  var date = toStr(day);
  var today = getToday();
  var values = getSymbolAtDate(ticker, date);
  while (values == null) {
    if (date == today) {
      return null;
    }
    date = nextDay(date);
    values = getSymbolAtDate(ticker, date);
  }
  return values;
}

// Returns the stock value from the start day until today or 1 year if it's too much.
function getDayValuesUpToToday(ticker, startDate) {
  // The challenge is that we don't know in advance how many lines will be returned.
  var startD = toDate(startDate);
  var end = getToday();
  var endD = toDate(end);
  if (startD.getTime() >= endD.getTime()) {
    return null;
  }
  // Javascript #closeenough
  var oneYear = 366 * 24 * 60 * 60 * 1000;
  if ((endD.getTime() - startD.getTime()) > oneYear) {
    end = toStr(new Date(startD.getTime() + oneYear));
  }

  var values = runGoogleFinanceRange([ticker, "all", startDate, end], 366);
  if (values == null) {
    return null;
  }
  // Change the date to be only YYYY-MM-DD, not 16:00.
  for (var i in values) {
    values[i][0] = toStr(new Date(values[i][0]));
  }
  return values;
}

// Returns true if the ticker is a currency exchange.
// This function doesn't do RPC.
function isCurrency(ticker) {
  return ticker.substr(0, 9) == "CURRENCY:";
}

function getSymbolAtDate(ticker, startDate) {
  var values = runGoogleFinanceRange([ticker, "all", startDate], 1);
  if (values == null || values[0][0] == "") {
    return null;
  }
  // Change the date to be only YYYY-MM-DD, not 16:00.
  values[0][0] = toStr(new Date(values[0][0]));
  return values;
}

// Utilities: low level GOOGLEFINANCE functions.
// These could be replaced by something else than GOOGLEFINANCE, e.g. UrlFetchApp.fetch() to another data provider.

// Returns a list of list, e.g. [[]].
function runGoogleFinanceRange(args, nbLines) {
  var scratchSheet = SpreadsheetApp.getActive().getSheetByName("Explanations");
  var cell = runGoogleFinanceInternal(scratchSheet, "A50", args);
  var values = scratchSheet.getRange(51, 1, 51+nbLines, 6).getValues();
  // Trim the extra lines.
  while (values.length && values[values.length-1][0] == "") {
    values.pop();
  }
  if (values.length == 0) {
    values = null;
  } else {
    for (var j in values) {
      for (var i in values[j]) {
        if (values[j][i] == "#N/A") {
          values[j][i] = ""
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
