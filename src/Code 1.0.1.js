// ```Code.js

var CACHE = CacheService.getScriptCache();
var PROPERTIES = PropertiesService.getScriptProperties();

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Trello Integration")
        .addItem("Configure Settings", "showSettingsDialog")
        .addItem("Sync with Trello", "syncWithTrello")
        .addItem("Generate Charts", "generateChartsWrapper")
        .addItem("Generate Reports", "generateReports")
        .addItem("Analyze Sales", "analyzeSales")
        .addToUi();
}

function syncWithTrello() {
    var settings = getSettings();
    if (!validateSettings(settings)) {
        SpreadsheetApp.getUi().alert(
            "Error: Invalid or missing settings. Please configure the settings first.",
        );
        return;
    }

    try {
        var boardData = fetchTrelloBoardData(settings);
        var customFields = boardData.customFields || [];
        var lists = boardData.lists || [];
        var cards = boardData.cards || [];

        updateSpreadsheetWithTrelloData(cards, customFields, lists);

        SpreadsheetApp.getUi().alert(
            "Sync with Trello completed successfully!",
        );
    } catch (error) {
        Logger.log("Error syncing with Trello: " + error.toString());
        SpreadsheetApp.getUi().alert(
            "Error syncing with Trello: " + error.toString(),
        );
    }
}

function fetchTrelloBoardData(settings) {
    var boardUrl = `https://api.trello.com/1/boards/${settings.TRELLO_BOARD_ID}?key=${settings.TRELLO_API_KEY}&token=${settings.TRELLO_API_TOKEN}&cards=all&customFields=true&lists=all`;
    var boardData = fetchTrelloData(boardUrl, "board data");

    return {
        customFields: boardData.customFields,
        lists: boardData.lists,
        cards: boardData.cards,
    };
}

function fetchTrelloData(url, dataType) {
    var options = {
        muteHttpExceptions: true,
        headers: {
            "Accept-Language": "en-US,en;q=0.9",
        },
    };
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode !== 200) {
        throw new Error(
            `Error fetching ${dataType}: HTTP ${responseCode}: ${responseBody}`,
        );
    }

    return JSON.parse(responseBody);
}

function updateSpreadsheetWithTrelloData(cards, customFields, lists) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var trelloSheet =
        sheet.getSheetByName("Trello Data") || sheet.insertSheet("Trello Data");
    trelloSheet.clear();

    var headers = [
        "ID",
        "Name",
        "List",
        "Creation Date",
        "Due Date",
        "Labels",
        "Description",
        "Sales Amount",
        "Client Type",
        "Material Category",
        "Entry Channel",
        "Productive Hours",
    ];
    customFields.forEach(function (field) {
        headers.push(field.name);
    });
    trelloSheet.appendRow(headers);

    var listMap = {};
    lists.forEach(function (list) {
        listMap[list.id] = list.name;
    });

    cards.forEach(function (card) {
        var rowData = [
            card.id,
            card.name,
            listMap[card.idList] || card.idList,
            new Date(card.dateLastActivity),
            card.due ? new Date(card.due) : "",
            card.labels.map((label) => label.name).join(", "),
            card.desc,
            getCustomFieldValue(card.customFieldItems, "Sales Amount"),
            getCustomFieldValue(card.customFieldItems, "Client Type"),
            getCustomFieldValue(card.customFieldItems, "Material Category"),
            getCustomFieldValue(card.customFieldItems, "Entry Channel"),
            getCustomFieldValue(card.customFieldItems, "Productive Hours"),
        ];

        customFields.forEach(function (field) {
            var fieldValue = getCustomFieldValue(
                card.customFieldItems,
                field.id,
            );
            rowData.push(fieldValue);
        });

        trelloSheet.appendRow(rowData);
    });

    applySheetFormatting(trelloSheet);
}

function getCustomFieldValue(customFields, fieldName) {
    if (!customFields) return "";

    for (var i = 0; i < customFields.length; i++) {
        var field = customFields[i];
        if (field.idCustomField === fieldName || field.name === fieldName) {
            if (field.value) {
                if (field.value.hasOwnProperty("text")) {
                    return field.value.text;
                } else if (field.value.hasOwnProperty("number")) {
                    return field.value.number;
                } else if (field.value.hasOwnProperty("checked")) {
                    return field.value.checked;
                } else if (field.value.hasOwnProperty("date")) {
                    return new Date(field.value.date);
                }
            }
        }
    }
    return "";
}

function applySheetFormatting(sheet) {
    var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground("#007aff");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    headerRange.setFontFamily("Inter");

    range.setFontSize(10);
    range.setFontFamily("Inter");
    range.setVerticalAlignment("middle");
    range.setHorizontalAlignment("left");

    for (var i = 2; i <= sheet.getLastRow(); i++) {
        var rowRange = sheet.getRange(i, 1, 1, sheet.getLastColumn());
        if (i % 2 == 0) {
            rowRange.setBackground("#f0f0f0");
        } else {
            rowRange.setBackground("#ffffff");
        }
    }

    sheet
        .getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn())
        .setWrap(true);

    for (var i = 1; i <= sheet.getLastColumn(); i++) {
        sheet.autoResizeColumn(i);
    }

    headerRange.createFilter();
}

function showSettingsDialog() {
    var html = HtmlService.createHtmlOutputFromFile("index")
        .setWidth(400)
        .setHeight(300)
        .setTitle("Configure Trello Integration Settings");
    SpreadsheetApp.getUi().showModalDialog(
        html,
        "Configure Trello Integration Settings",
    );
}

function saveSettings(settings) {
    for (var key in settings) {
        PROPERTIES.setProperty(key, settings[key]);
    }
    return { success: true, message: "Settings saved successfully!" };
}

function getSettings() {
    var settings = {};
    var keys = [
        "TRELLO_API_KEY",
        "TRELLO_SECRET_KEY",
        "TRELLO_API_TOKEN",
        "TRELLO_BOARD_ID",
        "SHEET_ID",
    ];
    keys.forEach(function (key) {
        settings[key] = PROPERTIES.getProperty(key) || "";
    });
    return settings;
}

function validateSettings(settings) {
    return (
        settings.TRELLO_API_KEY &&
        settings.TRELLO_API_TOKEN &&
        settings.TRELLO_BOARD_ID
    );
}

function analyzeSales() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var trelloDataSheet = sheet.getSheetByName("Trello Data");

    if (!trelloDataSheet) {
        SpreadsheetApp.getUi().alert(
            "Error: Trello Data sheet not found. Please sync with Trello first.",
        );
        return;
    }

    var data = trelloDataSheet.getDataRange().getValues();
    var headers = data[0];
    var salesData = data.slice(1);

    var salesAnalysisSheet =
        sheet.getSheetByName("Sales Analysis") ||
        sheet.insertSheet("Sales Analysis");
    salesAnalysisSheet.clear();

    // Ventas Totales
    var totalSales = calculateTotalSales(salesData, headers);
    var monthlySales = calculateMonthlySales(salesData, headers);
    var yearlySales = calculateYearlySales(salesData, headers);
    var quarterlySales = calculateQuarterlySales(salesData, headers);

    // Canal de Entrada
    var salesByEntryChannel = calculateSalesByEntryChannel(salesData, headers);

    // Tipología de Clientes
    var salesByClientType = calculateSalesByClientType(salesData, headers);

    // Categoría de Material
    var salesByMaterialCategory = calculateSalesByMaterialCategory(
        salesData,
        headers,
    );

    // Horas Productivas
    var productiveHours = calculateProductiveHours(salesData, headers);

    // Write results to Sales Analysis sheet
    writeSalesAnalysis(
        salesAnalysisSheet,
        totalSales,
        monthlySales,
        yearlySales,
        quarterlySales,
        salesByEntryChannel,
        salesByClientType,
        salesByMaterialCategory,
        productiveHours,
    );

    SpreadsheetApp.getUi().alert("Sales analysis completed successfully!");
}

function calculateTotalSales(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    return salesData.reduce(
        (total, row) => total + (row[salesAmountIndex] || 0),
        0,
    );
}

function calculateMonthlySales(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var dateIndex = headers.indexOf("Creation Date");
    var monthlySales = {};

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var monthYear = `${date.getFullYear()}-${date.getMonth() + 1}`;
        monthlySales[monthYear] =
            (monthlySales[monthYear] || 0) + (row[salesAmountIndex] || 0);
    });

    return monthlySales;
}

function calculateYearlySales(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var dateIndex = headers.indexOf("Creation Date");
    var yearlySales = {};

    salesData.forEach((row) => {
        var year = new Date(row[dateIndex]).getFullYear();
        yearlySales[year] =
            (yearlySales[year] || 0) + (row[salesAmountIndex] || 0);
    });

    return yearlySales;
}

function calculateQuarterlySales(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var dateIndex = headers.indexOf("Creation Date");
    var quarterlySales = {};

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var quarter = Math.floor(date.getMonth() / 3) + 1;
        var key = `${year}-Q${quarter}`;
        quarterlySales[key] =
            (quarterlySales[key] || 0) + (row[salesAmountIndex] || 0);
    });

    return quarterlySales;
}

function calculateSalesByEntryChannel(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var entryChannelIndex = headers.indexOf("Entry Channel");
    var salesByEntryChannel = {};

    salesData.forEach((row) => {
        var channel = row[entryChannelIndex];
        salesByEntryChannel[channel] =
            (salesByEntryChannel[channel] || 0) + (row[salesAmountIndex] || 0);
    });

    return salesByEntryChannel;
}

function calculateSalesByClientType(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var clientTypeIndex = headers.indexOf("Client Type");
    var salesByClientType = {};

    salesData.forEach((row) => {
        var clientType = row[clientTypeIndex];
        salesByClientType[clientType] =
            (salesByClientType[clientType] || 0) + (row[salesAmountIndex] || 0);
    });

    return salesByClientType;
}

function calculateSalesByMaterialCategory(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var materialCategoryIndex = headers.indexOf("Material Category");
    var salesByMaterialCategory = {};

    salesData.forEach((row) => {
        var category = row[materialCategoryIndex];
        salesByMaterialCategory[category] =
            (salesByMaterialCategory[category] || 0) +
            (row[salesAmountIndex] || 0);
    });

    return salesByMaterialCategory;
}

function calculateProductiveHours(salesData, headers) {
    var productiveHoursIndex = headers.indexOf("Productive Hours");
    var dateIndex = headers.indexOf("Creation Date");
    var productiveHours = {};

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var monthYear = `${date.getFullYear()}-${date.getMonth() + 1}`;
        productiveHours[monthYear] =
            (productiveHours[monthYear] || 0) +
            (row[productiveHoursIndex] || 0);
    });

    return productiveHours;
}

function writeSalesAnalysis(
    sheet,
    totalSales,
    monthlySales,
    yearlySales,
    quarterlySales,
    salesByEntryChannel,
    salesByClientType,
    salesByMaterialCategory,
    productiveHours,
) {
    sheet.getRange("A1").setValue("Sales Analysis Report");
    sheet.getRange("A1").setFontWeight("bold").setFontSize(14);

    // Ventas Totales
    sheet.getRange("A3").setValue("Total Sales:");
    sheet.getRange("B3").setValue(totalSales);

    sheet.getRange("A5").setValue("Monthly Sales:");
    var row = 6;
    for (var month in monthlySales) {
        sheet.getRange(row, 1).setValue(month);
        sheet.getRange(row, 2).setValue(monthlySales[month]);
        row++;
    }

    sheet.getRange("D5").setValue("Yearly Sales:");
    row = 6;
    for (var year in yearlySales) {
        sheet.getRange(row, 4).setValue(year);
        sheet.getRange(row, 5).setValue(yearlySales[year]);
        row++;
    }

    sheet.getRange("G5").setValue("Quarterly Sales:");
    row = 6;
    for (var quarter in quarterlySales) {
        sheet.getRange(row, 7).setValue(quarter);
        sheet.getRange(row, 8).setValue(quarterlySales[quarter]);
        row++;
    }

    // Canal de Entrada
    sheet.getRange("A" + (row + 2)).setValue("Sales by Entry Channel:");
    row += 3;
    for (var channel in salesByEntryChannel) {
        sheet.getRange(row, 1).setValue(channel);
        sheet.getRange(row, 2).setValue(salesByEntryChannel[channel]);
        row++;
    }

    // Tipología de Clientes
    sheet
        .getRange("D" + (row - salesByEntryChannel.length))
        .setValue("Sales by Client Type:");
    row = row - salesByEntryChannel.length + 1;
    for (var clientType in salesByClientType) {
        sheet.getRange(row, 4).setValue(clientType);
        sheet.getRange(row, 5).setValue(salesByClientType[clientType]);
        row++;
    }

    // Categoría de Material
    sheet
        .getRange("G" + (row - salesByClientType.length))
        .setValue("Sales by Material Category:");
    row = row - salesByClientType.length + 1;
    for (var category in salesByMaterialCategory) {
        sheet.getRange(row, 7).setValue(category);
        sheet.getRange(row, 8).setValue(salesByMaterialCategory[category]);
        row++;
    }

    // Horas Productivas
    sheet.getRange("A" + (row + 2)).setValue("Productive Hours:");
    row += 3;
    for (var month in productiveHours) {
        sheet.getRange(row, 1).setValue(month);
        sheet.getRange(row, 2).setValue(productiveHours[month]);
        row++;
    }

    // Apply formatting
    sheet.getRange(1, 1, row, 8).setNumberFormat("#,##0.00");
    sheet.getRange(1, 1, row, 8).applyRowBanding();
    sheet.autoResizeColumns(1, 8);
}

function generateChartsWrapper() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var trelloSheet = sheet.getSheetByName("Trello Data");
    var salesAnalysisSheet = sheet.getSheetByName("Sales Analysis");
    if (trelloSheet && salesAnalysisSheet) {
        generateCharts(trelloSheet, salesAnalysisSheet);
        SpreadsheetApp.getUi().alert("Charts generated successfully!");
    } else {
        SpreadsheetApp.getUi().alert(
            "Error: Required sheets not found. Please sync with Trello and analyze sales first.",
        );
    }
}

function generateCharts(trelloSheet, salesAnalysisSheet) {
    var chartsSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Charts") ||
        SpreadsheetApp.getActiveSpreadsheet().insertSheet("Charts");
    chartsSheet.clear();

    var trelloData = trelloSheet.getDataRange().getValues();
    var salesAnalysisData = salesAnalysisSheet.getDataRange().getValues();

    var chartRow = 1;

    // Create charts
    chartRow = createDistributionChart(chartsSheet, trelloData, chartRow);
    chartRow = createTaskCountChart(chartsSheet, trelloData, chartRow);
    chartRow = createTimelineChart(chartsSheet, trelloData, chartRow);
    chartRow = createCustomFieldChart(chartsSheet, trelloData, chartRow);
    chartRow = createSalesChart(chartsSheet, salesAnalysisData, chartRow);
    chartRow = createSalesComparisonChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createEntryChannelChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createClientTypeChart(chartsSheet, salesAnalysisData, chartRow);
    chartRow = createMaterialCategoryChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createProductiveHoursChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );

    chartsSheet.autoResizeColumns(1, chartsSheet.getLastColumn());
}

function createDistributionChart(sheet, data, startRow) {
    var headers = data[0];
    var listIndex = headers.indexOf("List");
    if (listIndex === -1) return startRow;

    var listCounts = {};
    for (var i = 1; i < data.length; i++) {
        var list = data[i][listIndex];
        listCounts[list] = (listCounts[list] || 0) + 1;
    }

    var chartData = [["List", "Count"]];
    for (var list in listCounts) {
        chartData.push([list, listCounts[list]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        chartData.length,
        chartData[0].length,
    );
    range.setValues(chartData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Task Distribution by List")
        .setOption("pieHole", 0.4)
        .setOption("legend", { position: "right" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + chartData.length + 2;
}

function createTaskCountChart(sheet, data, startRow) {
    var headers = data[0];
    var labelsIndex = headers.indexOf("Labels");
    if (labelsIndex === -1) return startRow;

    var labelCounts = {};
    for (var i = 1; i < data.length; i++) {
        var labels = data[i][labelsIndex].split(",");
        labels.forEach(function (label) {
            label = label.trim();
            if (label) {
                labelCounts[label] = (labelCounts[label] || 0) + 1;
            }
        });
    }

    var chartData = [["Label", "Count"]];
    for (var label in labelCounts) {
        chartData.push([label, labelCounts[label]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        chartData.length,
        chartData[0].length,
    );
    range.setValues(chartData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.BAR)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Task Count by Label")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Count" })
        .setOption("vAxis", { title: "Label" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + chartData.length + 2;
}

function createTimelineChart(sheet, data, startRow) {
    var headers = data[0];
    var dateIndex = headers.indexOf("Creation Date");
    if (dateIndex === -1) return startRow;

    var dateCounts = {};
    for (var i = 1; i < data.length; i++) {
        var date = new Date(data[i][dateIndex]);
        var dateString = Utilities.formatDate(
            date,
            Session.getScriptTimeZone(),
            "yyyy-MM-dd",
        );
        dateCounts[dateString] = (dateCounts[dateString] || 0) + 1;
    }

    var chartData = [["Date", "Tasks Created"]];
    var sortedDates = Object.keys(dateCounts).sort();
    var cumulativeCount = 0;
    sortedDates.forEach(function (date) {
        cumulativeCount += dateCounts[date];
        chartData.push([new Date(date), cumulativeCount]);
    });

    var range = sheet.getRange(
        startRow,
        1,
        chartData.length,
        chartData[0].length,
    );
    range.setValues(chartData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Cumulative Tasks Over Time")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Date", format: "MM/dd/yyyy" })
        .setOption("vAxis", { title: "Cumulative Tasks" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + chartData.length + 2;
}

function createCustomFieldChart(sheet, data, startRow) {
    var headers = data[0];
    var customFields = headers.slice(7);
    if (customFields.length === 0) return startRow;

    var selectedField = customFields[0];
    var fieldIndex = headers.indexOf(selectedField);

    var fieldCounts = {};
    for (var i = 1; i < data.length; i++) {
        var value = data[i][fieldIndex];
        if (value) {
            fieldCounts[value] = (fieldCounts[value] || 0) + 1;
        }
    }

    var chartData = [[selectedField, "Count"]];
    for (var value in fieldCounts) {
        chartData.push([value, fieldCounts[value]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        chartData.length,
        chartData[0].length,
    );
    range.setValues(chartData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", `Task Distribution by ${selectedField}`)
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: selectedField })
        .setOption("vAxis", { title: "Count" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + chartData.length + 2;
}

function createSalesChart(sheet, data, startRow) {
    var monthlySalesData = [["Month", "Sales"]];
    var monthlyIndex = data.findIndex((row) => row[0] === "Monthly Sales:");
    if (monthlyIndex === -1) return startRow;

    for (var i = monthlyIndex + 1; i < data.length && data[i][0] !== ""; i++) {
        monthlySalesData.push([data[i][0], data[i][1]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        monthlySalesData.length,
        monthlySalesData[0].length,
    );
    range.setValues(monthlySalesData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Monthly Sales")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Month" })
        .setOption("vAxis", { title: "Sales" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + monthlySalesData.length + 2;
}

function createSalesComparisonChart(sheet, data, startRow) {
    var yearlySalesData = [["Year", "Sales"]];
    var yearlyIndex = data.findIndex((row) => row[0] === "Yearly Sales:");
    if (yearlyIndex === -1) return startRow;

    for (var i = yearlyIndex + 1; i < data.length && data[i][3] !== ""; i++) {
        yearlySalesData.push([data[i][3], data[i][4]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        yearlySalesData.length,
        yearlySalesData[0].length,
    );
    range.setValues(yearlySalesData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Yearly Sales Comparison")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Year" })
        .setOption("vAxis", { title: "Sales" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + yearlySalesData.length + 2;
}

function createEntryChannelChart(sheet, data, startRow) {
    var entryChannelData = [["Entry Channel", "Sales"]];
    var entryChannelIndex = data.findIndex(
        (row) => row[0] === "Sales by Entry Channel:",
    );
    if (entryChannelIndex === -1) return startRow;

    for (
        var i = entryChannelIndex + 1;
        i < data.length && data[i][0] !== "";
        i++
    ) {
        entryChannelData.push([data[i][0], data[i][1]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        entryChannelData.length,
        entryChannelData[0].length,
    );
    range.setValues(entryChannelData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Sales by Entry Channel")
        .setOption("pieHole", 0.4)
        .setOption("legend", { position: "right" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + entryChannelData.length + 2;
}

function createClientTypeChart(sheet, data, startRow) {
    var clientTypeData = [["Client Type", "Sales"]];
    var clientTypeIndex = data.findIndex(
        (row) => row[3] === "Sales by Client Type:",
    );
    if (clientTypeIndex === -1) return startRow;

    for (
        var i = clientTypeIndex + 1;
        i < data.length && data[i][3] !== "";
        i++
    ) {
        clientTypeData.push([data[i][3], data[i][4]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        clientTypeData.length,
        clientTypeData[0].length,
    );
    range.setValues(clientTypeData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Sales by Client Type")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Client Type" })
        .setOption("vAxis", { title: "Sales" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + clientTypeData.length + 2;
}

function createMaterialCategoryChart(sheet, data, startRow) {
    var materialCategoryData = [["Material Category", "Sales"]];
    var materialCategoryIndex = data.findIndex(
        (row) => row[6] === "Sales by Material Category:",
    );
    if (materialCategoryIndex === -1) return startRow;

    for (
        var i = materialCategoryIndex + 1;
        i < data.length && data[i][6] !== "";
        i++
    ) {
        materialCategoryData.push([data[i][6], data[i][7]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        materialCategoryData.length,
        materialCategoryData[0].length,
    );
    range.setValues(materialCategoryData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Sales by Material Category")
        .setOption("pieHole", 0.4)
        .setOption("legend", { position: "right" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + materialCategoryData.length + 2;
}

function createProductiveHoursChart(sheet, data, startRow) {
    var productiveHoursData = [["Month", "Productive Hours"]];
    var productiveHoursIndex = data.findIndex(
        (row) => row[0] === "Productive Hours:",
    );
    if (productiveHoursIndex === -1) return startRow;

    for (
        var i = productiveHoursIndex + 1;
        i < data.length && data[i][0] !== "";
        i++
    ) {
        productiveHoursData.push([data[i][0], data[i][1]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        productiveHoursData.length,
        productiveHoursData[0].length,
    );
    range.setValues(productiveHoursData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Productive Hours per Month")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Month" })
        .setOption("vAxis", { title: "Hours" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + productiveHoursData.length + 2;
}

function generateReports() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var trelloDataSheet = ss.getSheetByName("Trello Data");
    var salesAnalysisSheet = ss.getSheetByName("Sales Analysis");

    if (!trelloDataSheet || !salesAnalysisSheet) {
        SpreadsheetApp.getUi().alert(
            "Error: Required sheets not found. Please sync with Trello and analyze sales first.",
        );
        return;
    }

    var reportSheet = ss.getSheetByName("Report") || ss.insertSheet("Report");
    reportSheet.clear();

    // Add report title
    reportSheet
        .getRange("A1")
        .setValue("Trello Board and Sales Analysis Report")
        .setFontWeight("bold")
        .setFontSize(16);

    // Add date of report
    reportSheet
        .getRange("A2")
        .setValue("Report generated on: " + new Date().toLocaleString());

    // Add summary statistics
    var trelloData = trelloDataSheet.getDataRange().getValues();
    var salesAnalysisData = salesAnalysisSheet.getDataRange().getValues();

    var totalCards = trelloData.length - 1; // Subtract 1 for header row
    var totalSales = salesAnalysisData[2][1]; // Assuming total sales is in B3

    reportSheet.getRange("A4").setValue("Total Cards: " + totalCards);
    reportSheet.getRange("A5").setValue("Total Sales: " + totalSales);

    // Add monthly sales summary
    var monthlySalesIndex = salesAnalysisData.findIndex(
        (row) => row[0] === "Monthly Sales:",
    );
    if (monthlySalesIndex !== -1) {
        reportSheet.getRange("A7").setValue("Monthly Sales Summary:");
        var row = 8;
        for (
            var i = monthlySalesIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][0] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][0]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][1]);
            row++;
        }
    }

    // Add sales by entry channel
    var entryChannelIndex = salesAnalysisData.findIndex(
        (row) => row[0] === "Sales by Entry Channel:",
    );
    if (entryChannelIndex !== -1) {
        reportSheet
            .getRange("A" + (row + 2))
            .setValue("Sales by Entry Channel:");
        row += 3;
        for (
            var i = entryChannelIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][0] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][0]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][1]);
            row++;
        }
    }

    // Add sales by client type
    var clientTypeIndex = salesAnalysisData.findIndex(
        (row) => row[3] === "Sales by Client Type:",
    );
    if (clientTypeIndex !== -1) {
        reportSheet.getRange("A" + (row + 2)).setValue("Sales by Client Type:");
        row += 3;
        for (
            var i = clientTypeIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][3] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][3]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][4]);
            row++;
        }
    }

    // Add sales by material category
    var materialCategoryIndex = salesAnalysisData.findIndex(
        (row) => row[6] === "Sales by Material Category:",
    );
    if (materialCategoryIndex !== -1) {
        reportSheet
            .getRange("A" + (row + 2))
            .setValue("Sales by Material Category:");
        row += 3;
        for (
            var i = materialCategoryIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][6] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][6]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][7]);
            row++;
        }
    }

    // Apply formatting
    reportSheet.getRange(1, 1, row, 2).setNumberFormat("#,##0.00");
    reportSheet.getRange(1, 1, row, 2).applyRowBanding();
    reportSheet.autoResizeColumns(1, 2);

    // Print the report
    var printOptions = {
        size: SpreadsheetApp.PaperSize.A4,
        orientation: SpreadsheetApp.PageOrientation.PORTRAIT,
        fitToPage: true,
        printNotes: false,
        printGridlines: false,
        repeatTopRows: 2,
    };

    reportSheet.setPageBreak(row, 1);

    SpreadsheetApp.getUi().alert(
        "Report generated successfully! Check the 'Report' sheet. You can now print or export the report as needed.",
    );
}
