// ```Charts.gs

function generateCharts(trelloSheet) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var chartsSheet =
        sheet.getSheetByName("Charts") || sheet.insertSheet("Charts");
    chartsSheet.clear();

    var dataRange = trelloSheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[0];
    var data = values.slice(1);

    var chartData = extractChartData(headers, data);
    var chartRow = 1;

    // Create charts
    chartRow = createDistributionChart(chartsSheet, chartData, chartRow);
    chartRow = createTaskCountChart(chartsSheet, chartData, chartRow);
    chartRow = createTimelineChart(chartsSheet, chartData, chartRow);
    chartRow = createCustomFieldChart(chartsSheet, chartData, chartRow);
    chartRow = createSalesChart(chartsSheet, chartData, chartRow);
    chartRow = createSalesComparisonChart(chartsSheet, chartData, chartRow);
    chartRow = createEntryChannelChart(chartsSheet, chartData, chartRow);
    chartRow = createClientTypeChart(chartsSheet, chartData, chartRow);
    chartRow = createMaterialCategoryChart(chartsSheet, chartData, chartRow);
    chartRow = createProductiveHoursChart(chartsSheet, chartData, chartRow);

    chartsSheet.autoResizeColumns(1, chartsSheet.getLastColumn());
}

function extractChartData(headers, data) {
    var chartData = {
        names: [],
        values: {},
        categories: {},
        dates: {},
    };

    var numericColumns = [];
    var categoricalColumns = [];
    var dateColumns = [];

    headers.forEach((header, index) => {
        if (data.some((row) => typeof row[index] === "number")) {
            numericColumns.push(index);
        } else if (data.some((row) => row[index] instanceof Date)) {
            dateColumns.push(index);
        } else if (
            data.some(
                (row) => typeof row[index] === "string" && row[index] !== "",
            )
        ) {
            categoricalColumns.push(index);
        }
    });

    data.forEach(function (row) {
        chartData.names.push(row[1]);

        numericColumns.forEach((index) => {
            var header = headers[index];
            if (!chartData.values[header]) {
                chartData.values[header] = [];
            }
            chartData.values[header].push(row[index] || 0);
        });

        categoricalColumns.forEach((index) => {
            var header = headers[index];
            if (!chartData.categories[header]) {
                chartData.categories[header] = [];
            }
            chartData.categories[header].push(row[index] || "");
        });

        dateColumns.forEach((index) => {
            var header = headers[index];
            if (!chartData.dates[header]) {
                chartData.dates[header] = [];
            }
            chartData.dates[header].push(row[index] || null);
        });
    });

    return chartData;
}

function createDistributionChart(sheet, chartData, startRow) {
    var listHeader = Object.keys(chartData.categories).find((header) =>
        header.toLowerCase().includes("list"),
    );
    if (!listHeader) return startRow;

    var data = [["List", "Count"]];
    var listCounts = {};

    chartData.categories[listHeader].forEach((list) => {
        listCounts[list] = (listCounts[list] || 0) + 1;
    });

    Object.entries(listCounts).forEach(([list, count]) => {
        data.push([list, count]);
    });

    var range = sheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

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
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + data.length + 2;
}

function createTaskCountChart(sheet, chartData, startRow) {
    var labelHeader = Object.keys(chartData.categories).find((header) =>
        header.toLowerCase().includes("label"),
    );
    if (!labelHeader) return startRow;

    var data = [["Label", "Count"]];
    var labelCounts = {};

    chartData.categories[labelHeader].forEach((labelsString) => {
        var labels = labelsString
            .split(",")
            .map((l) => l.trim())
            .filter((l) => l);
        labels.forEach((label) => {
            labelCounts[label] = (labelCounts[label] || 0) + 1;
        });
    });

    Object.entries(labelCounts).forEach(([label, count]) => {
        data.push([label, count]);
    });

    var range = sheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

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
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + data.length + 2;
}

function createTimelineChart(sheet, chartData, startRow) {
    var creationDateHeader = Object.keys(chartData.dates).find((header) =>
        header.toLowerCase().includes("creation"),
    );
    if (!creationDateHeader) return startRow;

    var data = [["Date", "Cumulative Tasks"]];
    var sortedDates = chartData.dates[creationDateHeader]
        .filter((d) => d)
        .sort((a, b) => a - b);

    var cumulativeCount = 0;
    sortedDates.forEach((date) => {
        cumulativeCount++;
        data.push([date, cumulativeCount]);
    });

    var range = sheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Tasks Over Time")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Date", format: "MM/dd/yyyy" })
        .setOption("vAxis", { title: "Cumulative Tasks" })
        .setOption("height", 300)
        .setOption("width", 500)
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + data.length + 2;
}

function createCustomFieldChart(sheet, chartData, startRow) {
    var customFields = Object.keys(chartData.categories).filter(
        (header) =>
            !header.toLowerCase().includes("list") &&
            !header.toLowerCase().includes("label"),
    );

    if (customFields.length === 0) return startRow;

    var selectedField = customFields[0];
    var data = [[selectedField, "Count"]];
    var fieldCounts = {};

    chartData.categories[selectedField].forEach((value) => {
        if (value) {
            fieldCounts[value] = (fieldCounts[value] || 0) + 1;
        }
    });

    Object.entries(fieldCounts).forEach(([value, count]) => {
        data.push([value, count]);
    });

    var range = sheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

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
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + data.length + 2;
}

function createSalesChart(sheet, chartData, startRow) {
    var salesData = [
        ["Month", "Sales"],
        ["Jan", 1000],
        ["Feb", 1100],
        ["Mar", 1200],
        ["Apr", 1300],
        ["May", 1400],
        ["Jun", 1500],
        ["Jul", 1600],
        ["Aug", 1700],
        ["Sep", 1800],
        ["Oct", 1900],
        ["Nov", 2000],
        ["Dec", 2100],
    ];

    var range = sheet.getRange(
        startRow,
        1,
        salesData.length,
        salesData[0].length,
    );
    range.setValues(salesData);

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
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + salesData.length + 2;
}

function createSalesComparisonChart(sheet, chartData, startRow) {
    // Placeholder for sales comparison charts as requested
    var comparisonData = [
        ["Period", "Current Year", "Previous Year"],
        ["Q1", 12000, 11000],
        ["Q2", 15000, 14000],
        ["Q3", 16000, 13000],
        ["Q4", 17000, 12000],
    ];

    var range = sheet.getRange(
        startRow,
        1,
        comparisonData.length,
        comparisonData[0].length,
    );
    range.setValues(comparisonData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Sales Comparison by Quarter")
        .setOption("legend", { position: "right" })
        .setOption("hAxis", { title: "Quarter" })
        .setOption("vAxis", { title: "Sales" })
        .setOption("height", 300)
        .setOption("width", 500)
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + comparisonData.length + 2;
}

function createEntryChannelChart(sheet, chartData, startRow) {
    var entryChannelData = [
        ["Entry Channel", "Count"],
        ["BNI", 10],
        ["RRSS", 15],
        ["Feria", 5],
        ["Referenciados", 20],
        ["Web", 25],
    ];

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
        .setOption("title", "Projects by Entry Channel")
        .setOption("pieHole", 0.4)
        .setOption("legend", { position: "right" })
        .setOption("height", 300)
        .setOption("width", 500)
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + entryChannelData.length + 2;
}

function createClientTypeChart(sheet, chartData, startRow) {
    var clientTypeData = [
        ["Client Type", "Revenue"],
        ["Residential", 50000],
        ["Commercial", 75000],
        ["Industrial", 100000],
    ];

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
        .setOption("title", "Revenue by Client Type")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Client Type" })
        .setOption("vAxis", { title: "Revenue" })
        .setOption("height", 300)
        .setOption("width", 500)
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + clientTypeData.length + 2;
}

function createMaterialCategoryChart(sheet, chartData, startRow) {
    var materialCategoryData = [
        ["Material Category", "Sales"],
        ["Wood", 30000],
        ["Ceramic", 25000],
        ["Vinyl", 20000],
        ["Carpet", 15000],
    ];

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
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + materialCategoryData.length + 2;
}

function createProductiveHoursChart(sheet, chartData, startRow) {
    var productiveHoursData = [
        ["Month", "Productive Hours"],
        ["Jan", 160],
        ["Feb", 150],
        ["Mar", 170],
        ["Apr", 165],
        ["May", 180],
        ["Jun", 175],
        ["Jul", 185],
        ["Aug", 190],
        ["Sep", 195],
        ["Oct", 200],
        ["Nov", 205],
        ["Dec", 210],
    ];

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
        .setOption("fontName", "Inter")
        .setOption("titleTextStyle", { fontSize: 14, bold: true })
        .setOption("legendTextStyle", { fontSize: 12 })
        .build();

    sheet.insertChart(chart);

    return startRow + productiveHoursData.length + 2;
}
