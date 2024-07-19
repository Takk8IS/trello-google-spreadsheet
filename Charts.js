// ```Charts.gs

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
    chartRow = createMonthlySalesChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createYearlySalesComparisonChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createQuarterlySalesChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createSalesByEntryChannelChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createSalesByClientTypeChart(
        chartsSheet,
        salesAnalysisData,
        chartRow,
    );
    chartRow = createSalesByMaterialCategoryChart(
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

function createMonthlySalesChart(sheet, data, startRow) {
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

function createYearlySalesComparisonChart(sheet, data, startRow) {
    var yearlySalesData = [["Year", "Sales", "Previous Year Sales"]];
    var yearlyIndex = data.findIndex((row) => row[0] === "Yearly Sales:");
    if (yearlyIndex === -1) return startRow;

    var years = [];
    var sales = [];

    for (var i = yearlyIndex + 1; i < data.length && data[i][3] !== ""; i++) {
        years.push(data[i][3]);
        sales.push(data[i][4]);
    }

    for (var i = 1; i < years.length; i++) {
        yearlySalesData.push([years[i], sales[i], sales[i - 1]]);
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
        .setOption("legend", { position: "bottom" })
        .setOption("hAxis", { title: "Year" })
        .setOption("vAxis", { title: "Sales" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + yearlySalesData.length + 2;
}

function createQuarterlySalesChart(sheet, data, startRow) {
    var quarterlySalesData = [["Quarter", "Sales"]];
    var quarterlyIndex = data.findIndex((row) => row[6] === "Quarterly Sales:");
    if (quarterlyIndex === -1) return startRow;

    for (
        var i = quarterlyIndex + 1;
        i < data.length && data[i][6] !== "";
        i++
    ) {
        quarterlySalesData.push([data[i][6], data[i][7]]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        quarterlySalesData.length,
        quarterlySalesData[0].length,
    );
    range.setValues(quarterlySalesData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.COLUMN)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Quarterly Sales")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Quarter" })
        .setOption("vAxis", { title: "Sales" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + quarterlySalesData.length + 2;
}

function createSalesByEntryChannelChart(sheet, data, startRow) {
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

function createSalesByClientTypeChart(sheet, data, startRow) {
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

function createSalesByMaterialCategoryChart(sheet, data, startRow) {
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
