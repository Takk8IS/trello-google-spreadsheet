// Charts.gs

// Function to generate charts
function generateCharts(trelloSheet) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var chartsSheet = sheet.getSheetByName("Charts");
    if (!chartsSheet) {
        chartsSheet = sheet.insertSheet("Charts");
    } else {
        chartsSheet.clear();
    }

    var dataRange = trelloSheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[0];
    var data = values.slice(1);

    // Extract relevant data for charts
    var chartData = {
        names: [],
        values: {},
        categories: {},
    };

    // Find numeric and categorical columns
    var numericColumns = [];
    var categoricalColumns = [];
    headers.forEach((header, index) => {
        if (data.some((row) => typeof row[index] === "number")) {
            numericColumns.push(index);
        } else if (
            data.some(
                (row) => typeof row[index] === "string" && row[index] !== "",
            )
        ) {
            categoricalColumns.push(index);
        }
    });

    // Populate chartData
    data.forEach(function (row) {
        // Assuming the name is in the second column
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
    });

    // Create charts
    var chartRow = 1;
    if (Object.keys(chartData.values).length > 0) {
        createChart(
            chartsSheet,
            chartData,
            "COLUMN",
            "Task Distribution by List",
            chartRow,
        );
        chartRow += 20;
        createChart(
            chartsSheet,
            chartData,
            "BAR",
            "Task Count by Label",
            chartRow,
        );
        chartRow += 20;
        createChart(
            chartsSheet,
            chartData,
            "LINE",
            "Tasks Over Time",
            chartRow,
        );
        chartRow += 20;
        createChart(
            chartsSheet,
            chartData,
            "PIE",
            "Task Distribution by Custom Field",
            chartRow,
        );
        chartRow += 20;
    }

    if (Object.keys(chartData.categories).length > 0) {
        createCategoricalChart(
            chartsSheet,
            chartData,
            "Task Status Distribution",
            chartRow,
        );
    }
}

// Function to create a chart
function createChart(sheet, chartData, chartType, title, startRow) {
    var data = [["Names"].concat(chartData.names)];
    for (var key in chartData.values) {
        data.push([key].concat(chartData.values[key]));
    }
    var range = sheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType[chartType])
        .addRange(range)
        .setPosition(startRow, 1, 0, 0)
        .setOption("title", title)
        .setOption("legend", { position: "top" })
        .setOption("height", 300)
        .setOption("width", 600)
        .build();
    sheet.insertChart(chart);
}

// Function to create a categorical chart
function createCategoricalChart(sheet, chartData, title, startRow) {
    // Use the first categorical column
    var categories = Object.keys(chartData.categories)[0];
    var data = [["Category", "Count"]];
    var categoryCount = {};

    chartData.categories[categories].forEach((category) => {
        if (category in categoryCount) {
            categoryCount[category]++;
        } else {
            categoryCount[category] = 1;
        }
    });

    for (var category in categoryCount) {
        data.push([category, categoryCount[category]]);
    }

    var range = sheet.getRange(startRow, 1, data.length, data[0].length);
    range.setValues(data);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(range)
        .setPosition(startRow, 1, 0, 0)
        .setOption("title", title)
        .setOption("legend", { position: "right" })
        .setOption("height", 300)
        .setOption("width", 400)
        .build();
    sheet.insertChart(chart);
}
