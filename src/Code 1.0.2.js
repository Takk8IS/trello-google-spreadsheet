// ```Code.gs

var CACHE = CacheService.getScriptCache();
var PROPERTIES = PropertiesService.getScriptProperties();

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Trello Integration")
        .addItem("Configure Settings", "showSettingsDialog")
        .addItem("Sync with Trello", "syncWithTrello")
        .addItem("Analyze Sales", "analyzeSales")
        .addItem("Generate Charts", "generateChartsWrapper")
        .addItem("Generate Reports", "generateReports")
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
    var otherCharges = calculateOtherCharges(salesData, headers);
    var averageTicket = calculateAverageTicket(salesData, headers);
    var newProjects = calculateNewProjects(salesData, headers);

    // Canal de Entrada
    var salesByEntryChannel = calculateSalesByEntryChannel(salesData, headers);
    var projectsByEntryChannel = calculateProjectsByEntryChannel(
        salesData,
        headers,
    );
    var proactivityDegree = calculateProactivityDegree(salesData, headers);

    // Tipología de Clientes
    var salesByClientType = calculateSalesByClientType(salesData, headers);
    var projectsByClientType = calculateProjectsByClientType(
        salesData,
        headers,
    );

    // Categoría de Material
    var salesByMaterialCategory = calculateSalesByMaterialCategory(
        salesData,
        headers,
    );
    var projectsByMaterialCategory = calculateProjectsByMaterialCategory(
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
        otherCharges,
        averageTicket,
        newProjects,
        salesByEntryChannel,
        projectsByEntryChannel,
        proactivityDegree,
        salesByClientType,
        projectsByClientType,
        salesByMaterialCategory,
        projectsByMaterialCategory,
        productiveHours,
    );

    SpreadsheetApp.getUi().alert("Sales analysis completed successfully!");
}

function calculateTotalSales(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    return salesData.reduce(
        (total, row) => total + (parseFloat(row[salesAmountIndex]) || 0),
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
            (monthlySales[monthYear] || 0) +
            (parseFloat(row[salesAmountIndex]) || 0);
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
            (yearlySales[year] || 0) + (parseFloat(row[salesAmountIndex]) || 0);
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
            (quarterlySales[key] || 0) +
            (parseFloat(row[salesAmountIndex]) || 0);
    });

    return quarterlySales;
}

function calculateOtherCharges(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var materialCategoryIndex = headers.indexOf("Material Category");
    var dateIndex = headers.indexOf("Creation Date");
    var otherCharges = { total: 0, monthly: {}, quarterly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor(month / 3) + 1;
        var salesAmount = parseFloat(row[salesAmountIndex]) || 0;
        var materialCategory = row[materialCategoryIndex];

        if (materialCategory === "Otros cobros") {
            otherCharges.total += salesAmount;

            var monthKey = `${year}-${month}`;
            otherCharges.monthly[monthKey] =
                (otherCharges.monthly[monthKey] || 0) + salesAmount;

            var quarterKey = `${year}-Q${quarter}`;
            otherCharges.quarterly[quarterKey] =
                (otherCharges.quarterly[quarterKey] || 0) + salesAmount;
        }
    });

    return otherCharges;
}

function calculateAverageTicket(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var materialCategoryIndex = headers.indexOf("Material Category");
    var totalSales = 0;
    var projectCount = 0;
    var ticketByMaterial = {};

    salesData.forEach((row) => {
        var salesAmount = parseFloat(row[salesAmountIndex]) || 0;
        var materialCategory = row[materialCategoryIndex];

        totalSales += salesAmount;
        projectCount++;

        if (!ticketByMaterial[materialCategory]) {
            ticketByMaterial[materialCategory] = { total: 0, count: 0 };
        }
        ticketByMaterial[materialCategory].total += salesAmount;
        ticketByMaterial[materialCategory].count++;
    });

    var averageTicket = totalSales / projectCount;

    for (var category in ticketByMaterial) {
        ticketByMaterial[category].average =
            ticketByMaterial[category].total / ticketByMaterial[category].count;
    }

    return { overall: averageTicket, byMaterial: ticketByMaterial };
}

function calculateNewProjects(salesData, headers) {
    var dateIndex = headers.indexOf("Creation Date");
    var newProjects = { monthly: {}, yearly: {}, quarterly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor(month / 3) + 1;

        var monthKey = `${year}-${month}`;
        var yearKey = `${year}`;
        var quarterKey = `${year}-Q${quarter}`;

        newProjects.monthly[monthKey] =
            (newProjects.monthly[monthKey] || 0) + 1;
        newProjects.yearly[yearKey] = (newProjects.yearly[yearKey] || 0) + 1;
        newProjects.quarterly[quarterKey] =
            (newProjects.quarterly[quarterKey] || 0) + 1;
    });

    return newProjects;
}

function calculateSalesByEntryChannel(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var entryChannelIndex = headers.indexOf("Entry Channel");
    var dateIndex = headers.indexOf("Creation Date");
    var salesByEntryChannel = {
        total: {},
        yearly: {},
        monthly: {},
        quarterly: {},
    };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor(month / 3) + 1;
        var channel = row[entryChannelIndex];
        var salesAmount = parseFloat(row[salesAmountIndex]) || 0;

        // Total sales by channel
        salesByEntryChannel.total[channel] =
            (salesByEntryChannel.total[channel] || 0) + salesAmount;

        // Yearly sales by channel
        if (!salesByEntryChannel.yearly[year])
            salesByEntryChannel.yearly[year] = {};
        salesByEntryChannel.yearly[year][channel] =
            (salesByEntryChannel.yearly[year][channel] || 0) + salesAmount;

        // Monthly sales by channel
        var monthKey = `${year}-${month}`;
        if (!salesByEntryChannel.monthly[monthKey])
            salesByEntryChannel.monthly[monthKey] = {};
        salesByEntryChannel.monthly[monthKey][channel] =
            (salesByEntryChannel.monthly[monthKey][channel] || 0) + salesAmount;

        // Quarterly sales by channel
        var quarterKey = `${year}-Q${quarter}`;
        if (!salesByEntryChannel.quarterly[quarterKey])
            salesByEntryChannel.quarterly[quarterKey] = {};
        salesByEntryChannel.quarterly[quarterKey][channel] =
            (salesByEntryChannel.quarterly[quarterKey][channel] || 0) +
            salesAmount;
    });

    return salesByEntryChannel;
}

function calculateProjectsByEntryChannel(salesData, headers) {
    var entryChannelIndex = headers.indexOf("Entry Channel");
    var dateIndex = headers.indexOf("Creation Date");
    var projectsByEntryChannel = { total: {}, yearly: {}, quarterly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var quarter = Math.floor((date.getMonth() + 3) / 3);
        var channel = row[entryChannelIndex];

        // Total projects by channel
        projectsByEntryChannel.total[channel] =
            (projectsByEntryChannel.total[channel] || 0) + 1;

        // Yearly projects by channel
        if (!projectsByEntryChannel.yearly[year])
            projectsByEntryChannel.yearly[year] = {};
        projectsByEntryChannel.yearly[year][channel] =
            (projectsByEntryChannel.yearly[year][channel] || 0) + 1;

        // Quarterly projects by channel
        var quarterKey = `${year}-Q${quarter}`;
        if (!projectsByEntryChannel.quarterly[quarterKey])
            projectsByEntryChannel.quarterly[quarterKey] = {};
        projectsByEntryChannel.quarterly[quarterKey][channel] =
            (projectsByEntryChannel.quarterly[quarterKey][channel] || 0) + 1;
    });

    return projectsByEntryChannel;
}

function calculateProactivityDegree(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var entryChannelIndex = headers.indexOf("Entry Channel");
    var dateIndex = headers.indexOf("Creation Date");
    var proactiveChannels = ["BNI", "RRSS", "feria", "flyer", "brevo"];
    var proactivityDegree = { yearly: {}, monthly: {}, quarterly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor((month + 2) / 3);
        var channel = row[entryChannelIndex];
        var salesAmount = parseFloat(row[salesAmountIndex]) || 0;

        var yearKey = year.toString();
        var monthKey = `${year}-${month}`;
        var quarterKey = `${year}-Q${quarter}`;

        // Initialize if not exist
        if (!proactivityDegree.yearly[yearKey])
            proactivityDegree.yearly[yearKey] = { proactive: 0, total: 0 };
        if (!proactivityDegree.monthly[monthKey])
            proactivityDegree.monthly[monthKey] = { proactive: 0, total: 0 };
        if (!proactivityDegree.quarterly[quarterKey])
            proactivityDegree.quarterly[quarterKey] = {
                proactive: 0,
                total: 0,
            };

        // Add to total
        proactivityDegree.yearly[yearKey].total += salesAmount;
        proactivityDegree.monthly[monthKey].total += salesAmount;
        proactivityDegree.quarterly[quarterKey].total += salesAmount;

        // Add to proactive if channel is proactive
        if (proactiveChannels.includes(channel)) {
            proactivityDegree.yearly[yearKey].proactive += salesAmount;
            proactivityDegree.monthly[monthKey].proactive += salesAmount;
            proactivityDegree.quarterly[quarterKey].proactive += salesAmount;
        }
    });

    // Calculate percentages
    for (let period in proactivityDegree) {
        for (let key in proactivityDegree[period]) {
            let data = proactivityDegree[period][key];
            data.percentage =
                data.total > 0 ? (data.proactive / data.total) * 100 : 0;
        }
    }

    return proactivityDegree;
}

function calculateSalesByClientType(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var clientTypeIndex = headers.indexOf("Client Type");
    var dateIndex = headers.indexOf("Creation Date");
    var salesByClientType = {
        total: {},
        yearly: {},
        monthly: {},
        quarterly: {},
    };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor((month + 2) / 3);
        var clientType = row[clientTypeIndex];
        var salesAmount = parseFloat(row[salesAmountIndex]) || 0;

        var yearKey = year.toString();
        var monthKey = `${year}-${month}`;
        var quarterKey = `${year}-Q${quarter}`;

        // Total sales by client type
        salesByClientType.total[clientType] =
            (salesByClientType.total[clientType] || 0) + salesAmount;

        // Yearly sales by client type
        if (!salesByClientType.yearly[yearKey])
            salesByClientType.yearly[yearKey] = {};
        salesByClientType.yearly[yearKey][clientType] =
            (salesByClientType.yearly[yearKey][clientType] || 0) + salesAmount;

        // Monthly sales by client type
        if (!salesByClientType.monthly[monthKey])
            salesByClientType.monthly[monthKey] = {};
        salesByClientType.monthly[monthKey][clientType] =
            (salesByClientType.monthly[monthKey][clientType] || 0) +
            salesAmount;

        // Quarterly sales by client type
        if (!salesByClientType.quarterly[quarterKey])
            salesByClientType.quarterly[quarterKey] = {};
        salesByClientType.quarterly[quarterKey][clientType] =
            (salesByClientType.quarterly[quarterKey][clientType] || 0) +
            salesAmount;
    });

    return salesByClientType;
}

function calculateProjectsByClientType(salesData, headers) {
    var clientTypeIndex = headers.indexOf("Client Type");
    var dateIndex = headers.indexOf("Creation Date");
    var projectsByClientType = { total: {}, yearly: {}, quarterly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var quarter = Math.floor((date.getMonth() + 3) / 3);
        var clientType = row[clientTypeIndex];

        var yearKey = year.toString();
        var quarterKey = `${year}-Q${quarter}`;

        // Total projects by client type
        projectsByClientType.total[clientType] =
            (projectsByClientType.total[clientType] || 0) + 1;

        // Yearly projects by client type
        if (!projectsByClientType.yearly[yearKey])
            projectsByClientType.yearly[yearKey] = {};
        projectsByClientType.yearly[yearKey][clientType] =
            (projectsByClientType.yearly[yearKey][clientType] || 0) + 1;

        // Quarterly projects by client type
        if (!projectsByClientType.quarterly[quarterKey])
            projectsByClientType.quarterly[quarterKey] = {};
        projectsByClientType.quarterly[quarterKey][clientType] =
            (projectsByClientType.quarterly[quarterKey][clientType] || 0) + 1;
    });

    return projectsByClientType;
}

function calculateSalesByMaterialCategory(salesData, headers) {
    var salesAmountIndex = headers.indexOf("Sales Amount");
    var materialCategoryIndex = headers.indexOf("Material Category");
    var dateIndex = headers.indexOf("Creation Date");
    var salesByMaterialCategory = {
        total: {},
        yearly: {},
        monthly: {},
        quarterly: {},
    };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor((month + 2) / 3);
        var category = row[materialCategoryIndex];
        var salesAmount = parseFloat(row[salesAmountIndex]) || 0;

        var yearKey = year.toString();
        var monthKey = `${year}-${month}`;
        var quarterKey = `${year}-Q${quarter}`;

        // Total sales by material category
        salesByMaterialCategory.total[category] =
            (salesByMaterialCategory.total[category] || 0) + salesAmount;

        // Yearly sales by material category
        if (!salesByMaterialCategory.yearly[yearKey])
            salesByMaterialCategory.yearly[yearKey] = {};
        salesByMaterialCategory.yearly[yearKey][category] =
            (salesByMaterialCategory.yearly[yearKey][category] || 0) +
            salesAmount;

        // Monthly sales by material category
        if (!salesByMaterialCategory.monthly[monthKey])
            salesByMaterialCategory.monthly[monthKey] = {};
        salesByMaterialCategory.monthly[monthKey][category] =
            (salesByMaterialCategory.monthly[monthKey][category] || 0) +
            salesAmount;

        // Quarterly sales by material category
        if (!salesByMaterialCategory.quarterly[quarterKey])
            salesByMaterialCategory.quarterly[quarterKey] = {};
        salesByMaterialCategory.quarterly[quarterKey][category] =
            (salesByMaterialCategory.quarterly[quarterKey][category] || 0) +
            salesAmount;
    });

    return salesByMaterialCategory;
}

function calculateProjectsByMaterialCategory(salesData, headers) {
    var materialCategoryIndex = headers.indexOf("Material Category");
    var dateIndex = headers.indexOf("Creation Date");
    var projectsByMaterialCategory = { total: {}, yearly: {}, quarterly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var quarter = Math.floor((date.getMonth() + 3) / 3);
        var category = row[materialCategoryIndex];

        var yearKey = year.toString();
        var quarterKey = `${year}-Q${quarter}`;

        // Total projects by material category
        projectsByMaterialCategory.total[category] =
            (projectsByMaterialCategory.total[category] || 0) + 1;

        // Yearly projects by material category
        if (!projectsByMaterialCategory.yearly[yearKey])
            projectsByMaterialCategory.yearly[yearKey] = {};
        projectsByMaterialCategory.yearly[yearKey][category] =
            (projectsByMaterialCategory.yearly[yearKey][category] || 0) + 1;

        // Quarterly projects by material category
        if (!projectsByMaterialCategory.quarterly[quarterKey])
            projectsByMaterialCategory.quarterly[quarterKey] = {};
        projectsByMaterialCategory.quarterly[quarterKey][category] =
            (projectsByMaterialCategory.quarterly[quarterKey][category] || 0) +
            1;
    });

    return projectsByMaterialCategory;
}

function calculateProductiveHours(salesData, headers) {
    var productiveHoursIndex = headers.indexOf("Productive Hours");
    var dateIndex = headers.indexOf("Creation Date");
    var productiveHours = { monthly: {}, quarterly: {}, yearly: {} };

    salesData.forEach((row) => {
        var date = new Date(row[dateIndex]);
        var year = date.getFullYear();
        var month = date.getMonth() + 1;
        var quarter = Math.floor((month + 2) / 3);
        var hours = parseFloat(row[productiveHoursIndex]) || 0;

        var yearKey = year.toString();
        var monthKey = `${year}-${month}`;
        var quarterKey = `${year}-Q${quarter}`;

        productiveHours.monthly[monthKey] =
            (productiveHours.monthly[monthKey] || 0) + hours;
        productiveHours.quarterly[quarterKey] =
            (productiveHours.quarterly[quarterKey] || 0) + hours;
        productiveHours.yearly[yearKey] =
            (productiveHours.yearly[yearKey] || 0) + hours;
    });

    return productiveHours;
}

function writeSalesAnalysis(
    sheet,
    totalSales,
    monthlySales,
    yearlySales,
    quarterlySales,
    otherCharges,
    averageTicket,
    newProjects,
    salesByEntryChannel,
    projectsByEntryChannel,
    proactivityDegree,
    salesByClientType,
    projectsByClientType,
    salesByMaterialCategory,
    projectsByMaterialCategory,
    productiveHours,
) {
    sheet.clear();
    sheet.getRange("A1").setValue("Sales Analysis Report");
    sheet.getRange("A1").setFontWeight("bold").setFontSize(14);

    var row = 3;

    // Ventas Totales
    sheet.getRange(`A${row}`).setValue("Total Sales:");
    sheet.getRange(`B${row}`).setValue(totalSales);
    row += 2;

    // Monthly Sales
    sheet.getRange(`A${row}`).setValue("Monthly Sales:");
    row++;
    for (var month in monthlySales) {
        sheet.getRange(`A${row}`).setValue(month);
        sheet.getRange(`B${row}`).setValue(monthlySales[month]);
        row++;
    }
    row++;

    // Yearly Sales
    sheet.getRange(`A${row}`).setValue("Yearly Sales:");
    row++;
    for (var year in yearlySales) {
        sheet.getRange(`A${row}`).setValue(year);
        sheet.getRange(`B${row}`).setValue(yearlySales[year]);
        row++;
    }
    row++;

    // Quarterly Sales
    sheet.getRange(`A${row}`).setValue("Quarterly Sales:");
    row++;
    for (var quarter in quarterlySales) {
        sheet.getRange(`A${row}`).setValue(quarter);
        sheet.getRange(`B${row}`).setValue(quarterlySales[quarter]);
        row++;
    }
    row++;

    // Other Charges
    sheet.getRange(`A${row}`).setValue("Other Charges:");
    sheet.getRange(`B${row}`).setValue(otherCharges.total);
    row++;
    sheet.getRange(`A${row}`).setValue("Other Charges (Monthly):");
    row++;
    for (var month in otherCharges.monthly) {
        sheet.getRange(`A${row}`).setValue(month);
        sheet.getRange(`B${row}`).setValue(otherCharges.monthly[month]);
        row++;
    }
    row++;

    // Average Ticket
    sheet.getRange(`A${row}`).setValue("Average Ticket:");
    sheet.getRange(`B${row}`).setValue(averageTicket.overall);
    row++;
    sheet.getRange(`A${row}`).setValue("Average Ticket by Material Category:");
    row++;
    for (var category in averageTicket.byMaterial) {
        sheet.getRange(`A${row}`).setValue(category);
        sheet
            .getRange(`B${row}`)
            .setValue(averageTicket.byMaterial[category].average);
        row++;
    }
    row++;

    // New Projects
    sheet.getRange(`A${row}`).setValue("New Projects (Monthly):");
    row++;
    for (var month in newProjects.monthly) {
        sheet.getRange(`A${row}`).setValue(month);
        sheet.getRange(`B${row}`).setValue(newProjects.monthly[month]);
        row++;
    }
    row++;

    // Sales by Entry Channel
    sheet.getRange(`A${row}`).setValue("Sales by Entry Channel:");
    row++;
    for (var channel in salesByEntryChannel.total) {
        sheet.getRange(`A${row}`).setValue(channel);
        sheet.getRange(`B${row}`).setValue(salesByEntryChannel.total[channel]);
        row++;
    }
    row++;

    // Projects by Entry Channel
    sheet.getRange(`A${row}`).setValue("Projects by Entry Channel:");
    row++;
    for (var channel in projectsByEntryChannel.total) {
        sheet.getRange(`A${row}`).setValue(channel);
        sheet
            .getRange(`B${row}`)
            .setValue(projectsByEntryChannel.total[channel]);
        row++;
    }
    row++;

    // Proactivity Degree
    sheet.getRange(`A${row}`).setValue("Proactivity Degree (Yearly):");
    row++;
    for (var year in proactivityDegree.yearly) {
        sheet.getRange(`A${row}`).setValue(year);
        sheet
            .getRange(`B${row}`)
            .setValue(proactivityDegree.yearly[year].percentage + "%");
        row++;
    }
    row++;

    // Sales by Client Type
    sheet.getRange(`A${row}`).setValue("Sales by Client Type:");
    row++;
    for (var clientType in salesByClientType.total) {
        sheet.getRange(`A${row}`).setValue(clientType);
        sheet.getRange(`B${row}`).setValue(salesByClientType.total[clientType]);
        row++;
    }
    row++;

    // Projects by Client Type
    sheet.getRange(`A${row}`).setValue("Projects by Client Type:");
    row++;
    for (var clientType in projectsByClientType.total) {
        sheet.getRange(`A${row}`).setValue(clientType);
        sheet
            .getRange(`B${row}`)
            .setValue(projectsByClientType.total[clientType]);
        row++;
    }
    row++;

    // Sales by Material Category
    sheet.getRange(`A${row}`).setValue("Sales by Material Category:");
    row++;
    for (var category in salesByMaterialCategory.total) {
        sheet.getRange(`A${row}`).setValue(category);
        sheet
            .getRange(`B${row}`)
            .setValue(salesByMaterialCategory.total[category]);
        row++;
    }
    row++;

    // Projects by Material Category
    sheet.getRange(`A${row}`).setValue("Projects by Material Category:");
    row++;
    for (var category in projectsByMaterialCategory.total) {
        sheet.getRange(`A${row}`).setValue(category);
        sheet
            .getRange(`B${row}`)
            .setValue(projectsByMaterialCategory.total[category]);
        row++;
    }
    row++;

    // Productive Hours
    sheet.getRange(`A${row}`).setValue("Productive Hours (Monthly):");
    row++;
    for (var month in productiveHours.monthly) {
        sheet.getRange(`A${row}`).setValue(month);
        sheet.getRange(`B${row}`).setValue(productiveHours.monthly[month]);
        row++;
    }

    // Apply formatting
    sheet.getRange(1, 1, row, 2).setNumberFormat("#,##0.00");
    sheet.getRange(1, 1, row, 2).applyRowBanding();
    sheet.autoResizeColumns(1, 2);
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
    chartRow = createProactivityDegreeChart(
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

    for (var i = yearlyIndex + 1; i < data.length && data[i][0] !== ""; i++) {
        years.push(data[i][0]);
        sales.push(data[i][1]);
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
    var quarterlyIndex = data.findIndex((row) => row[0] === "Quarterly Sales:");
    if (quarterlyIndex === -1) return startRow;

    for (
        var i = quarterlyIndex + 1;
        i < data.length && data[i][0] !== "";
        i++
    ) {
        quarterlySalesData.push([data[i][0], data[i][1]]);
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
        (row) => row[0] === "Sales by Client Type:",
    );
    if (clientTypeIndex === -1) return startRow;

    for (
        var i = clientTypeIndex + 1;
        i < data.length && data[i][0] !== "";
        i++
    ) {
        clientTypeData.push([data[i][0], data[i][1]]);
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
        (row) => row[0] === "Sales by Material Category:",
    );
    if (materialCategoryIndex === -1) return startRow;

    for (
        var i = materialCategoryIndex + 1;
        i < data.length && data[i][0] !== "";
        i++
    ) {
        materialCategoryData.push([data[i][0], data[i][1]]);
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
        (row) => row[0] === "Productive Hours (Monthly):",
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

function createProactivityDegreeChart(sheet, data, startRow) {
    var proactivityData = [["Year", "Proactivity Degree"]];
    var proactivityIndex = data.findIndex(
        (row) => row[0] === "Proactivity Degree (Yearly):",
    );
    if (proactivityIndex === -1) return startRow;

    for (
        var i = proactivityIndex + 1;
        i < data.length && data[i][0] !== "";
        i++
    ) {
        proactivityData.push([data[i][0], parseFloat(data[i][1])]);
    }

    var range = sheet.getRange(
        startRow,
        1,
        proactivityData.length,
        proactivityData[0].length,
    );
    range.setValues(proactivityData);

    var chart = sheet
        .newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(range)
        .setPosition(startRow, 5, 0, 0)
        .setOption("title", "Proactivity Degree by Year")
        .setOption("legend", { position: "none" })
        .setOption("hAxis", { title: "Year" })
        .setOption("vAxis", { title: "Proactivity Degree (%)" })
        .setOption("height", 300)
        .setOption("width", 500)
        .build();

    sheet.insertChart(chart);

    return startRow + proactivityData.length + 2;
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
        (row) => row[0] === "Sales by Client Type:",
    );
    if (clientTypeIndex !== -1) {
        reportSheet.getRange("A" + (row + 2)).setValue("Sales by Client Type:");
        row += 3;
        for (
            var i = clientTypeIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][0] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][0]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][1]);
            row++;
        }
    }

    // Add sales by material category
    var materialCategoryIndex = salesAnalysisData.findIndex(
        (row) => row[0] === "Sales by Material Category:",
    );
    if (materialCategoryIndex !== -1) {
        reportSheet
            .getRange("A" + (row + 2))
            .setValue("Sales by Material Category:");
        row += 3;
        for (
            var i = materialCategoryIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][0] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][0]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][1]);
            row++;
        }
    }

    // Add productive hours
    var productiveHoursIndex = salesAnalysisData.findIndex(
        (row) => row[0] === "Productive Hours (Monthly):",
    );
    if (productiveHoursIndex !== -1) {
        reportSheet.getRange("A" + (row + 2)).setValue("Productive Hours:");
        row += 3;
        for (
            var i = productiveHoursIndex + 1;
            i < salesAnalysisData.length && salesAnalysisData[i][0] !== "";
            i++
        ) {
            reportSheet.getRange(row, 1).setValue(salesAnalysisData[i][0]);
            reportSheet.getRange(row, 2).setValue(salesAnalysisData[i][1]);
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

// Auxiliary functions

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
