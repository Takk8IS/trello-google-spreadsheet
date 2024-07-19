// ```Code.gs

var CACHE = CacheService.getScriptCache();
var PROPERTIES = PropertiesService.getScriptProperties();

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Trello Integration")
        .addItem("Configure Settings", "showSettingsDialog")
        .addItem("Sync with Trello", "syncWithTrello")
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
    generateCharts(trelloSheet);
}

function getCustomFieldValue(customFields, fieldId) {
    if (!customFields) return "";

    for (var i = 0; i < customFields.length; i++) {
        var field = customFields[i];
        if (field.idCustomField === fieldId) {
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

function generateChartsWrapper() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var trelloSheet = sheet.getSheetByName("Trello Data");
    if (trelloSheet) {
        generateCharts(trelloSheet);
        SpreadsheetApp.getUi().alert("Charts generated successfully!");
    } else {
        SpreadsheetApp.getUi().alert(
            "Error: Trello Data sheet not found. Please sync with Trello first.",
        );
    }
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

function generateReports() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Trello Data");

    if (!sheet) {
        SpreadsheetApp.getUi().alert(
            "Error: Trello Data sheet not found. Please sync with Trello first.",
        );
        return;
    }

    var reportSheet = ss.getSheetByName("Report") || ss.insertSheet("Report");
    reportSheet.clear();

    // Add report title
    reportSheet
        .getRange("A1")
        .setValue("Trello Board Report")
        .setFontWeight("bold")
        .setFontSize(16);

    // Add date of report
    reportSheet
        .getRange("A2")
        .setValue("Report generated on: " + new Date().toLocaleString());

    // Add summary statistics
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);

    var totalCards = rows.length;
    var listsCount = {};
    var labelCount = {};

    rows.forEach(function (row) {
        var list = row[headers.indexOf("List")];
        var labels = row[headers.indexOf("Labels")].split(", ");

        listsCount[list] = (listsCount[list] || 0) + 1;
        labels.forEach(function (label) {
            if (label) {
                labelCount[label] = (labelCount[label] || 0) + 1;
            }
        });
    });

    reportSheet.getRange("A4").setValue("Total Cards: " + totalCards);

    var row = 6;
    reportSheet
        .getRange("A" + row)
        .setValue("Cards per List:")
        .setFontWeight("bold");
    row++;
    for (var list in listsCount) {
        reportSheet
            .getRange("A" + row)
            .setValue(list + ": " + listsCount[list]);
        row++;
    }

    row += 2;
    reportSheet
        .getRange("A" + row)
        .setValue("Cards per Label:")
        .setFontWeight("bold");
    row++;
    for (var label in labelCount) {
        reportSheet
            .getRange("A" + row)
            .setValue(label + ": " + labelCount[label]);
        row++;
    }

    reportSheet.autoResizeColumns(1, reportSheet.getLastColumn());

    // Print the report
    var printOptions = {
        landscape: false,
        showGridlines: false,
        size: SpreadsheetApp.PrintSize.FIT_TO_WIDTH,
        topMargin: 0.3,
        bottomMargin: 0.3,
        leftMargin: 0.3,
        rightMargin: 0.3,
    };

    ss.getSheetByName("Report").print(printOptions);

    SpreadsheetApp.getUi().alert(
        "Report generated and printed successfully! Check the 'Report' sheet.",
    );
}
