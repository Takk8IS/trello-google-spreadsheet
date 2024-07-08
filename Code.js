// Code.gs

// Global variables
var CACHE = CacheService.getScriptCache();

// Function executed when the spreadsheet is opened
function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Trello Integration")
        .addItem("Sync with Trello", "syncWithTrello")
        .addItem("Generate Charts", "generateChartsWrapper")
        .addItem("Configure Settings", "showSettingsDialog")
        .addToUi();
}

// Function to sync data with Trello
function syncWithTrello() {
    var settings = getSettings();
    var apiKey = settings.TRELLO_API_KEY;
    var secretKey = settings.TRELLO_SECRET_KEY;
    var apiToken = settings.TRELLO_API_TOKEN;
    var boardId = settings.TRELLO_BOARD_ID;
    var url = `https://api.trello.com/1/boards/${boardId}/cards?key=${apiKey}&token=${apiToken}&customFieldItems=true`;

    try {
        var response = UrlFetchApp.fetch(url);
        var cards = JSON.parse(response.getContentText());
        var sheet = SpreadsheetApp.openById(settings.SHEET_ID);
        var trelloSheet = sheet.getSheetByName("Trello Data");

        if (!trelloSheet) {
            trelloSheet = sheet.insertSheet("Trello Data");
        } else {
            trelloSheet.clear();
        }

        // Get custom field names
        var customFieldsUrl = `https://api.trello.com/1/boards/${boardId}/customFields?key=${apiKey}&token=${apiToken}`;
        var customFieldsResponse = UrlFetchApp.fetch(customFieldsUrl);
        var customFields = JSON.parse(customFieldsResponse.getContentText());

        var headers = [
            "ID",
            "Name",
            "List",
            "Creation Date",
            "Due Date",
            "Labels",
        ];
        customFields.forEach(function (field) {
            headers.push(field.name);
        });
        trelloSheet.appendRow(headers);

        // Get list names
        var listsUrl = `https://api.trello.com/1/boards/${boardId}/lists?key=${apiKey}&token=${apiToken}`;
        var listsResponse = UrlFetchApp.fetch(listsUrl);
        var lists = JSON.parse(listsResponse.getContentText());
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
        SpreadsheetApp.getUi().alert(
            "Sync with Trello completed successfully!",
        );
    } catch (error) {
        SpreadsheetApp.getUi().alert(
            "Error syncing with Trello: " + error.toString(),
        );
    }
}

// Helper function to get custom field values
function getCustomFieldValue(customFields, fieldId) {
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

// Function to apply formatting to the sheet
function applySheetFormatting(sheet) {
    var range = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());

    var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
    headerRange.setBackground("#007aff");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    headerRange.setFontFamily("Arial");

    range.setFontSize(10);
    range.setFontFamily("Arial");
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

    // Automatically adjust column widths
    for (var i = 1; i <= sheet.getLastColumn(); i++) {
        sheet.autoResizeColumn(i);
    }

    // Add filters to the header row
    headerRange.createFilter();
}

// Wrapper function to generate charts
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

// Function to show settings dialog
function showSettingsDialog() {
    var html = HtmlService.createHtmlOutputFromFile("index")
        .setWidth(400)
        .setHeight(300);
    SpreadsheetApp.getUi().showModalDialog(
        html,
        "Configure Trello Integration Settings",
    );
}

// Function to save settings
function saveSettings(settings) {
    for (var key in settings) {
        // Cache for 6 hours
        CACHE.put(key, settings[key], 21600);
    }
    return { success: true, message: "Settings saved successfully!" };
}

// Function to get settings
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
        settings[key] = CACHE.get(key) || "";
    });
    return settings;
}

// Login to Trello (placeholder function)
function loginToTrello(username, password) {
    // We need implement proper authentication here
    return { success: true };
}
