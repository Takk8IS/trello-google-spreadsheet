# Trello-Google Spreadsheet Integration

![Trello-Google Spreadsheet Integration](./assets/screenshot-01.png?raw=true)
![Trello-Google Spreadsheet Integration](./assets/screenshot-02.png?raw=true)
![Trello-Google Spreadsheet Integration](./assets/screenshot-03.png?raw=true)
![Trello-Google Spreadsheet Integration](./assets/screenshot-04.png?raw=true)

## Overview

This project provides a robust integration between Trello and Google Spreadsheets, allowing users to synchronise their Trello board data with a Google Spreadsheet and generate insightful charts. It's designed to enhance project management and data visualisation capabilities for Trello users.

## Features

-   **Data Synchronisation**: Automatically fetch and update Trello board data in a Google Spreadsheet.
-   **Custom Field Support**: Capture and display custom fields from Trello cards.
-   **Chart Generation**: Create various charts based on Trello data for visual analysis.
-   **Configurable Settings**: Easy-to-use interface for managing API keys and other settings.
-   **Responsive Design**: Material Design 3 compliant user interface for settings management.

## Prerequisites

Before you begin, ensure you have the following:

-   A Google account with access to Google Sheets and Google Apps Script
-   A Trello account with API access enabled
-   Trello API Key and Token
-   The ID of the Trello board you wish to integrate

## Installation

1. Create a new Google Spreadsheet.
2. Open the Script Editor (Tools > Script editor).
3. Copy the contents of `Code.gs` and `Charts.gs` into separate files in the Script Editor.
4. Create a new HTML file named `index.html` and copy the provided HTML content into it.
5. Save all files.

## Configuration

1. In the Google Spreadsheet, refresh the page. You should see a new menu item "Trello Integration".
2. Click on "Trello Integration" > "Configure Settings".
3. Enter your Trello API Key, API Token, Board ID, and the current Spreadsheet ID.

Retrieved via URL "https://trello.com/power-ups/admin"

-   TRELLO_API_KEY = "{TRELLO_API_KEY}";

Retrieved via URL "https://trello.com/power-ups/admin"

-   TRELLO_SECRET_KEY = "{TRELLO_SECRET_KEY}";

Retrieved via URL "https://trello.com/1/authorize?expiration=never&scope=read,write,account&response_type=token&key={TRELLO_API_KEY}"

-   TRELLO_API_TOKEN = "{TRELLO_API_TOKEN}";

Retrieved via URL "https://trello.com/b/{TRELLO_BOARD_ID}/card-name"

-   TRELLO_BOARD_ID = "{TRELLO_BOARD_ID}";

Retrieved via URL "https://docs.google.com/spreadsheets/d/{SHEET_ID}"

-   SHEET_ID = "{SHEET_ID}"

4. Click "Save Settings" to store your configuration.

## Usage

After configuration, you can use the following features:

1. **Sync with Trello**: Click "Trello Integration" > "Sync with Trello" to fetch the latest data from your Trello board.
2. **Generate Charts**: After syncing, click "Trello Integration" > "Generate Charts" to create visual representations of your data.

## File Structure

-   `Code.gs`: Main script file containing core functionality.
-   `Charts.gs`: Script file for chart generation functions.
-   `index.html`: HTML file for the settings dialogue.

## Acknowledgments

-   Trello API documentation
-   Google Apps Script documentation
-   Material Design 3 guidelines

## License

This project is licensed under the Attribution 4.0 International License.

## Support

If you need help with this project, please contact via email at say@takk.ag.

## Donations

If this script has been helpful for you, consider making a donation to support our work:

-   $USDT (TRC-20): TGpiWetnYK2VQpxNGPR27D9vfM6Mei5vNA

Your donations help us continue developing useful and innovative tools.

## Takk™ Innovate Studio

Leading the Digital Revolution as the Pioneering 100% Artificial Intelligence Team.

-   Copyright (c)
-   License: Attribution 4.0 International (CC BY 4.0)
-   Author: David C Cavalcante
-   Email: say@takk.ag
-   LinkedIn: https://www.linkedin.com/in/hellodav/
-   Medium: https://medium.com/@davcavalcante/
-   Positive results, rapid innovation
-   URL: https://takk.ag/
-   X: https://twitter.com/takk8is/
-   Medium: https://takk8is.medium.com/
