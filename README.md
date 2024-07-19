# Trello-Google Spreadsheet Integration

![Trello-Google Spreadsheet Integration](./assets/screenshot-01.png?raw=true)
![Trello-Google Spreadsheet Integration](./assets/screenshot-02.png?raw=true)
![Trello-Google Spreadsheet Integration](./assets/screenshot-03.png?raw=true)
![Trello-Google Spreadsheet Integration](./assets/screenshot-04.png?raw=true)

## Overview

This project provides a robust integration between Trello and Google Spreadsheets, enabling users to synchronize Trello board data with a Google Spreadsheet and generate insightful charts and reports. It's designed to enhance project management, data visualization, and sales analysis capabilities for Trello users.

## Features

-   **Data Synchronization**: Automatically fetch and update Trello board data in a Google Spreadsheet.
-   **Custom Field Support**: Capture and display custom fields from Trello cards.
-   **Sales Analysis**: Generate comprehensive sales reports and analytics.
-   **Chart Generation**: Create various charts based on Trello data for visual analysis.
-   **Report Generation**: Produce detailed reports summarizing Trello board and sales data.
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

Use the following URLs to retrieve the necessary information:

-   TRELLO_API_KEY: https://trello.com/power-ups/admin
-   TRELLO_SECRET_KEY: https://trello.com/power-ups/admin
-   TRELLO_API_TOKEN: https://trello.com/1/authorize?expiration=never&scope=read,write,account&response_type=token&key={TRELLO_API_KEY}
-   TRELLO_BOARD_ID: https://trello.com/b/{TRELLO_BOARD_ID}/card-name
-   SHEET_ID: https://docs.google.com/spreadsheets/d/{SHEET_ID}

4. Click "Save Settings" to store your configuration.

## Usage

After configuration, you can use the following features:

1. **Sync with Trello**: Click "Trello Integration" > "Sync with Trello" to fetch the latest data from your Trello board.
2. **Analyze Sales**: Click "Trello Integration" > "Analyze Sales" to generate sales analytics based on your Trello data.
3. **Generate Charts**: Click "Trello Integration" > "Generate Charts" to create visual representations of your data.
4. **Generate Reports**: Click "Trello Integration" > "Generate Reports" to produce a comprehensive report of your Trello board and sales data.

## File Structure

-   `Code.gs`: Main script file containing core functionality, including data synchronization, sales analysis, and report generation.
-   `Charts.gs`: Script file for chart generation functions.
-   `index.html`: HTML file for the settings dialogue, featuring a Material Design 3 compliant interface.

## Troubleshooting

If you encounter any issues:

1. Ensure all API keys and tokens are correctly entered in the settings.
2. Check that your Trello board has the necessary custom fields for sales data.
3. Verify that you have the required permissions for both Trello and Google Sheets.
4. If charts or reports fail to generate, ensure you've synced with Trello and analyzed sales first.

## Acknowledgments

-   Trello API documentation
-   Google Apps Script documentation
-   Material Design 3 guidelines

## License

This project is licensed under the Attribution 4.0 International License (CC BY 4.0).

## Support

If you need help with this project, please contact us via email at say@takk.ag.

## Donations

If this script has been helpful for you, consider making a donation to support our work:

-   $USDT (TRC-20): TGpiWetnYK2VQpxNGPR27D9vfM6Mei5vNA

Your donations help us continue developing useful and innovative tools.

## About Takk™ Innovate Studio

Leading the Digital Revolution as the Pioneering 100% Artificial Intelligence Team.

-   Copyright (c) Takk™ Innovate Studio
-   Author: David C Cavalcante
-   Email: say@takk.ag
-   LinkedIn: https://www.linkedin.com/in/hellodav/
-   Medium: https://medium.com/@davcavalcante/
-   Website: https://takk.ag/
-   Twitter: https://twitter.com/takk8is/
-   Medium: https://takk8is.medium.com/
