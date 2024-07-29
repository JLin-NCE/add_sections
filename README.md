# Puppeteer Automation Script

This script uses Puppeteer to automate the process of logging into a web application, selecting a database, and updating fields with data extracted from an Excel sheet.

## Prerequisites

- Node.js installed on your machine
- NPM (Node Package Manager)
- Required dependencies: `puppeteer`, `fs`, `path`, `xlsx`, `luxon`, `string-similarity`

## Setup

1. **Install Node.js and NPM**:
   - Download and install Node.js from [nodejs.org](https://nodejs.org/).

2. **Install Dependencies**:
   Run the following command in your project directory to install the necessary packages:
   ```bash
   npm install puppeteer fs path xlsx luxon string-similarity
   ```

3. **Create Configuration File**:
   Create a `config.json` file in the root directory of your project with the following structure:
   ```json
   {
     "username": "your_username",
     "password": "your_password",
     "excelPath": "path_to_your_excel_file.xlsx",
     "dataBase": "your_database_name"
   }
   ```

4. **Prepare Excel File**:
   Ensure your Excel file is correctly formatted with at least ten columns: "Street/Lot ID", "Section ID", "Street Name/Lot Location", "Area", "Beg Location", "End Location", "Lanes", "Functional Class", "Length", "Width", "Surface Type", "Originally Constructed". The script reads data from these columns.

## Script Explanation

The script performs the following steps:

1. **Read Credentials and Excel Path**:
   Reads the login credentials, path to the Excel file, and database name from `config.json`.

2. **Read Data from Excel Sheet**:
   Reads the data from the specified Excel sheet and extracts values from the columns.

3. **Launch Puppeteer and Open Browser**:
   Launches Puppeteer with a visible browser window and maximizes it to full screen.

4. **Navigate and Login to the Website**:
   Navigates to the specified URL, enters the username and password, and clicks the login button.

5. **Select the Database**:
   Expands the "Systems Admin" menu, opens the database selection dropdown, and selects the closest matching database.

6. **Process Each Entry from Excel Data**:
   Loops through each entry in the Excel file, performs the necessary actions on the website, and updates the fields with values from the Excel file.

7. **Logout and Close Browser**:
   Logs out of the web application and closes the browser.

## Usage

To run the script, use the following command:
```bash
node index.js
```

## Error Handling

The script includes basic error handling. If an error occurs during execution, it will be logged to the console.

## Important Notes

- The script is configured to run in non-headless mode (visible browser window). For headless mode, change `headless: false` to `headless: true` in the Puppeteer launch options.
- Ensure the selectors used in the script match the actual elements on the target web page. Modify them if necessary.

## File Descriptions

### index.js
The main script file that orchestrates the entire process:
- Clears unmatched rows at the start.
- Reads configuration and Excel data.
- Logs in to the web application.
- Selects the database.
- Processes each entry from the Excel file.
- Logs timestamps and errors for each action.

### puppeteerActions.js
Contains functions to perform specific actions using Puppeteer:
- Calculating Levenshtein distance for string similarity.
- Selecting the database.
- Handling retries for clicking elements.
- Adding road names.
- Filling out section details.

### excelReader.js
A module to read data from an Excel file:
- Reads data from the specified Excel file and returns it as a JSON object.

### configReader.js
A module to read configuration details from a JSON file:
- Reads and parses the configuration file to extract necessary details.

## Conclusion

This README provides a comprehensive guide to setting up and running the Puppeteer automation script. Follow the steps carefully to ensure correct execution and modify the script as needed for your specific use case.
