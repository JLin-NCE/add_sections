const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { readConfig } = require('./configReader');
const { readExcelData } = require('./excelReader');
const { selectDatabase, retryClick } = require('./puppeteerActions');

// Function to preprocess a string value
function preprocessString(value) {
  const firstDashIndex = value.indexOf('-');
  if (firstDashIndex === -1) {
    return value.trim(); // No dash found
  }
  
  const secondDashIndex = value.indexOf('-', firstDashIndex + 1);
  if (secondDashIndex === -1) {
    return value.trim(); // Only one dash found
  }
  
  return value.substring(0, secondDashIndex).trim(); // Erase all characters after the second dash
}

// Function to preprocess a numeric value
function preprocessNumber(value) {
  const parsedValue = parseFloat(value);
  if (isNaN(parsedValue)) {
    return "0.00"; // Default to 0.00 if value is not a valid number
  }
  return parsedValue.toFixed(2);
}

// Function to clear the unmatched rows file
function clearUnmatchedRows() {
  const filePath = './unmatched_rows.xlsx';
  if (fs.existsSync(filePath)) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
      ['StreetLotId', 'SectionId', 'StreetNameLotLocation', 'Area']
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Unmatched Rows');
    XLSX.writeFile(workbook, filePath);
    console.log('Cleared unmatched rows file.');
  }
}

(async () => {
  try {
    // Clear unmatched rows file
    clearUnmatchedRows();

    // Read the credentials and Excel path from config.json
    const configPath = path.join(__dirname, 'config.json');
    const { username, password, excelPath, dataBase } = readConfig(configPath);

    // Read data from the Excel sheet
    const data = readExcelData(excelPath);

    // Launch the browser and open a new page
    const browser = await puppeteer.launch({
      headless: false, // Set headless: true for headless mode
      defaultViewport: null // Disable the default viewport
    });
    const page = await browser.newPage();

    // Open the page in full screen before logging in
    await page.setViewport({ width: 1920, height: 1080 });

    // Navigate to the login page and login
    console.log('Navigating to the login page...');
    await page.goto('https://demo.streetsaver.com/Forms/PavementSections/Section?linkid=linkSection');

    console.log('Entering username and password...');
    await page.type('#Email', username);
    await page.type('#Password', password);

    console.log('Clicking the login button...');
    await page.click('#ContentPlaceHolder1_btnLogin');
    await page.waitForNavigation();
    console.log("Waiting 8 seconds for rendering...");
    await new Promise(r => setTimeout(r, 8000));

    // Select the database
    console.log('Selecting database...');
    await selectDatabase(page, dataBase);

    console.log('Database selected. Ready to add sections.');

    // Loop through each entry in the Excel data
    for (const row of data) {
      const streetLotId = row['Street/Lot ID'];
      const sectionId = preprocessString(row['Section ID']);
      const streetNameLotLocation = preprocessString(row['Street Name/Lot Location']);
      const area = preprocessNumber(row['Area']);

      console.log(`Processing Street/Lot ID: ${streetLotId}, Section ID: ${sectionId}, Area: ${area}`);

      try {
        // Step one: click on "Add Record" button
        const addRecordSelector = '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl00_lbtnAddRecord';
        console.log('Waiting for "Add Record" button to be visible...');
        await page.waitForSelector(addRecordSelector, { visible: true });

        console.log('Clicking on "Add Record" button...');
        await page.evaluate((selector) => {
          document.querySelector(selector).click();
        }, addRecordSelector);
        console.log('"Add Record" button clicked.');
        await new Promise(r => setTimeout(r, 2000)); // Wait for the form to be ready

        // Step two: type streetNameLotLocation into the "Road Name" field
        const roadNameSelector = '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_TB_RoadName';
        console.log(`Typing ${streetNameLotLocation} into the "Road Name" field (${roadNameSelector})...`);
        await page.waitForSelector(roadNameSelector, { visible: true });
        await page.type(roadNameSelector, streetNameLotLocation);
        await new Promise(r => setTimeout(r, 500));
        console.log(`Typed ${streetNameLotLocation} into the "Road Name" field.`);

        // Extract all characters before last dash in the third column for Street/Lot Number
        console.log(`Typing ${streetLotId} into the "Street/Lot Number" field...`);
        await page.type('#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_TB_RoadNumber', streetLotId);
        await new Promise(r => setTimeout(r, 500));
        console.log(`Typed ${streetLotId} into the "Street/Lot Number" field.`);

        console.log('Form filled out.');

        // Step four: click on "Save" button to save
        console.log('Clicking on "Save" button...');
        await retryClick(page, '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_PerformInsertButton');
        await new Promise(r => setTimeout(r, 2000)); // Wait for 2 seconds to ensure the save operation completes
        console.log('"Save" button clicked.');
      } catch (error) {
        console.error(`Error processing Street/Lot ID ${streetLotId}:`, error);
      }

      // Brief pause before processing the next entry
      await new Promise(r => setTimeout(r, 500));
    }

    console.log('Script completed successfully for all entries.');
    await browser.close();
  } catch (error) {
    console.error('Error:', error);
  }
})();
