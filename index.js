const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { readConfig } = require('./configReader');
const { readExcelData } = require('./excelReader');
const { selectDatabase, retryClick } = require('./puppeteerActions');
const stringSimilarity = require('string-similarity');
const { DateTime } = require('luxon');  // Importing DateTime from luxon

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

function logTimestamp(action, streetLotId = null, sectionId = null, streetNameLotLocation = null, area = null) {
  const timestamp = DateTime.now().setZone('America/Los_Angeles').toFormat('yyyy-LL-dd HH:mm:ss');
  const logEntry = [timestamp, action, streetLotId, sectionId, streetNameLotLocation, area];

  console.log(`Timestamp: ${timestamp}, Action: ${action}`);

  const filePath = './time_log.xlsx';
  let workbook;
  let worksheet;

  if (fs.existsSync(filePath)) {
    workbook = XLSX.readFile(filePath);
    worksheet = workbook.Sheets['Log'];
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.aoa_to_sheet([
      ['Timestamp', 'Action', 'StreetLotId', 'SectionId', 'StreetNameLotLocation', 'Area']
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Log');
  }

  XLSX.utils.sheet_add_aoa(worksheet, [logEntry], { origin: -1 });
  XLSX.writeFile(workbook, filePath);
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
      const beginLocation = row['Beg Location'];
      const endLocation = row['End Location'];
      const lanes = row['Lanes'];
      const functionalClass = row['Functional Class'];
      const length = preprocessNumber(row['Length']);
      const width = preprocessNumber(row['Width']);
      const surfaceType = row['Surface Type'];
      const originallyConstructed = row['Originally Constructed'];

      console.log(`Processing Street/Lot ID: ${streetLotId}, Section ID: ${sectionId}, Area: ${area}`);

      try {
        console.log('Expanding "Pavement Sections" menu...');
        await retryClick(page, '#togglePavementSections');
        await page.waitForSelector('#pavementSections.menu-dropdown.collapse.show', { visible: true });

        console.log('Clicking on "Road Names"...');
        await retryClick(page, '#linkRdNames');
        await page.waitForSelector('#ctl00_ContentPlaceHolder1_grdEDIT_grdData', { visible: true });

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

        // Check if the error message is displayed
        try {
          await page.waitForSelector('.swal2-html-container', { visible: true, timeout: 2000 });
          console.log('Error message detected. Skipping this entry.');
          await page.click('.swal2-confirm');
          await new Promise(r => setTimeout(r, 1000)); // Wait for the alert to be handled
          console.log('Clicked OK on the alert.');

          // Click the cancel button to leave the page
          const cancelButtonSelector = '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_CancelButton';
          console.log('Clicking the cancel button...');
          await page.click(cancelButtonSelector);
          await new Promise(r => setTimeout(r, 1000)); // Wait for the action to complete
          console.log('Clicked cancel button and staying on the same entry.');
        } catch (error) {
          // No error message, proceed as usual
        }

        // Step five: click on "Pavement Sections" id=togglePavementSections, dropdown menu
        console.log('Expanding "Pavement Sections" menu again...');
        await retryClick(page, '#togglePavementSections');
        await page.waitForSelector('#pavementSections.menu-dropdown.collapse.show', { visible: true });

        // Step six: click on "Edit Section" id = linkSectionCreate
        console.log('Clicking on "Edit Section"...');
        await retryClick(page, '#linkSectionCreate');
        await new Promise(r => setTimeout(r, 1000));

        // Step seven: fill out the form
        console.log('Filling out form fields...');
        await page.type('#StreetID', streetLotId);
        await page.type('#BegLocation', beginLocation);
        await page.type('#SectionID', sectionId);
        await page.type('#EndLocation', endLocation);

        // Select the closest matching street name/lot location
        const rdNamesKeySelector = '#RdNames_Key';
        console.log(`Selecting closest match for ${streetNameLotLocation} from the dropdown...`);
        await page.waitForSelector(rdNamesKeySelector, { visible: true });

        const options = await page.evaluate((selector) => {
          const optionsList = Array.from(document.querySelectorAll(`${selector} option`));
          return optionsList.map(option => ({ value: option.value, text: option.textContent }));
        }, rdNamesKeySelector);

        console.log('Dropdown options:', options);

        let bestMatch = '';
        let highestSimilarity = 0;

        for (const option of options) {
          const similarity = stringSimilarity.compareTwoStrings(streetNameLotLocation, option.text);
          if (similarity > highestSimilarity) {
            highestSimilarity = similarity;
            bestMatch = option.value;
          }
        }

        console.log(`Best match found: ${bestMatch}`);
        await page.select(rdNamesKeySelector, bestMatch);
        await new Promise(r => setTimeout(r, 500));
        console.log('Selected the best match from the dropdown.');

        await page.type('#Lanes', lanes.toString());

        // Select the closest matching functional class
        const functionalClassSelector = '#FCDetail_Key';
        console.log(`Selecting closest match for ${functionalClass} from the dropdown...`);
        await page.waitForSelector(functionalClassSelector, { visible: true });

        const fcOptions = await page.evaluate((selector) => {
          const optionsList = Array.from(document.querySelectorAll(`${selector} option`));
          return optionsList.map(option => ({ value: option.value, text: option.textContent }));
        }, functionalClassSelector);

        console.log('Functional Class Dropdown options:', fcOptions);

        let fcBestMatch = '';
        let fcHighestSimilarity = 0;

        for (const option of fcOptions) {
          const similarity = stringSimilarity.compareTwoStrings(functionalClass, option.text);
          if (similarity > fcHighestSimilarity) {
            fcHighestSimilarity = similarity;
            fcBestMatch = option.value;
          }
        }

        console.log(`Best Functional Class match found: ${fcBestMatch}`);
        await page.select(functionalClassSelector, fcBestMatch);
        await new Promise(r => setTimeout(r, 500));
        console.log('Selected the best Functional Class from the dropdown.');

        await page.type('#SectionLength', length.toString());
        await page.type('#SectionWidth', width.toString());
        await page.type('#SectionArea', area.toString());

        // Select the closest matching surface type
        const surfaceTypeSelector = '#SurfaceType_Key';
        console.log(`Selecting closest match for ${surfaceType} from the dropdown...`);
        await page.waitForSelector(surfaceTypeSelector, { visible: true });

        const stOptions = await page.evaluate((selector) => {
          const optionsList = Array.from(document.querySelectorAll(`${selector} option`));
          return optionsList.map(option => ({ value: option.value, text: option.textContent }));
        }, surfaceTypeSelector);

        console.log('Surface Type Dropdown options:', stOptions);

        let stBestMatch = '';
        let stHighestSimilarity = 0;

        for (const option of stOptions) {
          const similarity = stringSimilarity.compareTwoStrings(surfaceType, option.text);
          if (similarity > stHighestSimilarity) {
            stHighestSimilarity = similarity;
            stBestMatch = option.value;
          }
        }

        console.log(`Best Surface Type match found: ${stBestMatch}`);
        await page.select(surfaceTypeSelector, stBestMatch);
        await new Promise(r => setTimeout(r, 500));
        console.log('Selected the best Surface Type from the dropdown.');

        await page.type('#DateOfOriginalConstruction', originallyConstructed);

        console.log('Form filled out.');

        // Step eight: click on "Save" button to save the section
        console.log('Clicking on the input element and deleting contents...');
        await retryClick(page, 'input[name="ctl00$ContentPlaceHolder1$SectionArea"]');
        await page.focus('input[name="ctl00$ContentPlaceHolder1$SectionArea"]');
        await page.keyboard.down('Control');
        await page.keyboard.press('A');
        await page.keyboard.up('Control');
        await page.keyboard.press('Backspace');

        await new Promise(r => setTimeout(r, 300)); // Shorter delay before typing

        await page.keyboard.type(area.toString());

        console.log('Clicking another element to ensure the input is registered...');
        await retryClick(page, 'input[name="ctl00$ContentPlaceHolder1$SectionLength"]');
        await new Promise(r => setTimeout(r, 500)); // Add delay to ensure input registration

        console.log('Pressing F6 to save...');
        await page.keyboard.press('F6'); // Press F6 to save

        console.log('Clicking the execDash button...');
        await Promise.all([
          page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 }), // Adjusted timeout
          retryClick(page, '#execDash')
        ]);

        console.log('Returning to the dashboard and starting the next entry...');
        logTimestamp('Returning to Dashboard');

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