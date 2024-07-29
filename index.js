const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { readConfig } = require('./configReader');
const { readExcelData } = require('./excelReader');
const { selectDatabase, retryClick } = require('./puppeteerActions');
const stringSimilarity = require('string-similarity');
const { DateTime } = require('luxon');  // Importing DateTime from luxon

function preprocessString(value) {
  const firstDashIndex = value.indexOf('-');
  if (firstDashIndex === -1) return value.trim(); // No dash found
  const secondDashIndex = value.indexOf('-', firstDashIndex + 1);
  if (secondDashIndex === -1) return value.trim(); // Only one dash found
  return value.substring(0, secondDashIndex).trim(); // Erase all characters after the second dash
}

function preprocessNumber(value) {
  const parsedValue = parseFloat(value);
  return isNaN(parsedValue) ? "0.00" : parsedValue.toFixed(2); // Default to 0.00 if value is not a valid number
}

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

async function loginToPage(page, username, password) {
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
}

async function processRow(page, row) {
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
    await addRoadName(page, streetNameLotLocation, streetLotId);
    await fillSectionDetails(page, {
      streetLotId, beginLocation, sectionId, endLocation, streetNameLotLocation,
      lanes, functionalClass, length, width, area, surfaceType, originallyConstructed
    });
    logTimestamp('Returning to Dashboard');
  } catch (error) {
    console.error(`Error processing Street/Lot ID ${streetLotId}:`, error);
  }

  await new Promise(r => setTimeout(r, 500)); // Brief pause before processing the next entry
}

async function addRoadName(page, streetNameLotLocation, streetLotId) {
  console.log('Expanding "Pavement Sections" menu...');
  await retryClick(page, '#togglePavementSections');
  await page.waitForSelector('#pavementSections.menu-dropdown.collapse.show', { visible: true });
  console.log('Clicking on "Road Names"...');
  await retryClick(page, '#linkRdNames');
  await page.waitForSelector('#ctl00_ContentPlaceHolder1_grdEDIT_grdData', { visible: true });
  const addRecordSelector = '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl00_lbtnAddRecord';
  console.log('Waiting for "Add Record" button to be visible...');
  await page.waitForSelector(addRecordSelector, { visible: true });
  console.log('Clicking on "Add Record" button...');
  await page.evaluate((selector) => {
    document.querySelector(selector).click();
  }, addRecordSelector);
  console.log('"Add Record" button clicked.');
  await new Promise(r => setTimeout(r, 2000)); // Wait for the form to be ready
  console.log(`Typing ${streetNameLotLocation} into the "Road Name" field...`);
  await page.type('#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_TB_RoadName', streetNameLotLocation);
  await new Promise(r => setTimeout(r, 500));
  console.log(`Typed ${streetNameLotLocation} into the "Road Name" field.`);
  console.log(`Typing ${streetLotId} into the "Street/Lot Number" field...`);
  await page.type('#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_TB_RoadNumber', streetLotId);
  await new Promise(r => setTimeout(r, 500));
  console.log(`Typed ${streetLotId} into the "Street/Lot Number" field.`);
  console.log('Form filled out.');
  console.log('Clicking on "Save" button...');
  await retryClick(page, '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_PerformInsertButton');
  await new Promise(r => setTimeout(r, 2000)); // Wait for 2 seconds to ensure the save operation completes
  console.log('"Save" button clicked.');
  try {
    await page.waitForSelector('.swal2-html-container', { visible: true, timeout: 2000 });
    console.log('Error message detected. Skipping this entry.');
    await page.click('.swal2-confirm');
    await new Promise(r => setTimeout(r, 1000)); // Wait for the alert to be handled
    console.log('Clicked OK on the alert.');
    const cancelButtonSelector = '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_CancelButton';
    console.log('Clicking the cancel button...');
    await page.click(cancelButtonSelector);
    await new Promise(r => setTimeout(r, 1000)); // Wait for the action to complete
    console.log('Clicked cancel button and staying on the same entry.');
  } catch (error) {
    // No error message, proceed as usual
  }
}

async function fillSectionDetails(page, details) {
  const {
    streetLotId, beginLocation, sectionId, endLocation, streetNameLotLocation,
    lanes, functionalClass, length, width, area, surfaceType, originallyConstructed
  } = details;

  console.log('Expanding "Pavement Sections" menu again...');
  await retryClick(page, '#togglePavementSections');
  await page.waitForSelector('#pavementSections.menu-dropdown.collapse.show', { visible: true });
  console.log('Clicking on "Edit Section"...');
  await retryClick(page, '#linkSectionCreate');
  await new Promise(r => setTimeout(r, 1000));
  console.log('Filling out form fields...');
  await page.type('#StreetID', streetLotId);
  await page.type('#BegLocation', beginLocation);
  await page.type('#SectionID', sectionId);
  await page.type('#EndLocation', endLocation);
  console.log(`Selecting closest match for ${streetNameLotLocation} from the dropdown...`);
  await selectClosestMatch(page, '#RdNames_Key', streetNameLotLocation);
  await page.type('#Lanes', lanes.toString());
  console.log(`Selecting closest match for ${functionalClass} from the dropdown...`);
  await selectClosestMatch(page, '#FCDetail_Key', functionalClass);
  await page.type('#SectionLength', length.toString());
  await page.type('#SectionWidth', width.toString());
  await page.type('#SectionArea', area.toString());
  console.log(`Selecting closest match for ${surfaceType} from the dropdown...`);
  await selectClosestMatch(page, '#SurfaceType_Key', surfaceType);
  await page.type('#DateOfOriginalConstruction', originallyConstructed);
  console.log('Form filled out.');
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
}

async function selectClosestMatch(page, selector, text) {
  await page.waitForSelector(selector, { visible: true });
  const options = await page.evaluate((selector) => {
    const optionsList = Array.from(document.querySelectorAll(`${selector} option`));
    return optionsList.map(option => ({ value: option.value, text: option.textContent }));
  }, selector);
  console.log('Dropdown options:', options);
  let bestMatch = '';
  let highestSimilarity = 0;
  for (const option of options) {
    const similarity = stringSimilarity.compareTwoStrings(text, option.text);
    if (similarity > highestSimilarity) {
      highestSimilarity = similarity;
      bestMatch = option.value;
    }
  }
  console.log(`Best match found: ${bestMatch}`);
  await page.select(selector, bestMatch);
  await new Promise(r => setTimeout(r, 500));
  console.log('Selected the best match from the dropdown.');
}

(async () => {
  try {
    clearUnmatchedRows();
    const configPath = path.join(__dirname, 'config.json');
    const { username, password, excelPath, dataBase } = readConfig(configPath);
    const data = readExcelData(excelPath);
    const browser = await puppeteer.launch({
      headless: false, // Set headless: true for headless mode
      defaultViewport: null // Disable the default viewport
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 1920, height: 1080 });
    await loginToPage(page, username, password);
    await selectDatabase(page, dataBase);
    console.log('Database selected. Ready to add sections.');
    for (const row of data) {
      await processRow(page, row);
    }
    console.log('Script completed successfully for all entries.');
    await browser.close();
  } catch (error) {
    console.error('Error:', error);
  }
})();
