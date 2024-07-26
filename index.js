const puppeteer = require('puppeteer');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { readConfig } = require('./configReader');
const { readExcelData } = require('./excelReader');
const { selectDatabase, addSection } = require('./puppeteerActions');

// Function to preprocess a string value
function preprocessString(value) {
  return value ? value.trim() : '';
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
      const streetLotId = preprocessString(row['Street/Lot ID']);
      const sectionId = preprocessString(row['Section ID']);
      const streetNameLotLocation = preprocessString(row['Street Name/Lot Location']);
      const area = preprocessNumber(row['Area']);
      // Other fields
      const beginLocation = preprocessString(row['Begin Location']);
      const beginPoint = preprocessString(row['Begin Point']);
      const endLocation = preprocessString(row['End Location']);
      const endPoint = preprocessString(row['End Point']);
      const numLanes = preprocessNumber(row['# of Lanes']);
      const functionalClass = preprocessString(row['Functional Class']);
      const length = preprocessNumber(row['Length (ft.)']);
      const width = preprocessNumber(row['Width (ft.)']);
      const surfaceType = preprocessString(row['Surface Type']);
      const parkingLotType = preprocessString(row['Parking Lot Type']);
      const slabLength = preprocessNumber(row['Slab Length']);
      const slabWidth = preprocessNumber(row['Slab Width']);
      const numSlabs = preprocessNumber(row['# of Slabs']);
      const trafficIndex = preprocessNumber(row['Traffic Index']);
      const adt = preprocessNumber(row['ADT']);
      const areaId = preprocessString(row['Area ID']);
      const shoulderWidth = preprocessNumber(row['Shoulder Width']);
      const fundSource = preprocessString(row['Fund Source']);
      const effectiveDate = preprocessString(row['Effective Date']);
      const generalCode = preprocessString(row['General Code']);
      const comments = preprocessString(row['Comments']);

      console.log(`Processing Street/Lot ID: ${streetLotId}, Section ID: ${sectionId}, Area: ${area}`);

      // Perform Puppeteer actions for each entry
      await addSection(
        page,
        streetLotId,
        sectionId,
        streetNameLotLocation,
        beginLocation,
        beginPoint,
        endLocation,
        endPoint,
        numLanes,
        functionalClass,
        length,
        width,
        surfaceType,
        parkingLotType,
        slabLength,
        slabWidth,
        numSlabs,
        trafficIndex,
        adt,
        areaId,
        shoulderWidth,
        fundSource,
        effectiveDate,
        generalCode,
        comments
      );

      // Brief pause before processing the next entry
      await new Promise(r => setTimeout(r, 500));
    }

    console.log('Script completed successfully for all entries.');
    await browser.close();
  } catch (error) {
    console.error('Error:', error);
  }
})();
