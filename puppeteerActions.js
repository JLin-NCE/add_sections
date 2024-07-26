const puppeteer = require('puppeteer');
const fs = require('fs');
const XLSX = require('xlsx');

// Function to calculate Levenshtein distance
function levenshtein(a, b) {
  const an = a ? a.length : 0;
  const bn = b ? b.length : 0;
  if (an === 0) {
    return bn;
  }
  if (bn === 0) {
    return an;
  }
  const matrix = [];
  for (let i = 0; i <= bn; i++) {
    matrix[i] = [i];
  }
  for (let j = 0; j <= an; j++) {
    matrix[0][j] = j;
  }
  for (let i = 1; i <= bn; i++) {
    for (let j = 1; i <= an; j++) {
      if (b.charAt(i - 1) === a.charAt(j - 1)) {
        matrix[i][j] = matrix[i - 1][j - 1];
      } else {
        matrix[i][j] = Math.min(matrix[i - 1][j - 1] + 1, Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1));
      }
    }
  }
  return matrix[bn][an];
}

async function selectDatabase(page, dataBase, streetName) {
  try {
    console.log('Expanding "Systems Admin" menu...');
    await retryClick(page, '#toggleSysAdmin');
    await page.waitForSelector('#sysadmin.menu-dropdown.collapse.show', { visible: true });

    console.log('Clicking on "Open Database"...');
    await retryClick(page, '#linkDBOpen');
    await new Promise(r => setTimeout(r, 1000)); // Wait for 1 second to ensure the page fully renders

    console.log('Clicking on Database Dropdown...');
    await retryClick(page, '#cboDBName');

    // Output the dropdown options and find the closest match
    const options = await page.evaluate(() => {
      const select = document.querySelector('#cboDBName');
      return Array.from(select.options).map(option => ({
        value: option.value,
        text: option.text
      }));
    });
    console.log('Dropdown options:', options);

    // Find the closest match to the dataBase from config.json
    let closestMatch = options[0];
    let smallestDistance = levenshtein(dataBase, closestMatch.text);
    for (const option of options) {
      const distance = levenshtein(dataBase, option.text);
      if (distance < smallestDistance) {
        smallestDistance = distance;
        closestMatch = option;
      }
    }
    console.log(`Selecting closest database match: ${closestMatch.text}`);

    // Select the closest match
    await page.select('#cboDBName', closestMatch.value);

    // Click on the closest match to ensure it's selected
    await page.evaluate((value) => {
      const option = document.querySelector(`#cboDBName option[value="${value}"]`);
      if (option) {
        option.selected = true;
        const changeEvent = new Event('change', { bubbles: true });
        option.dispatchEvent(changeEvent);
      }
    }, closestMatch.value);

    console.log("Clicked on ", closestMatch.value);

    console.log('Expanding "Pavement Sections" menu...');
    await retryClick(page, '#togglePavementSections');
    await page.waitForSelector('#pavementSections.menu-dropdown.collapse.show', { visible: true });

    console.log('Clicking on "Road Names"...');
    await retryClick(page, '#linkRdNames');
    await page.waitForSelector('#ctl00_ContentPlaceHolder1_grdEDIT_grdData', { visible: true });

    console.log("Clicking on 'Add New' Button...")
    await page.waitForSelector('#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl00_lbtnAddRecord')

    // Type in Street Name/Lot Location Text Field
    console.log('Typing in Street Name/Lot Location...');
    await page.type('#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_TB_RoadName', streetName);

    // Extract all characters before last dash in the third column for Street/Lot Number
    const streetLotNumber = streetName.substring(0, streetName.lastIndexOf('-'));

    console.log(`Typing in Street/Lot Number: ${streetLotNumber}`);
    await page.type('#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_TB_RoadNumber', streetLotNumber);

    // Confirm add street name/lot location by clicking the button
    console.log('Confirming add street name/lot location...');
    await retryClick(page, '#ctl00_ContentPlaceHolder1_grdEDIT_grdData_ctl00_ctl02_ctl04_PerformInsertButton');

    console.log('Street name/lot location added.');
  } catch (error) {
    console.error('Error during database selection:', error.message);
  }
}

async function addSection(
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
) {
  try {
    console.log('Filling out the "Create Section" form...');

    console.log(`StreetLotId: ${streetLotId}`);
    await page.type('#StreetLotId', streetLotId);

    console.log(`SectionId: ${sectionId}`);
    await page.type('#SectionId', sectionId);

    console.log(`StreetNameLotLocation: ${streetNameLotLocation}`);
    await page.type('#StreetNameLotLocation', streetNameLotLocation);

    console.log(`BeginLocation: ${beginLocation}`);
    await page.type('#BeginLocation', beginLocation);

    console.log(`BeginPoint: ${beginPoint}`);
    await page.type('#BeginPoint', beginPoint);

    console.log(`EndLocation: ${endLocation}`);
    await page.type('#EndLocation', endLocation);

    console.log(`EndPoint: ${endPoint}`);
    await page.type('#EndPoint', endPoint);

    console.log(`NumLanes: ${numLanes}`);
    await page.type('#NumLanes', numLanes);

    console.log(`FunctionalClass: ${functionalClass}`);
    await page.select('#FunctionalClass', functionalClass);

    console.log(`Length: ${length}`);
    await page.type('#Length', length);

    console.log(`Width: ${width}`);
    await page.type('#Width', width);

    console.log(`SurfaceType: ${surfaceType}`);
    await page.select('#SurfaceType', surfaceType);

    console.log(`ParkingLotType: ${parkingLotType}`);
    await page.select('#ParkingLotType', parkingLotType);

    console.log(`SlabLength: ${slabLength}`);
    await page.type('#SlabLength', slabLength);

    console.log(`SlabWidth: ${slabWidth}`);
    await page.type('#SlabWidth', slabWidth);

    console.log(`NumSlabs: ${numSlabs}`);
    await page.type('#NumSlabs', numSlabs);

    console.log(`TrafficIndex: ${trafficIndex}`);
    await page.type('#TrafficIndex', trafficIndex);

    console.log(`ADT: ${adt}`);
    await page.type('#ADT', adt);

    console.log(`AreaId: ${areaId}`);
    await page.select('#AreaId', areaId);

    console.log(`ShoulderWidth: ${shoulderWidth}`);
    await page.type('#ShoulderWidth', shoulderWidth);

    console.log(`FundSource: ${fundSource}`);
    await page.select('#FundSource', fundSource);

    console.log(`EffectiveDate: ${effectiveDate}`);
    await page.type('#EffectiveDate', effectiveDate);

    console.log(`GeneralCode: ${generalCode}`);
    await page.select('#GeneralCode', generalCode);

    console.log(`Comments: ${comments}`);
    await page.type('#Comments', comments);

    console.log('Form filled out.');

    console.log('Saving the section...');
    await retryClick(page, '#SaveSectionButton'); // Assuming there is a button with this ID to save the section
    await page.waitForNavigation();
    console.log('Section saved.');

  } catch (error) {
    console.error('Error during section addition:', error.message);
  }
}

async function retryClick(page, selector, retries = 3, waitTime = 500) {
  for (let attempt = 0; attempt < retries; attempt++) {
    try {
      await page.waitForSelector(selector, { visible: true });
      await page.click(selector);
      return;
    } catch (error) {
      if (attempt < retries - 1) {
        console.log(`Retry ${attempt + 1} for clicking ${selector}`);
        await new Promise(r => setTimeout(r, waitTime));
      } else {
        console.error(`Failed to click ${selector} after ${retries} attempts:`, error);
        throw error;
      }
    }
  }
}

module.exports = { selectDatabase, addSection };
