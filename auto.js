// auto.js - Final Version (with Electron File Dialog for File Selection)
require('dotenv').config();
const logger = require('./src/utils/logger');
const ExcelParser = require('./src/utils/excelParser');
const { retry } = require('./src/utils/retry');
const ErrorHandler = require('./src/utils/errorHandler');
const EnhancedProgressTracker = require('./src/utils/enhancedProgress');
const config = require('./src/config');
const { chromium } = require('@playwright/test');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const fsPromises = fs.promises;
const { spawn } = require('child_process');

/* ===================== Utility: Select Excel File via Electron Dialog ===================== */
/**
 * Spawns an Electron process that opens a file selection dialog.
 * The selected file path is written to temp/selectedExcel.json.
 * Returns the full path of the selected file.
 */
async function selectExcelFile(rootDir) {
  const electronBinary = require('electron');
  const selectionModulePath = path.join(__dirname, 'src', 'selection', 'main.js');
  
  // Spawn the selection module (it will open the dialog and write the selection)
  await new Promise((resolve, reject) => {
    const selectionProcess = spawn(electronBinary, [selectionModulePath, rootDir], {
      stdio: 'inherit',
      shell: true
    });
    selectionProcess.on('close', (code) => {
      if (code === 0) resolve();
      else reject(new Error('File selection cancelled or failed.'));
    });
    selectionProcess.on('error', (err) => reject(err));
  });
  
  // Read the selected file path from the temporary JSON file
  const tempFile = path.join(rootDir, 'temp', 'selectedExcel.json');
  try {
    const rawData = await fsPromises.readFile(tempFile, 'utf-8');
    const data = JSON.parse(rawData);
    if (!data.filePath) throw new Error('No file path found in selection data.');
    logger.info('User selected Excel file:', { filePath: data.filePath });
    return data.filePath;
  } catch (error) {
    throw new Error('Failed to load selected Excel file: ' + error.message);
  }
}

/* ===================== Date Formatting Helper ===================== */
function formatDateForInput(dateInput) {
  let date = dateInput instanceof Date ? dateInput : new Date(dateInput);
  const pad = n => n.toString().padStart(2, '0');
  const month = pad(date.getMonth() + 1);
  const day = pad(date.getDate());
  const year = date.getFullYear();
  let hours = date.getHours();
  const minutes = pad(date.getMinutes());
  const ampm = hours >= 12 ? 'PM' : 'AM';
  hours = hours % 12;
  if (hours === 0) hours = 12;
  hours = pad(hours);
  return `${month}-${day}-${year} ${hours}:${minutes} ${ampm}`;
}

/* ===================== New Transfer-Level Extraction ===================== */
function extractField(worksheet, fieldName) {
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const maxRow = Math.min(range.s.r + 30, range.e.r);
  const maxCol = Math.min(range.s.c + 10, range.e.c);
  for (let row = range.s.r; row <= maxRow; row++) {
    for (let col = range.s.c; col < maxCol; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = worksheet[cellAddress];
      if (cell && cell.v != null && cell.v.toString().trim().toLowerCase() === fieldName.toLowerCase()) {
        const nextCellAddress = XLSX.utils.encode_cell({ r: row, c: col + 1 });
        const nextCell = worksheet[nextCellAddress];
        return nextCell && nextCell.v != null ? nextCell.v.toString().trim() : null;
      }
    }
  }
  return null;
}

function extractTransferInfo(worksheet) {
  return {
    store: extractField(worksheet, "Store"),
    driver1: extractField(worksheet, "Driver 1"),
    driver2: extractField(worksheet, "Driver 2"),
    vehicle: extractField(worksheet, "Vehicle"),
    departureDate: extractField(worksheet, "Departure Time") || extractField(worksheet, "Departure Date"),
    arrivalDate: extractField(worksheet, "Arrival Time") || extractField(worksheet, "Arrival Date"),
    route: extractField(worksheet, "Route") || extractField(worksheet, "Route/Description")
  };
}

/* ===================== New Product Table Extraction ===================== */
function extractProductTable(worksheet) {
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  let headerRowIndex = -1;
  let headers = [];

  const requiredHeaders = {
    productName: ["product name", "label/description", "product"],
    barcode: ["barcode", "value"],
    externalCode: ["external code", "ext code"],
    qty: ["qty", "quantity", "qnty"]
  };

  for (let row = range.s.r; row <= Math.min(range.s.r + 200, range.e.r); row++) {
    let rowHeaders = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = worksheet[cellAddress];
      rowHeaders.push(cell && cell.v != null ? cell.v.toString().trim().toLowerCase() : "");
    }
    if (rowHeaders.some(val => requiredHeaders.productName.includes(val))) {
      headerRowIndex = row;
      headers = rowHeaders;
      break;
    }
  }

  if (headerRowIndex === -1) {
    return [];
  }

  const headerMap = {};
  for (let key in requiredHeaders) {
    for (let i = 0; i < headers.length; i++) {
      if (requiredHeaders[key].includes(headers[i])) {
        headerMap[key] = i;
        break;
      }
    }
  }

  if (
    headerMap.productName === undefined ||
    headerMap.barcode === undefined ||
    headerMap.externalCode === undefined ||
    headerMap.qty === undefined
  ) {
    throw new Error(
      "Product table is missing one or more required headers: Product Name, Barcode, External Code, QTY/Quantity"
    );
  }

  const products = [];
  for (let row = headerRowIndex + 1; row <= range.e.r; row++) {
    const prodCellAddress = XLSX.utils.encode_cell({ r: row, c: headerMap.productName });
    const prodCell = worksheet[prodCellAddress];
    const prodValue = prodCell && prodCell.v != null ? prodCell.v.toString().trim() : "";
    if (prodValue === "") {
      break;
    }
    const barcodeCell = worksheet[XLSX.utils.encode_cell({ r: row, c: headerMap.barcode })];
    const externalCodeCell = worksheet[XLSX.utils.encode_cell({ r: row, c: headerMap.externalCode })];
    const qtyCell = worksheet[XLSX.utils.encode_cell({ r: row, c: headerMap.qty })];
    const product = {
      productName: prodValue,
      barcode: barcodeCell && barcodeCell.v != null ? barcodeCell.v.toString().trim() : "N/A",
      externalCode: externalCodeCell && externalCodeCell.v != null ? externalCodeCell.v.toString().trim() : "N/A",
      qty: qtyCell && qtyCell.v != null ? qtyCell.v.toString().trim() : "N/A"
    };
    products.push(product);
  }
  return products;
}

/* ===================== Read Transfer Data ===================== */
async function readTransferData(filePath) {
  try {
    logger.info('Reading Excel file:', { filePath });
    const workbook = XLSX.readFile(filePath, {
      cellDates: true,
      cellNF: true,
      cellText: false,
      cellStyles: true,
      cellFormula: true
    });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const transferInfo = extractTransferInfo(worksheet);
    const products = extractProductTable(worksheet);
    const transferData = {
      ...transferInfo,
      products: products
    };
    const validationErrors = validateTransferData(transferData);
    if (validationErrors.length > 0) {
      throw new Error(`Validation errors: ${validationErrors.join(', ')}`);
    }
    logger.info('Transfer data validated successfully');
    return transferData;
  } catch (error) {
    logger.error('Error reading Excel file:', error);
    throw error;
  }
}

/* ===================== Validate Transfer Data ===================== */
function validateTransferData(transferData) {
  const errors = [];
  const requiredFields = config.get('validation.requiredFields');
  const dateRules = config.get('validation.dateRules');
  requiredFields.forEach(field => {
    if (!transferData[field]) {
      errors.push(`Missing required field: ${field}`);
    }
  });
  if (transferData.departureDate && transferData.arrivalDate) {
    if (!(transferData.departureDate instanceof Date)) {
      transferData.departureDate = new Date(transferData.departureDate);
    }
    if (!(transferData.arrivalDate instanceof Date)) {
      transferData.arrivalDate = new Date(transferData.arrivalDate);
    }
    if (isNaN(transferData.departureDate.getTime())) {
      errors.push('Invalid departure date format');
    }
    if (isNaN(transferData.arrivalDate.getTime())) {
      errors.push('Invalid arrival date format');
    }
    const timeDiff = transferData.arrivalDate.getTime() - transferData.departureDate.getTime();
    if (timeDiff < dateRules.minAdvanceTime) {
      errors.push('Departure and arrival times must be at least 30 minutes apart');
    }
    if (timeDiff > dateRules.maxTripDuration) {
      errors.push('Trip duration cannot exceed 24 hours');
    }
    const departureDay = transferData.departureDate.toLocaleString('en-US', { weekday: 'long' });
    if (!dateRules.allowedDays.includes(departureDay)) {
      errors.push(`Transfers not allowed on ${departureDay}`);
    }
    const departureHour = transferData.departureDate.getHours();
    const [startHour] = dateRules.allowedHours.start.split(':').map(Number);
    const [endHour] = dateRules.allowedHours.end.split(':').map(Number);
    if (departureHour < startHour || departureHour > endHour) {
      errors.push('Transfer must be scheduled during business hours');
    }
  }
  if (!transferData.products || transferData.products.length === 0) {
    errors.push('No products found');
  } else if (transferData.products.length > config.get('excel.dataValidation.maxProducts')) {
    errors.push(`Number of products exceeds maximum limit of ${config.get('excel.dataValidation.maxProducts')}`);
  }
  return errors;
}

/* ===================== Write Temporary Data ===================== */
async function writeTempData(data) {
  const tempDir = path.join(__dirname, 'temp');
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }
  const filePath = path.join(tempDir, 'transferData.json');
  await fsPromises.writeFile(filePath, JSON.stringify(data, null, 2), 'utf-8');
  return filePath;
}

/* ===================== Electron Confirmation ===================== */
async function openConfirmationWindow(rootDir) {
  return new Promise((resolve, reject) => {
    const electronBinary = require('electron');
    const electronMainPath = path.join(__dirname, 'src', 'confirmation', 'main.js');
    const electronProcess = spawn(electronBinary, [electronMainPath, rootDir], {
      stdio: 'inherit',
      shell: true
    });
    electronProcess.on('close', (code) => {
      resolve(code === 0);
    });
    electronProcess.on('error', (err) => {
      reject(err);
    });
  });
}

async function loadApprovedTransferData(rootDir) {
  const tempFile = path.join(rootDir, 'temp', 'transferData.json');
  try {
    const rawData = await fsPromises.readFile(tempFile, 'utf-8');
    return JSON.parse(rawData);
  } catch (error) {
    logger.error('Error loading approved transfer data:', error);
    throw error;
  }
}

/* ===================== Main Automation Process ===================== */
async function createTransferProcess(transferData) {
  const progress = new EnhancedProgressTracker(8, 'Transfer Creation');
  const browser = await chromium.launch({
    channel: config.get('browser.channel'),
    headless: config.get('browser.headless')
  });
  const context = await browser.newContext();
  const page = await context.newPage();
  try {
    // LOGIN PROCESS
    progress.addStep('Login Process');
    await retry(async () => {
      await ErrorHandler.withErrorHandler(page, 'login', async () => {
        logger.info('Starting login process');
        progress.updateStepProgress(10, 'Navigating to login page');
        await page.goto(config.get('urls.login'));
        await page.click('text=Log in with Microsoft');
        progress.updateStepProgress(30, 'Entering SweedPos credentials');
        await page.waitForSelector('input[placeholder="Email"]', { timeout: config.get('browser.timeout.element') });
        await page.fill('input[placeholder="Email"]', process.env.MS_USERNAME);
        await page.click('text=Verify');
        progress.updateStepProgress(60, 'Completing Microsoft login');
        await page.waitForURL(/login\.microsoftonline\.com/, { timeout: config.get('browser.timeout.navigation') });
        await page.waitForSelector('input[name="loginfmt"]');
        await page.fill('input[name="loginfmt"]', process.env.MS_USERNAME);
        await page.click('#idSIButton9');
        progress.updateStepProgress(80, 'Entering password');
        await page.waitForSelector('input[type="password"]');
        await page.fill('input[type="password"]', process.env.MS_PASSWORD);
        await page.click('#idSIButton9');
        progress.updateStepProgress(100, 'Login completed');
      });
    }, {
      retries: config.get('retry.attempts'),
      delay: config.get('retry.delay'),
      name: 'login process',
      onRetry: async (error) => {
        progress.addWarning('Login', 'Retrying login process');
        await ErrorHandler.handleError(page, error, 'login-retry');
      }
    });
    progress.completeStep('Login Process');

    // SESSION SETUP
    progress.addStep('Session Setup');
    await ErrorHandler.withErrorHandler(page, 'stay-signed-in', async () => {
      try {
        await page.waitForSelector('#idSIButton9', { timeout: config.get('browser.timeout.element') });
        await page.click('#idSIButton9');
        progress.updateStepProgress(100, 'Session setup completed');
      } catch {
        logger.info('Stay signed in prompt not shown');
        progress.updateStepProgress(100, 'No session prompt needed');
      }
    });
    progress.completeStep('Session Setup');

    // NAVIGATION
    progress.addStep('Navigation');
    await retry(async () => {
      await ErrorHandler.withErrorHandler(page, 'navigation', async () => {
        progress.updateStepProgress(20, 'Waiting for dashboard');
        await page.waitForSelector('.rac-header__user-avatar', { timeout: config.get('browser.timeout.navigation') });
        await page.waitForLoadState('networkidle');
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        progress.updateStepProgress(50, 'Accessing inventory menu');
        try {
          await page.waitForSelector('div[title="Inventory"]', { timeout: config.get('browser.timeout.element') });
          await page.click('div[title="Inventory"]');
        } catch {
          logger.info('Inventory menu might already be expanded');
        }
        progress.updateStepProgress(80, 'Navigating to transfers');
        await page.waitForSelector('a[data-id="portal:dashboards:main_link_inventory-transfers"]');
        await page.click('a[data-id="portal:dashboards:main_link_inventory-transfers"]');
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        progress.updateStepProgress(100, 'Navigation completed');
      });
    }, {
      retries: config.get('retry.attempts'),
      delay: config.get('retry.delay'),
      name: 'navigation',
      onRetry: async (error) => {
        progress.addWarning('Navigation', 'Retrying navigation');
        await ErrorHandler.handleError(page, error, 'navigation-retry');
      }
    });
    progress.completeStep('Navigation');

    // TRANSFER CREATION
    progress.addStep('Transfer Creation');
    await ErrorHandler.withErrorHandler(page, 'create-transfer', async () => {
      progress.updateStepProgress(30, 'Creating new transfer');
      await page.getByRole('button', { name: 'New transfer' }).click();
      await page.waitForTimeout(config.get('browser.timeout.animation'));
      progress.updateStepProgress(70, 'Selecting BioTrack');
      await page.waitForSelector('div.rac-field__self');
      await page.evaluate(() => {
        const radios = document.querySelectorAll('input[type="checkbox"]');
        for (const radio of radios) {
          if (!radio.checked) radio.click();
        }
      });
      await page.waitForTimeout(config.get('browser.timeout.animation'));
      progress.updateStepProgress(100, 'Transfer initialized');
    });
    progress.completeStep('Transfer Creation');

    // DETAILS ENTRY
    progress.addStep('Details Entry');
    await retry(async () => {
      await ErrorHandler.withErrorHandler(page, 'fill-details', async () => {
        await page.locator('div').filter({ hasText: /^To store\*$/ }).locator('div').nth(1).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('li').filter({ hasText: transferData.store }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('input[name="\\$\\$biotrackFl\\$biotrackDriver1"]').click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('li').filter({ hasText: transferData.driver1 }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        if (transferData.driver2) {
          await page.locator('input[name="\\$\\$biotrackFl\\$biotrackDriver2"]').click();
          await page.waitForTimeout(config.get('browser.timeout.animation'));
          await page.locator('li').filter({ hasText: transferData.driver2 }).click();
          await page.waitForTimeout(config.get('browser.timeout.animation'));
        }
        await page.locator('input[name="\\$\\$biotrackFl\\$biotrackVehicle"]').click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('li').filter({ hasText: transferData.vehicle }).first().click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        const departureDateFormatted = formatDateForInput(transferData.departureDate);
        const arrivalDateFormatted = formatDateForInput(transferData.arrivalDate);
        await page.locator('div').filter({ hasText: /^Approximate departure date\*$/ }).locator('div').nth(1).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.getByPlaceholder('MM-DD-YYYY hh:mm A').fill(departureDateFormatted);
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('button').filter({ hasText: 'Ok' }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('div').filter({ hasText: /^Approximate arrival date\*$/ }).locator('div').nth(1).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.getByPlaceholder('MM-DD-YYYY hh:mm A').fill(arrivalDateFormatted);
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('button').filter({ hasText: 'Ok' }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('textarea[name="\\$\\$biotrackFl\\$plannedRoad"]').click();
        await page.locator('textarea[name="\\$\\$biotrackFl\\$plannedRoad"]').fill(transferData.route);
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        progress.updateStepProgress(100, 'Details completed');
      });
    }, {
      retries: config.get('retry.attempts'),
      delay: config.get('retry.delay'),
      name: 'fill-details',
      onRetry: async (error) => {
        progress.addWarning('Details', 'Retrying details entry');
        await ErrorHandler.handleError(page, error, 'fill-details-retry');
      }
    });
    progress.completeStep('Details Entry');

    // -------------------- PRODUCT PROCESSING --------------------
    progress.addStep('Product Processing');
    await ErrorHandler.withErrorHandler(page, 'process-products', async () => {
      await page.waitForSelector('button:has-text("Create transfer")', { timeout: config.get('browser.timeout.element') });
      await page.getByRole('button', { name: 'Create transfer' }).click();
      await page.waitForTimeout(config.get('browser.timeout.animation'));
      await page.waitForTimeout(config.get('browser.timeout.navigation'));

      const maxProducts = config.get('processing.maxProducts') || transferData.products.length;
      for (let i = 0; i < Math.min(maxProducts, transferData.products.length); i++) {
        const product = transferData.products[i];
        logger.info(`Processing product ${i + 1} of ${maxProducts}:`, { productName: product.productName });
        
        if (!product.externalCode || product.externalCode === 'N/A') {
          logger.warn(`Skipping product ${i + 1}: External code not found`);
          continue;
        }
        
        const searchQuery = product.externalCode;
        
        await page.waitForSelector('button:has-text("Add product manually")', { timeout: config.get('browser.timeout.element') });
        await page.getByRole('button', { name: 'Add product manually' }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        
        const searchBox = await page.getByRole('textbox', { name: 'Search' });
        await searchBox.click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await searchBox.fill(searchQuery);
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        
        if (await page.getByText('No items found').isVisible()) {
          await searchBox.fill('');
          await page.waitForTimeout(config.get('browser.timeout.animation'));
          await searchBox.fill(product.barcode);
          await page.waitForTimeout(config.get('browser.timeout.animation'));
        }
        
        await page.getByRole('listitem').first().click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.getByRole('button', { name: 'Apply' }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.mouse.click(1200, 600);
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        const qtyInput = page.locator('input[name="qty"]').last();
        await qtyInput.click();
        await qtyInput.fill(product.qty);
        await page.waitForTimeout(config.get('browser.timeout.animation'));
      }
      progress.updateStepProgress(100, 'Products processed');
    });
    progress.completeStep('Product Processing');

    // FINALIZATION
    progress.addStep('Finalization');
    await retry(async () => {
      progress.updateStepProgress(50, 'Applying changes');
      await page.getByRole('button', { name: 'Apply' }).click();
      await page.waitForTimeout(config.get('browser.timeout.animation'));
      progress.updateStepProgress(90, 'Updating draft');
      await page.getByRole('button', { name: 'Update draft' }).click();
      await page.waitForTimeout(config.get('browser.timeout.animation'));
      progress.updateStepProgress(100, 'Transfer finalized');
    }, {
      retries: config.get('retry.attempts'),
      delay: config.get('retry.delay'),
      name: 'finalization',
      onRetry: async (error) => {
        progress.addWarning('Finalization', 'Retrying finalization');
        await ErrorHandler.handleError(page, error, 'finalization-retry');
      }
    });
    progress.completeStep('Finalization');
  } catch (error) {
    progress.addError(progress.steps[progress.currentStep]?.name || 'Unknown Step', error);
    await ErrorHandler.handleError(page, error, 'transfer-creation');
    throw error;
  } finally {
    try {
      progress.addStep('Cleanup');
      await browser.close();
      progress.updateStepProgress(100, 'Browser closed');
      progress.completeStep('Cleanup');
      progress.finish();
    } catch (error) {
      logger.error('Error closing browser:', error);
      progress.addError('Cleanup', error);
    }
  }
}

/* ===================== Main Execution Block ===================== */
(async () => {
  const progress = new EnhancedProgressTracker(4, 'Overall Process');
  try {
    const rootDir = process.cwd();
    progress.addStep('File Detection');
    // Instead of automatically picking the first file, call the Electron file dialog module.
    const excelPath = await selectExcelFile(rootDir);
    progress.completeStep('File Detection');

    progress.addStep('Data Reading');
    let transferData = await readTransferData(excelPath);
    progress.completeStep('Data Reading');

    await writeTempData(transferData);

    progress.addStep('User Confirmation');
    const userConfirmed = await openConfirmationWindow(rootDir);
    progress.completeStep('User Confirmation');

    if (!userConfirmed) {
      logger.info('Transfer cancelled by user from confirmation window.');
      process.exit(0);
    }

    const approvedData = await loadApprovedTransferData(rootDir);

    progress.addStep('Transfer Creation');
    logger.info('Starting transfer process...');
    await createTransferProcess(approvedData);
    progress.completeStep('Transfer Creation');
    logger.info('Transfer process completed successfully');
  } catch (error) {
    logger.error('Failed to process transfer:', error);
    progress.addError('Process', error);
    process.exit(1);
  } finally {
    progress.finish();
  }
  if (process.argv.includes('--run-tests')) {
    await runTests();
  }
})();

async function runTests() {
  const assert = require('assert');
  const os = require('os');
  logger.info('\n--- Running Unit Tests ---\n');
  function createTempDir() {
    return fs.mkdtempSync(path.join(os.tmpdir(), 'auto-manifest-test-'));
  }
  function createDummyExcelFile(filePath, sheetName = 'Sheet1') {
    const XLSX = require('xlsx');
    const workbook = XLSX.utils.book_new();
    const ws_data = [['Header1', 'Header2'], ['Data1', 'Data2']];
    const worksheet = XLSX.utils.aoa_to_sheet(ws_data);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, filePath);
  }
  (function testNoExcelFiles() {
    const tempDir = createTempDir();
    let errorCaught = false;
    selectExcelFile(tempDir)
      .then(() => { assert.fail('Expected error was not thrown.'); })
      .catch((error) => { errorCaught = true; assert.strictEqual(error.message, 'File selection cancelled or failed.'); })
      .finally(() => {
        fs.rmdirSync(tempDir, { recursive: true });
        assert.strictEqual(errorCaught, true, 'Error was not caught for no Excel files.');
        logger.info('Test 1 (No Excel files) passed.');
      });
  })();
  await (async function testValidExcelFile() {
    const tempDir = createTempDir();
    const dummyFileName = 'test.xlsx';
    const dummyFilePath = path.join(tempDir, dummyFileName);
    const requiredSheet = config.get('excel.requiredSheet') || 'Sheet1';
    createDummyExcelFile(dummyFilePath, requiredSheet);
    // Write the dummy file path to the temporary selection file
    const tempSelectionFile = path.join(tempDir, 'selectedExcel.json');
    await fsPromises.writeFile(tempSelectionFile, JSON.stringify({ filePath: dummyFilePath }), 'utf-8');
    try {
      const foundFile = await selectExcelFile(tempDir);
      assert.strictEqual(foundFile, dummyFilePath, 'The found file does not match the expected file.');
      logger.info('Test 2 (Valid Excel file) passed.');
    } catch (error) {
      assert.fail('Unexpected error for valid Excel file: ' + error.message);
    } finally {
      fs.rmdirSync(tempDir, { recursive: true });
    }
  })();
  logger.info('--- All Unit Tests Completed ---\n');
}
