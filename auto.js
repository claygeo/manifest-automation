// auto.js - Updated Version for Store Transfer Automation with Split Functionality
require('dotenv').config();
const logger = require('./src/utils/logger');
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

// Store-specific configurations (for reference, though data comes from Electron)
const storeConfigs = {
  FTP: {
    toStore: 'Ft Pierce Warehouse Hub',
    drivers: { driver1: 'Ange', driver2: 'James Roberts' },
    departureTime: '07:00 AM',
    arrivalTime: '12:00 PM'
  },
  Ocala: {
    toStore: 'Ocala Warehouse Hub',
    drivers: { driver1: 'Sergio Hervis', driver2: 'Courtney Bruce' },
    departureTime: '12:00 PM',
    arrivalTime: '07:00 PM'
  },
  Homestead: {
    toStore: 'Homestead Processing Hub'
  },
  'Mt. Dora': {
    toStore: 'Mt. Dora Processing Hub'
  }
};

/* ===================== Utility: Select Excel File via Electron Dialog ===================== */
async function selectExcelFile(rootDir) {
  const electronBinary = require('electron');
  const selectionModulePath = path.join(__dirname, 'src', 'selection', 'main.js');
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
function formatDateForInput(dateInput, time) {
  let date = dateInput instanceof Date ? dateInput : new Date(dateInput);
  const pad = n => n.toString().padStart(2, '0');
  const month = pad(date.getMonth() + 1);
  const day = pad(date.getDate());
  const year = date.getFullYear();
  return `${month}-${day}-${year} ${time}`;
}

/* ===================== Extract Data from Excel ===================== */
function extractTransferData(worksheet) {
  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const data = [];
  const headers = [];
  
  // Dynamically read headers from the first row
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
    const cell = worksheet[cellAddress];
    headers.push(cell && cell.v ? cell.v.toString().trim() : '');
  }

  // Extract store names from headers (assuming D-G are [Store] Units and [Store] Cases)
  const store1 = headers[3].replace(' Units', ''); // e.g., "FTP Units" -> "FTP"
  const store2 = headers[5].replace(' Units', ''); // e.g., "Ocala Units" -> "Ocala"

  // Extract data rows
  for (let row = range.s.r + 1; row <= range.e.r; row++) {
    const rowData = {};
    headers.forEach((header, col) => {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = worksheet[cellAddress];
      rowData[header] = cell && cell.v != null ? cell.v.toString().trim() : '';
    });
    if (rowData['#']) data.push(rowData);
  }

  return { data, stores: [store1, store2], headers };
}

/* ===================== Electron Confirmation Window ===================== */
async function openConfirmationWindow(rootDir) {
  return new Promise((resolve, reject) => {
    const electronBinary = require('electron');
    const electronMainPath = path.join(__dirname, 'src', 'confirmation', 'main.js');
    const electronProcess = spawn(electronBinary, [electronMainPath, rootDir], {
      stdio: 'inherit',
      shell: true
    });
    electronProcess.on('close', (code) => {
      if (code === 0) {
        fs.readFile(path.join(rootDir, 'temp', 'transferConfig.json'), 'utf-8', (err, data) => {
          if (err) reject(err);
          else resolve(JSON.parse(data));
        });
      } else {
        // Check for cancellation flag
        const cancelFlagPath = path.join(rootDir, 'temp', 'cancelFlag.json');
        if (fs.existsSync(cancelFlagPath)) {
          const cancelData = JSON.parse(fs.readFileSync(cancelFlagPath, 'utf-8'));
          if (cancelData.canceled) {
            resolve(null); // Indicate cancellation
          } else {
            resolve(null); // Non-zero exit without explicit cancel
          }
        } else {
          resolve(null); // Non-zero exit without cancel flag
        }
      }
    });
    electronProcess.on('error', (err) => reject(err));
  });
}

/* ===================== Write Temporary Data ===================== */
async function writeTempData(data) {
  const tempDir = path.join(__dirname, 'temp');
  if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
  const filePath = path.join(tempDir, 'transferData.json');
  await fsPromises.writeFile(filePath, JSON.stringify(data, null, 2), 'utf-8');
  return filePath;
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
        // Validate transferData before proceeding
        if (!transferData.toStore || !transferData.driver1 || !transferData.departureDate || !transferData.departureTime || !transferData.arrivalTime) {
          logger.error('Missing required fields in transferData:', transferData);
          throw new Error('Missing required fields in transferData');
        }

        // Click the "To store*" dropdown
        await page.locator('div').filter({ hasText: /^To store\*$/ }).locator('div').nth(1).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        await page.locator('li').filter({ hasText: transferData.toStore }).click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));

        // Driver 1
        await page.locator('input[name="\\$\\$biotrackFl\\$biotrackDriver1"]').click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        const driver1Options = await page.locator('li').allTextContents();
        const driver1Option = page.locator('li').filter({ hasText: transferData.driver1 });
        if (await driver1Option.count() === 0) {
          logger.error(`Driver 1 "${transferData.driver1}" not found in options:`, driver1Options);
          throw new Error(`Driver 1 "${transferData.driver1}" not found`);
        }
        await driver1Option.first().click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));

        // Driver 2
        if (transferData.driver2) {
          await page.locator('input[name="\\$\\$biotrackFl\\$biotrackDriver2"]').click();
          await page.waitForTimeout(config.get('browser.timeout.animation'));
          const driver2Options = await page.locator('li').allTextContents();
          const driver2Option = page.locator('li').filter({ hasText: transferData.driver2 });
          if (await driver2Option.count() === 0) {
            logger.error(`Driver 2 "${transferData.driver2}" not found in options:`, driver2Options);
            throw new Error(`Driver 2 "${transferData.driver2}" not found`);
          }
          await driver2Option.first().click();
          await page.waitForTimeout(config.get('browser.timeout.animation'));
        }

        // Vehicle
        await page.locator('input[name="\\$\\$biotrackFl\\$biotrackVehicle"]').click();
        await page.waitForTimeout(config.get('browser.timeout.animation'));
        if (transferData.vehicle) {
          const vehicleOptions = await page.locator('li').allTextContents();
          const vehicleOption = page.locator('li').filter({ hasText: transferData.vehicle });
          if (await vehicleOption.count() === 0) {
            logger.error(`Vehicle "${transferData.vehicle}" not found in options:`, vehicleOptions);
            throw new Error(`Vehicle "${transferData.vehicle}" not found`);
          }
          await vehicleOption.first().click();
        } else {
          logger.info('No vehicle specified, skipping vehicle selection');
        }
        await page.waitForTimeout(config.get('browser.timeout.animation'));

        // Departure and Arrival Dates/Times
        const departureDateFormatted = formatDateForInput(transferData.departureDate, transferData.departureTime);
        const arrivalDateFormatted = formatDateForInput(transferData.departureDate, transferData.arrivalTime);

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

        // Route
        await page.locator('textarea[name="\\$\\$biotrackFl\\$plannedRoad"]').click();
        await page.locator('textarea[name="\\$\\$biotrackFl\\$plannedRoad"]').fill(transferData.route || 'Default Route');
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

      const maxProducts = config.get('processing.maxProducts') || transferData.transferData.length;
      const selectedStore = transferData.store;
      const unitColumn = `${selectedStore} Units`;
      const caseColumn = `${selectedStore} Cases`;

      // Check if the selected store has data in the Excel sheet
      if (!transferData.headers.includes(unitColumn) && !transferData.headers.includes(caseColumn)) {
        logger.info(`No data found for store ${selectedStore} in Excel sheet. Skipping product processing.`);
        progress.updateStepProgress(100, 'No products to process for this store');
        return;
      }

      for (let i = 0; i < Math.min(maxProducts, transferData.transferData.length); i++) {
        const row = transferData.transferData[i];
        logger.info(`Processing row ${i + 1} of ${maxProducts}:`, { row });

        const qty = row[unitColumn] || row[caseColumn] || '0'; // Default to 0 if no units or cases

        if (!row['Last 4 of Barcode'] || row['Last 4 of Barcode'] === 'N/A' || qty === '0') {
          logger.warn(`Skipping row ${i + 1}: Barcode not found or no quantity`);
          continue;
        }

        const searchQuery = row['Last 4 of Barcode'];

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
          await searchBox.fill(row['Last 4 of Barcode']);
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
        await qtyInput.fill(qty);
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

      // Check if split is requested and handle the Split All action
      if (transferData.split) {
        logger.info('Split option selected, clicking "Split All" button');
        await page.getByRole('button', { name: 'Split All' }).click(); // Left-click (default)
        logger.info('Waiting 30 seconds after Split All action');
        await page.waitForTimeout(30000); // Wait 30 seconds as required
      } else {
        logger.info('Split option not selected, skipping Split All');
      }

      progress.updateStepProgress(90, 'Updating draft');
      await page.getByRole('button', { name: 'Update draft' }).click(); // Left-click (default)
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
    const excelPath = await selectExcelFile(rootDir);
    progress.completeStep('File Detection');

    progress.addStep('Data Reading');
    const workbook = XLSX.readFile(excelPath, { cellDates: true, cellNF: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const { data, stores, headers } = extractTransferData(worksheet);
    const transferData = { transferData: data, stores, headers };
    progress.completeStep('Data Reading');

    progress.addStep('User Confirmation');
    const userConfig = await openConfirmationWindow(rootDir);
    if (!userConfig) {
      logger.info('Transfer cancelled by user from confirmation window.');
      progress.completeStep('User Confirmation');
      progress.finish();
      // Clean up temp files
      const tempDir = path.join(rootDir, 'temp');
      if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
      }
      process.exit(0);
    }
    await writeTempData({ ...userConfig, ...transferData });
    progress.completeStep('User Confirmation');

    progress.addStep('Transfer Creation');
    logger.info('Starting transfer process with data:', userConfig);
    await createTransferProcess({ ...userConfig, ...transferData });
    progress.completeStep('Transfer Creation');
    logger.info('Transfer process completed successfully');
  } catch (error) {
    logger.error('Failed to process transfer:', error);
    progress.addError('Process', error);
    process.exit(1);
  } finally {
    progress.finish();
  }
})();