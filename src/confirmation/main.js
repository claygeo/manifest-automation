// src/confirmation/main.js
const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const config = require('../config');

let mainWindow;

/**
 * Scans the given directory for Excel files based on the supported formats
 * specified in your config.
 */
function scanExcelFiles(rootDir) {
  const supportedFormats = config.get('excel.supportedFormats') || ['.xlsx', '.xls', '.xlsm'];
  console.log("Scanning directory for Excel files:", rootDir);
  let files = [];
  try {
    files = fs.readdirSync(rootDir);
  } catch (err) {
    console.error("Error reading directory:", err);
  }
  const excelFiles = files
    .filter(file => supportedFormats.some(ext => file.toLowerCase().endsWith(ext)))
    .map(file => path.join(rootDir, file));
  console.log("Found Excel files:", excelFiles);
  return excelFiles;
}

/**
 * Legacy parser: Parses an Excel file using XLSX.
 * This version reads the first sheet, uses the first row as headers,
 * and converts subsequent rows into objects.
 * (This parser is used only if no approved JSON is present.)
 */
function parseExcelFile(filePath) {
  try {
    console.log("Parsing Excel file using legacy parser:", filePath);
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const transferData = { products: [] };

    // Extract headers from the first row
    const range = XLSX.utils.decode_range(worksheet['!ref']);
    const headers = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
      const cell = worksheet[cellAddress];
      headers.push(cell ? cell.v : '');
    }
    
    // Parse each row (starting at row 2)
    for (let row = range.s.r + 1; row <= range.e.r; row++) {
      const rowData = {};
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[cellAddress];
        rowData[headers[col - range.s.c]] = cell ? cell.v : '';
      }
      transferData.products.push(rowData);
    }
    
    return transferData;
  } catch (error) {
    console.error("Error parsing Excel file:", error);
    return null;
  }
}

/**
 * Creates the confirmation window.
 * If an approved JSON file exists, load that data; otherwise, show the Excel file list.
 */
function createWindow(rootDir) {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 700,
    resizable: false,
    title: 'Transfer Data Confirmation',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'renderer.js')
    }
  });
  
  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  mainWindow.webContents.on('did-finish-load', () => {
    console.log("Confirmation window finished loading.");
    // Check if approved JSON exists.
    const tempFile = path.join(rootDir, 'temp', 'transferData.json');
    if (fs.existsSync(tempFile)) {
      console.log("Approved JSON found. Loading approved transfer data from:", tempFile);
      try {
        const rawData = fs.readFileSync(tempFile, 'utf-8');
        const transferData = JSON.parse(rawData);
        mainWindow.webContents.send('transfer-data', transferData);
      } catch (err) {
        console.error("Error loading approved transfer data:", err);
        // If error, fall back to showing Excel file list.
        const excelFiles = scanExcelFiles(rootDir);
        mainWindow.webContents.send('excel-files', excelFiles);
      }
    } else {
      // No approved JSON file—send the Excel files list.
      const excelFiles = scanExcelFiles(rootDir);
      mainWindow.webContents.send('excel-files', excelFiles);
    }
  });
  
  // When the user selects an Excel file, parse it using the legacy parser.
  ipcMain.on('select-excel', (event, filePath) => {
    console.log("Excel file selected (legacy parse):", filePath);
    const transferData = parseExcelFile(filePath);
    mainWindow.webContents.send('transfer-data', transferData);
  });
  
  // When the user approves, write the approved JSON file.
  ipcMain.on('approval', (event, finalData) => {
    const tempDir = path.join(process.argv[2] || process.cwd(), 'temp');
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
    const tempFile = path.join(tempDir, 'transferData.json');
    fs.writeFileSync(tempFile, JSON.stringify(finalData, null, 2), 'utf-8');
    console.log("Approved transfer data written to:", tempFile);
    app.exit(0);
  });
  
  ipcMain.on('cancel', () => {
    app.exit(1);
  });
}

app.whenReady().then(() => {
  // Use process.argv[2] if provided; otherwise use the current working directory.
  const rootDir = process.argv[2] || process.cwd();
  console.log("Confirmation window root directory:", rootDir);
  createWindow(rootDir);
});

app.on('activate', function () {
  if (BrowserWindow.getAllWindows().length === 0)
    createWindow(process.argv[2] || process.cwd());
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});
