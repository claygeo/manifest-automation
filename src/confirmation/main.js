const { app, BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');

let mainWindow;

function createWindow(rootDir) {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 800,
    resizable: false,
    title: 'Transfer Confirmation',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'index.html'));

  mainWindow.webContents.on('did-finish-load', () => {
    const excelPath = path.join(rootDir, 'temp', 'selectedExcel.json');
    if (fs.existsSync(excelPath)) {
      const rawData = fs.readFileSync(excelPath, 'utf-8');
      const { filePath } = JSON.parse(rawData);
      const workbook = XLSX.readFile(filePath);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const transferData = [];
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      
      // Dynamically read headers from the first row
      const headers = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const cell = worksheet[cellAddress];
        headers.push(cell && cell.v ? cell.v.toString().trim() : '');
      }

      // Extract store names from headers
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
        if (rowData['#']) transferData.push(rowData);
      }

      const currentDate = new Date().toLocaleDateString('en-US');
      mainWindow.webContents.send('initial-data', { transferData, currentDate, stores: [store1, store2], headers });
    }
  });

  ipcMain.on('approve-transfer', (event, config) => {
    const tempDir = path.join(rootDir, 'temp');
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
    fs.writeFileSync(path.join(tempDir, 'transferConfig.json'), JSON.stringify(config));
    mainWindow.close();
  });

  ipcMain.on('cancel-transfer', () => {
    const tempDir = path.join(rootDir, 'temp');
    if (!fs.existsSync(tempDir)) fs.mkdirSync(tempDir, { recursive: true });
    fs.writeFileSync(path.join(tempDir, 'cancelFlag.json'), JSON.stringify({ canceled: true }));
    mainWindow.close();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

app.whenReady().then(() => {
  const rootDir = process.argv[2] || process.cwd();
  createWindow(rootDir);
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (mainWindow === null) createWindow(process.argv[2] || process.cwd());
});