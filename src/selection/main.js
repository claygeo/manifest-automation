// src/selection/main.js
const { app, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const fsPromises = fs.promises;

let rootDir = process.argv[2] || process.cwd();
let tempDir = path.join(rootDir, 'temp');

async function selectFile() {
  // Ensure the temporary directory exists.
  if (!fs.existsSync(tempDir)) {
    fs.mkdirSync(tempDir, { recursive: true });
  }

  // Open the file dialog with filters for Excel files.
  const result = await dialog.showOpenDialog({
    title: 'Select an Excel File',
    defaultPath: rootDir,
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  });

  if (result.canceled || result.filePaths.length === 0) {
    app.exit(1); // Indicate cancellation/failure.
  } else {
    // Write the selected file path to a temporary JSON file.
    const selectedData = { filePath: result.filePaths[0] };
    const tempFilePath = path.join(tempDir, 'selectedExcel.json');
    await fsPromises.writeFile(tempFilePath, JSON.stringify(selectedData, null, 2), 'utf-8');
    app.exit(0);
  }
}

app.whenReady().then(selectFile);
