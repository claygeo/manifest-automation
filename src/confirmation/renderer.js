// src/confirmation/renderer.js
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  onExcelFiles: (callback) => ipcRenderer.on('excel-files', (event, files) => callback(files)),
  selectExcel: (filePath) => ipcRenderer.send('select-excel', filePath),
  onTransferData: (callback) => ipcRenderer.on('transfer-data', (event, data) => callback(data)),
  sendApproval: (data) => ipcRenderer.send('approval', data),
  sendCancel: () => ipcRenderer.send('cancel'),
  onError: (callback) => ipcRenderer.on('error', (event, error) => callback(error))
});
