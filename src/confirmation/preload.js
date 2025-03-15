// src/confirmation/preload.js
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  sendApproval: (data) => ipcRenderer.send('approve-transfer', data),
  sendCancel: () => ipcRenderer.send('cancel-transfer'),
  onInitialData: (callback) => ipcRenderer.on('initial-data', (event, data) => callback(data))
});