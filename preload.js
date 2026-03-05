'use strict';

const {contextBridge, ipcRenderer} = require('electron');

contextBridge.exposeInMainWorld('api', {
	listRecords: () => ipcRenderer.invoke('records:list'),
	importApplicants: () => ipcRenderer.invoke('records:import'),
	exportToExcel: () => ipcRenderer.invoke('records:export'),
	downloadTemplate: () => ipcRenderer.invoke('records:download-template'),
});
