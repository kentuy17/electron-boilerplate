'use strict';

const {contextBridge, ipcRenderer} = require('electron');

contextBridge.exposeInMainWorld('api', {
	listRecords: () => ipcRenderer.invoke('records:list'),
	createRecord: record => ipcRenderer.invoke('records:create', record),
	updateRecord: record => ipcRenderer.invoke('records:update', record),
	deleteRecord: id => ipcRenderer.invoke('records:delete', id),
	exportToExcel: () => ipcRenderer.invoke('records:export'),
});
