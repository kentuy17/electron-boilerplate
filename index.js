'use strict';

const path = require('path');
const fs = require('fs/promises');
const {app, BrowserWindow, Menu, ipcMain, dialog} = require('electron');
/// const {autoUpdater} = require('electron-updater');
const {is} = require('electron-util');
const unhandled = require('electron-unhandled');
const debug = require('electron-debug');
const contextMenu = require('electron-context-menu');
const sqlite3 = require('sqlite3').verbose();
const xlsx = require('xlsx');
const menu = require('./menu.js');

unhandled();
debug();
contextMenu();

app.setAppUserModelId('com.company.AppName');

let mainWindow;
let db;

const initDatabase = () => new Promise((resolve, reject) => {
	const databasePath = path.join(app.getPath('userData'), 'records.sqlite');
	db = new sqlite3.Database(databasePath, error => {
		if (error) {
			reject(error);
			return;
		}

		db.run(
			'CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, email TEXT NOT NULL)',
			tableError => {
				if (tableError) {
					reject(tableError);
					return;
				}

				resolve();
			},
		);
	});
});

const runQuery = (query, parameters = []) => new Promise((resolve, reject) => {
	db.run(query, parameters, function (error) {
		if (error) {
			reject(error);
			return;
		}

		resolve({id: this.lastID, changes: this.changes});
	});
});

const allQuery = (query, parameters = []) => new Promise((resolve, reject) => {
	db.all(query, parameters, (error, rows) => {
		if (error) {
			reject(error);
			return;
		}

		resolve(rows);
	});
});

ipcMain.handle('records:list', async () => allQuery('SELECT id, name, email FROM records ORDER BY id DESC'));

ipcMain.handle('records:create', async (_event, record) => {
	const result = await runQuery('INSERT INTO records (name, email) VALUES (?, ?)', [record.name, record.email]);
	return {id: result.id};
});

ipcMain.handle('records:update', async (_event, record) => {
	await runQuery('UPDATE records SET name = ?, email = ? WHERE id = ?', [record.name, record.email, record.id]);
	return {ok: true};
});

ipcMain.handle('records:delete', async (_event, id) => {
	await runQuery('DELETE FROM records WHERE id = ?', [id]);
	return {ok: true};
});

ipcMain.handle('records:export', async () => {
	const rows = await allQuery('SELECT id, name, email FROM records ORDER BY id ASC');
	const worksheet = xlsx.utils.json_to_sheet(rows);
	const workbook = xlsx.utils.book_new();
	xlsx.utils.book_append_sheet(workbook, worksheet, 'Records');

	const defaultPath = path.join(app.getPath('documents'), 'records.xlsx');
	const {canceled, filePath} = await dialog.showSaveDialog({
		title: 'Export records to Excel',
		defaultPath,
		filters: [{name: 'Excel Workbook', extensions: ['xlsx']}],
	});

	if (canceled || !filePath) {
		return {canceled: true};
	}

	const buffer = xlsx.write(workbook, {type: 'buffer', bookType: 'xlsx'});
	await fs.writeFile(filePath, buffer);
	return {canceled: false, filePath};
});

const createMainWindow = async () => {
	const window_ = new BrowserWindow({
		title: app.name,
		show: false,
		width: 900,
		height: 680,
		webPreferences: {
			preload: path.join(__dirname, 'preload.js'),
		},
	});

	window_.on('ready-to-show', () => {
		window_.show();
	});

	window_.on('closed', () => {
		mainWindow = undefined;
	});

	await window_.loadFile(path.join(__dirname, 'index.html'));
	return window_;
};

if (!app.requestSingleInstanceLock()) {
	app.quit();
}

app.on('second-instance', () => {
	if (mainWindow) {
		if (mainWindow.isMinimized()) {
			mainWindow.restore();
		}

		mainWindow.show();
	}
});

app.on('window-all-closed', () => {
	if (!is.macos) {
		app.quit();
	}
});

app.on('activate', async () => {
	if (!mainWindow) {
		mainWindow = await createMainWindow();
	}
});

app.on('before-quit', () => {
	if (db) {
		db.close();
	}
});

(async () => {
	await app.whenReady();
	await initDatabase();
	Menu.setApplicationMenu(menu);
	mainWindow = await createMainWindow();
})();
