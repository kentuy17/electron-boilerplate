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

		db.serialize(() => {
			db.run(
				'CREATE TABLE IF NOT EXISTS records (id INTEGER PRIMARY KEY AUTOINCREMENT, applicant_name TEXT NOT NULL, applicant_code TEXT NOT NULL)',
				tableError => {
					if (tableError) {
						reject(tableError);
					}
				},
			);

			db.all('PRAGMA table_info(records)', (schemaError, columns) => {
				if (schemaError) {
					reject(schemaError);
					return;
				}

				const hasExpectedSchema = columns.some(column => column.name === 'applicant_name')
					&& columns.some(column => column.name === 'applicant_code');

				if (hasExpectedSchema) {
					resolve();
					return;
				}

				db.run('DROP TABLE IF EXISTS records', dropError => {
					if (dropError) {
						reject(dropError);
						return;
					}

					db.run(
						'CREATE TABLE records (id INTEGER PRIMARY KEY AUTOINCREMENT, applicant_name TEXT NOT NULL, applicant_code TEXT NOT NULL)',
						recreateError => {
							if (recreateError) {
								reject(recreateError);
								return;
							}

							resolve();
						},
					);
				});
			});
		});
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

ipcMain.handle('records:list', async () => allQuery('SELECT id, applicant_name, applicant_code FROM records ORDER BY id DESC'));

const templateRows = [
	['applicant_name', 'applicant_code'],
	['Jane Doe', 'APP-1001'],
	['John Smith', 'APP-1002'],
];

const saveTemplateFile = async filePath => {
	const worksheet = xlsx.utils.aoa_to_sheet(templateRows);
	const workbook = xlsx.utils.book_new();
	xlsx.utils.book_append_sheet(workbook, worksheet, 'Applicants');
	const buffer = xlsx.write(workbook, {type: 'buffer', bookType: 'xlsx'});
	await fs.writeFile(filePath, buffer);
};

ipcMain.handle('records:download-template', async () => {
	const defaultPath = path.join(app.getPath('documents'), 'applicant-import-template.xlsx');
	const {canceled, filePath} = await dialog.showSaveDialog({
		title: 'Download applicant import template',
		defaultPath,
		filters: [{name: 'Excel Workbook', extensions: ['xlsx']}],
	});

	if (canceled || !filePath) {
		return {canceled: true};
	}

	await saveTemplateFile(filePath);
	return {canceled: false, filePath};
});

ipcMain.handle('records:import', async () => {
	const {canceled, filePaths} = await dialog.showOpenDialog({
		title: 'Import applicant file',
		properties: ['openFile'],
		filters: [{name: 'Spreadsheet Files', extensions: ['xlsx', 'xls', 'csv']}],
	});

	if (canceled || !filePaths || filePaths.length === 0) {
		return {canceled: true};
	}

	const workbook = xlsx.readFile(filePaths[0]);
	const firstSheetName = workbook.SheetNames[0];
	if (!firstSheetName) {
		return {ok: false, message: 'The selected file has no sheets.'};
	}

	const worksheet = workbook.Sheets[firstSheetName];
	const rows = xlsx.utils.sheet_to_json(worksheet, {defval: ''});
	const requiredColumns = ['applicant_name', 'applicant_code'];
	const availableColumns = Object.keys(rows[0] || {});
	const missingColumns = requiredColumns.filter(column => !availableColumns.includes(column));

	if (missingColumns.length > 0) {
		const defaultTemplatePath = path.join(app.getPath('documents'), 'applicant-import-template.xlsx');
		const templateSaveResult = await dialog.showSaveDialog({
			title: 'Missing required columns. Save sample template',
			defaultPath: defaultTemplatePath,
			filters: [{name: 'Excel Workbook', extensions: ['xlsx']}],
		});

		if (!templateSaveResult.canceled && templateSaveResult.filePath) {
			await saveTemplateFile(templateSaveResult.filePath);
		}

		return {
			ok: false,
			message: 'Import failed. File must include applicant_name and applicant_code columns. A sample template is available for download.',
		};
	}

	const normalizedRows = rows
		.map(row => ({
			applicantName: String(row.applicant_name).trim(),
			applicantCode: String(row.applicant_code).trim(),
		}))
		.filter(row => row.applicantName && row.applicantCode);

	for (const row of normalizedRows) {
		// eslint-disable-next-line no-await-in-loop
		await runQuery('INSERT INTO records (applicant_name, applicant_code) VALUES (?, ?)', [row.applicantName, row.applicantCode]);
	}

	return {ok: true, insertedCount: normalizedRows.length};
});

ipcMain.handle('records:export', async () => {
	const rows = await allQuery('SELECT id, applicant_name, applicant_code FROM records ORDER BY id ASC');
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
