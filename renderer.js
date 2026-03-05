'use strict';

const importButton = document.querySelector('#import-btn');
const exportButton = document.querySelector('#export-btn');
const templateButton = document.querySelector('#template-btn');
const tableBody = document.querySelector('#records-body');
const statusText = document.querySelector('#status-text');

const setStatus = message => {
	statusText.textContent = message;
};

const renderRecords = records => {
	tableBody.innerHTML = '';

	for (const record of records) {
		const row = document.createElement('tr');
		row.innerHTML = `
			<td>${record.id}</td>
			<td>${record.applicant_name}</td>
			<td>${record.applicant_code}</td>
		`;
		tableBody.append(row);
	}
};

const loadRecords = async () => {
	const records = await window.api.listRecords();
	renderRecords(records);
};

importButton.addEventListener('click', async () => {
	const result = await window.api.importApplicants();
	if (result.canceled) {
		setStatus('Import canceled.');
		return;
	}

	if (!result.ok) {
		setStatus(result.message);
		return;
	}

	await loadRecords();
	setStatus(`Imported ${result.insertedCount} applicant record(s).`);
});

exportButton.addEventListener('click', async () => {
	const result = await window.api.exportToExcel();
	if (result.canceled) {
		setStatus('Export canceled.');
		return;
	}

	setStatus(`Excel exported to ${result.filePath}`);
});

templateButton.addEventListener('click', async () => {
	const result = await window.api.downloadTemplate();
	if (result.canceled) {
		setStatus('Template download canceled.');
		return;
	}

	setStatus(`Template saved to ${result.filePath}`);
});

(async () => {
	await loadRecords();
	setStatus('Ready.');
})();
