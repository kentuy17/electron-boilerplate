'use strict';

const form = document.querySelector('#record-form');
const idField = document.querySelector('#record-id');
const nameField = document.querySelector('#name');
const emailField = document.querySelector('#email');
const submitButton = document.querySelector('#submit-btn');
const cancelEditButton = document.querySelector('#cancel-edit-btn');
const exportButton = document.querySelector('#export-btn');
const tableBody = document.querySelector('#records-body');
const statusText = document.querySelector('#status-text');

const setStatus = message => {
	statusText.textContent = message;
};

const resetForm = () => {
	idField.value = '';
	form.reset();
	submitButton.textContent = 'Add Record';
	cancelEditButton.hidden = true;
};

const fillFormForEdit = record => {
	idField.value = String(record.id);
	nameField.value = record.name;
	emailField.value = record.email;
	submitButton.textContent = 'Update Record';
	cancelEditButton.hidden = false;
};

const renderRecords = records => {
	tableBody.innerHTML = '';

	for (const record of records) {
		const row = document.createElement('tr');
		row.innerHTML = `
			<td>${record.id}</td>
			<td>${record.name}</td>
			<td>${record.email}</td>
			<td class="actions-cell"></td>
		`;

		const actionsCell = row.querySelector('.actions-cell');
		const editButton = document.createElement('button');
		editButton.textContent = 'Edit';
		editButton.className = 'small-btn';
		editButton.addEventListener('click', () => fillFormForEdit(record));

		const deleteButton = document.createElement('button');
		deleteButton.textContent = 'Delete';
		deleteButton.className = 'small-btn danger-btn';
		deleteButton.addEventListener('click', async () => {
			await window.api.deleteRecord(record.id);
			await loadRecords();
			if (Number(idField.value) === record.id) {
				resetForm();
			}

			setStatus(`Deleted record #${record.id}`);
		});

		actionsCell.append(editButton, deleteButton);
		tableBody.append(row);
	}
};

const loadRecords = async () => {
	const records = await window.api.listRecords();
	renderRecords(records);
};

form.addEventListener('submit', async event => {
	event.preventDefault();

	const record = {
		name: nameField.value.trim(),
		email: emailField.value.trim(),
	};

	if (!record.name || !record.email) {
		setStatus('Name and email are required.');
		return;
	}

	if (idField.value) {
		await window.api.updateRecord({
			id: Number(idField.value),
			...record,
		});
		setStatus(`Updated record #${idField.value}`);
	} else {
		const created = await window.api.createRecord(record);
		setStatus(`Created record #${created.id}`);
	}

	resetForm();
	await loadRecords();
});

cancelEditButton.addEventListener('click', () => {
	resetForm();
	setStatus('Edit canceled.');
});

exportButton.addEventListener('click', async () => {
	const result = await window.api.exportToExcel();
	if (result.canceled) {
		setStatus('Export canceled.');
		return;
	}

	setStatus(`Excel exported to ${result.filePath}`);
});

(async () => {
	await loadRecords();
	setStatus('Ready.');
})();
