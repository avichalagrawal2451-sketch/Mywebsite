const expenseForm = document.getElementById('expenseForm');
const expenseTableBody = document.querySelector('#expenseTable tbody');
const totalAmountEl = document.getElementById('totalAmount');
const fileUploader = document.getElementById('fileUploader');
const dropzone = document.getElementById('dropzone');
const attachmentList = document.getElementById('attachmentList');

const expenses = [];
const uploadedFiles = [];

document.getElementById('expenseDate').valueAsDate = new Date();

function renderAttachments() {
  attachmentList.innerHTML = '';
  uploadedFiles.forEach((file, i) => {
    const li = document.createElement('li');
    li.textContent = `${i + 1}. ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    attachmentList.appendChild(li);
  });
}

function filesToDataEntries(files) {
  return Promise.all([...files].map((file) => new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = () => resolve({
      name: file.name,
      type: file.type || 'application/octet-stream',
      size: file.size,
      dataUrl: reader.result,
    });
    reader.readAsDataURL(file);
  })));
}

fileUploader.addEventListener('change', async (e) => {
  const entries = await filesToDataEntries(e.target.files);
  uploadedFiles.push(...entries);
  renderAttachments();
  expenseForm.receiptAttached.value = uploadedFiles.length ? 'Yes' : 'No';
});

['dragenter', 'dragover'].forEach((eventName) => {
  dropzone.addEventListener(eventName, (e) => {
    e.preventDefault();
    dropzone.classList.add('dragging');
  });
});
['dragleave', 'drop'].forEach((eventName) => {
  dropzone.addEventListener(eventName, (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragging');
  });
});
dropzone.addEventListener('drop', async (e) => {
  const entries = await filesToDataEntries(e.dataTransfer.files);
  uploadedFiles.push(...entries);
  renderAttachments();
  expenseForm.receiptAttached.value = uploadedFiles.length ? 'Yes' : 'No';
});

function renderTable() {
  expenseTableBody.innerHTML = '';
  let total = 0;
  expenses.forEach((item, index) => {
    total += Number(item.amount);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${item.date}</td>
      <td>${item.type}</td>
      <td>${item.description}</td>
      <td>${item.paymentMode}</td>
      <td>${Number(item.amount).toFixed(2)}</td>
      <td>${item.receiptAttached}</td>
    `;
    expenseTableBody.appendChild(tr);
  });
  totalAmountEl.textContent = total.toFixed(2);
}

expenseForm.addEventListener('submit', (e) => {
  e.preventDefault();
  const attachments = uploadedFiles.map((file) => ({ ...file }));
  const entry = {
    type: expenseForm.expenseType.value,
    description: expenseForm.description.value.trim(),
    amount: expenseForm.amount.value,
    date: expenseForm.expenseDate.value,
    paymentMode: expenseForm.paymentMode.value,
    receiptAttached: expenseForm.receiptAttached.value,
    attachmentCount: attachments.length,
    attachments,
  };
  expenses.push(entry);
  renderTable();

  uploadedFiles.length = 0;
  renderAttachments();

  expenseForm.reset();
  expenseForm.expenseDate.valueAsDate = new Date();
  expenseForm.receiptAttached.value = 'No';
});

document.getElementById('clearBtn').addEventListener('click', () => {
  expenseForm.reset();
  expenseForm.expenseDate.valueAsDate = new Date();
  expenseForm.receiptAttached.value = uploadedFiles.length ? 'Yes' : 'No';
});

function buildExportRows() {
  return expenses.map((item, index) => ({
    '#': index + 1,
    Date: item.date,
    Type: item.type,
    Description: item.description,
    'Payment Mode': item.paymentMode,
    Amount: Number(item.amount).toFixed(2),
    'Receipt Attached': item.receiptAttached,
    'Attached Files Count': item.attachmentCount,
    'Attachment Names': (item.attachments || []).map((file) => file.name).join(', '),
  }));
}

function exportCSV() {
  const rows = buildExportRows();
  if (!rows.length) return alert('Add at least one expense to export.');
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(',')]
    .concat(rows.map((row) => headers.map((h) => `"${String(row[h]).replace(/"/g, '""')}"`).join(',')))
    .join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8;' }));
  a.download = 'expense-report.csv';
  a.click();
}

function exportExcel() {
  if (!expenses.length) return alert('Add at least one expense to export.');
  const wb = XLSX.utils.book_new();

  const reportSheet = XLSX.utils.json_to_sheet(buildExportRows());
  XLSX.utils.book_append_sheet(wb, reportSheet, 'Expense Report');

  const allAttachments = expenses.flatMap((expense, expenseIndex) =>
    (expense.attachments || []).map((file, attachmentIndex) => ({
      '#': attachmentIndex + 1,
      'Expense #': expenseIndex + 1,
      'Expense Date': expense.date,
      'Expense Type': expense.type,
      Filename: file.name,
      MIME: file.type,
      SizeBytes: file.size,
      DataURL: file.dataUrl,
      Note: 'DataURL contains encoded file content for merged export.',
    }))
  );

  const attachmentRows = allAttachments.length
    ? allAttachments
    : [{ Note: 'No attachments uploaded.' }];

  const attachmentSheet = XLSX.utils.json_to_sheet(attachmentRows);
  XLSX.utils.book_append_sheet(wb, attachmentSheet, 'Attachments');

  XLSX.writeFile(wb, 'expense-report-merged.xlsx');
}

function exportPDF() {
  if (!expenses.length) return alert('Add at least one expense to export.');
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  doc.setFontSize(16);
  doc.text('Expense Report', 14, 16);
  doc.autoTable({
    startY: 24,
    head: [["#", "Date", "Type", "Description", "Payment", "Amount", "Receipt"]],
    body: expenses.map((e, i) => [i + 1, e.date, e.type, e.description, e.paymentMode, Number(e.amount).toFixed(2), e.receiptAttached]),
  });
  const y = doc.lastAutoTable.finalY + 10;
  doc.text(`Total Amount: ₹${totalAmountEl.textContent}`, 14, y);
  doc.text(`Uploaded Attachments: ${uploadedFiles.length}`, 14, y + 8);
  doc.save('expense-report.pdf');
}

function printReport() {
  if (!expenses.length) return alert('Add at least one expense to print.');
  window.print();
}

document.querySelectorAll('[data-export]').forEach((btn) => {
  btn.addEventListener('click', () => {
    const type = btn.dataset.export;
    if (type === 'csv') exportCSV();
    if (type === 'excel') exportExcel();
    if (type === 'pdf') exportPDF();
    if (type === 'print') printReport();
  });
});
