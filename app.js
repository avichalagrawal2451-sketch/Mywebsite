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
  if (typeof XLSX === 'undefined') {
    alert('Excel export library is not available. Please refresh and try again.');
    return;
  }
  if (!expenses.length) return alert('Add at least one expense to export.');
  const wb = XLSX.utils.book_new();

  const reportSheet = XLSX.utils.json_to_sheet(buildExportRows());
  XLSX.utils.book_append_sheet(wb, reportSheet, 'Expense Report');

  const attachments = expenses.flatMap((expense, expenseIndex) =>
    (expense.attachments || []).map((file, attachmentIndex) => {
      const [header, base64 = ''] = String(file.dataUrl || '').split(',');
      return {
        expenseIndex,
        attachmentIndex,
        expense,
        file,
        mimeHeader: header,
        base64,
      };
    })
  );

  const usedSheetNames = new Set(['Expense Report', 'Attachments']);
  const makeSheetName = (expenseIndex, attachmentIndex, fileName) => {
    const normalized = String(fileName || 'Attachment').replace(/[\\/?*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim();
    const prefix = `E${expenseIndex + 1}-A${attachmentIndex + 1}`;
    const maxBaseLength = 31 - prefix.length - 1;
    const trimmed = normalized.slice(0, Math.max(1, maxBaseLength));
    let name = `${prefix}-${trimmed}`;
    let duplicateCounter = 1;
    while (usedSheetNames.has(name)) {
      const suffix = `-${duplicateCounter++}`;
      name = `${prefix}-${trimmed}`.slice(0, 31 - suffix.length) + suffix;
    }
    usedSheetNames.add(name);
    return name;
  };

  const chunkBase64 = (value, size = 30000) => {
    const chunks = [];
    for (let i = 0; i < value.length; i += size) {
      chunks.push(value.slice(i, i + size));
    }
    return chunks.length ? chunks : [''];
  };

  const attachmentRows = attachments.length
    ? attachments.map((item) => {
      const sheetName = makeSheetName(item.expenseIndex, item.attachmentIndex, item.file.name);
      const detailRows = [
        { Field: 'Filename', Value: item.file.name },
        { Field: 'MIME Type', Value: item.file.type || 'application/octet-stream' },
        { Field: 'Size (Bytes)', Value: item.file.size },
        { Field: 'Expense #', Value: item.expenseIndex + 1 },
        { Field: 'Expense Date', Value: item.expense.date },
        { Field: 'Expense Type', Value: item.expense.type },
        { Field: 'Data URL Header', Value: item.mimeHeader },
        { Field: 'Base64 Chunks', Value: chunkBase64(item.base64).length },
        { Field: 'How to rebuild', Value: 'Join all Base64 chunks in this sheet and prepend Data URL Header + comma.' },
      ];
      const detailSheet = XLSX.utils.json_to_sheet(detailRows);
      XLSX.utils.book_append_sheet(wb, detailSheet, sheetName);
      const chunks = chunkBase64(item.base64);
      XLSX.utils.sheet_add_json(
        detailSheet,
        chunks.map((chunk, i) => ({ 'Chunk #': i + 1, Base64: chunk })),
        { origin: 'A12' }
      );

      return {
        '#': item.attachmentIndex + 1,
        'Expense #': item.expenseIndex + 1,
        'Expense Date': item.expense.date,
        'Expense Type': item.expense.type,
        Filename: item.file.name,
        MIME: item.file.type,
        SizeBytes: item.file.size,
        'Stored In Sheet': sheetName,
      };
    })
    : [{ Note: 'No attachments uploaded.' }];

  const attachmentSheet = XLSX.utils.json_to_sheet(attachmentRows);
  XLSX.utils.book_append_sheet(wb, attachmentSheet, 'Attachments');

  XLSX.writeFile(wb, 'expense-report-merged.xlsx', { compression: true });
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
