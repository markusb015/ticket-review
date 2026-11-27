/* global XLSX */

const CONFIG = {
  sourceSheet: 'All Tickets',
  alertKeywords: [
    ['OPM availability alerts', 'OPM'],
    ['OPM Infra Alerts', 'OPM'],
    ['OPM Network Availability Devices Alerts', 'OPM'],
    ['OPM Network Firewall Devices Alerts', 'OPM'],
    ['[Site24x7]', 'Site24x7'],
    ['Azure monitor alert', 'Azure'],
  ],
  pendingStatuses: new Set([
    'pending user information', 'pending user validation', 'pending 3rd party',
    'pending user information np', 'pending user validation np', 'pending 3rd party np',
  ]),
  phTechnicians: new Set([
    'abigail.cruz@inchcape.com', 'adrian.pascua@inchcape.com', 'adrian.catubay@inchcape.com', 'ann.mercado@inchcape.com', 'david.flores@inchcape.com', 'david.libao@inchcape.com', 'hazel.calingasan@inchcape.com', 'hazel.lumabad@inchcape.com', 'jasper.rodriguez@inchcape.com', 'kristine.reyes@inchcape.com', 'markus.bautista@inchcape.com', 'michole.delvalle@inchcape.com', 'mika.manansala@inchcape.com', 'patricia.sampag@inchcape.com', 'rayann.gamutlong@inchcape.com', 'robin.reyes@inchcape.com', 'rosebhil.salapi@inchcape.com', 'staceymarie.estrada@inchcape.com', 'carmelamarie.ko@inchcape.com',
  ]),
  colTechnicians: new Set([
    'andrea.beleno@inchcape.com', 'andres.toro@inchcape.com', 'cesar.lancheros@inchcape.com', 'damian.gil@inchcape.com', 'daniela.pelaez@inchcape.com', 'felix.uriza@inchcape.com', 'javierl.suarez@inchcape.com', 'jhonatan.baron@inchcape.com', 'juan.sierrap@inchcape.com', 'juana.plaza@inchcape.com', 'julian.prias@inchcape.com', 'katerin.amariles@inchcape.com', 'kevin.alvarez@inchcape.com', 'kevin.rincons@inchcape.com', 'lina.andrade@inchcape.com', 'maria.tobon@inchcape.com', 'paula.bustos@inchcape.com',
  ]),
  columns: {
    classified: ['#', 'RequestID', 'Technician', 'Status', 'Region', 'No Update', 'Update needed', 'OnSheet', 'Subject', 'Priority', 'Last Updated Time', 'Last Updated By Email', 'UnReplied Count', 'Created Time', 'Due By Time'],
    status: ['#', 'RequestID', 'Technician', 'Status', 'Region', 'No Update', 'Update needed', 'OnSheet', 'Subject', 'Priority', 'Last Updated Time', 'Last Updated By Email', 'UnReplied Count', 'Created Time', 'Due By Time'],
    alert: ['#', 'RequestID', 'Technician', 'Status', 'Region', 'No Update', 'Update needed', 'OnSheet', 'Subject', 'Priority', 'Last Updated Time', 'Last Updated By Email', 'UnReplied Count', 'Created Time', 'Due By Time'],
    master: ['#', 'RequestID', 'Technician', 'Status', 'Region', 'No Update', 'Update needed', 'OnSheet', 'Subject', 'Priority', 'Last Updated Time', 'Last Updated By Email', 'UnReplied Count', 'Created Time', 'Due By Time'],
  },
};

CONFIG.allowedTechnicians = new Set([...CONFIG.phTechnicians, ...CONFIG.colTechnicians]);

const state = { sheets: {}, activeSheet: 'All Tickets', technicianGroup: 'all', visibleColumns: {}, done: new Set(), noUpdate: new Set(), updateNeeded: new Map(), edits: new Map(), reportTime: null, sort: { column: null, direction: 'asc' } };
const $ = (selector) => document.querySelector(selector);
const text = (value) => value === undefined || value === null ? '' : String(value);
const normalise = (value) => text(value).trim().replace(/\s+/g, ' ').toLowerCase();
const getDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') return XLSX.SSF.parse_date_code(value) ? new Date((value - 25569) * 86400 * 1000) : null;
  const date = new Date(value); return Number.isNaN(date) ? null : date;
};
const formatDate = (value) => { const date = getDate(value); return date ? date.toLocaleString('sv-SE', { hour12: false }).replace('T', ' ') : text(value); };
const isAlert = (ticket) => CONFIG.alertKeywords.some(([key]) => normalise(ticket.Subject).includes(normalise(key)));
const classification = (ticket) => {
  const match = CONFIG.alertKeywords.find(([key]) => normalise(ticket.Subject).includes(normalise(key)));
  if (match) return match[1];
  const status = normalise(ticket.Status);
  if (status === 'work in progress') return 'Work in Progress';
  if (CONFIG.pendingStatuses.has(status)) return 'Pending';
  if (status === 'waiting approval') return 'Waiting Approval';
  return 'Unclassified';
};
const isOverdue = (ticket) => { const date = getDate(ticket['Due By Time']); return date && date < state.reportTime; };
const hasUnreplied = (ticket) => Number(ticket['UnReplied Count'] || 0) > 0;
const isNP = (ticket) => { const s = normalise(ticket.Status); return s.startsWith('pending') && s.endsWith(' np'); };
const technicianEmail = (ticket) => normalise(ticket['Technician Email'] || ticket.Technician);
const matchesTechnicianGroup = (ticket) => state.technicianGroup === 'ph' ? CONFIG.phTechnicians.has(technicianEmail(ticket)) : state.technicianGroup === 'col' ? CONFIG.colTechnicians.has(technicianEmail(ticket)) : true;
const tableSort = (rows, field, newestFirst) => [...rows].sort((a, b) => {
  const left = getDate(a[field]), right = getDate(b[field]);
  if (!left) return right ? 1 : 0; if (!right) return -1;
  return newestFirst ? right - left : left - right;
});

function updateClock() { $('#currentTime').textContent = new Date().toLocaleString('sv-SE', { hour12: false }).replace('T', ' '); }
setInterval(updateClock, 1000); updateClock();

function findHeaderRow(rows) {
  const required = ['RequestID', 'Technician', 'Request Status', 'Region', 'Subject'];
  return rows.findIndex((row, i) => i < 20 && required.every((key) => row.map(text).includes(key)));
}

function loadFile(file) {
  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true });
      if (!workbook.SheetNames.includes(CONFIG.sourceSheet)) throw new Error(`The workbook does not contain an “${CONFIG.sourceSheet}” sheet.`);
      const sheet = workbook.Sheets[CONFIG.sourceSheet];
      const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: true });
      const headerIndex = findHeaderRow(rawRows);
      if (headerIndex < 0) throw new Error('Could not find the required All Tickets column headers in the first 20 rows.');
      const headers = rawRows[headerIndex].map(text);
      const range = XLSX.utils.decode_range(sheet['!ref']);
      const sourceRows = rawRows.slice(headerIndex + 1).map((row, offset) => ({ row, offset })).filter(({ row }) => row.some((value) => value !== '')).map(({ row, offset }) => {
        const ticket = Object.fromEntries(headers.map((header, index) => [header, row[index] ?? '']));
        const requestCell = sheet[XLSX.utils.encode_cell({ r: headerIndex + 1 + offset, c: headers.indexOf('RequestID') })];
        ticket.Status = ticket['Request Status'];
        ticket.url = requestCell?.l?.Target || ''; return ticket;
      });
      buildAudit(sourceRows);
      $('#uploadPanel').close();
      $('#fileDetails').className = 'mt-5 flex items-center justify-between rounded-lg bg-emerald-50 px-4 py-3 text-sm text-emerald-800';
      $('#fileDetails').innerHTML = `<span><strong>${escapeHtml(file.name)}</strong> · ${sourceRows.length.toLocaleString()} tickets read</span><span class="font-medium">Ready</span>`;
      $('#uploadError').classList.add('hidden');
    } catch (error) { showError(error.message || 'This workbook could not be read.'); }
  };
  reader.readAsArrayBuffer(file);
}

function buildAudit(tickets) {
  state.reportTime = new Date(); state.sheets = {}; state.technicianGroup = 'all'; state.edits.clear(); state.done.clear(); state.noUpdate.clear(); state.updateNeeded.clear();
  $('#technicianGroupFilter').value = state.technicianGroup;
  const withMeta = tickets.map((ticket) => ({ ...ticket, OnSheet: classification(ticket), 'Updated By': '', 'Update Type': '' }));
  const makeRows = (items, columns, sortBy, newest) => tableSort(items, sortBy, newest).map((ticket, index) => ({ ...ticket, '#': index + 1, _type: normalise(ticket['Request Type']) }));
  state.sheets['All Tickets'] = { columns: CONFIG.columns.master, rows: makeRows(withMeta, CONFIG.columns.master, 'Last Updated Time', false) };
  [...new Set(CONFIG.alertKeywords.map(([, name]) => name))].forEach((name) => {
    state.sheets[name] = { columns: CONFIG.columns.alert, rows: makeRows(withMeta.filter((ticket) => classification(ticket) === name), CONFIG.columns.alert, 'Created Time', true) };
  });
  state.sheets['Non Alerts'] = { columns: CONFIG.columns.master, rows: makeRows(withMeta.filter((ticket) => !isAlert(ticket)), CONFIG.columns.master, 'Last Updated Time', false) };
  state.sheets['Work in Progress'] = { columns: CONFIG.columns.status, rows: makeRows(withMeta.filter((ticket) => !isAlert(ticket) && normalise(ticket.Status) === 'work in progress'), CONFIG.columns.status, 'Due By Time', false) };
  state.sheets.Pending = { columns: CONFIG.columns.status, rows: makeRows(withMeta.filter((ticket) => !isAlert(ticket) && CONFIG.pendingStatuses.has(normalise(ticket.Status))), CONFIG.columns.status, 'Due By Time', false) };
  state.sheets['Waiting Approval'] = { columns: CONFIG.columns.status, rows: makeRows(withMeta.filter((ticket) => !isAlert(ticket) && normalise(ticket.Status) === 'waiting approval'), CONFIG.columns.status, 'Due By Time', false) };
  state.visibleColumns = Object.fromEntries(Object.entries(state.sheets).map(([name, table]) => [name, new Set(table.columns)]));
  state.activeSheet = 'All Tickets';
  state.sort = { column: null, direction: 'asc' };
  renderApp();
}

function renderApp() {
  $('#auditPanel').classList.remove('hidden'); $('#auditPanel').classList.add('flex', 'flex-col');
  $('#auditStatus').textContent = `Created from your selected workbook at ${formatDate(state.reportTime)}.`;
  renderTabs(); renderTable();
}

function renderTabs() {
  $('#sheetTabs').innerHTML = Object.entries(state.sheets).map(([name, table]) => `<button data-sheet="${escapeAttr(name)}" class="${name === state.activeSheet ? 'active-tab' : 'border-transparent text-slate-500 hover:border-slate-300 hover:text-slate-700'} shrink-0 border-b-2 px-3 py-3 text-sm font-medium">${escapeHtml(name)} <span class="ml-1 rounded-full bg-slate-100 px-1.5 py-0.5 text-xs text-slate-900">${table.rows.filter(matchesTechnicianGroup).length}</span></button>`).join('');
  $('#sheetTabs').querySelectorAll('button').forEach((button) => button.addEventListener('click', () => { state.activeSheet = button.dataset.sheet; state.sort = { column: null, direction: 'asc' }; $('#searchInput').value = ''; renderTabs(); renderTable(); }));
}

function renderTable() {
  const table = state.sheets[state.activeSheet], query = normalise($('#searchInput').value), visible = table.columns.filter((column) => state.visibleColumns[state.activeSheet].has(column));
  const rows = table.rows.filter((row) => matchesTechnicianGroup(row) && (!query || visible.some((column) => normalise(displayValue(row[column])).includes(query))));
  $('#ticketTotal').textContent = `${rows.length.toLocaleString()} shown`;
  $('#tableHead').innerHTML = `<tr><th class="w-12">Done</th>${visible.map((column) => sortHeader(column)).join('')}</tr>`;
  $('#tableBody').innerHTML = sortRows(rows).map((row) => renderRow(row, visible)).join('');
  $('#emptyState').classList.toggle('hidden', rows.length > 0); document.querySelector('.table-wrap table').classList.toggle('hidden', rows.length === 0);
  $('#clearDoneButton').classList.toggle('hidden', state.done.size === 0); renderWorkflowStatus();
  document.querySelectorAll('[data-done]').forEach((box) => box.addEventListener('change', () => { box.checked ? state.done.add(box.dataset.done) : state.done.delete(box.dataset.done); renderTable(); }));
  document.querySelectorAll('[data-workflow]').forEach((select) => select.addEventListener('change', () => stageEdit(select.dataset.ticket, select.dataset.workflow, select.value)));
  document.querySelectorAll('[data-no-update]').forEach((box) => box.addEventListener('change', () => {
    box.checked ? state.noUpdate.add(box.dataset.noUpdate) : state.noUpdate.delete(box.dataset.noUpdate);
    if (state.sort.column === 'No Update') renderTable();
  }));
  document.querySelectorAll('[data-update-needed]').forEach((select) => select.addEventListener('change', () => {
    state.updateNeeded.set(select.dataset.updateNeeded, select.value);
    if (state.sort.column === 'Update needed') renderTable();
  }));
  document.querySelectorAll('[data-sort]').forEach((button) => button.addEventListener('click', () => {
    const column = button.dataset.sort;
    state.sort = { column, direction: state.sort.column === column && state.sort.direction === 'asc' ? 'desc' : 'asc' };
    renderTable();
  }));
  document.querySelectorAll('[data-copy]').forEach((button) => button.addEventListener('click', async () => {
    const copied = await copyToClipboard(button.dataset.copy);
    if (!copied) return;
    button.innerHTML = '✓'; button.setAttribute('aria-label', 'Copied'); button.title = 'Copied';
    setTimeout(() => { button.innerHTML = copyIcon(); button.setAttribute('aria-label', 'Copy cell value'); button.title = 'Copy cell value'; }, 1200);
  }));
}

function sortHeader(column) {
  const active = state.sort.column === column;
  const glyph = active ? (state.sort.direction === 'asc' ? '↑' : '↓') : '↕';
  return `<th class="${frozenClass(column)}"><button class="sort-button" type="button" data-sort="${escapeAttr(column)}">${escapeHtml(column)} <span aria-hidden="true">${glyph}</span></button></th>`;
}

function sortRows(rows) {
  const { column, direction } = state.sort;
  if (!column) return rows;
  const multiplier = direction === 'asc' ? 1 : -1;
  return [...rows].sort((left, right) => {
    const leftValue = column === 'No Update' ? (state.noUpdate.has(text(left.RequestID)) ? 'Yes' : 'No') : column === 'Update needed' ? (state.updateNeeded.get(text(left.RequestID)) || '') : left[column];
    const rightValue = column === 'No Update' ? (state.noUpdate.has(text(right.RequestID)) ? 'Yes' : 'No') : column === 'Update needed' ? (state.updateNeeded.get(text(right.RequestID)) || '') : right[column];
    const isTimeColumn = column.includes('Time');
    const leftDate = isTimeColumn ? getDate(leftValue) : null, rightDate = isTimeColumn ? getDate(rightValue) : null;
    let result;
    if (leftDate && rightDate) result = leftDate - rightDate;
    else if (typeof leftValue === 'number' && typeof rightValue === 'number') result = leftValue - rightValue;
    else result = text(leftValue).localeCompare(text(rightValue), undefined, { numeric: true, sensitivity: 'base' });
    return result * multiplier;
  });
}

function renderRow(row, visible) {
  const key = `${state.activeSheet}:${row['#']}`;
  return `<tr><td><input data-done="${escapeAttr(key)}" type="checkbox" ${state.done.has(key) ? 'checked' : ''} class="h-4 w-4 rounded border-slate-300 text-ocean focus:ring-ocean" aria-label="Mark ticket ${escapeAttr(row.RequestID)} done"></td>${visible.map((column) => renderCell(row, column)).join('')}</tr>`;
}

function renderCell(row, column) {
  const frozen = Boolean(frozenClass(column));
  const copyButton = frozen ? '' : `<button type="button" class="copy-button" data-copy="${escapeAttr(displayValue(row[column]))}" aria-label="Copy cell value" title="Copy cell value">${copyIcon()}</button>`;
  return `<td class="${cellClass(row, column)} ${frozen ? '' : 'copyable-cell'}">${cellContent(row, column)}${copyButton}</td>`;
}

function frozenClass(column) {
  return column === 'Update needed' ? 'frozen-update-needed' : column === 'No Update' ? 'frozen-no-update' : column === 'Updated By' ? 'frozen-updated-by' : column === 'Update Type' ? 'frozen-update-type' : '';
}

function cellClass(row, column) {
  const classes = [frozenClass(column)];
  const lastUpdated = getDate(row['Last Updated Time']);
  if ((column === 'Due By Time' && isOverdue(row)) || (column === 'UnReplied Count' && hasUnreplied(row)) || (column === 'Status' && isNP(row)) || (column === 'Last Updated Time' && lastUpdated && lastUpdated.toDateString() !== state.reportTime.toDateString()) || (column === 'Last Updated By Email' && text(row[column]) && !CONFIG.allowedTechnicians.has(normalise(row[column])))) classes.push('warning-cell');
  if (column === 'Subject') classes.push('subject-cell');
  return classes.join(' ');
}

function cellContent(row, column) {
  const ticketKey = text(row.RequestID);
  if (column === 'RequestID') return row.url ? `<a class="font-medium text-blue-700 underline decoration-blue-300 underline-offset-2 hover:text-blue-900" href="${escapeAttr(row.url)}" target="_blank" rel="noopener">${escapeHtml(row[column])}</a>` : escapeHtml(row[column]);
  if (column === 'Update needed') return selectControl(ticketKey, column, state.updateNeeded.get(ticketKey) || '', ['', 'Next step', 'Replied - Next step', 'Approved - Next step', 'Replied - Next step', 'New Alert', 'Follow up - Approval', 'Follow up - Validation', 'Follow up - Information','New alert', 'Strike 1', 'Strike 2', 'Strike 3', 'Monitoring 1', 'Monitoring 2', 'Monitoring 3', 'Monitoring 4', 'Monitoring 5', 'Resolved', 'Canceled'], 'update-needed');
  if (column === 'No Update') return `<input data-no-update="${escapeAttr(ticketKey)}" type="checkbox" ${state.noUpdate.has(ticketKey) ? 'checked' : ''} class="h-4 w-4 rounded border-slate-300 text-ocean focus:ring-ocean" aria-label="No Update for ${escapeAttr(ticketKey)}">`;
  if (column === 'Updated By') return selectControl(ticketKey, column, getValue(row, column), ['', 'Technician', 'Requestor', 'System', 'User']);
  if (column === 'Update Type') return selectControl(ticketKey, column, getValue(row, column), ['', 'New alert', 'New reply', 'New approval']);
  return escapeHtml(displayValue(row[column]));
}

function selectControl(ticket, field, current, options, mode = 'workflow') { const data = mode === 'update-needed' ? `data-update-needed="${escapeAttr(ticket)}"` : `data-ticket="${escapeAttr(ticket)}" data-workflow="${escapeAttr(field)}"`; return `<select ${data} class="rounded-md border border-slate-300 bg-white px-2 py-1 text-xs text-slate-700 focus:border-ocean focus:outline-none focus:ring-1 focus:ring-ocean">${options.map((option) => `<option value="${escapeAttr(option)}" ${option === current ? 'selected' : ''}>${option || '—'}</option>`).join('')}</select>`; }
function getValue(row, field) { return state.edits.get(text(row.RequestID))?.[field] ?? text(row[field]); }
function stageEdit(ticket, field, value) { const changed = state.edits.get(ticket) || { 'Updated By': '', 'Update Type': '' }; changed[field] = value; state.edits.set(ticket, changed); renderWorkflowStatus(); }
function renderWorkflowStatus() { /* Workflow values are saved silently into the downloaded workbook. */ }
function displayValue(value) { return value instanceof Date ? formatDate(value) : text(value); }
function copyIcon() { return '<svg aria-hidden="true" viewBox="0 0 24 24" class="h-3.5 w-3.5 fill-none stroke-current stroke-2"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M15 9V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v8a2 2 0 0 0 2 2h4"/></svg>'; }
async function copyToClipboard(value) {
  try { await navigator.clipboard.writeText(value); return true; }
  catch {
    const field = document.createElement('textarea'); field.value = value; field.style.position = 'fixed'; field.style.opacity = '0';
    document.body.append(field); field.select(); const copied = document.execCommand('copy'); field.remove(); return copied;
  }
}

function toggleColumnMenu() {
  const menu = $('#columnMenu');
  if (!menu.classList.contains('hidden')) { menu.classList.add('hidden'); return; }
  const table = state.sheets[state.activeSheet], selected = state.visibleColumns[state.activeSheet];
  $('#columnOptions').innerHTML = table.columns.map((column) => `<label class="flex cursor-pointer items-center gap-3 rounded-lg px-2 py-2 hover:bg-slate-50"><input type="checkbox" value="${escapeAttr(column)}" ${selected.has(column) ? 'checked' : ''} class="h-4 w-4 rounded border-slate-300 text-ocean focus:ring-ocean"><span class="text-sm text-slate-700">${escapeHtml(column)}</span></label>`).join('');
  $('#columnOptions').querySelectorAll('input').forEach((input) => input.addEventListener('change', () => {
    const next = new Set([...$('#columnOptions').querySelectorAll('input:checked')].map((item) => item.value));
    if (!next.size) { input.checked = true; return; }
    state.visibleColumns[state.activeSheet] = next; renderTable();
  }));
  menu.classList.remove('hidden');
}

function exportWorkbook() {
  const workbook = XLSX.utils.book_new();
  Object.entries(state.sheets).forEach(([name, table]) => {
    const rows = table.rows.map((row) => Object.fromEntries(table.columns.map((column) => [column, column === 'Update needed' ? (state.updateNeeded.get(text(row.RequestID)) || '') : column === 'No Update' ? (state.noUpdate.has(text(row.RequestID)) ? 'Yes' : 'No') : column === 'Updated By' || column === 'Update Type' ? getValue(row, column) : row[column]])));
    const sheet = XLSX.utils.json_to_sheet(rows, { header: table.columns });
    sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: table.columns.length - 1 } }) }; sheet['!freeze'] = { xSplit: 0, ySplit: 1 };
    sheet['!cols'] = table.columns.map((column) => ({ wch: Math.min(Math.max(column.length + 2, ...rows.map((row) => displayValue(row[column]).length + 2)), 45) }));
    applyWorkbookFormatting(sheet, table, rows);
    XLSX.utils.book_append_sheet(workbook, sheet, name);
  });
  XLSX.writeFile(workbook, 'Ticket_Audit.xlsx');
}

function fillStyle(color) {
  return { fill: { patternType: 'solid', fgColor: { rgb: color } } };
}

function applyWorkbookFormatting(sheet, table, outputRows) {
  const headerStyle = {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '1F4E78' } },
  };
  const incidentStyle = fillStyle('FFF2CC');
  const serviceRequestStyle = fillStyle('DDEBF7');
  const warningStyle = fillStyle('F4CCCC');
  const linkStyle = { font: { color: { rgb: '0563C1' }, underline: true } };
  const dateColumns = new Set(['Created Time', 'Due By Time', 'Last Updated Time']);

  table.columns.forEach((column, columnIndex) => {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c: columnIndex })];
    if (cell) cell.s = headerStyle;
  });

  table.rows.forEach((ticket, rowIndex) => {
    const excelRow = rowIndex + 1;
    const isIncident = normalise(ticket['Request Type']) === 'incident';
    const isServiceRequest = normalise(ticket['Request Type']) === 'service request';
    const rowStyle = isIncident ? incidentStyle : isServiceRequest ? serviceRequestStyle : null;

    table.columns.forEach((column, columnIndex) => {
      const address = XLSX.utils.encode_cell({ r: excelRow, c: columnIndex });
      const cell = sheet[address];
      if (!cell) return;
      if (rowStyle) cell.s = rowStyle;
      if (dateColumns.has(column)) cell.z = 'yyyy-mm-dd hh:mm';
    });

    const shadeWarning = (column) => {
      const columnIndex = table.columns.indexOf(column);
      if (columnIndex < 0) return;
      const cell = sheet[XLSX.utils.encode_cell({ r: excelRow, c: columnIndex })];
      if (cell) cell.s = warningStyle;
    };
    if (isNP(ticket)) shadeWarning('Status');
    if (isOverdue(ticket)) shadeWarning('Due By Time');
    if (hasUnreplied(ticket)) shadeWarning('UnReplied Count');

    const lastUpdated = getDate(ticket['Last Updated Time']);
    if (lastUpdated && lastUpdated.toDateString() !== state.reportTime.toDateString()) shadeWarning('Last Updated Time');
    if (text(ticket['Last Updated By Email']) && !CONFIG.allowedTechnicians.has(normalise(ticket['Last Updated By Email']))) shadeWarning('Last Updated By Email');

    if (ticket.url) {
      const ticketNumberIndex = table.columns.indexOf('RequestID');
      [ticketNumberIndex].filter((index) => index >= 0).forEach((columnIndex) => {
        const cell = sheet[XLSX.utils.encode_cell({ r: excelRow, c: columnIndex })];
        if (cell) { cell.l = { Target: ticket.url }; cell.s = { ...(rowStyle || {}), ...linkStyle }; }
      });
    }
  });
}

function showError(message) { $('#uploadError').textContent = message; $('#uploadError').classList.remove('hidden'); }
function escapeHtml(value) { return text(value).replace(/[&<>'"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[c])); }
function escapeAttr(value) { return escapeHtml(value); }

$('#fileInput').addEventListener('change', (event) => event.target.files[0] && loadFile(event.target.files[0]));
['dragenter', 'dragover'].forEach((eventName) => $('#dropZone').addEventListener(eventName, (event) => { event.preventDefault(); $('#dropZone').classList.add('is-dragging'); }));
['dragleave', 'drop'].forEach((eventName) => $('#dropZone').addEventListener(eventName, (event) => { event.preventDefault(); $('#dropZone').classList.remove('is-dragging'); }));
$('#dropZone').addEventListener('drop', (event) => event.dataTransfer.files[0] && loadFile(event.dataTransfer.files[0]));
$('#searchInput').addEventListener('input', renderTable); $('#technicianGroupFilter').addEventListener('change', (event) => { state.technicianGroup = event.target.value; renderTabs(); renderTable(); }); $('#columnsButton').addEventListener('click', toggleColumnMenu); $('#downloadButton').addEventListener('click', exportWorkbook);
document.addEventListener('click', (event) => { if (!event.target.closest('#columnsButton') && !event.target.closest('#columnMenu')) $('#columnMenu').classList.add('hidden'); });
$('#clearDoneButton').addEventListener('click', () => { state.done.clear(); renderTable(); });
$('#newAuditButton').addEventListener('click', () => { $('#fileInput').value = ''; $('#fileDetails').classList.add('hidden'); $('#uploadError').classList.add('hidden'); $('#uploadPanel').showModal(); });
$('#closeUploadDialog').addEventListener('click', () => $('#uploadPanel').close());

function showEmptyAudit() {
  state.sheets = { 'All Tickets': { columns: CONFIG.columns.master, rows: [] } };
  state.visibleColumns = { 'All Tickets': new Set(CONFIG.columns.master) };
  state.activeSheet = 'All Tickets';
  state.reportTime = new Date();
  renderApp();
  $('#auditStatus').textContent = 'Choose a workbook with New audit to begin.';
}

showEmptyAudit();
