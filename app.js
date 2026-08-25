const input = document.querySelector('#file-input');
const dropzone = document.querySelector('#dropzone');
const status = document.querySelector('#status');
const search = document.querySelector('#search');
const columnSelect = document.querySelector('#column-select');
const sortAsc = document.querySelector('#sort-asc');
const sortDesc = document.querySelector('#sort-desc');
const download = document.querySelector('#download');
const tableWrap = document.querySelector('#table-wrap');
const rowCount = document.querySelector('#row-count');
let headers = [], rows = [];

input.addEventListener('change', () => input.files[0] && loadFile(input.files[0]));
['dragenter','dragover'].forEach(event => dropzone.addEventListener(event, e => { e.preventDefault(); dropzone.classList.add('dragging'); }));
['dragleave','drop'].forEach(event => dropzone.addEventListener(event, e => { e.preventDefault(); dropzone.classList.remove('dragging'); }));
dropzone.addEventListener('drop', e => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));
search.addEventListener('input', render);
columnSelect.addEventListener('change', render);
sortAsc.addEventListener('click', () => sortRows(1)); sortDesc.addEventListener('click', () => sortRows(-1)); download.addEventListener('click', exportCsv);

async function loadFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();
  setStatus(`Обрабатываю «${file.name}»…`, 'loading');
  try {
    let matrix;
    if (['xlsx','xls','csv'].includes(extension)) matrix = await readExcel(file);
    else if (extension === 'docx') matrix = await readWord(file);
    else if (['jpg','jpeg','png'].includes(extension)) matrix = await readImage(file);
    else throw new Error('Этот формат пока не поддерживается.');
    setData(matrix);
    setStatus(`Готово: «${file.name}» обработан. При необходимости исправьте распознанные данные прямо в исходном файле и загрузите его снова.`);
  } catch (error) { console.error(error); setStatus(`Не удалось обработать файл: ${error.message}`, 'error'); }
}
function readExcel(file) { return new Promise((resolve, reject) => { const reader = new FileReader(); reader.onload = e => { try { const book = XLSX.read(e.target.result, {type:'array'}); resolve(XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]], {header:1, defval:''})); } catch (err) { reject(err); } }; reader.onerror = reject; reader.readAsArrayBuffer(file); }); }
async function readWord(file) { const result = await mammoth.extractRawText({arrayBuffer:await file.arrayBuffer()}); const lines = result.value.split(/\r?\n/).map(x=>x.trim()).filter(Boolean); return lines.map(line => line.split(/\t| {2,}|\s*;\s*/)); }
async function readImage(file) { const result = await Tesseract.recognize(file, 'rus+eng', { logger: m => { if (m.status === 'recognizing text') setStatus(`Распознаю текст: ${Math.round(m.progress * 100)}%`, 'loading'); } }); const lines = result.data.text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean); return lines.map(line => line.split(/\t| {2,}|\s*;\s*/)); }
function setData(matrix) { const clean = matrix.filter(row => row.some(cell => String(cell).trim() !== '')); if (!clean.length) throw new Error('В файле не найдено строк.'); headers = clean.shift().map((h,i) => String(h).trim() || `Столбец ${i+1}`); rows = clean.map(row => headers.map((_,i) => row[i] ?? '')); columnSelect.innerHTML = headers.map((h,i) => `<option value="${i}">${escapeHtml(h)}</option>`).join(''); [search,columnSelect,sortAsc,sortDesc,download].forEach(el=>el.disabled=false); render(); }
function render() { const term = search.value.trim().toLowerCase(); const visible = rows.filter(row => row.some(cell => String(cell).toLowerCase().includes(term))); rowCount.textContent = `Строк: ${visible.length} из ${rows.length}`; tableWrap.classList.remove('empty'); tableWrap.innerHTML = `<table><thead><tr>${headers.map(h=>`<th>${escapeHtml(h)}</th>`).join('')}</tr></thead><tbody>${visible.length ? visible.map(row=>`<tr>${row.map(cell=>`<td>${escapeHtml(String(cell))}</td>`).join('')}</tr>`).join('') : `<tr><td colspan="${headers.length}">Ничего не найдено.</td></tr>`}</tbody></table>`; }
function sortRows(direction) { const col = Number(columnSelect.value); rows.sort((a,b) => String(a[col]).localeCompare(String(b[col]), 'ru', {numeric:true, sensitivity:'base'}) * direction); render(); }
function exportCsv() { const csv = [headers,...rows].map(row => row.map(v => `"${String(v).replaceAll('"','""')}"`).join(';')).join('\n'); const url = URL.createObjectURL(new Blob([csv],{type:'text/csv;charset=utf-8;'})); const a=document.createElement('a'); a.href=url; a.download='таблица.csv'; a.click(); URL.revokeObjectURL(url); }
function setStatus(message, type='') { status.textContent=message; status.className=`status ${type}`; }
function escapeHtml(value) { const d=document.createElement('div'); d.textContent=value; return d.innerHTML; }
