const state = { preview: null };

const idfInput = document.querySelector('#idf-file');
const templateInput = document.querySelector('#template-file');
const templatePanel = document.querySelector('#template-panel');
const previewHead = document.querySelector('#preview-head');
const previewBody = document.querySelector('#preview-body');
const previewBtn = document.querySelector('#preview-btn');
const exportBtn = document.querySelector('#export-btn');
const message = document.querySelector('#message');
const stats = document.querySelector('#stats');

function selectedMode() {
  return document.querySelector('input[name="mode"]:checked').value;
}

function showMessage(text, type = '') {
  message.textContent = text;
  message.className = `message ${type}`.trim();
}

function formatFile(input, target) {
  const files = [...(input.files || [])];
  target.textContent = files.length ? `${files.length} 个文件：${files.map((file) => file.name).join('、')}` : '尚未选择';
}

function renderRows(rows, headers = []) {
  previewHead.innerHTML = '';
  const headRow = document.createElement('tr');
  for (const header of headers) {
    const th = document.createElement('th');
    th.textContent = header;
    headRow.appendChild(th);
  }
  if (!headers.length) {
    const th = document.createElement('th');
    th.textContent = '管件明细';
    headRow.appendChild(th);
  }
  previewHead.appendChild(headRow);
  previewBody.innerHTML = '';
  if (!rows.length) {
    previewBody.innerHTML = `<tr><td colspan="${Math.max(headers.length, 1)}" class="empty-state">没有找到组件记录</td></tr>`;
    return;
  }
  for (const row of rows) {
    const tr = document.createElement('tr');
    for (const header of headers) {
      const td = document.createElement('td');
      td.textContent = row[header] ?? '';
      tr.appendChild(td);
    }
    previewBody.appendChild(tr);
  }
}

async function identify() {
  const files = [...(idfInput.files || [])];
  if (!files.length) {
    showMessage('请先选择 IDF 文件。', 'error');
    return;
  }
  const form = new FormData();
  for (const file of files) form.append('idf', file);
  previewBtn.disabled = true;
  exportBtn.disabled = true;
  showMessage('正在读取 IDF 材料定义…');
  try {
    const response = await fetch('/api/preview', { method: 'POST', body: form });
    const data = await response.json();
    if (!response.ok || !data.ok) throw new Error(data.error || '识别失败');
    state.preview = data;
    renderRows(data.rows, data.headers || Object.keys(data.rows[0] || {}));
    stats.textContent = `${data.file_count || files.length} 个文件 · ${data.material_count} 条材料汇总 · ${data.detail_count} 条管件明细`;
    exportBtn.disabled = false;
    showMessage('识别完成，可以导出 Excel。', 'success');
  } catch (error) {
    state.preview = null;
    stats.textContent = '识别失败';
    renderRows([], state.preview?.headers || []);
    showMessage(error.message, 'error');
  } finally {
    previewBtn.disabled = false;
  }
}

async function exportWorkbook() {
  const idf = idfInput.files?.[0];
  if (!idf || !state.preview) {
    showMessage('请先识别材料。', 'error');
    return;
  }
  const mode = selectedMode();
  const form = new FormData();
  for (const file of idfInput.files) form.append('idf', file);
  form.append('mode', mode);
  if (mode === 'custom') {
    const template = templateInput.files?.[0];
    if (!template) {
      showMessage('自定义模板模式需要选择 Excel 模板。', 'error');
      return;
    }
    form.append('template', template);
    form.append('sheet_name', document.querySelector('#sheet-name').value.trim());
    form.append('header_row', document.querySelector('#header-row').value.trim());
  }
  exportBtn.disabled = true;
  showMessage('正在生成 Excel…');
  try {
    const response = await fetch('/api/export', { method: 'POST', body: form });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      throw new Error(data.error || '导出失败');
    }
    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const match = disposition.match(/filename="([^"]+)"/i);
    const filename = match?.[1] || 'IDF_材料清单.xlsx';
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    showMessage(`已下载：${filename}`, 'success');
  } catch (error) {
    showMessage(error.message, 'error');
  } finally {
    exportBtn.disabled = false;
  }
}

idfInput.addEventListener('change', () => {
  formatFile(idfInput, document.querySelector('#idf-name'));
  state.preview = null;
  exportBtn.disabled = true;
  stats.textContent = '等待识别';
  showMessage('文件已选择，请点击“识别材料”。');
});

templateInput.addEventListener('change', () => formatFile(templateInput, document.querySelector('#template-name')));
previewBtn.addEventListener('click', identify);
exportBtn.addEventListener('click', exportWorkbook);
document.querySelectorAll('input[name="mode"]').forEach((input) => {
  input.addEventListener('change', () => {
    document.querySelectorAll('.mode-option').forEach((option) => option.classList.toggle('active', option.querySelector('input').checked));
    templatePanel.classList.toggle('hidden', selectedMode() !== 'custom');
  });
});
