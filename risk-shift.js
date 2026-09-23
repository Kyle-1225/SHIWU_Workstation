/**
 * 风险推移功能模块
 * 处理Excel数据：计算总装缺口、涂装缺口、风险判断
 * 生成图片预览、文字输出、Excel下载
 */

const COLUMN_NAMES = {
  factory:          ['工厂'],
  partNo:           ['零件号'],
  jianHao:          ['简号'],
  supplier:         ['供应商名称'],
  theoryStock:      ['理论在库'],
  assemblyWip:      ['总装在制数'],
  paintWip:         ['涂装在制数'],
  secondLockDemand: ['二次锁定需求'],
  logisticsMode:    ['物流模式'],
};

const state = {
  headers: null,
  rawData: null,
  newHeaders: null,
  newRows: null,
  origIdx: null,
  newIdx: null,
  filename: '',
  timeStr: '',
  factory: '',
  stats: { total: 0, risk: 0, assembly: 0, paint: 0 },
  currentPage: 1,
  pageSize: 50,
  filterRisk: false,
  imageScale: 1,
  imageDataUrl: null,
};

const el = {};

function initElements() {
  el.uploadZone = document.getElementById('uploadZone');
  el.fileInput = document.getElementById('fileInput');
  el.uploadText = document.getElementById('uploadText');
  el.resultsSection = document.getElementById('resultsSection');
  el.loadingOverlay = document.getElementById('loadingOverlay');
  el.loadingText = document.getElementById('loadingText');
  el.tableHead = document.getElementById('tableHead');
  el.tableBody = document.getElementById('tableBody');
  el.pagination = document.getElementById('pagination');
  el.filterRisk = document.getElementById('filterRisk');
  el.statTotal = document.getElementById('statTotal');
  el.statRisk = document.getElementById('statRisk');
  el.statAssembly = document.getElementById('statAssembly');
  el.statPaint = document.getElementById('statPaint');
  el.statDisplay = document.getElementById('statDisplay');
  el.downloadBtn = document.getElementById('downloadBtn');
  el.imagePreviewArea = document.getElementById('imagePreviewArea');
  el.textOutput = document.getElementById('textOutput');
  el.copyTextBtn = document.getElementById('copyTextBtn');
  el.regenTextBtn = document.getElementById('regenTextBtn');
  el.downloadImgBtn = document.getElementById('downloadImgBtn');
  el.zoomIn = document.getElementById('zoomIn');
  el.zoomOut = document.getElementById('zoomOut');
  el.zoomLevel = document.getElementById('zoomLevel');
  el.captureTable = document.getElementById('captureTable');
}

function init() {
  initElements();

  el.uploadZone.addEventListener('click', () => el.fileInput.click());
  el.uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.uploadZone.classList.add('dragover');
  });
  el.uploadZone.addEventListener('dragleave', () => {
    el.uploadZone.classList.remove('dragover');
  });
  el.uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    el.uploadZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleFileSelect(file);
  });
  el.fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFileSelect(file);
  });

  el.filterRisk.addEventListener('change', () => {
    state.filterRisk = el.filterRisk.checked;
    state.currentPage = 1;
    renderTable();
  });

  el.downloadBtn.addEventListener('click', downloadExcel);
  el.copyTextBtn.addEventListener('click', copyText);
  el.regenTextBtn.addEventListener('click', () => {
    el.textOutput.value = generateText();
  });
  el.downloadImgBtn.addEventListener('click', downloadImage);
  el.zoomIn.addEventListener('click', () => {
    state.imageScale = Math.min(state.imageScale + 0.25, 3);
    updateZoom();
  });
  el.zoomOut.addEventListener('click', () => {
    state.imageScale = Math.max(state.imageScale - 0.25, 0.25);
    updateZoom();
  });
}

function showLoading(text) {
  el.loadingText.textContent = text || '正在处理...';
  el.loadingOverlay.style.display = 'flex';
}

function hideLoading() {
  el.loadingOverlay.style.display = 'none';
}

function handleFileSelect(file) {
  const validExts = ['.xlsx', '.xls'];
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!validExts.includes(ext)) {
    alert('请上传 .xlsx 或 .xls 格式的文件');
    return;
  }

  showLoading('正在读取文件...');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      showLoading('正在解析数据...');
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });

      if (data.length < 2) {
        throw new Error('文件数据不足，请检查文件内容');
      }

      state.headers = data[0];
      state.rawData = data.slice(1);
      state.filename = file.name;

      showLoading('正在查找列索引...');
      state.origIdx = findColumnIndices(state.headers);

      const required = ['theoryStock', 'assemblyWip', 'paintWip', 'secondLockDemand'];
      for (const key of required) {
        if (state.origIdx[key] === -1) {
          throw new Error(`缺少必要列: ${COLUMN_NAMES[key].join(' 或 ')}`);
        }
      }

      showLoading('正在计算风险数据...');
      const { newHeaders, newRows } = buildNewData();
      state.newHeaders = newHeaders;
      state.newRows = newRows;
      state.newIdx = calculateNewIndices();

      state.timeStr = extractTimeFromFilename(file.name);
      const factoryRow = state.rawData[0];
      state.factory = factoryRow ? String(factoryRow[state.origIdx.factory] || '未知工厂') : '未知工厂';

      calculateStats();

      showLoading('正在生成表格...');
      renderTable();

      showLoading('正在生成图片...');
      generateImage();

      el.textOutput.value = generateText();

      el.resultsSection.style.display = 'block';
      el.uploadText.innerHTML = `<span class="file-info">✓ ${file.name}</span>`;
      el.uploadText.nextElementSibling.textContent = '点击重新上传文件';

      hideLoading();
      el.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
      hideLoading();
      alert('处理文件时出错: ' + err.message);
      console.error(err);
    }
  };
  reader.onerror = () => {
    hideLoading();
    alert('文件读取失败');
  };
  reader.readAsArrayBuffer(file);
}

function findColumnIndices(headers) {
  const idx = {};
  for (const [key, names] of Object.entries(COLUMN_NAMES)) {
    idx[key] = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (names.includes(h)) {
        idx[key] = i;
        break;
      }
    }
  }
  return idx;
}

function buildNewData() {
  const oi = state.origIdx;
  const headers = state.headers;
  const rows = state.rawData;

  const insertions = [
    { after: oi.secondLockDemand, name: '风险判断' },
    { after: oi.paintWip, name: '涂装缺口' },
    { after: oi.assemblyWip, name: '总装缺口' },
  ].filter(ins => ins.after >= 0).sort((a, b) => b.after - a.after);

  let newHeaders = [...headers];
  for (const ins of insertions) {
    newHeaders = [
      ...newHeaders.slice(0, ins.after + 1),
      ins.name,
      ...newHeaders.slice(ins.after + 1)
    ];
  }

  const newRows = rows.map(row => {
    let newRow = [...row];
    for (const ins of insertions) {
      let value;
      const theory = Number(row[oi.theoryStock]) || 0;
      const assemblyWip = Number(row[oi.assemblyWip]) || 0;
      const paintWip = Number(row[oi.paintWip]) || 0;
      const assemblyGap = theory - assemblyWip;
      const paintGap = assemblyGap - paintWip;

      if (ins.name === '总装缺口') {
        value = Math.round(assemblyGap * 100) / 100;
      } else if (ins.name === '涂装缺口') {
        value = Math.round(paintGap * 100) / 100;
      } else if (ins.name === '风险判断') {
        if (assemblyGap < 0) value = '不满足总装';
        else if (paintGap < 0) value = '不满足涂装';
        else value = '';
      }

      newRow = [
        ...newRow.slice(0, ins.after + 1),
        value,
        ...newRow.slice(ins.after + 1)
      ];
    }
    return newRow;
  });

  return { newHeaders, newRows };
}

function calculateNewIndices() {
  const oi = state.origIdx;
  let shift1 = 0;
  let shift2 = 0;
  let shift3 = 0;

  if (oi.assemblyWip >= 0) shift1 = 1;
  if (oi.paintWip >= 0 && oi.paintWip > oi.assemblyWip) shift2 = 1;
  if (oi.secondLockDemand >= 0 && oi.secondLockDemand > oi.paintWip) shift3 = 1;

  return {
    partNo:           oi.partNo,
    jianHao:          oi.jianHao,
    supplier:         oi.supplier,
    theoryStock:      oi.theoryStock,
    assemblyWip:      oi.assemblyWip,
    assemblyGap:      oi.assemblyWip + 1,
    paintWip:         oi.paintWip + shift1,
    paintGap:         oi.paintWip + shift1 + 1,
    secondLockDemand: oi.secondLockDemand + shift1 + shift2,
    riskJudge:        oi.secondLockDemand + shift1 + shift2 + 1,
    logisticsMode:    oi.logisticsMode >= 0 ? oi.logisticsMode + shift1 + shift2 + shift3 : -1,
  };
}

function getDisplayColumns() {
  const ni = state.newIdx;
  const cols = [];
  const startIdx = ni.partNo >= 0 ? ni.partNo : 0;
  const endIdx = ni.riskJudge >= 0 ? ni.riskJudge : (state.newHeaders.length - 1);
  for (let i = startIdx; i <= endIdx; i++) {
    cols.push({
      index: i,
      name: state.newHeaders[i],
      isNew: i === ni.assemblyGap || i === ni.paintGap || i === ni.riskJudge,
      type: i === ni.assemblyGap ? 'assemblyGap' :
            i === ni.paintGap ? 'paintGap' :
            i === ni.riskJudge ? 'riskJudge' : 'normal'
    });
  }
  if (ni.logisticsMode >= 0 && ni.logisticsMode < state.newHeaders.length) {
    cols.push({
      index: ni.logisticsMode,
      name: state.newHeaders[ni.logisticsMode],
      isNew: false,
      type: 'logistics'
    });
  }
  return cols;
}

function calculateStats() {
  const ni = state.newIdx;
  state.stats.total = state.newRows.length;
  state.stats.assembly = state.newRows.filter(r => r[ni.riskJudge] === '不满足总装').length;
  state.stats.paint = state.newRows.filter(r => r[ni.riskJudge] === '不满足涂装').length;
  state.stats.risk = state.stats.assembly + state.stats.paint;

  el.statTotal.textContent = state.stats.total.toLocaleString();
  el.statRisk.textContent = state.stats.risk.toLocaleString();
  el.statAssembly.textContent = state.stats.assembly.toLocaleString();
  el.statPaint.textContent = state.stats.paint.toLocaleString();
}

function getFilteredRows() {
  const ni = state.newIdx;
  if (state.filterRisk) {
    return state.newRows.filter(r => r[ni.riskJudge] === '不满足总装' || r[ni.riskJudge] === '不满足涂装');
  }
  return state.newRows;
}

function renderTable() {
  const cols = getDisplayColumns();

  let thead = '<tr>';
  cols.forEach(col => {
    let cls = '';
    if (col.isNew) cls += ' col-new';
    if (col.type === 'assemblyGap') cls += ' col-highlight-assembly';
    else if (col.type === 'paintGap') cls += ' col-highlight-paint';
    thead += `<th class="${cls.trim()}">${col.name}</th>`;
  });
  thead += '</tr>';
  el.tableHead.innerHTML = thead;

  const filtered = getFilteredRows();
  const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
  if (state.currentPage > totalPages) state.currentPage = totalPages;

  const start = (state.currentPage - 1) * state.pageSize;
  const end = Math.min(start + state.pageSize, filtered.length);
  const pageRows = filtered.slice(start, end);

  const ni = state.newIdx;
  let tbody = '';
  for (const row of pageRows) {
    tbody += '<tr>';
    for (const col of cols) {
      const value = row[col.index];
      let cls = '';
      if (col.isNew) cls += ' col-new';
      if (col.type === 'assemblyGap' && Number(value) < 0) cls += ' highlight-assembly';
      else if (col.type === 'paintGap' && Number(value) < 0) cls += ' highlight-paint';
      else if (col.type === 'riskJudge') {
        if (value === '不满足总装') cls += ' risk-assembly';
        else if (value === '不满足涂装') cls += ' risk-paint';
      }
      tbody += `<td class="${cls.trim()}">${formatCellValue(value)}</td>`;
    }
    tbody += '</tr>';
  }
  el.tableBody.innerHTML = tbody;

  el.statDisplay.textContent = filtered.length.toLocaleString();

  let pagi = '';
  pagi += `<button ${state.currentPage <= 1 ? 'disabled' : ''} onclick="window._rsGoPage(${state.currentPage - 1})">上一页</button>`;
  pagi += `<span class="page-info">第 ${state.currentPage} / ${totalPages} 页</span>`;
  pagi += `<button ${state.currentPage >= totalPages ? 'disabled' : ''} onclick="window._rsGoPage(${state.currentPage + 1})">下一页</button>`;
  el.pagination.innerHTML = pagi;
}

window._rsGoPage = function(page) {
  state.currentPage = page;
  renderTable();
};

function formatCellValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number') {
    return Number.isInteger(value) ? String(value) : value.toFixed(2);
  }
  return String(value);
}

function generateText() {
  const ni = state.newIdx;
  const factoryShort = state.factory.replace('工厂', '厂');

  const assemblyRiskRows = state.newRows.filter(r => r[ni.riskJudge] === '不满足总装');
  const paintRiskRows = state.newRows.filter(r => r[ni.riskJudge] === '不满足涂装');
  const riskCount = assemblyRiskRows.length + paintRiskRows.length;

  let text = `${factoryShort}风险推移：\n`;
  text += `时间：${state.timeStr}\n`;
  text += `风险推移风险点数：${riskCount}\n`;
  text += `风险零件跟进结果如下：\n`;

  if (assemblyRiskRows.length > 0) {
    text += `拉动：不满足总装\n`;
    const modeGroups = {};
    assemblyRiskRows.forEach(r => {
      const mode = (ni.logisticsMode >= 0 ? r[ni.logisticsMode] : null) || '未分类';
      if (!modeGroups[mode]) modeGroups[mode] = {};
      const supplier = (ni.supplier >= 0 ? r[ni.supplier] : null) || '未知供应商';
      if (!modeGroups[mode][supplier]) modeGroups[mode][supplier] = new Set();
      const jh = (ni.jianHao >= 0 ? r[ni.jianHao] : null) || '';
      if (jh) modeGroups[mode][supplier].add(jh);
    });

    const modes = Object.keys(modeGroups);
    modes.forEach((mode, mi) => {
      if (mi > 0) text += '\n';
      text += `${mode}\n`;
      Object.keys(modeGroups[mode]).forEach(supplier => {
        text += `${supplier}\n`;
        Array.from(modeGroups[mode][supplier]).forEach(jh => {
          text += `        ${jh}\n`;
        });
      });
    });
  }

  if (paintRiskRows.length > 0) {
    if (assemblyRiskRows.length > 0) text += '\n';
    text += `循环：不满足涂装\n`;
    const supplierGroups = {};
    paintRiskRows.forEach(r => {
      const supplier = (ni.supplier >= 0 ? r[ni.supplier] : null) || '未知供应商';
      if (!supplierGroups[supplier]) supplierGroups[supplier] = new Set();
      const jh = (ni.jianHao >= 0 ? r[ni.jianHao] : null) || '';
      if (jh) supplierGroups[supplier].add(jh);
    });

    const suppliers = Object.keys(supplierGroups);
    suppliers.forEach((supplier, si) => {
      if (si > 0) text += '\n';
      text += `${supplier}\n`;
      Array.from(supplierGroups[supplier]).forEach(jh => {
        text += `        ${jh}\n`;
      });
    });
  }

  text += '\n';
  text += `以上零件有缺件风险，麻烦跟进到货，谢谢`;

  return text;
}

async function copyText() {
  const textarea = el.textOutput;
  try {
    await navigator.clipboard.writeText(textarea.value);
    const original = el.copyTextBtn.textContent;
    el.copyTextBtn.textContent = '已复制 ✓';
    setTimeout(() => { el.copyTextBtn.textContent = original; }, 2000);
  } catch (err) {
    textarea.select();
    document.execCommand('copy');
    const original = el.copyTextBtn.textContent;
    el.copyTextBtn.textContent = '已复制 ✓';
    setTimeout(() => { el.copyTextBtn.textContent = original; }, 2000);
  }
}

function downloadExcel() {
  const aoa = [state.newHeaders, ...state.newRows];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '风险推移');
  const outName = `风险推移_${state.timeStr.replace(':', '')}.xlsx`;
  XLSX.writeFile(wb, outName);
}

function extractTimeFromFilename(filename) {
  const match = filename.match(/(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/);
  if (match) {
    return `${match[4]}:${match[5]}`;
  }
  const now = new Date();
  const h = String(now.getHours()).padStart(2, '0');
  const m = String(now.getMinutes()).padStart(2, '0');
  return `${h}:${m}`;
}

function generateImage() {
  const ni = state.newIdx;
  const riskRows = state.newRows.filter(r =>
    r[ni.riskJudge] === '不满足总装' || r[ni.riskJudge] === '不满足涂装'
  );

  if (riskRows.length === 0) {
    el.imagePreviewArea.innerHTML = `
      <div class="image-preview-placeholder">
        <svg viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 15.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93z"/>
        </svg>
        <p>暂无风险数据</p>
      </div>`;
    return;
  }

  const cols = [];
  const imgStart = ni.partNo >= 0 ? ni.partNo : 0;
  const imgEnd = ni.riskJudge >= 0 ? ni.riskJudge : (state.newHeaders.length - 1);
  for (let i = imgStart; i <= imgEnd; i++) {
    cols.push({ index: i, name: state.newHeaders[i] });
  }

  const fontSize = 12;
  const headerHeight = 30;
  const cellHeight = 24;
  const padding = 8;
  const scale = 2;

  const measureCanvas = document.createElement('canvas');
  const mctx = measureCanvas.getContext('2d');
  mctx.font = `${fontSize}px "Microsoft YaHei", "PingFang SC", sans-serif`;

  const colWidths = cols.map(col => {
    let maxWidth = mctx.measureText(col.name).width;
    const sampleSize = Math.min(riskRows.length, 200);
    for (let i = 0; i < sampleSize; i++) {
      const text = formatCellValue(riskRows[i][col.index]);
      const width = mctx.measureText(text).width;
      if (width > maxWidth) maxWidth = width;
    }
    return Math.min(Math.ceil(maxWidth + padding * 2), 220);
  });

  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  const totalHeight = headerHeight + riskRows.length * cellHeight;

  const canvas = document.createElement('canvas');
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;
  canvas.style.width = totalWidth + 'px';
  canvas.style.height = totalHeight + 'px';

  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);
  ctx.font = `${fontSize}px "Microsoft YaHei", "PingFang SC", sans-serif`;
  ctx.textBaseline = 'middle';

  ctx.fillStyle = '#2563EB';
  ctx.fillRect(0, 0, totalWidth, headerHeight);

  let x = 0;
  cols.forEach((col, i) => {
    let headerBg = '#2563EB';
    if (col.index === ni.assemblyGap) headerBg = '#FDD835';
    else if (col.index === ni.paintGap) headerBg = '#EF9A9A';

    if (headerBg !== '#2563EB') {
      ctx.fillStyle = headerBg;
      ctx.fillRect(x, 0, colWidths[i], headerHeight);
    }

    ctx.fillStyle = col.index === ni.assemblyGap || col.index === ni.paintGap ? '#1E293B' : 'white';
    ctx.textAlign = 'center';
    ctx.fillText(col.name, x + colWidths[i] / 2, headerHeight / 2);

    ctx.strokeStyle = col.index === ni.assemblyGap || col.index === ni.paintGap ? '#B0BEC5' : '#1D4ED8';
    ctx.lineWidth = 1;
    ctx.strokeRect(x, 0, colWidths[i], headerHeight);

    x += colWidths[i];
  });

  riskRows.forEach((row, rowIdx) => {
    const y = headerHeight + rowIdx * cellHeight;
    let x = 0;

    cols.forEach((col, colIdx) => {
      const value = row[col.index];
      let bg = rowIdx % 2 === 0 ? '#FFFFFF' : '#F8FAFC';

      if (col.index === ni.assemblyGap && Number(value) < 0) {
        bg = '#FFF9C4';
      } else if (col.index === ni.paintGap && Number(value) < 0) {
        bg = '#FFCDD2';
      }

      ctx.fillStyle = bg;
      ctx.fillRect(x, y, colWidths[colIdx], cellHeight);

      ctx.strokeStyle = '#CBD5E1';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(x, y, colWidths[colIdx], cellHeight);

      let textColor = '#1E293B';
      if (col.index === ni.riskJudge) {
        if (value === '不满足总装') textColor = '#DC2626';
        else if (value === '不满足涂装') textColor = '#F59E0B';
      }

      ctx.fillStyle = textColor;
      ctx.textAlign = 'center';
      const displayText = formatCellValue(value);
      const maxWidth = colWidths[colIdx] - padding * 2;
      let text = displayText;
      if (mctx.measureText(text).width > maxWidth) {
        while (mctx.measureText(text + '...').width > maxWidth && text.length > 0) {
          text = text.slice(0, -1);
        }
        text = text + '...';
      }
      ctx.fillText(text, x + colWidths[colIdx] / 2, y + cellHeight / 2);

      x += colWidths[colIdx];
    });
  });

  const dataUrl = canvas.toDataURL('image/png');
  state.imageDataUrl = dataUrl;
  state.imageScale = 1;

  el.imagePreviewArea.innerHTML = `<img src="${dataUrl}" id="previewImage" style="transform: scale(${state.imageScale}); transform-origin: top center;">`;
  updateZoom();
}

function updateZoom() {
  const img = document.getElementById('previewImage');
  if (img) {
    img.style.transform = `scale(${state.imageScale})`;
  }
  el.zoomLevel.textContent = Math.round(state.imageScale * 100) + '%';
}

function downloadImage() {
  if (!state.imageDataUrl) {
    alert('暂无图片可下载');
    return;
  }
  const link = document.createElement('a');
  link.download = `风险推移图片_${state.timeStr.replace(':', '')}.png`;
  link.href = state.imageDataUrl;
  link.click();
}

document.addEventListener('DOMContentLoaded', init);
