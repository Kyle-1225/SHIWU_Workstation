/**
 * 风险推移功能模块 v2
 * 
 * 计算规则：
 *   总装缺口 = 理论在库(K) - 总装在制数(N) - 未审核不良数量(AJ)
 *   涂装缺口 = 总装缺口 - 涂装在制数(P)
 *   风险判断：
 *     物流模式=循环取货 且 涂装缺口<0 且 二次锁定需求≠0 → 不满足涂装
 *     物流模式=厂外拉动 且 总装缺口<0 且 二次锁定需求≠0 → 不满足总装
 * 
 * 车型排除：行的使用车型全部在排除集合中才剔除
 * 区域筛选：根据分工表匹配的所属区域筛选文字输出
 */

// ========== Constants ==========

const RISK_COL_NAMES = {
  factory:           ['工厂'],
  workshop:          ['车间'],
  partNo:            ['零件号'],
  jianHao:           ['简号'],
  partName:          ['零件名称'],
  supplierCode:      ['供应商代码'],
  supplierName:      ['供应商名称'],
  vehicleType:       ['使用车型'],
  line:              ['线体', '线体'],
  capacity:          ['收容数'],
  theoryStock:       ['理论在库'],
  endFlowNo:         ['截止流水号'],
  endVIN:            ['截止VIN'],
  assemblyWip:       ['总装在制数'],
  midPaintWip:       ['中涂在制数'],
  paintWip:          ['涂装在制数'],
  weldWip:           ['焊装在制数'],
  blueCardWip:       ['蓝卡在制数'],
  estQueueDemand:    ['推算队列需求'],
  firstLockDemand:   ['一次锁定需求'],
  secondLockDemand:  ['二次锁定需求'],
  unauditedBad:      ['未审核不良数量'],
  logisticsMode:     ['物流模式'],
};

const ASSIGN_COL_NAMES = {
  supplierCode:    ['供应商代码', '供应商编码', '供应商号', '代码'],
  region:          ['所属区域', '区域'],
  manager:         ['管理员'],
  assemblyClerk:   ['总装事务员', '事务员'],
};

const ASSIGN_STORAGE_KEY = 'riskShift_assignment_v2';
const VEHICLE_EXCLUDE_KEY = 'riskShift_vehicleExclude';

// Display columns: from 零件号 to 风险判断, plus 物流模式
const DISPLAY_SPEC = [
  { key: 'partNo',         label: '零件号' },
  { key: 'jianHao',         label: '简号' },
  { key: 'partName',        label: '零件名称' },
  { key: 'supplierCode',    label: '供应商代码' },
  { key: 'region',          label: '所属区域',     isNew: true, type: 'region' },
  { key: 'manager',        label: '管理员',       isNew: true, type: 'manager' },
  { key: 'assemblyClerk',  label: '总装事务员',   isNew: true, type: 'clerk' },
  { key: 'supplierName',   label: '供应商名称' },
  { key: 'vehicleType',    label: '使用车型' },
  { key: 'theoryStock',    label: '理论在库' },
  { key: 'assemblyWip',    label: '总装在制数' },
  { key: 'assemblyGap',    label: '总装缺口',     isNew: true, type: 'assemblyGap' },
  { key: 'midPaintWip',    label: '中涂在制数' },
  { key: 'paintWip',        label: '涂装在制数' },
  { key: 'paintGap',        label: '涂装缺口',     isNew: true, type: 'paintGap' },
  { key: 'firstLockDemand', label: '一次锁定需求' },
  { key: 'secondLockDemand',label: '二次锁定需求' },
  { key: 'riskJudge',       label: '风险判断',     isNew: true, type: 'riskJudge' },
  { key: 'logisticsMode',   label: '物流模式' },
];

// Image columns: same as display but without 物流模式
const IMAGE_SPEC = DISPLAY_SPEC.filter(c => c.key !== 'logisticsMode');

// Full column spec for Excel export: includes ALL original + new columns
const EXCEL_SPEC = [
  { key: 'factory',          label: '工厂',           hidden: true },
  { key: 'workshop',         label: '车间',           hidden: true },
  { key: 'partNo',           label: '零件号' },
  { key: 'jianHao',          label: '简号' },
  { key: 'partName',         label: '零件名称' },
  { key: 'supplierCode',     label: '供应商代码' },
  { key: 'region',           label: '所属区域',     isNew: true, type: 'region' },
  { key: 'manager',          label: '管理员',       isNew: true, type: 'manager' },
  { key: 'assemblyClerk',    label: '总装事务员',   isNew: true, type: 'clerk' },
  { key: 'supplierName',     label: '供应商名称' },
  { key: 'vehicleType',      label: '使用车型' },
  { key: 'line',             label: '线体',         hidden: true },
  { key: 'capacity',         label: '收容数',       hidden: true },
  { key: 'theoryStock',      label: '理论在库' },
  { key: 'endFlowNo',        label: '截止流水号',   hidden: true },
  { key: 'endVIN',           label: '截止VIN',      hidden: true },
  { key: 'assemblyWip',      label: '总装在制数' },
  { key: 'assemblyGap',      label: '总装缺口',     isNew: true, type: 'assemblyGap' },
  { key: 'midPaintWip',      label: '中涂在制数' },
  { key: 'paintWip',         label: '涂装在制数' },
  { key: 'paintGap',         label: '涂装缺口',     isNew: true, type: 'paintGap' },
  { key: 'weldWip',          label: '焊装在制数',   hidden: true },
  { key: 'blueCardWip',      label: '蓝卡在制数',   hidden: true },
  { key: 'estQueueDemand',   label: '推算队列需求', hidden: true },
  { key: 'firstLockDemand',  label: '一次锁定需求' },
  { key: 'secondLockDemand', label: '二次锁定需求' },
  { key: 'unauditedBad',     label: '未审核不良数量',hidden: true },
  { key: 'riskJudge',        label: '风险判断',     isNew: true, type: 'riskJudge' },
  { key: 'logisticsMode',    label: '物流模式' },
];

// ========== State ==========

const state = {
  origHeaders: null,
  origRows: null,
  origIdx: null,
  processedRows: null,
  filename: '',
  timeStr: '',
  factory: '',
  assignmentMap: null,
  assignmentInfo: null,
  vehicleTypes: [],
  selectedVehicles: new Set(getPersistentVehicleExclude()),
  regions: [],
  selectedRegions: new Set(),
  currentPage: 1,
  pageSize: 50,
  filterRisk: true,
  imageScale: 1,
  imageDataUrl: null,
  vehicleSelectRef: null,
  regionSelectRef: null,
};

function getPersistentVehicleExclude() {
  try {
    const raw = localStorage.getItem(VEHICLE_EXCLUDE_KEY);
    if (raw) return JSON.parse(raw);
  } catch (e) {}
  return ['AS9', 'AS9-G'];
}

const el = {};

// ========== Element Init ==========

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
  el.downloadImgBtn = document.getElementById('downloadImgBtn');
  el.generateImgBtn = document.getElementById('generateImgBtn');
  el.copyImgBtn = document.getElementById('copyImgBtn');
  el.zoomIn = document.getElementById('zoomIn');
  el.zoomOut = document.getElementById('zoomOut');
  el.zoomLevel = document.getElementById('zoomLevel');
  el.vehicleSelectContainer = document.getElementById('vehicleSelectContainer');
  el.regionSelectContainer = document.getElementById('regionSelectContainer');
  el.assignFileInput = document.getElementById('assignFileInput');
  el.assignUploadBtn = document.getElementById('assignUploadBtn');
  el.assignUpdateBtn = document.getElementById('assignUpdateBtn');
  el.assignStatus = document.getElementById('assignStatus');
  el.assignFilename = document.getElementById('assignFilename');
  el.assignDate = document.getElementById('assignDate');
  el.assignFileRow = document.getElementById('assignFileRow');
  el.assignDateRow = document.getElementById('assignDateRow');
}

// ========== Init ==========

function init() {
  initElements();

  // Risk shift file upload
  el.uploadZone.addEventListener('click', () => el.fileInput.click());
  el.uploadZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    el.uploadZone.classList.add('dragover');
  });
  el.uploadZone.addEventListener('dragleave', () => el.uploadZone.classList.remove('dragover'));
  el.uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    el.uploadZone.classList.remove('dragover');
    if (e.dataTransfer.files[0]) handleRiskFile(e.dataTransfer.files[0]);
  });
  el.fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleRiskFile(e.target.files[0]);
  });

  // Assignment table upload
  el.assignUploadBtn.addEventListener('click', () => el.assignFileInput.click());
  el.assignUpdateBtn.addEventListener('click', () => el.assignFileInput.click());
  el.assignFileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) handleAssignmentFile(e.target.files[0]);
    el.assignFileInput.value = '';
  });

  // Filter toggle
  el.filterRisk.addEventListener('change', () => {
    state.filterRisk = el.filterRisk.checked;
    state.currentPage = 1;
    renderTable();
  });

  // Download buttons
  el.downloadBtn.addEventListener('click', downloadExcel);
  el.downloadImgBtn.addEventListener('click', downloadImage);
  el.generateImgBtn.addEventListener('click', () => {
    if (state.selectedRegions.size === 0) {
      alert('请先在输出结果标题右侧选择区域');
      return;
    }
    generateImage();
  });
  el.copyImgBtn.addEventListener('click', copyImage);
  el.copyTextBtn.addEventListener('click', copyText);

  // Zoom
  el.zoomIn.addEventListener('click', () => {
    state.imageScale = Math.min(state.imageScale + 0.25, 3);
    updateZoom();
  });
  el.zoomOut.addEventListener('click', () => {
    state.imageScale = Math.max(state.imageScale - 0.25, 0.25);
    updateZoom();
  });

  // Load assignment table from localStorage
  loadAssignmentFromStorage();
}

// ========== Assignment Table (localStorage) ==========

function loadAssignmentFromStorage() {
  try {
    const raw = localStorage.getItem(ASSIGN_STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    state.assignmentMap = new Map(data.map);
    state.assignmentInfo = { filename: data.filename, importDate: data.importDate };
    updateAssignmentDisplay();
  } catch (e) {
    console.error('Failed to load assignment table:', e);
  }
}

function saveAssignmentToStorage(map, filename) {
  const data = {
    map: Array.from(map.entries()),
    filename: filename,
    importDate: new Date().toLocaleString('zh-CN'),
  };
  try {
    localStorage.setItem(ASSIGN_STORAGE_KEY, JSON.stringify(data));
  } catch (e) {
    console.error('Failed to save assignment table:', e);
    alert('分工表保存失败，可能是浏览器存储空间不足');
  }
}

function updateAssignmentDisplay() {
  if (state.assignmentMap && state.assignmentMap.size > 0) {
    el.assignStatus.textContent = '已导入';
    el.assignStatus.className = 'info-status active';
    el.assignFilename.textContent = state.assignmentInfo.filename;
    el.assignDate.textContent = state.assignmentInfo.importDate;
    el.assignFileRow.style.display = 'flex';
    el.assignDateRow.style.display = 'flex';
    el.assignUploadBtn.style.display = 'none';
    el.assignUpdateBtn.style.display = 'inline-flex';
  } else {
    el.assignStatus.textContent = '未导入';
    el.assignStatus.className = 'info-status empty';
    el.assignFileRow.style.display = 'none';
    el.assignDateRow.style.display = 'none';
    el.assignUploadBtn.style.display = 'inline-flex';
    el.assignUpdateBtn.style.display = 'none';
  }
}

function handleAssignmentFile(file) {
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!['.xlsx', '.xls'].includes(ext)) {
    alert('请上传 .xlsx 或 .xls 格式的文件');
    return;
  }

  showLoading('正在读取分工表...');
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
      if (data.length < 2) throw new Error('分工表数据不足');

      const headers = data[0];
      const aIdx = findAssignmentColIndices(headers);

      if (aIdx.supplierCode === -1) throw new Error('分工表中未找到"供应商代码"列');

      const map = new Map();
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const code = String(row[aIdx.supplierCode] || '').trim();
        if (!code) continue;
        if (map.has(code)) continue;
        map.set(code, {
          region: aIdx.region >= 0 ? String(row[aIdx.region] || '').trim() : '',
          manager: aIdx.manager >= 0 ? String(row[aIdx.manager] || '').trim() : '',
          assemblyClerk: aIdx.assemblyClerk >= 0 ? String(row[aIdx.assemblyClerk] || '').trim() : '',
        });
      }

      state.assignmentMap = map;
      state.assignmentInfo = { filename: file.name, importDate: new Date().toLocaleString('zh-CN') };
      saveAssignmentToStorage(map, file.name);
      updateAssignmentDisplay();
      hideLoading();

      // If risk data already loaded, reprocess
      if (state.origRows) {
        processData();
      }
    } catch (err) {
      hideLoading();
      alert('分工表处理失败: ' + err.message);
    }
  };
  reader.onerror = () => { hideLoading(); alert('文件读取失败'); };
  reader.readAsArrayBuffer(file);
}

function findAssignmentColIndices(headers) {
  const idx = {};
  for (const [key, names] of Object.entries(ASSIGN_COL_NAMES)) {
    idx[key] = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (names.includes(h)) { idx[key] = i; break; }
    }
  }
  return idx;
}

// ========== Risk Shift File Handling ==========

function handleRiskFile(file) {
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!['.xlsx', '.xls'].includes(ext)) {
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

      if (data.length < 2) throw new Error('文件数据不足');

      state.origHeaders = data[0];
      state.origRows = data.slice(1);
      state.filename = file.name;

      showLoading('正在查找列索引...');
      state.origIdx = findRiskColIndices(state.origHeaders);

      const required = ['theoryStock', 'assemblyWip', 'paintWip', 'secondLockDemand', 'supplierCode', 'vehicleType'];
      for (const key of required) {
        if (state.origIdx[key] === -1) throw new Error(`缺少必要列: ${RISK_COL_NAMES[key].join(' 或 ')}`);
      }

      processData();

      el.uploadText.innerHTML = `<span class="file-info">✓ ${file.name}</span>`;
      el.uploadText.nextElementSibling.textContent = '点击重新上传文件';

      hideLoading();
      el.resultsSection.style.display = 'block';
      el.resultsSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
    } catch (err) {
      hideLoading();
      alert('处理文件时出错: ' + err.message);
      console.error(err);
    }
  };
  reader.onerror = () => { hideLoading(); alert('文件读取失败'); };
  reader.readAsArrayBuffer(file);
}

function findRiskColIndices(headers) {
  const idx = {};
  for (const [key, names] of Object.entries(RISK_COL_NAMES)) {
    idx[key] = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i] || '').trim();
      if (names.includes(h)) { idx[key] = i; break; }
    }
  }
  return idx;
}

// ========== Data Processing ==========

function processData() {
  const oi = state.origIdx;
  const rows = state.origRows;
  const amap = state.assignmentMap || new Map();

  showLoading('正在计算风险数据...');

  state.processedRows = rows.map(row => {
    const theoryStock = Number(row[oi.theoryStock]) || 0;
    const assemblyWip = Number(row[oi.assemblyWip]) || 0;
    const paintWip = Number(row[oi.paintWip]) || 0;
    const secondLockDemand = Number(row[oi.secondLockDemand]) || 0;
    const unauditedBad = oi.unauditedBad >= 0 ? (Number(row[oi.unauditedBad]) || 0) : 0;
    const logisticsMode = oi.logisticsMode >= 0 ? String(row[oi.logisticsMode] || '').trim() : '';
    const supplierCode = String(row[oi.supplierCode] || '').trim();
    const vehicleType = String(row[oi.vehicleType] || '').trim();

    const assemblyGap = theoryStock - assemblyWip - unauditedBad;
    const paintGap = assemblyGap - paintWip;

    const assign = amap.get(supplierCode) || {};
    const region = assign.region || '';
    const manager = assign.manager || '';
    const assemblyClerk = assign.assemblyClerk || '';

    let riskJudge = '';
    if (region) {
      if (logisticsMode === '循环取货' && paintGap < 0 && secondLockDemand !== 0) {
        riskJudge = '不满足涂装';
      } else if (logisticsMode === '厂外拉动' && assemblyGap < 0 && secondLockDemand !== 0) {
        riskJudge = '不满足总装';
      }
    }

    return {
      factory:          oi.factory >= 0 ? row[oi.factory] : '',
      workshop:         oi.workshop >= 0 ? row[oi.workshop] : '',
      partNo:           oi.partNo >= 0 ? row[oi.partNo] : '',
      jianHao:          oi.jianHao >= 0 ? row[oi.jianHao] : '',
      partName:         oi.partName >= 0 ? row[oi.partName] : '',
      supplierCode:     supplierCode,
      supplierName:     oi.supplierName >= 0 ? row[oi.supplierName] : '',
      vehicleType:      vehicleType,
      line:             oi.line >= 0 ? row[oi.line] : '',
      capacity:         oi.capacity >= 0 ? row[oi.capacity] : '',
      theoryStock:      row[oi.theoryStock],
      endFlowNo:        oi.endFlowNo >= 0 ? row[oi.endFlowNo] : '',
      endVIN:           oi.endVIN >= 0 ? row[oi.endVIN] : '',
      assemblyWip:     row[oi.assemblyWip],
      midPaintWip:     oi.midPaintWip >= 0 ? row[oi.midPaintWip] : '',
      paintWip:         row[oi.paintWip],
      weldWip:          oi.weldWip >= 0 ? row[oi.weldWip] : '',
      blueCardWip:      oi.blueCardWip >= 0 ? row[oi.blueCardWip] : '',
      estQueueDemand:   oi.estQueueDemand >= 0 ? row[oi.estQueueDemand] : '',
      firstLockDemand:  oi.firstLockDemand >= 0 ? row[oi.firstLockDemand] : '',
      secondLockDemand: row[oi.secondLockDemand],
      unauditedBad:     oi.unauditedBad >= 0 ? row[oi.unauditedBad] : '',
      logisticsMode:    logisticsMode,
      region:           region,
      manager:          manager,
      assemblyClerk:    assemblyClerk,
      assemblyGap:      Math.round(assemblyGap * 100) / 100,
      paintGap:         Math.round(paintGap * 100) / 100,
      riskJudge:        riskJudge,
    };
  });

  // Extract time from filename
  state.timeStr = extractTimeFromFilename(state.filename);
  // Factory
  const firstRow = state.processedRows[0];
  state.factory = firstRow ? String(firstRow.factory || '未知工厂') : '未知工厂';

  // Extract vehicle types
  state.vehicleTypes = [];
  const vset = new Set();
  for (const r of state.processedRows) {
    const types = String(r.vehicleType || '').split(',').map(t => t.trim()).filter(Boolean);
    for (const t of types) vset.add(t);
  }
  state.vehicleTypes = Array.from(vset).sort();

  // Extract regions (non-null)
  state.regions = [];
  const rset = new Set();
  for (const r of state.processedRows) {
    if (r.region) rset.add(r.region);
  }
  state.regions = Array.from(rset).sort();

  // Default: no regions selected
  state.selectedRegions = new Set();

  calculateStats();
  setupMultiSelects();
  renderTable();
  el.imagePreviewArea.innerHTML = `
    <div class="image-preview-placeholder">
      <svg viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
        <path d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/>
      </svg>
      <p>选择区域后点击「生成图片」</p>
    </div>`;
  state.imageDataUrl = null;
  el.copyImgBtn.style.display = 'none';
  el.downloadImgBtn.style.display = 'none';
  el.textOutput.value = generateText();
}

function calculateStats() {
  const allRows = getFilteredRows(false);
  const riskRows = getFilteredRows(true);
  // 风险点数按已选区域筛选（未选区域时显示0）
  const regionFiltered = riskRows.filter(r => {
    if (state.selectedRegions.size === 0) return false;
    return state.selectedRegions.has(r.region);
  });

  state.stats = {
    total: allRows.length,
    risk: regionFiltered.filter(r => r.riskJudge).length,
    assembly: regionFiltered.filter(r => r.riskJudge === '不满足总装').length,
    paint: regionFiltered.filter(r => r.riskJudge === '不满足涂装').length,
  };

  el.statTotal.textContent = state.stats.total.toLocaleString();
  el.statRisk.textContent = state.stats.risk.toLocaleString();
  el.statAssembly.textContent = state.stats.assembly.toLocaleString();
  el.statPaint.textContent = state.stats.paint.toLocaleString();
}

// ========== Multi-Select Dropdowns ==========

function setupMultiSelects() {
  const persistentVehicles = getPersistentVehicleExclude();

  // Vehicle type exclusion dropdown
  if (state.vehicleSelectRef) {
    state.vehicleSelectRef.updateOptions(state.vehicleTypes);
    state.vehicleSelectRef.setSelected(persistentVehicles);
    state.selectedVehicles = new Set(persistentVehicles);
  } else {
    state.vehicleSelectRef = createMultiSelect(
      el.vehicleSelectContainer,
      state.vehicleTypes,
      {
        label: '排除车型',
        defaultSelected: persistentVehicles,
        showText: true,
        wide: true,
        onChange: (selected) => {
          const newSet = new Set(selected);
          const persistentSet = new Set(getPersistentVehicleExclude());
          const newlyAdded = Array.from(newSet).filter(v => !persistentSet.has(v));
          const newlyRemoved = Array.from(persistentSet).filter(v => !newSet.has(v));

          if (newlyAdded.length > 0 || newlyRemoved.length > 0) {
            const parts = [];
            if (newlyAdded.length > 0) parts.push(`新增：${newlyAdded.join(', ')}`);
            if (newlyRemoved.length > 0) parts.push(`取消：${newlyRemoved.join(', ')}`);
            const keep = confirm(`检测到排除车型变更：\n${parts.join('\n')}\n\n点击「确定」永久保留此选择（下次默认生效）\n点击「取消」仅本次生效（下次恢复默认）`);
            if (keep) {
              try { localStorage.setItem(VEHICLE_EXCLUDE_KEY, JSON.stringify(selected)); } catch (e) {}
            }
          }

          state.selectedVehicles = newSet;
          calculateStats();
          renderTable();
          if (state.imageDataUrl) generateImage();
          el.textOutput.value = generateText();
        }
      }
    );
    state.selectedVehicles = new Set(persistentVehicles);
  }

  // Region filter dropdown
  if (state.regionSelectRef) {
    state.regionSelectRef.updateOptions(state.regions);
    state.regionSelectRef.setSelected([]);
  } else {
    state.regionSelectRef = createMultiSelect(
      el.regionSelectContainer,
      state.regions,
      {
        label: '筛选区域',
        defaultSelected: [],
        showText: true,
        wide: true,
        onChange: (selected) => {
          state.selectedRegions = new Set(selected);
          calculateStats();
          state.currentPage = 1;
          renderTable();
          if (state.imageDataUrl) generateImage();
          el.textOutput.value = generateText();
        }
      }
    );
    state.selectedRegions = new Set();
  }
}

function createMultiSelect(container, options, config) {
  const { label, defaultSelected = [], onChange, showText = false, wide = false } = config;
  const selected = new Set(defaultSelected);

  const btnClass = 'ms-btn' + (wide ? ' ms-btn-wide' : '');
  container.innerHTML = `
    <div class="multi-select">
      <button class="${btnClass}" type="button">
        <span class="ms-label">${label}</span>
        ${showText ? '<span class="ms-text"></span>' : '<span class="ms-count">0</span>'}
        <svg class="ms-arrow" viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
          <path d="M7 10l5 5 5-5z"/>
        </svg>
      </button>
      <div class="ms-dropdown">
        <div class="ms-actions">
          <button class="ms-action-btn" data-action="all" type="button">全选</button>
          <button class="ms-action-btn" data-action="none" type="button">清除</button>
        </div>
        <div class="ms-options"></div>
      </div>
    </div>
  `;

  const btn = container.querySelector('.ms-btn');
  const dropdown = container.querySelector('.ms-dropdown');
  const optionsContainer = container.querySelector('.ms-options');
  const displayEl = container.querySelector(showText ? '.ms-text' : '.ms-count');

  function updateDisplay() {
    if (showText) {
      displayEl.textContent = selected.size > 0 ? Array.from(selected).join(', ') : '无';
    } else {
      displayEl.textContent = selected.size;
    }
  }

  function renderOptions() {
    optionsContainer.innerHTML = options.map(opt => `
      <label class="ms-option">
        <input type="checkbox" value="${escapeHtml(opt)}" ${selected.has(opt) ? 'checked' : ''}>
        <span>${escapeHtml(opt)}</span>
      </label>
    `).join('');

    optionsContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      cb.addEventListener('change', () => {
        if (cb.checked) selected.add(cb.value);
        else selected.delete(cb.value);
        updateDisplay();
        if (onChange) onChange(Array.from(selected));
      });
    });
    updateDisplay();
  }

  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    btn.classList.toggle('open');
    dropdown.classList.toggle('show');
  });

  document.addEventListener('click', (e) => {
    if (!container.contains(e.target)) {
      btn.classList.remove('open');
      dropdown.classList.remove('show');
    }
  });

  container.querySelector('[data-action="all"]').addEventListener('click', (e) => {
    e.stopPropagation();
    options.forEach(opt => selected.add(opt));
    renderOptions();
    updateDisplay();
    if (onChange) onChange(Array.from(selected));
  });

  container.querySelector('[data-action="none"]').addEventListener('click', (e) => {
    e.stopPropagation();
    selected.clear();
    renderOptions();
    updateDisplay();
    if (onChange) onChange(Array.from(selected));
  });

  renderOptions();

  return {
    getSelected: () => Array.from(selected),
    setSelected: (vals) => {
      selected.clear();
      vals.forEach(v => selected.add(v));
      renderOptions();
      updateDisplay();
    },
    updateOptions: (newOptions) => {
      options = newOptions;
      const optionSet = new Set(options);
      for (const val of selected) {
        if (!optionSet.has(val)) selected.delete(val);
      }
      renderOptions();
      updateDisplay();
    }
  };
}

// ========== Vehicle Type Exclusion ==========

function isRowExcluded(row) {
  const types = String(row.vehicleType || '').split(',').map(t => t.trim()).filter(Boolean);
  if (types.length === 0) return false;
  return types.every(t => state.selectedVehicles.has(t));
}

function getFilteredRows(includeRiskOnly) {
  if (!state.processedRows) return [];
  return state.processedRows.filter(r => {
    if (isRowExcluded(r)) return false;
    if (includeRiskOnly) {
      if (!r.riskJudge) return false;
      if (!r.region) return false;
    }
    return true;
  });
}

function getDisplayRows() {
  const riskOnly = el.filterRisk ? el.filterRisk.checked : true;
  let rows = getFilteredRows(riskOnly);
  if (state.selectedRegions.size > 0) {
    rows = rows.filter(r => state.selectedRegions.has(r.region));
  }
  return rows;
}

// ========== Table Rendering ==========

function renderTable() {
  const cols = DISPLAY_SPEC;

  // Header
  let thead = '<tr>';
  cols.forEach(col => {
    let cls = '';
    if (col.isNew) cls += ' col-new';
    if (col.type === 'assemblyGap') cls += ' col-highlight-assembly';
    else if (col.type === 'paintGap') cls += ' col-highlight-paint';
    thead += `<th class="${cls.trim()}">${col.label}</th>`;
  });
  thead += '</tr>';
  el.tableHead.innerHTML = thead;

  // Body
  const filtered = getDisplayRows();
  const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
  if (state.currentPage > totalPages) state.currentPage = totalPages;

  const start = (state.currentPage - 1) * state.pageSize;
  const end = Math.min(start + state.pageSize, filtered.length);
  const pageRows = filtered.slice(start, end);

  let tbody = '';
  for (const row of pageRows) {
    tbody += '<tr>';
    for (const col of cols) {
      const value = row[col.key];
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

  // Pagination
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

// ========== Image Generation ==========

function generateImage() {
  const riskRows = getFilteredRows(true).filter(r => {
    if (!r.riskJudge) return false;
    if (!r.region) return false;
    if (state.selectedRegions.size > 0 && !state.selectedRegions.has(r.region)) return false;
    return true;
  });

  if (riskRows.length === 0) {
    el.imagePreviewArea.innerHTML = `
      <div class="image-preview-placeholder">
        <svg viewBox="0 0 24 24" fill="currentColor" xmlns="http://www.w3.org/2000/svg">
          <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 15.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93z"/>
        </svg>
        <p>暂无风险数据</p>
      </div>`;
    state.imageDataUrl = null;
    el.copyImgBtn.style.display = 'none';
    el.downloadImgBtn.style.display = 'none';
    return;
  }

  const cols = IMAGE_SPEC;
  const fontSize = 12;
  const headerHeight = 28;
  const cellHeight = 22;
  const padding = 8;
  const maxRows = 800;

  const displayRows = riskRows.slice(0, maxRows);
  const truncated = riskRows.length > maxRows;

  // Measure column widths
  const measureCanvas = document.createElement('canvas');
  const mctx = measureCanvas.getContext('2d');
  mctx.font = `${fontSize}px "Microsoft YaHei", "PingFang SC", sans-serif`;

  const colWidths = cols.map(col => {
    let maxWidth = mctx.measureText(col.label).width;
    const sampleSize = Math.min(displayRows.length, 100);
    for (let i = 0; i < sampleSize; i++) {
      const text = formatCellValue(displayRows[i][col.key]);
      const width = mctx.measureText(text).width;
      if (width > maxWidth) maxWidth = width;
    }
    return Math.min(Math.ceil(maxWidth + padding * 2), 200);
  });

  const totalWidth = colWidths.reduce((a, b) => a + b, 0);
  const totalHeight = headerHeight + displayRows.length * cellHeight + (truncated ? 30 : 0);

  // Determine scale: use 2x if within browser limits, else 1x
  const maxCanvasDim = 32767;
  let scale = 2;
  if (totalWidth * scale > maxCanvasDim || totalHeight * scale > maxCanvasDim) {
    scale = 1;
  }

  const canvas = document.createElement('canvas');
  canvas.width = totalWidth * scale;
  canvas.height = totalHeight * scale;
  canvas.style.width = totalWidth + 'px';
  canvas.style.height = totalHeight + 'px';

  const ctx = canvas.getContext('2d');
  ctx.scale(scale, scale);
  ctx.font = `${fontSize}px "Microsoft YaHei", "PingFang SC", sans-serif`;
  ctx.textBaseline = 'middle';

  // Draw header
  let x = 0;
  cols.forEach((col, i) => {
    let headerBg = '#2563EB';
    if (col.type === 'assemblyGap') headerBg = '#FDD835';
    else if (col.type === 'paintGap') headerBg = '#EF9A9A';
    else if (col.isNew) headerBg = '#1D4ED8';

    ctx.fillStyle = headerBg;
    ctx.fillRect(x, 0, colWidths[i], headerHeight);

    ctx.fillStyle = (col.type === 'assemblyGap' || col.type === 'paintGap') ? '#1E293B' : 'white';
    ctx.textAlign = 'center';
    ctx.fillText(col.label, x + colWidths[i] / 2, headerHeight / 2);

    ctx.strokeStyle = col.type === 'assemblyGap' || col.type === 'paintGap' ? '#B0BEC5' : '#1D4ED8';
    ctx.lineWidth = 1;
    ctx.strokeRect(x, 0, colWidths[i], headerHeight);
    x += colWidths[i];
  });

  // Draw data rows
  displayRows.forEach((row, rowIdx) => {
    const y = headerHeight + rowIdx * cellHeight;
    let x = 0;

    cols.forEach((col, colIdx) => {
      const value = row[col.key];
      let bg = rowIdx % 2 === 0 ? '#FFFFFF' : '#F8FAFC';

      if (col.type === 'assemblyGap' && Number(value) < 0) bg = '#FFF9C4';
      else if (col.type === 'paintGap' && Number(value) < 0) bg = '#FFCDD2';

      ctx.fillStyle = bg;
      ctx.fillRect(x, y, colWidths[colIdx], cellHeight);

      ctx.strokeStyle = '#CBD5E1';
      ctx.lineWidth = 0.5;
      ctx.strokeRect(x, y, colWidths[colIdx], cellHeight);

      let textColor = '#1E293B';
      if (col.type === 'riskJudge') {
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

  // Truncation notice
  if (truncated) {
    const noticeY = headerHeight + displayRows.length * cellHeight + 5;
    ctx.fillStyle = '#64748B';
    ctx.font = `11px "Microsoft YaHei", sans-serif`;
    ctx.textAlign = 'left';
    ctx.fillText(`仅显示前 ${maxRows} 行，共 ${riskRows.length} 行风险数据，请下载Excel查看全部`, 10, noticeY + 10);
  }

  try {
    const dataUrl = canvas.toDataURL('image/png');
    state.imageDataUrl = dataUrl;
    state.imageScale = 1;

    el.imagePreviewArea.innerHTML = `<img src="${dataUrl}" id="previewImage" style="transform: scale(${state.imageScale}); transform-origin: top center;">`;
    el.copyImgBtn.style.display = 'inline-flex';
    el.downloadImgBtn.style.display = 'inline-flex';
    updateZoom();
  } catch (err) {
    console.error('Image generation failed:', err);
    el.imagePreviewArea.innerHTML = `
      <div class="image-preview-placeholder">
        <p>图片生成失败，请尝试减少数据量</p>
      </div>`;
    state.imageDataUrl = null;
  }
}

function updateZoom() {
  const img = document.getElementById('previewImage');
  if (img) img.style.transform = `scale(${state.imageScale})`;
  el.zoomLevel.textContent = Math.round(state.imageScale * 100) + '%';
}

function downloadImage() {
  if (!state.imageDataUrl) { alert('暂无图片可下载'); return; }
  const link = document.createElement('a');
  link.download = `风险推移图片_${state.timeStr.replace(':', '')}.png`;
  link.href = state.imageDataUrl;
  link.click();
}

async function copyImage() {
  if (!state.imageDataUrl) { alert('暂无图片可复制'); return; }
  try {
    const resp = await fetch(state.imageDataUrl);
    const blob = await resp.blob();
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
    const original = el.copyImgBtn.textContent;
    el.copyImgBtn.textContent = '已复制 ✓';
    setTimeout(() => { el.copyImgBtn.textContent = original; }, 2000);
  } catch (err) {
    alert('复制图片失败，请改用「下载图片」按钮\n（部分浏览器不支持图片复制）');
  }
}

// ========== Text Generation ==========

function generateText() {
  const factoryShort = state.factory.replace('工厂', '厂');
  const riskRows = getFilteredRows(true).filter(r => r.riskJudge);

  // Filter by selected regions
  const regionFiltered = riskRows.filter(r => {
    if (!r.region) return false;
    return state.selectedRegions.has(r.region);
  });

  // Build region string
  const regionStr = Array.from(state.selectedRegions).filter(r => r).join('，');

  let text = `${factoryShort}风险推移：\n`;
  text += `时间：${state.timeStr}\n`;
  text += `区域：${regionStr}\n`;
  text += `推移人：\n`;
  text += `风险推移风险点数：${regionFiltered.length}\n`;
  text += `风险零件跟进结果如下：\n`;

  // 拉动：不满足总装
  const assemblyRows = regionFiltered.filter(r => r.riskJudge === '不满足总装');
  if (assemblyRows.length > 0) {
    text += `拉动：不满足总装\n`;
    text += `厂外拉动\n`;
    const suppliers = {};
    assemblyRows.forEach(r => {
      const name = r.supplierName || '未知供应商';
      if (!suppliers[name]) suppliers[name] = new Set();
      if (r.jianHao) suppliers[name].add(r.jianHao);
    });
    Object.keys(suppliers).forEach((name, si) => {
      if (si > 0) text += '\n';
      text += `${name}\n`;
      Array.from(suppliers[name]).forEach(jh => {
        text += `        ${jh}\n`;
      });
    });
  }

  // 循环：不满足涂装
  const paintRows = regionFiltered.filter(r => r.riskJudge === '不满足涂装');
  if (paintRows.length > 0) {
    if (assemblyRows.length > 0) text += '\n';
    text += `循环：不满足涂装\n`;
    const suppliers = {};
    paintRows.forEach(r => {
      const name = r.supplierName || '未知供应商';
      if (!suppliers[name]) suppliers[name] = new Set();
      if (r.jianHao) suppliers[name].add(r.jianHao);
    });
    Object.keys(suppliers).forEach((name, si) => {
      if (si > 0) text += '\n';
      text += `${name}\n`;
      Array.from(suppliers[name]).forEach(jh => {
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
  } catch (err) {
    textarea.select();
    document.execCommand('copy');
  }
  const original = el.copyTextBtn.textContent;
  el.copyTextBtn.textContent = '已复制 ✓';
  setTimeout(() => { el.copyTextBtn.textContent = original; }, 2000);
}

// ========== Excel Download (ExcelJS) ==========

async function downloadExcel() {
  showLoading('正在生成Excel文件...');
  try {
    // All rows after vehicle exclusion, filtered by selected regions
    const allRows = getFilteredRows(false).filter(r => {
      if (state.selectedRegions.size === 0) return true;
      return state.selectedRegions.has(r.region);
    });
    if (allRows.length === 0) {
      hideLoading();
      alert('没有数据可下载');
      return;
    }

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('风险推移', {
      views: [{ state: 'frozen', ySplit: 1 }],
    });

    // Use full column spec (includes hidden columns)
    const cols = EXCEL_SPEC;
    ws.columns = cols.map(col => ({
      header: col.label,
      key: col.key,
      width: 18,
      hidden: col.hidden || false,
    }));

    // Style header row
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    headerRow.alignment = { horizontal: 'center', vertical: 'middle' };
    headerRow.height = 24;

    cols.forEach((col, i) => {
      const cell = headerRow.getCell(i + 1);
      if (col.type === 'assemblyGap') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDD835' } };
        cell.font = { bold: true, color: { argb: 'FF1E293B' }, size: 11 };
      } else if (col.type === 'paintGap') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEF9A9A' } };
        cell.font = { bold: true, color: { argb: 'FF1E293B' }, size: 11 };
      } else if (col.isNew) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2563EB' } };
      } else {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1D4ED8' } };
      }
      cell.border = {
        top: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        bottom: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        left: { style: 'thin', color: { argb: 'FFCBD5E1' } },
        right: { style: 'thin', color: { argb: 'FFCBD5E1' } },
      };
    });

    // Add ALL data rows
    allRows.forEach(rowData => {
      const row = ws.addRow(rowData);
      row.alignment = { horizontal: 'center', vertical: 'middle' };
      row.font = { size: 10 };

      cols.forEach((col, i) => {
        const cell = row.getCell(i + 1);

        if (col.type === 'assemblyGap') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF9C4' } };
        } else if (col.type === 'paintGap') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFCDD2' } };
        } else if (col.type === 'region' || col.type === 'manager' || col.type === 'clerk') {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0F2FE' } };
        }

        if (col.type === 'riskJudge') {
          if (cell.value === '不满足总装') {
            cell.font = { bold: true, color: { argb: 'FFDC2626' }, size: 10 };
          } else if (cell.value === '不满足涂装') {
            cell.font = { bold: true, color: { argb: 'FFF59E0B' }, size: 10 };
          }
        }

        if (col.type === 'assemblyGap' && Number(cell.value) < 0) {
          cell.font = { bold: true, color: { argb: 'FFDC2626' }, size: 10 };
        }
        if (col.type === 'paintGap' && Number(cell.value) < 0) {
          cell.font = { bold: true, color: { argb: 'FFF59E0B' }, size: 10 };
        }

        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
          right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
        };
      });
    });

    // Auto-filter + default filter to show only risk rows
    const lastColLetter = ws.getColumn(cols.length).letter;
    const riskColIdx = cols.findIndex(c => c.type === 'riskJudge');
    const riskColLetter = ws.getColumn(riskColIdx + 1).letter;
    ws.autoFilter = {
      from: `A1`,
      to: `${lastColLetter}${allRows.length + 1}`,
      // Default: filter riskJudge column to non-blank values
      columnFilter: {},
    };
    // Set column filter for riskJudge (non-empty values)
    ws.autoFilter.columnFilter[riskColIdx + 1] = {
      filters: ['不满足总装', '不满足涂装'],
    };

    // Auto-fit column widths
    ws.columns.forEach(col => {
      let maxLen = col.header.length;
      col.eachCell({ includeEmpty: true }, cell => {
        const val = cell.value;
        const len = val ? String(val).length : 0;
        if (len > maxLen) maxLen = len;
      });
      col.width = Math.min(maxLen + 4, 35);
    });

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `风险推移_${state.timeStr.replace(':', '')}.xlsx`;
    a.click();
    URL.revokeObjectURL(url);

    hideLoading();
  } catch (err) {
    hideLoading();
    alert('Excel生成失败: ' + err.message);
    console.error(err);
  }
}

// ========== Utility ==========

function showLoading(text) {
  el.loadingText.textContent = text || '正在处理...';
  el.loadingOverlay.style.display = 'flex';
}

function hideLoading() {
  el.loadingOverlay.style.display = 'none';
}

function extractTimeFromFilename(filename) {
  const match = filename.match(/(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})/);
  if (match) return `${match[4]}:${match[5]}`;
  const now = new Date();
  const h = String(now.getHours()).padStart(2, '0');
  const m = String(now.getMinutes()).padStart(2, '0');
  return `${h}:${m}`;
}

function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = String(str);
  return div.innerHTML;
}

document.addEventListener('DOMContentLoaded', init);
