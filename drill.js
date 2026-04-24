/* global d3, tableau */

// ==================== CONFIGURATION ====================
const barWidth = 200;
const barHeight = 18;
const margin = { top: 60, right: 180, bottom: 40, left: 20 };
const palette = ['#007bff', '#28a745', '#dc3545', '#ffc107', '#17a2b8'];

const MEASURE_COLORS = {
  primary: '#007bff',
  target:  '#000000'
};

let svg, g, root, i = 0;
let width, height, dx, dy;
const duration = 500;
let expandedPaths = new Set();
let userDimensionOrder = null;

let currentMeasure = null;
let currentTarget = null;

// ==================== EXTENSION TYPE DETECTION ====================
let extensionMode = null;

function detectExtensionMode() {
  try {
    if (tableau.extensions.worksheetContent && tableau.extensions.worksheetContent.worksheet) {
      extensionMode = 'worksheet';
      console.log('✅ Extension Mode: WORKSHEET');
      return;
    }
  } catch (e) {}

  try {
    if (tableau.extensions.dashboardContent && tableau.extensions.dashboardContent.dashboard) {
      extensionMode = 'dashboard';
      console.log('✅ Extension Mode: DASHBOARD');
      return;
    }
  } catch (e) {}

  extensionMode = null;
  console.warn('⚠️ Could not detect extension mode');
}

function getWorksheet() {
  if (extensionMode === 'worksheet') {
    return tableau.extensions.worksheetContent.worksheet;
  }
  if (extensionMode === 'dashboard') {
    const dashboard = tableau.extensions.dashboardContent.dashboard;
    return dashboard.worksheets[0];
  }
  return null;
}

function getAllWorksheetsForFilter() {
  if (extensionMode === 'dashboard') {
    return tableau.extensions.dashboardContent.dashboard.worksheets;
  }
  if (extensionMode === 'worksheet') {
    return [tableau.extensions.worksheetContent.worksheet];
  }
  return [];
}

// ==================== DOM SETUP ====================
function setupDOM() {
  if (!document.getElementById('toolbar')) {
    const toolbar = document.createElement('div');
    toolbar.id = 'toolbar';
    toolbar.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      width: 100%;
      height: 44px;
      background: #fff;
      border-bottom: 1px solid #e0e0e0;
      z-index: 9999;
      display: flex;
      align-items: center;
      padding: 0 12px;
      box-sizing: border-box;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    `;
    document.body.appendChild(toolbar);
  }

  if (!document.getElementById('chart-container')) {
    const div = document.createElement('div');
    div.id = 'chart-container';
    document.body.appendChild(div);
  }

  const el = document.getElementById('chart-container');
  el.style.cssText = `
    position: fixed;
    top: 44px;
    left: 0;
    right: 0;
    bottom: 0;
    overflow-x: auto;
    overflow-y: auto;
    background: #fff;
  `;

  // BLANK AREA CLICK → AUTO CLEAR SELECTION
  el.addEventListener('click', function (event) {
    // Check if click was on blank area (not on any node element)
    const target = event.target;
    const isNode = target.closest && target.closest('g.node');
    const isToggle = target.closest && target.closest('.toggle-icon');
    const isToolbar = target.closest && target.closest('#toolbar');
    const isOverlay = target.closest && target.closest('#config-overlay');

    if (!isNode && !isToggle && !isToolbar && !isOverlay) {
      console.log('🖱️ Blank area clicked → Clearing selection');
      clearNodeHighlights();
      clearAllSelections();
    }
  });

  document.body.style.cssText = 'margin:0;padding:0;overflow:hidden;width:100%;height:100%;';
  document.documentElement.style.cssText = 'margin:0;padding:0;overflow:hidden;width:100%;height:100%;';
}

// ==================== INITIALIZATION ====================
window.onload = function () {
  setupDOM();
  if (typeof tableau === 'undefined') {
    showError('Tableau Extensions API not found.');
    return;
  }

  tableau.extensions.initializeAsync().then(() => {
    detectExtensionMode();

    if (!extensionMode) {
      showError('Could not detect extension mode. Make sure extension is added properly.');
      return;
    }

    loadSettings();
    loadData();

    const ws = getWorksheet();
    if (ws) {
      ws.addEventListener(
        tableau.TableauEventType.SummaryDataChanged,
        () => {
          saveExpandedState();
          loadData();
        }
      );
    }

    window.addEventListener('resize', () => {
      if (root) {
        updateSize();
        update(root);
      }
    });
  }).catch(err => showError(err.message));
};

// ==================== SETTINGS ====================
function loadSettings() {
  try {
    const saved = tableau.extensions.settings.get('dimensionOrder');
    if (saved) userDimensionOrder = JSON.parse(saved);
  } catch (e) {
    console.warn('⚠️ Could not load settings:', e.message);
  }
}

function saveSettings(dimensionOrder) {
  try {
    userDimensionOrder = dimensionOrder;
    tableau.extensions.settings.set('dimensionOrder', JSON.stringify(dimensionOrder));
    return tableau.extensions.settings.saveAsync();
  } catch (e) {
    return Promise.resolve();
  }
}

// ==================== CHART INIT ====================
function initChart() {
  const container = document.getElementById('chart-container');
  container.innerHTML = '';

  width = container.clientWidth || 800;
  height = container.clientHeight || 600;
  dx = 85;
  dy = Math.max(barWidth + 160, Math.floor(width / 5));

  const svgWrap = document.createElement('div');
  svgWrap.id = 'svg-wrap';
  svgWrap.style.cssText = 'position:relative;';
  container.appendChild(svgWrap);

  svg = d3.select('#svg-wrap').append('svg')
    .style('font', '12px sans-serif')
    .style('user-select', 'none')
    .style('display', 'block');

  // SVG BLANK AREA CLICK → AUTO CLEAR
  svg.on('click', function (event) {
    const target = event.target;
    const isNode = target.closest && target.closest('g.node');
    const isToggle = target.closest && target.closest('.toggle-icon');

    if (!isNode && !isToggle) {
      console.log('🖱️ SVG blank area clicked → Clearing selection');
      clearNodeHighlights();
      clearAllSelections();
    }
  });

  g = svg.append('g');
}

function updateSize() {
  const c = document.getElementById('chart-container');
  width = c.clientWidth || 800;
  height = c.clientHeight || 600;
  dy = Math.max(barWidth + 160, Math.floor(width / 5));
}

// ==================== MEASURE NAME MATCHING ====================
function stripOneAggregationLayer(s) {
  if (!s) return s;
  const m = String(s).trim().match(/^(SUM|AVG|COUNTD?|MIN|MAX|ATTR|MEDIAN|AGG|COUNT)\s*\(([\s\S]*)\)\s*$/i);
  return m ? m[2].trim() : s;
}

function normalizeMeasureKey(name) {
  if (name == null || name === '') return '';
  let s = String(name).trim();

  while (/^\[[\s\S]+\]$/.test(s)) s = s.slice(1, -1).trim();

  let prev;
  do {
    prev = s;
    s = stripOneAggregationLayer(s);
    while (/^\[[\s\S]+\]$/.test(s)) s = s.slice(1, -1).trim();
  } while (s !== prev);

  return s.toLowerCase().replace(/\s+/g, ' ').trim();
}

function findMeasureMatchingField(fieldName, allMeasures) {
  if (!fieldName || !allMeasures.length) return null;
  const raw = String(fieldName).trim();

  const exact = allMeasures.find(m => m === raw);
  if (exact) return exact;

  const key = normalizeMeasureKey(raw);
  if (!key) return null;

  const candidates = allMeasures.filter(m => normalizeMeasureKey(m) === key);
  if (candidates.length === 1) return candidates[0];
  if (candidates.length === 0) return null;

  const ordered = [...candidates].sort(
    (a, b) => allMeasures.indexOf(a) - allMeasures.indexOf(b)
  );
  const byExact = ordered.find(m => m === raw);
  return byExact || ordered[0];
}

function orderDimensionsByVisualShelf(allDimensions, vizSpec) {
  if (!vizSpec || !allDimensions.length) return [...allDimensions];
  const combined = [
    ...(vizSpec.rowFields || []),
    ...(vizSpec.columnFields || [])
  ];
  if (!combined.length) return [...allDimensions];

  const remaining = new Set(allDimensions);
  const ordered = [];
  for (const f of combined) {
    if (!f) continue;
    const fieldLabel = f.name != null && String(f.name).trim() !== ''
      ? String(f.name).trim()
      : '';
    if (!fieldLabel) continue;
    const match = findMeasureMatchingField(fieldLabel, [...remaining]);
    if (match) {
      ordered.push(match);
      remaining.delete(match);
    }
  }
  for (const d of allDimensions) {
    if (remaining.has(d)) ordered.push(d);
  }
  return ordered.length ? ordered : [...allDimensions];
}

function findDetailMeasureInVisualSpec(vizSpec, allMeasures) {
  if (!vizSpec || !allMeasures.length) return null;

  if (vizSpec.encodingBindings) {
    for (const binding of vizSpec.encodingBindings) {
      const encType = (binding.encodingType || binding.type || '').toLowerCase();
      const field = binding?.field || {};
      const fieldName = field.fieldName || field.name || field.caption || '';
      if (encType === 'detail') {
        const matched = findMeasureMatchingField(fieldName, allMeasures);
        if (matched) return matched;
      }
    }
  }

  const marksList = vizSpec.marksSpecifications || [];
  const tableauDetail =
    typeof tableau !== 'undefined' && tableau.EncodingType && tableau.EncodingType.Detail;

  for (const ms of marksList) {
    const encodings = ms.encodings || [];
    for (const enc of encodings) {
      const rawType = enc.type;
      const isDetail = tableauDetail != null && rawType === tableauDetail
        ? true
        : (() => {
          const encType = rawType != null ? String(rawType).toLowerCase() : '';
          return encType === 'detail' || encType.endsWith('detail');
        })();

      if (!isDetail) continue;

      const field = enc.field || {};
      const fieldName = field.name || field.fieldName || '';
      const matched = findMeasureMatchingField(fieldName, allMeasures);
      if (matched) return matched;
    }
  }

  return null;
}

// ==================== FIELD DETECTION ====================
async function detectAllFields(columns, worksheet) {
  const SKIP = [
    /^measure names$/i,
    /^measure values$/i,
    /^number of records$/i
  ];

  const TARGET_PATTERNS = [
    /target/i, /budget/i, /goal/i, /plan/i,
    /forecast/i, /benchmark/i, /expected/i
  ];

  const tagged = columns.map((col, idx) => {
    const name = col.fieldName || `Col_${idx}`;
    const dtype = (col.dataType || '').toLowerCase();
    const role = (col.role || '').toLowerCase();

    if (SKIP.some(p => p.test(name))) return { name, kind: 'skip', idx };
    if (role === 'dimension' || ['string', 'bool', 'date', 'datetime'].includes(dtype)) return { name, kind: 'dimension', idx };
    if (role === 'measure' || ['int', 'integer', 'float', 'double'].includes(dtype)) return { name, kind: 'measure', idx };
    return { name, kind: 'unknown', idx };
  });

  let allDimensions = tagged.filter(t => t.kind === 'dimension').map(t => t.name);
  const allMeasures = tagged.filter(t => t.kind === 'measure').map(t => t.name);

  let vizSpec = null;
  try {
    if (worksheet && typeof worksheet.getVisualSpecificationAsync === 'function') {
      vizSpec = await worksheet.getVisualSpecificationAsync();
    }
  } catch (e) {
    console.warn('⚠️ Visual spec not available:', e.message);
  }

  if (vizSpec) {
    const shelfOrdered = orderDimensionsByVisualShelf(allDimensions, vizSpec);
    if (shelfOrdered.length) allDimensions = shelfOrdered;
  }

  console.log('🔍 Dimensions (drill order):', allDimensions);
  console.log('🔍 All measures:', allMeasures);

  let measure = null;
  let targetMeasure = null;

  if (allMeasures.length === 0) {
    // nothing
  } else if (allMeasures.length === 1) {
    measure = allMeasures[0];
  } else {
    let detailMeasure = null;

    if (vizSpec) {
      detailMeasure = findDetailMeasureInVisualSpec(vizSpec, allMeasures);
      if (detailMeasure) {
        console.log(`✅ Detail shelf measure found: "${detailMeasure}"`);
      }
    }

    if (detailMeasure) {
      targetMeasure = detailMeasure;
      measure = allMeasures.find(m => m !== detailMeasure) || allMeasures[0];
    } else {
      const tIdx = allMeasures.findIndex(m => TARGET_PATTERNS.some(p => p.test(m)));

      if (tIdx >= 0) {
        targetMeasure = allMeasures[tIdx];
        measure = allMeasures.find((_, idx) => idx !== tIdx) || allMeasures[0];
      } else {
        measure = allMeasures[0];
        targetMeasure = allMeasures.length > 1 ? allMeasures[1] : null;
      }
    }
  }

  console.log(`✅ FINAL → Value bar: "${measure}" | Target line: "${targetMeasure}"`);
  return { allDimensions, measure, targetMeasure };
}

// ==================== DIMENSION ORDER ====================
function determineDimensionOrder(allDimensions) {
  if (userDimensionOrder && userDimensionOrder.length > 0) {
    const cs = new Set(allDimensions);
    const ss = new Set(userDimensionOrder);
    const ordered = userDimensionOrder.filter(d => cs.has(d));
    allDimensions.forEach(d => {
      if (!ss.has(d)) ordered.push(d);
    });
    return ordered;
  }
  return allDimensions;
}

// ==================== DATA LOADING ====================
async function loadData() {
  try {
    const ws = getWorksheet();

    if (!ws) {
      throw new Error('No worksheet found. Make sure extension is properly configured.');
    }

    console.log(`📋 Using worksheet: "${ws.name}" (mode: ${extensionMode})`);

    const dataTable = await ws.getSummaryDataAsync({
      ignoreSelection: true,
      maxRows: 0
    });

    if (!dataTable?.data?.length || !dataTable?.columns?.length) {
      throw new Error('No data found.');
    }

    console.log(`📊 ${dataTable.data.length} rows, ${dataTable.columns.length} cols`);

    const { allDimensions, measure, targetMeasure } =
      await detectAllFields(dataTable.columns, ws);

    currentMeasure = measure;
    currentTarget = targetMeasure;

    if (!allDimensions.length) {
      showError('No dimensions detected.');
      return;
    }

    const dimensions = determineDimensionOrder(allDimensions);
    const rawData = parseDataTable(dataTable);
    const hierarchy = buildHierarchy(rawData, dimensions, measure, targetMeasure);

    if (!hierarchy.children?.length) {
      showError('Empty hierarchy.');
      return;
    }

    renderChart(hierarchy, dimensions);
    console.log('✅ Chart rendered!');
  } catch (err) {
    console.error('❌ Error:', err);
    showError('Error: ' + err.message);
  }
}

// ==================== DATA PARSING ====================
function parseDataTable(dataTable) {
  return dataTable.data.map(row =>
    Object.fromEntries(dataTable.columns.map((col, c) => {
      const name = col.fieldName || `Col_${c}`;
      const cell = row[c];
      let val;
      if (cell !== null && typeof cell === 'object') {
        val = ('nativeValue' in cell && cell.nativeValue != null) ? cell.nativeValue : cell.value;
      } else {
        val = cell;
      }
      return [name, val];
    }))
  );
}

// ==================== HIERARCHY ====================
function buildHierarchy(data, dimensions, measure, targetMeasure) {
  const rootNode = {
    name: cleanName(measure) || 'Total',
    children: [],
    value: 0,
    target: 0,
    path: 'root',
    level: 0
  };

  const map = new Map([['root', rootNode]]);

  data.forEach(row => {
    const rv = measure ? (parseFloat(row[measure]) || 0) : 1;
    const rt = targetMeasure ? (parseFloat(row[targetMeasure]) || 0) : 0;
    if (rv === 0 && rt === 0) return;

    let parentPath = 'root';

    dimensions.forEach((dim, levelIndex) => {
      const raw = row[dim];
      const label = (raw == null || String(raw).trim() === '' || String(raw) === 'null')
        ? `(Unknown ${dim})`
        : String(raw).trim();

      const path = `${parentPath}|L${levelIndex}:${dim}:${label}`;

      if (!map.has(path)) {
        map.set(path, {
          name: label,
          type: dim,
          level: levelIndex + 1,
          path,
          isLastLevel: levelIndex === dimensions.length - 1,
          children: [],
          value: 0,
          target: 0
        });
        map.get(parentPath).children.push(map.get(path));
      }

      map.get(path).value += rv;
      map.get(path).target += rt;
      parentPath = path;
    });
  });

  rootNode.value = rootNode.children.reduce((s, c) => s + c.value, 0);
  rootNode.target = rootNode.children.reduce((s, c) => s + c.target, 0);

  function sortChildren(node) {
    if (node.children && node.children.length > 0) {
      node.children.sort((a, b) => {
        const dv = (b.value || 0) - (a.value || 0);
        if (dv !== 0) return dv;
        return String(a.name).localeCompare(String(b.name), undefined, { sensitivity: 'base' });
      });
      node.children.forEach(sortChildren);
    }
  }

  sortChildren(rootNode);
  return rootNode;
}

// ==================== HELPERS ====================
function showError(msg) {
  const c = document.getElementById('chart-container');
  if (!c) return;
  svg = null;
  g = null;
  c.innerHTML = `
    <div style="font-family:Arial;text-align:center;position:absolute;top:50%;left:50%;
      transform:translate(-50%,-50%);padding:32px 40px;background:#fff8f8;
      border:1px solid #f5c6c6;border-radius:10px;max-width:520px;">
      <div style="font-size:36px;margin-bottom:8px;">⚠️</div>
      <h3 style="margin:0 0 10px;color:#c0392b;">Chart Not Ready</h3>
      <p style="margin:0;font-size:13px;color:#444;">${msg}</p>
    </div>`;
}

function getNodePath(d) {
  return d.data.path || 'unknown';
}

function collapseAll(node) {
  if (node.children) {
    node._children = node.children;
    node.children = null;
  }
  if (node._children) node._children.forEach(collapseAll);
}

function saveExpandedState() {
  expandedPaths.clear();
  if (root) {
    root.descendants().forEach(d => {
      if (d.children) expandedPaths.add(getNodePath(d));
    });
  }
}

function cleanName(n) {
  if (!n) return '';
  return n.replace(/^(SUM|AVG|COUNT|COUNTD|MIN|MAX|ATTR|MEDIAN|AGG)\(/i, '').replace(/\)$/, '');
}

function clearNodeHighlights() {
  if (!g) return;
  g.selectAll('.bar-bg')
    .style('stroke', 'none')
    .style('stroke-width', null)
    .style('fill', '#efefef');
}

function highlightNode(d) {
  if (!g || !d) return;

  clearNodeHighlights();

  const nodeSelection = g.selectAll('g.node').filter(node => node === d);
  nodeSelection.select('.bar-bg')
    .style('stroke', '#ff6600')
    .style('stroke-width', '2px')
    .style('fill', '#fff3e0');
}

function highlightAndFilter(event, d) {
  event.stopPropagation();
  if (!d || d.depth === 0) return;
  highlightNode(d);
  applyFilter(d);
}

// ==================== COLLECT ALL ANCESTOR FILTER PAIRS ====================
function getAncestorFilterPairs(d) {
  const pairs = [];
  let current = d;

  while (current && current.depth > 0) {
    if (current.data.type && current.data.name) {
      pairs.unshift({
        fieldName: current.data.type,
        value: current.data.name
      });
    }
    current = current.parent;
  }

  return pairs;
}

// ==================== RENDER ====================
function renderChart(data, dimensions) {
  initChart();
  addConfigButton(dimensions);

  root = d3.hierarchy(data);
  root.descendants().forEach(d => {
    if (d.children) {
      d._children = d.children;
      d.children = null;
    }
  });

  if (root._children) {
    root.children = root._children;
    root._children = null;
  }

  if (expandedPaths.size > 0) {
    root.descendants().forEach(d => {
      if (expandedPaths.has(getNodePath(d)) && d._children) {
        d.children = d._children;
        d._children = null;
      }
    });
  }

  root.x0 = 0;
  root.y0 = 0;
  update(root);
}

// ==================== CONFIG BUTTON (NO CLEAR BUTTON) ====================
function addConfigButton(dimensions) {
  const toolbar = document.getElementById('toolbar');
  if (!toolbar) return;

  const existing = document.getElementById('config-btn');
  if (existing) existing.remove();

  const btn = document.createElement('button');
  btn.id = 'config-btn';
  btn.innerHTML = '⚙️ Hierarchy Order';
  btn.style.cssText = `
    padding: 6px 14px;
    border: 1px solid #ccc;
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
    font-size: 12px;
    color: #555;
    font-family: Arial;
    box-shadow: 0 1px 4px rgba(0,0,0,0.1);
  `;
  btn.addEventListener('mouseenter', () => {
    btn.style.background = '#f0f4ff';
    btn.style.borderColor = '#007bff';
  });
  btn.addEventListener('mouseleave', () => {
    btn.style.background = '#fff';
    btn.style.borderColor = '#ccc';
  });
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    showConfigDialog(dimensions);
  });
  toolbar.appendChild(btn);
}

// ==================== CONFIG DIALOG ====================
function showConfigDialog(dimensions) {
  const existing = document.getElementById('config-overlay');
  if (existing) existing.remove();

  const overlay = document.createElement('div');
  overlay.id = 'config-overlay';
  overlay.style.cssText = `
    position:fixed;top:0;left:0;width:100%;height:100%;
    background:rgba(0,0,0,0.5);z-index:10000;
    display:flex;align-items:center;justify-content:center;font-family:Arial;
  `;

  const dialog = document.createElement('div');
  dialog.style.cssText = `
    background:#fff;border-radius:12px;padding:24px 28px;
    min-width:320px;max-width:420px;box-shadow:0 8px 32px rgba(0,0,0,0.3);
  `;
  dialog.innerHTML = `
    <h3 style="margin:0 0 4px;font-size:16px;color:#333;">📊 Set Hierarchy Order</h3>
    <p style="margin:0 0 16px;font-size:12px;color:#888;">
      Drag or use arrows. Top = first drill-down level.
    </p>
    <div id="dim-list"></div>
    <div style="margin-top:20px;display:flex;gap:10px;justify-content:flex-end;">
      <button id="config-cancel" style="padding:8px 18px;border:1px solid #ccc;
        border-radius:6px;background:#fff;cursor:pointer;font-size:13px;">Cancel</button>
      <button id="config-apply" style="padding:8px 18px;border:none;border-radius:6px;
        background:#007bff;color:#fff;cursor:pointer;font-size:13px;font-weight:bold;">Apply</button>
    </div>`;
  overlay.appendChild(dialog);
  document.body.appendChild(overlay);

  let currentOrder = [...dimensions];
  let dragIdx = null;

  function renderList() {
    const listEl = document.getElementById('dim-list');
    listEl.innerHTML = '';

    currentOrder.forEach((dim, idx) => {
      const item = document.createElement('div');
      item.draggable = true;
      item.style.cssText = `
        padding:10px 14px;margin:4px 0;background:#f8f9fa;border:2px solid #e9ecef;
        border-radius:8px;cursor:grab;display:flex;align-items:center;gap:10px;font-size:13px;
      `;
      item.innerHTML = `
        <span style="color:#aaa;font-size:16px;">☰</span>
        <span style="background:${palette[idx % palette.length]};color:#fff;border-radius:50%;
          width:22px;height:22px;display:flex;align-items:center;justify-content:center;
          font-size:11px;font-weight:bold;">${idx + 1}</span>
        <span style="font-weight:500;">${dim}</span>`;

      item.addEventListener('dragstart', e => {
        dragIdx = idx;
        item.style.opacity = '0.4';
        e.dataTransfer.effectAllowed = 'move';
      });
      item.addEventListener('dragend', () => {
        item.style.opacity = '1';
        dragIdx = null;
        renderList();
      });
      item.addEventListener('dragover', e => {
        e.preventDefault();
        item.style.borderColor = '#007bff';
      });
      item.addEventListener('dragleave', () => {
        item.style.borderColor = '#e9ecef';
      });
      item.addEventListener('drop', e => {
        e.preventDefault();
        if (dragIdx !== null && dragIdx !== idx) {
          const moved = currentOrder.splice(dragIdx, 1)[0];
          currentOrder.splice(idx, 0, moved);
        }
        renderList();
      });

      const btns = document.createElement('span');
      btns.style.cssText = 'margin-left:auto;display:flex;gap:2px;';

      if (idx > 0) {
        const u = document.createElement('button');
        u.textContent = '▲';
        u.style.cssText = 'border:none;background:#e9ecef;border-radius:4px;cursor:pointer;padding:2px 6px;font-size:10px;';
        u.onclick = e => {
          e.stopPropagation();
          [currentOrder[idx - 1], currentOrder[idx]] = [currentOrder[idx], currentOrder[idx - 1]];
          renderList();
        };
        btns.appendChild(u);
      }

      if (idx < currentOrder.length - 1) {
        const dn = document.createElement('button');
        dn.textContent = '▼';
        dn.style.cssText = 'border:none;background:#e9ecef;border-radius:4px;cursor:pointer;padding:2px 6px;font-size:10px;';
        dn.onclick = e => {
          e.stopPropagation();
          [currentOrder[idx], currentOrder[idx + 1]] = [currentOrder[idx + 1], currentOrder[idx]];
          renderList();
        };
        btns.appendChild(dn);
      }

      item.appendChild(btns);
      listEl.appendChild(item);
    });
  }

  renderList();

  document.getElementById('config-cancel').onclick = () => overlay.remove();
  document.getElementById('config-apply').onclick = () => {
    overlay.remove();
    saveSettings(currentOrder).then(() => {
      saveExpandedState();
      loadData();
    });
  };
  overlay.onclick = e => {
    if (e.target === overlay) overlay.remove();
  };
}

// ==================== D3 UPDATE ====================
function update(source) {
  d3.tree().nodeSize([dx, dy])(root);
  root.each(d => {
    d.y = d.depth * dy;
  });

  function getSubtreeLeaves(node) {
    if (!node.children || node.children.length === 0) return 1;
    return node.children.reduce((sum, child) => sum + getSubtreeLeaves(child), 0);
  }

  function layoutNodes(node) {
    if (!node.children || node.children.length === 0) return;

    const totalLeaves = getSubtreeLeaves(node);
    const totalHeight = (totalLeaves - 1) * dx;
    let currentX = node.x - totalHeight / 2;

    node.children.forEach(child => {
      const childLeaves = getSubtreeLeaves(child);
      child.x = currentX + ((childLeaves - 1) * dx) / 2;
      layoutNodes(child);
      currentX += childLeaves * dx;
    });
  }

  root.x = 0;
  layoutNodes(root);

  const nodes = root.descendants();
  const links = root.links();

  const minX = d3.min(nodes, d => d.x) || 0;
  const maxX = d3.max(nodes, d => d.x) || 0;
  const maxY = d3.max(nodes, d => d.y) || 0;

  const finalHeight = Math.max(height, (maxX - minX) + margin.top + margin.bottom + dx + 40);
  const finalWidth = Math.max(width, maxY + barWidth + 500 + margin.left + margin.right);

  svg
    .attr('width', finalWidth)
    .attr('height', finalHeight);

  const svgWrapEl = document.getElementById('svg-wrap');
  if (svgWrapEl) {
    svgWrapEl.style.width = finalWidth + 'px';
    svgWrapEl.style.height = finalHeight + 'px';
  }

  const hasTarget = !!(currentTarget && root.data.target > 0);

  const nodeSelection = g.selectAll('g.node').data(nodes, d => d.id || (d.id = ++i));

  const nodeEnter = nodeSelection.enter().append('g')
    .attr('class', 'node')
    .attr('transform', () => `translate(${source.y0 || 0},${source.x0 || 0})`)
    .style('opacity', 0);

  nodeEnter.append('rect').attr('class', 'bar-bg')
    .attr('rx', 5).attr('ry', 5)
    .attr('width', barWidth).attr('height', barHeight)
    .attr('y', -barHeight / 2)
    .style('fill', '#efefef')
    .style('cursor', 'pointer')
    .on('click', highlightAndFilter);

  nodeEnter.append('rect').attr('class', 'bar-fill')
    .attr('rx', 5).attr('ry', 5)
    .attr('height', barHeight).attr('y', -barHeight / 2)
    .attr('x', 0).attr('width', 0)
    .style('fill', MEASURE_COLORS.primary)
    .style('cursor', 'pointer')
    .on('click', highlightAndFilter);

  nodeEnter.append('rect').attr('class', 'target-line')
    .attr('width', 3).attr('height', barHeight)
    .attr('y', -barHeight / 2)
    .attr('rx', 1).attr('ry', 1).attr('x', 0)
    .style('fill', MEASURE_COLORS.target)
    .style('opacity', 0)
    .style('cursor', 'pointer')
    .on('click', highlightAndFilter);

  nodeEnter.append('text').attr('class', 'node-label')
    .attr('dy', `${barHeight / 2 + 16}px`).attr('x', 0)
    .style('font-weight', 'bold').style('fill', '#333').style('font-size', '12px')
    .style('cursor', 'pointer')
    .on('click', highlightAndFilter);

  nodeEnter.append('text').attr('class', 'node-value')
    .attr('dy', `${barHeight / 2 + 30}px`).attr('x', 0)
    .style('font-size', '10px').style('fill', '#999')
    .style('cursor', 'pointer')
    .on('click', highlightAndFilter);

  nodeEnter.append('text').attr('class', 'achievement-badge')
    .attr('x', barWidth + 8).attr('dy', '0.35em')
    .style('font-size', '10px').style('font-weight', 'bold');

  const toggleX = hasTarget ? barWidth + 65 : barWidth + 15;

  const icon = nodeEnter.append('g').attr('class', 'toggle-icon')
    .attr('transform', `translate(${toggleX},0)`)
    .style('cursor', 'pointer')
    .on('click', function (event, d) {
      event.stopPropagation();
      if (d.data.isLastLevel) return;

      if (d.children) {
        d._children = d.children;
        d.children = null;
      } else if (d._children) {
        d.children = d._children;
        d._children = null;
      }

      saveExpandedState();
      update(d);
    });

  icon.append('circle').attr('r', 12)
    .style('fill', '#fff').style('stroke', '#007bff').style('stroke-width', '2px');

  icon.append('text').attr('dy', '0.35em').attr('text-anchor', 'middle')
    .style('font-size', '16px').style('font-weight', 'bold').style('fill', '#007bff');

  const nodeUpdate = nodeEnter.merge(nodeSelection);

  nodeUpdate.transition().duration(duration)
    .attr('transform', d => `translate(${d.y},${d.x})`)
    .style('opacity', 1);

  nodeUpdate.select('.bar-fill')
    .transition().duration(duration)
    .attr('width', d => {
      let maxVal;
      if (d.parent) {
        const siblings = d.parent.children || d.parent._children || [];
        maxVal = d3.max(siblings, s => Math.max(s.data.value, s.data.target || 0)) || 1;
      } else {
        maxVal = Math.max(d.data.value, d.data.target || 0) || 1;
      }
      const w = maxVal > 0 ? (d.data.value / maxVal) * barWidth : 0;
      return Math.max(0, Math.min(w, barWidth));
    })
    .style('fill', MEASURE_COLORS.primary);

  nodeUpdate.select('.target-line')
    .transition().duration(duration)
    .attr('x', d => {
      if (!hasTarget || !d.data.target) return 0;

      let maxVal;
      if (d.parent) {
        const siblings = d.parent.children || d.parent._children || [];
        maxVal = d3.max(siblings, s => Math.max(s.data.value, s.data.target || 0)) || 1;
      } else {
        maxVal = Math.max(d.data.value, d.data.target) || 1;
      }

      const ratio = d.data.target / maxVal;
      const x = Math.max(1, Math.min(ratio * barWidth, barWidth - 3));
      return x - 1.5;
    })
    .style('opacity', d => (hasTarget && d.data.target > 0) ? 1 : 0);

  nodeUpdate.select('.node-label')
    .text(d => {
      const pct = root.data.value > 0
        ? ((d.data.value / root.data.value) * 100).toFixed(0)
        : 0;
      return `${d.data.name} (${pct}%)`;
    });

  nodeUpdate.select('.node-value')
    .text(d => {
      const formatVal = v => {
        if (v >= 1e6) return d3.format('.1f')(v / 1e6) + 'M';
        if (v >= 1e3) return d3.format('.1f')(v / 1e3) + 'K';
        return d3.format(',.0f')(v);
      };

      let txt = `${cleanName(currentMeasure)}: ${formatVal(d.data.value)}`;
      if (hasTarget && d.data.target > 0) {
        txt += ` | ${cleanName(currentTarget)}: ${formatVal(d.data.target)}`;
      }
      return txt;
    });

  nodeUpdate.select('.achievement-badge')
    .text(d => {
      if (!hasTarget || !d.data.target || d.data.target === 0) return '';
      const pct = ((d.data.value / d.data.target) * 100).toFixed(0);
      return d.data.value >= d.data.target ? `✓ ${pct}%` : `${pct}%`;
    })
    .style('fill', d => {
      if (!hasTarget || !d.data.target || d.data.target === 0) return '#999';
      const r = d.data.value / d.data.target;
      if (r >= 1) return '#28a745';
      if (r >= 0.8) return '#e6a817';
      return '#dc3545';
    });

  nodeUpdate.select('.toggle-icon')
    .attr('transform', `translate(${toggleX},0)`)
    .style('opacity', d => (!d.data.isLastLevel && (d._children || d.children)) ? 1 : 0)
    .style('pointer-events', d => (!d.data.isLastLevel && (d._children || d.children)) ? 'all' : 'none');

  nodeUpdate.select('.toggle-icon text')
    .text(d => d.data.isLastLevel ? '' : d.children ? '−' : '+');

  nodeSelection.exit().transition().duration(duration)
    .attr('transform', `translate(${source.y || 0},${source.x || 0})`)
    .style('opacity', 0)
    .remove();

  const diagonal = ({ source: s, target: t }) => {
    const sx = s.y + toggleX + 12;
    const ex = t.y - 5;
    const mid = (sx + ex) / 2;
    return `M${sx},${s.x} C${mid},${s.x} ${mid},${t.x} ${ex},${t.x}`;
  };

  const link = g.selectAll('path.link').data(links, d => d.target.id);

  link.enter().insert('path', 'g')
    .attr('class', 'link')
    .attr('d', () => {
      const o = { x: source.x0 || 0, y: source.y0 || 0 };
      return diagonal({ source: o, target: o });
    })
    .style('fill', 'none')
    .style('stroke', '#ccc')
    .style('stroke-width', '1.5px')
    .merge(link)
    .transition().duration(duration)
    .attr('d', diagonal);

  link.exit().transition().duration(duration)
    .attr('d', () => {
      const o = { x: source.x || 0, y: source.y || 0 };
      return diagonal({ source: o, target: o });
    })
    .remove();

  nodes.forEach(d => {
    d.x0 = d.x;
    d.y0 = d.y;
  });

  const vOff = margin.top + (minX < 0 ? Math.abs(minX) : 0);
  g.attr('transform', `translate(${margin.left},${vOff})`);
}

// ==================== ACTION FILTER (MARK SELECTION) ====================
async function applyFilter(d) {
  try {
    if (!d.data.type || d.depth === 0) return;

    const ws = getWorksheet();
    if (!ws) {
      console.error('❌ No worksheet available for selection');
      return;
    }

    const filterPairs = getAncestorFilterPairs(d);

    console.log('🔍 Applying Action Filter (Mark Selection):');
    filterPairs.forEach(p => console.log(`   ${p.fieldName} = "${p.value}"`));

    // METHOD 1: Try selectMarksByValueAsync (proper action filter)
    try {
      const selectionCriteria = filterPairs.map(pair => ({
        fieldName: pair.fieldName,
        value: [pair.value]
      }));

      await ws.selectMarksByValueAsync(
        selectionCriteria,
        tableau.SelectionUpdateType.Replace
      );

      console.log('✅ Marks selected via selectMarksByValueAsync (Action Filter)');
      return;
    } catch (e1) {
      console.warn('⚠️ selectMarksByValueAsync failed:', e1.message);
    }

    // METHOD 2: Fallback - applyFilterAsync
    console.log('🔄 Fallback: Using applyFilterAsync...');

    const worksheets = getAllWorksheetsForFilter();

    for (const targetWs of worksheets) {
      for (const pair of filterPairs) {
        try {
          await targetWs.applyFilterAsync(
            pair.fieldName,
            [pair.value],
            tableau.FilterUpdateType.Replace
          );
          console.log(`✅ Filter applied: "${pair.fieldName}" = "${pair.value}" on "${targetWs.name}"`);
        } catch (wsErr) {
          console.warn(`⚠️ Skipped "${targetWs.name}" for "${pair.fieldName}": ${wsErr.message}`);
        }
      }
    }

    console.log('✅ Fallback filters applied');

  } catch (err) {
    console.error('❌ Action Filter error:', err);
  }
}

// ==================== CLEAR SELECTION (AUTO ON BLANK CLICK) ====================
async function clearAllSelections() {
  try {
    const ws = getWorksheet();

    // METHOD 1: Clear mark selection
    if (ws) {
      try {
        await ws.clearSelectedMarksAsync();
        console.log('✅ Mark selection cleared');
      } catch (e) {
        console.warn('⚠️ clearSelectedMarksAsync not available:', e.message);
      }
    }

    // METHOD 2: Clear applied filters
    const worksheets = getAllWorksheetsForFilter();

    for (const targetWs of worksheets) {
      try {
        const filters = await targetWs.getFiltersAsync();
        for (const f of filters) {
          try {
            await targetWs.clearFilterAsync(f.fieldName);
          } catch (clearErr) {
            // Some filters can't be cleared
          }
        }
        console.log(`✅ Filters cleared on: "${targetWs.name}"`);
      } catch (wsErr) {
        console.warn(`⚠️ Could not clear "${targetWs.name}": ${wsErr.message}`);
      }
    }

    console.log('✅ All selections cleared (blank area click)');
  } catch (err) {
    console.error('❌ Clear selection error:', err);
  }
}