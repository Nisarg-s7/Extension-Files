/* global d3, tableau */

// ==================== CONFIGURATION ====================
const CONFIG = {
  margin: 40,
  colors: d3.scaleOrdinal(d3.schemeCategory10),
  arcOpacity: 0.9,
  arcOpacityHover: 1,
  transitionDuration: 750,
  minAngle: 0.005
};

let svg, g, partition, arc, tooltip;
let width, height, radius;
let currentData = null;
let originalData = null; // ✅ Fix #2 - Store original root
let currentRoot = null;
let detectedDimensions = [];
let selectedDimensionsOrder = [];
let currentWs = null; // ✅ Fix #10 - Track worksheet for cleanup

// ==================== DOM SETUP ====================
function setupDOM() {
  // Main container
  let container = document.getElementById("sunburst-container");
  if (!container) {
    container = document.createElement("div");
    container.id = "sunburst-container";
    container.style.cssText = `
      width: 100%;
      height: 100%;
      position: absolute;
      top: 0;
      left: 0;
      overflow: hidden;
      background: #fff;
    `;
    document.body.appendChild(container);
  }

  // ✅ Fix #1 - Hierarchy button OUTSIDE the if block
  let hierarchyBtn = document.getElementById("hierarchy-btn");
  if (!hierarchyBtn) {
    hierarchyBtn = document.createElement("button");
    hierarchyBtn.id = "hierarchy-btn";
    hierarchyBtn.textContent = "Set Hierarchy Order";
    hierarchyBtn.style.cssText = `
      position: absolute;
      top: 15px;
      right: 20px;
      padding: 8px 14px;
      background: #007bff;
      color: #fff;
      border: none;
      border-radius: 6px;
      cursor: pointer;
      font-size: 13px;
      z-index: 200;
    `;
    container.appendChild(hierarchyBtn);
    hierarchyBtn.onclick = openHierarchyModal;
  }

  // Chart host
  let chartHost = document.getElementById("sunburst-chart-host");
  if (!chartHost) {
    chartHost = document.createElement("div");
    chartHost.id = "sunburst-chart-host";
    chartHost.style.cssText = `
      width: 100%;
      height: 100%;
      position: absolute;
      inset: 0;
    `;
    container.appendChild(chartHost);
  }

  // Tooltip
  let tip = document.getElementById("sunburst-tooltip");
  if (!tip) {
    tip = document.createElement("div");
    tip.id = "sunburst-tooltip";
    tip.style.cssText = `
      position: fixed;
      padding: 12px 16px;
      background: rgba(0, 0, 0, 0.9);
      color: #fff;
      border-radius: 6px;
      font-size: 13px;
      font-family: system-ui, -apple-system, sans-serif;
      pointer-events: none;
      opacity: 0;
      transition: opacity 0.2s;
      z-index: 10000;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.3);
      line-height: 1.6;
      max-width: 320px;
    `;
    document.body.appendChild(tip);
  }
  tooltip = d3.select("#sunburst-tooltip");

  // Center label container
  let centerLabel = document.getElementById("center-label");
  if (!centerLabel) {
    centerLabel = document.createElement("div");
    centerLabel.id = "center-label";
    centerLabel.style.cssText = `
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      text-align: center;
      pointer-events: none;
      font-family: system-ui, -apple-system, sans-serif;
      z-index: 20;
    `;
    centerLabel.innerHTML = `
      <div id="center-name" style="font-size:18px; font-weight:600; color:#333; margin-bottom:4px;"></div>
      <div id="center-value" style="font-size:24px; font-weight:700; color:#007bff;"></div>
      <div id="center-percentage" style="font-size:14px; color:#666; margin-top:4px;"></div>
    `;
    container.appendChild(centerLabel);
  }

  // Error overlay
  let err = document.getElementById("sunburst-error");
  if (!err) {
    err = document.createElement("div");
    err.id = "sunburst-error";
    err.style.cssText = `
      position: absolute;
      inset: 0;
      display: none;
      align-items: center;
      justify-content: center;
      z-index: 30;
      background: rgba(255,255,255,0.75);
      backdrop-filter: blur(1px);
    `;
    err.innerHTML = `
      <div style="
        font-family: system-ui, -apple-system, sans-serif;
        text-align: center;
        padding: 40px 48px;
        background: #fff8f8;
        border: 2px solid #f5c6c6;
        border-radius: 12px;
        max-width: 560px;
        box-shadow: 0 4px 16px rgba(0,0,0,0.1);
      ">
        <div style="font-size:48px; margin-bottom:16px;">☀️</div>
        <h3 style="margin:0 0 12px; color:#c0392b; font-size:20px;">Sunburst Chart Error</h3>
        <p id="sunburst-error-msg" style="margin:0; font-size:14px; color:#555; line-height:1.6;"></p>
        <p style="margin:16px 0 0; font-size:12px; color:#999;">
          Required: At least 1 dimension and 1 measure (or row count).<br>
          Press F12 for console details.
        </p>
      </div>
    `;
    container.appendChild(err);
  }
}

// ==================== ERROR HELPERS ====================
function showError(message) {
  const err = document.getElementById("sunburst-error");
  const msg = document.getElementById("sunburst-error-msg");
  if (msg) msg.textContent = String(message || "Unknown error");
  if (err) err.style.display = "flex";
}

function hideError() {
  const err = document.getElementById("sunburst-error");
  if (err) err.style.display = "none";
}

// ==================== WORKSHEET PICKER ====================
function getTargetWorksheet() {
  try {
    const dash = tableau.extensions.dashboardContent?.dashboard;
    if (dash?.worksheets?.length) return dash.worksheets[0];
  } catch (e) { /* ignore */ }
  return tableau.extensions.worksheetContent?.worksheet || null;
}

// ==================== INITIALIZATION ====================
window.onload = function () {
  setupDOM();

  if (typeof tableau === "undefined") {
    showError("Tableau Extensions API not found. Please load this extension in Tableau.");
    return;
  }

  tableau.extensions.initializeAsync().then(() => {
    console.log("✅ Tableau Extension initialized");
    initChart();
    loadData();

    window.addEventListener("resize", () => {
      updateSize();
      if (currentData) renderSunburst(currentData);
    });

  }).catch(err => {
    console.error("❌ Initialization failed:", err);
    showError("Failed to initialize: " + err.message);
  });
};

// ==================== CHART INITIALIZATION ====================
function initChart() {
  const host = document.getElementById("sunburst-chart-host");
  host.innerHTML = "";

  width = host.clientWidth || 800;
  height = host.clientHeight || 600;
  radius = Math.min(width, height) / 2 - CONFIG.margin;

  svg = d3.select("#sunburst-chart-host")
    .append("svg")
    .attr("width", width)
    .attr("height", height)
    .style("position", "absolute")
    .style("top", "0")
    .style("left", "0");

  g = svg.append("g")
    .attr("transform", `translate(${width / 2},${height / 2})`);

  partition = d3.partition().size([2 * Math.PI, radius]);

  arc = d3.arc()
    .startAngle(d => d.x0)
    .endAngle(d => d.x1)
    .innerRadius(d => d.y0)
    .outerRadius(d => d.y1);

  // ✅ Fix #2 - Back click uses originalData, not currentData
  svg.on("click", () => {
    if (originalData) {
      currentData = originalData;
      renderSunburst(originalData);
    }
  });

  console.log(`✅ Chart initialized (${width}x${height}, radius: ${radius})`);
}

function updateSize() {
  const host = document.getElementById("sunburst-chart-host");
  width = host.clientWidth || 800;
  height = host.clientHeight || 600;
  radius = Math.min(width, height) / 2 - CONFIG.margin;

  if (svg) svg.attr("width", width).attr("height", height);
  if (g) g.attr("transform", `translate(${width / 2},${height / 2})`);
  if (partition) partition.size([2 * Math.PI, radius]);
}

// ==================== DATA LOADING ====================
async function loadData() {
  try {
    hideError();

    const ws = getTargetWorksheet();
    if (!ws) {
      showError("No worksheet found. Add this extension to a dashboard with at least one worksheet.");
      return;
    }

    // ✅ Fix #10 - Remove old listeners before adding new ones
    if (currentWs) {
      try {
        currentWs.removeEventListener(
          tableau.TableauEventType.SummaryDataChanged, loadData
        );
        currentWs.removeEventListener(
          tableau.TableauEventType.FilterChanged, loadData
        );
      } catch (e) { /* ignore */ }
    }

    currentWs = ws;
    currentWs.addEventListener(tableau.TableauEventType.SummaryDataChanged, loadData);
    currentWs.addEventListener(tableau.TableauEventType.FilterChanged, loadData);

    console.log("📥 Loading data from:", ws.name);

    const dataTable = await ws.getSummaryDataAsync({
      ignoreSelection: false,
      maxRows: 100000
    });

    if (!dataTable || !dataTable.columns || dataTable.columns.length === 0) {
      showError(`Worksheet "${ws.name}" returned no columns.`);
      return;
    }

    if (!dataTable.data || dataTable.data.length === 0) {
      showError(`Worksheet "${ws.name}" returned 0 rows. Check filters.`);
      return;
    }

    // ✅ Fix #5 - Improved field detection
    const { dimensions, measures } = detectFields(dataTable.columns);

    if (!dimensions.length) {
      showError("No dimensions found. Add at least 1 dimension to the worksheet.");
      return;
    }

    detectedDimensions = dimensions.slice();

    // ✅ Fix #7 - Only reset order if dimensions changed
    const orderIsValid = selectedDimensionsOrder.length > 0 &&
      selectedDimensionsOrder.every(d => dimensions.includes(d)) &&
      dimensions.every(d => selectedDimensionsOrder.includes(d));

    if (!orderIsValid) {
      selectedDimensionsOrder = dimensions.slice();
    }

    // ✅ Fix #6 - Build measure selector UI
    buildMeasureSelector(measures);

    const activeMeasure = getActiveMeasure(measures);

    const rawData = parseTableauData(dataTable);
    const hierarchyData = buildHierarchy(rawData, selectedDimensionsOrder, activeMeasure);

    if (!hierarchyData.children || hierarchyData.children.length === 0) {
      showError("No valid hierarchy produced. Check dimension values and measure > 0.");
      return;
    }

    // ✅ Fix #2 - Save both original and current
    originalData = hierarchyData;
    currentData = hierarchyData;

    renderSunburst(hierarchyData);

  } catch (err) {
    console.error("❌ Error loading data:", err);
    showError(err?.message || String(err));
  }
}

// ==================== MEASURE SELECTOR ====================
// ✅ Fix #6 - Measure selector dropdown

let activeMeasureName = null;

function buildMeasureSelector(measures) {
  const container = document.getElementById("sunburst-container");
  if (!container || measures.length <= 1) {
    // If only one measure, just use it directly
    if (measures.length === 1) activeMeasureName = measures[0].fieldName;
    return;
  }

  let selector = document.getElementById("measure-selector");
  if (!selector) {
    selector = document.createElement("select");
    selector.id = "measure-selector";
    selector.style.cssText = `
      position: absolute;
      top: 15px;
      left: 20px;
      padding: 6px 10px;
      border: 1px solid #ccc;
      border-radius: 6px;
      font-size: 13px;
      background: #fff;
      cursor: pointer;
      z-index: 200;
    `;
    container.appendChild(selector);

    selector.onchange = function () {
      activeMeasureName = this.value;
      loadData();
    };
  }

  // Rebuild options if measures changed
  const existing = Array.from(selector.options).map(o => o.value);
  const incoming = measures.map(m => m.fieldName);
  const changed = existing.join(",") !== incoming.join(",");

  if (changed) {
    selector.innerHTML = "";
    measures.forEach(m => {
      const opt = document.createElement("option");
      opt.value = m.fieldName;
      opt.textContent = m.fieldName;
      selector.appendChild(opt);
    });

    // Keep previous selection if still valid
    if (activeMeasureName && incoming.includes(activeMeasureName)) {
      selector.value = activeMeasureName;
    } else {
      activeMeasureName = measures[0].fieldName;
      selector.value = activeMeasureName;
    }
  }
}

function getActiveMeasure(measures) {
  if (!measures || measures.length === 0) return null;
  if (activeMeasureName && measures.find(m => m.fieldName === activeMeasureName)) {
    return activeMeasureName;
  }
  activeMeasureName = measures[0].fieldName;
  return activeMeasureName;
}

// ==================== FIELD DETECTION ====================
// ✅ Fix #5 - Role-based detection only, no dtype ambiguity

function detectFields(columns) {
  const SKIP = [/^measure names$/i, /^measure values$/i, /^number of records$/i];

  const dimensions = columns.filter(col => {
    const name = col.fieldName || "";
    if (SKIP.some(rx => rx.test(name))) return false;
    const role = (col.role || "").toLowerCase();
    return role === "dimension";
  });

  const measures = columns.filter(col => {
    const name = col.fieldName || "";
    if (SKIP.some(rx => rx.test(name))) return false;
    const role = (col.role || "").toLowerCase();
    return role === "measure";
  });

  // Fallback if role info not available
  const fallbackDimensions = dimensions.length ? dimensions : columns.filter(col => {
    const name = col.fieldName || "";
    if (SKIP.some(rx => rx.test(name))) return false;
    const dtype = (col.dataType || "").toLowerCase();
    return ["string", "bool", "date", "datetime"].includes(dtype);
  });

  const fallbackMeasures = measures.length ? measures : columns.filter(col => {
    const name = col.fieldName || "";
    if (SKIP.some(rx => rx.test(name))) return false;
    const dtype = (col.dataType || "").toLowerCase();
    return ["int", "integer", "float", "double"].includes(dtype);
  });

  return {
    dimensions: fallbackDimensions.map(col => col.fieldName),
    measures: fallbackMeasures
  };
}

// ==================== DATA PARSING ====================
function parseTableauData(dataTable) {
  return dataTable.data.map(row => {
    const parsed = {};
    dataTable.columns.forEach((col, idx) => {
      const fieldName = col.fieldName;
      const cell = row[idx];
      let value;
      if (cell !== null && typeof cell === "object") {
        value = ("nativeValue" in cell && cell.nativeValue != null)
          ? cell.nativeValue
          : cell.value;
      } else {
        value = cell;
      }
      parsed[fieldName] = value;
    });
    return parsed;
  });
}

// ==================== BUILD HIERARCHY ====================
function buildHierarchy(data, dimensions, measure) {
  const root = { name: "Total", children: [] };
  const map = new Map([["root", root]]);

  data.forEach(row => {
    const rowValue = measure ? (parseFloat(row[measure]) || 0) : 1;
    if (!(rowValue > 0)) return;

    let parentPath = "root";

    dimensions.forEach((dim, index) => {
      const raw = row[dim];
      const label = (raw == null || String(raw).trim() === "")
        ? `(Unknown ${dim})`
        : String(raw).trim();

      const path = `${parentPath}|L${index}:${dim}:${label}`;

      if (!map.has(path)) {
        const node = {
          name: label,
          dimension: dim,
          level: index + 1,
          children: [],
          _leafValue: 0
        };
        map.set(path, node);
        map.get(parentPath).children.push(node);
      }

      if (index === dimensions.length - 1) {
        map.get(path)._leafValue += rowValue;
      }

      parentPath = path;
    });
  });

  return root;
}

// ==================== SUNBURST RENDERING ====================
function renderSunburst(data) {
  if (!svg || !g) initChart();

  g.selectAll("*").remove();

  const root = d3.hierarchy(data)
    .sum(d => (!d.children || d.children.length === 0) ? (d._leafValue || 0) : 0)
    .sort((a, b) => b.value - a.value);

  currentRoot = root;
  partition(root);

  const nodes = root.descendants().filter(d => (d.x1 - d.x0) > CONFIG.minAngle);
  const color = d3.scaleOrdinal(d3.schemeCategory10);

  // Arcs
  g.selectAll("path")
    .data(nodes)
    .enter()
    .append("path")
    .attr("d", arc)
    .style("fill", d => {
      if (d.depth === 0) return "transparent";
      let top = d;
      while (top.depth > 1) top = top.parent;
      return color(top.data.name);
    })
    .style("stroke", "#fff")
    .style("stroke-width", "2px")
    .style("opacity", CONFIG.arcOpacity)
    .style("cursor", d => d.depth === 0 ? "default" : "pointer")
    .on("mouseover", function (event, d) {
      d3.select(this).style("opacity", CONFIG.arcOpacityHover);
      g.selectAll("path")
        .filter(node => !isAncestor(node, d) && !isDescendant(node, d))
        .style("opacity", 0.25);

      updateCenterLabel(d);

      const percentage = d.parent
        ? ((d.value / d.parent.value) * 100).toFixed(1)
        : "100.0";

      showTooltip(event, `
        <strong>${escapeHtml(d.data.name)}</strong><br>
        ${d.data.dimension ? `<em>${escapeHtml(d.data.dimension)}</em><br>` : ""}
        Value: <strong>${d3.format(",.0f")(d.value)}</strong><br>
        Percentage: <strong>${percentage}%</strong><br>
        Level: ${d.depth}
      `);
    })
    .on("mousemove", moveTooltip)
    .on("mouseout", function () {
      g.selectAll("path").style("opacity", CONFIG.arcOpacity);
      updateCenterLabel(currentRoot);
      hideTooltip();
    })
    .on("click", function (event, d) {
      if (d.depth === 0) return;
      event.stopPropagation();
      zoomTo(d);
    });

  // ✅ Fix #8 & #9 - Better label threshold + valid font-family
  const labelNodes = nodes.filter(d => {
    if (d.depth === 0) return false;
    const arcLength = (d.x1 - d.x0) * (d.y0 + d.y1) / 2;
    return arcLength > 30;
  });

  const labels = g.selectAll("text.arc-label")
    .data(labelNodes)
    .enter()
    .append("text")
    .attr("class", "arc-label")
    .attr("transform", d => {
      const a = (d.x0 + d.x1) / 2;
      const r = (d.y0 + d.y1) / 2;
      const x = Math.cos(a - Math.PI / 2) * r;
      const y = Math.sin(a - Math.PI / 2) * r;
      let rot = (a * 180 / Math.PI) - 90;
      if (a > Math.PI) rot += 180;
      return `translate(${x},${y}) rotate(${rot})`;
    })
    .attr("text-anchor", "middle")
    .attr("dominant-baseline", "central")
    .style("pointer-events", "none");

  // Line 1: Name
  labels.append("tspan")
    .attr("x", 0)
    .attr("dy", "-0.35em")
    .style("font-family", "system-ui, -apple-system, sans-serif") // ✅ Fix #9
    .style("font-size", "12px")
    .style("font-weight", "700")
    .style("fill", "#1b1616")
    .text(d => truncate(d.data.name, 18));

  // Line 2: Value
  labels.append("tspan")
    .attr("x", 0)
    .attr("dy", "1.2em")
    .style("font-family", "system-ui, -apple-system, sans-serif") // ✅ Fix #9
    .style("font-size", "11px")
    .style("font-weight", "600")
    .style("fill", "#1b1616")
    .text(d => d3.format(",.0f")(d.value));

  updateCenterLabel(root);
}

// ==================== ZOOM ====================
// ✅ Fix #2 - zoomTo only changes currentData, originalData stays intact
function zoomTo(node) {
  currentData = node.data;
  renderSunburst(node.data);
}

// ==================== CENTER LABEL ====================
function ensureCenterLabel() {
  const container = document.getElementById("sunburst-container");
  if (!container) return;

  if (!document.getElementById("center-label")) {
    const centerLabel = document.createElement("div");
    centerLabel.id = "center-label";
    centerLabel.style.cssText = `
      position: absolute;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      text-align: center;
      pointer-events: none;
      font-family: system-ui, -apple-system, sans-serif;
      z-index: 20;
    `;
    centerLabel.innerHTML = `
      <div id="center-name" style="font-size:18px; font-weight:600; color:#333; margin-bottom:4px;"></div>
      <div id="center-value" style="font-size:24px; font-weight:700; color:#007bff;"></div>
      <div id="center-percentage" style="font-size:14px; color:#666; margin-top:4px;"></div>
    `;
    container.appendChild(centerLabel);
  }
}

function updateCenterLabel(d) {
  ensureCenterLabel();

  const nameEl = document.getElementById("center-name");
  const valueEl = document.getElementById("center-value");
  const pctEl = document.getElementById("center-percentage");

  if (!nameEl || !valueEl || !pctEl) return;

  nameEl.textContent = d?.data?.name ?? "Total";
  valueEl.textContent = d3.format(",.0f")(d?.value || 0);

  if (d?.parent) {
    const percentage = d.parent.value
      ? ((d.value / d.parent.value) * 100).toFixed(1)
      : "0.0";
    pctEl.textContent = `${percentage}% of ${d.parent.data.name}`;
  } else {
    pctEl.textContent = "Total";
  }
}

// ==================== HIERARCHY MODAL ====================
// ✅ Fix #3 & #4 - Modal uses selectedDimensionsOrder directly

function openHierarchyModal() {
  closeModal();

  const modal = document.createElement("div");
  modal.id = "hierarchy-modal";
  modal.style.cssText = `
    position: fixed;
    inset: 0;
    background: rgba(0,0,0,0.4);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 500;
  `;
  document.body.appendChild(modal);

  const box = document.createElement("div");
  box.style.cssText = `
    background: #fff;
    padding: 24px;
    width: 420px;
    border-radius: 12px;
    box-shadow: 0 5px 20px rgba(0,0,0,0.3);
    font-family: system-ui, -apple-system, sans-serif;
  `;

  const title = document.createElement("h3");
  title.textContent = "Set Hierarchy Order";
  title.style.cssText = "margin: 0 0 16px; font-size:18px; color:#333;";
  box.appendChild(title);

  const hint = document.createElement("p");
  hint.textContent = "Drag or use arrows to reorder dimensions from outermost to innermost ring.";
  hint.style.cssText = "margin:0 0 14px; font-size:12px; color:#888;";
  box.appendChild(hint);

  // ✅ Fix #4 - Render directly from selectedDimensionsOrder, no tempOrder
  function renderRows() {
    const listContainer = box.querySelector(".dim-list") || (() => {
      const d = document.createElement("div");
      d.className = "dim-list";
      box.appendChild(d);
      return d;
    })();

    listContainer.innerHTML = "";

    selectedDimensionsOrder.forEach((dim, index) => {
      const row = document.createElement("div");
      row.style.cssText = `
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 10px 12px;
        margin-bottom: 6px;
        background: #f7f9fc;
        border: 1px solid #e0e6ef;
        border-radius: 8px;
      `;

      const badge = document.createElement("span");
      badge.textContent = `${index + 1}`;
      badge.style.cssText = `
        background: #007bff;
        color: #fff;
        border-radius: 50%;
        width: 22px;
        height: 22px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-size: 11px;
        font-weight: 700;
        margin-right: 10px;
        flex-shrink: 0;
      `;

      const label = document.createElement("span");
      label.textContent = dim;
      label.style.cssText = "flex:1; font-size:14px; color:#333;";

      const controls = document.createElement("div");
      controls.style.cssText = "display:flex; gap:6px;";

      const up = document.createElement("button");
      up.textContent = "▲";
      up.disabled = index === 0;
      up.style.cssText = `
        padding: 4px 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
        cursor: ${index === 0 ? "not-allowed" : "pointer"};
        background: ${index === 0 ? "#eee" : "#fff"};
        font-size: 12px;
      `;

      const down = document.createElement("button");
      down.textContent = "▼";
      down.disabled = index === selectedDimensionsOrder.length - 1;
      down.style.cssText = `
        padding: 4px 8px;
        border: 1px solid #ccc;
        border-radius: 4px;
        cursor: ${index === selectedDimensionsOrder.length - 1 ? "not-allowed" : "pointer"};
        background: ${index === selectedDimensionsOrder.length - 1 ? "#eee" : "#fff"};
        font-size: 12px;
      `;

      // ✅ Fix #4 - Mutate selectedDimensionsOrder directly, re-render list only
      up.onclick = () => {
        if (index > 0) {
          [selectedDimensionsOrder[index - 1], selectedDimensionsOrder[index]] =
          [selectedDimensionsOrder[index], selectedDimensionsOrder[index - 1]];
          renderRows(); // ✅ Only re-render rows, not whole modal
        }
      };

      down.onclick = () => {
        if (index < selectedDimensionsOrder.length - 1) {
          [selectedDimensionsOrder[index + 1], selectedDimensionsOrder[index]] =
          [selectedDimensionsOrder[index], selectedDimensionsOrder[index + 1]];
          renderRows(); // ✅ Only re-render rows, not whole modal
        }
      };

      controls.appendChild(up);
      controls.appendChild(down);

      row.appendChild(badge);
      row.appendChild(label);
      row.appendChild(controls);
      listContainer.appendChild(row);
    });
  }

  renderRows();

  // Footer buttons
  const footer = document.createElement("div");
  footer.style.cssText = `
    display: flex;
    justify-content: flex-end;
    gap: 10px;
    margin-top: 20px;
    padding-top: 16px;
    border-top: 1px solid #eee;
  `;

  const cancelBtn = document.createElement("button");
  cancelBtn.textContent = "Cancel";
  cancelBtn.style.cssText = `
    padding: 8px 18px;
    border: 1px solid #ccc;
    border-radius: 6px;
    background: #fff;
    cursor: pointer;
    font-size: 13px;
  `;

  const resetBtn = document.createElement("button");
  resetBtn.textContent = "Reset";
  resetBtn.style.cssText = `
    padding: 8px 18px;
    border: 1px solid #ccc;
    border-radius: 6px;
    background: #f5f5f5;
    cursor: pointer;
    font-size: 13px;
  `;

  const applyBtn = document.createElement("button");
  applyBtn.textContent = "Apply";
  applyBtn.style.cssText = `
    padding: 8px 18px;
    border: none;
    border-radius: 6px;
    background: #007bff;
    color: #fff;
    cursor: pointer;
    font-size: 13px;
    font-weight: 600;
  `;

  cancelBtn.onclick = () => {
    // Restore original order on cancel
    selectedDimensionsOrder = detectedDimensions.slice();
    closeModal();
  };

  resetBtn.onclick = () => {
    selectedDimensionsOrder = detectedDimensions.slice();
    renderRows();
  };

  applyBtn.onclick = () => {
    closeModal();
    loadData();
  };

  footer.appendChild(cancelBtn);
  footer.appendChild(resetBtn);
  footer.appendChild(applyBtn);
  box.appendChild(footer);
  modal.appendChild(box);

  // Close on backdrop click
  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      selectedDimensionsOrder = detectedDimensions.slice();
      closeModal();
    }
  });
}

function closeModal() {
  const modal = document.getElementById("hierarchy-modal");
  if (modal) modal.remove();
}

// ==================== HELPERS ====================
function isAncestor(ancestor, node) {
  let current = node;
  while (current) {
    if (current === ancestor) return true;
    current = current.parent;
  }
  return false;
}

function isDescendant(descendant, node) {
  return isAncestor(node, descendant);
}

function showTooltip(event, html) {
  tooltip
    .html(html)
    .style("opacity", 1)
    .style("left", (event.clientX + 15) + "px")
    .style("top", (event.clientY - 28) + "px");
}

function moveTooltip(event) {
  tooltip
    .style("left", (event.clientX + 15) + "px")
    .style("top", (event.clientY - 28) + "px");
}

function hideTooltip() {
  tooltip.style("opacity", 0);
}

function truncate(str, n) {
  const s = String(str ?? "");
  return s.length > n ? s.slice(0, n - 1) + "…" : s;
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = String(str ?? "");
  return div.innerHTML;
}