import './style.css';
import Plotly from 'plotly.js-dist-min';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';

document.head.insertAdjacentHTML('beforeend', `<style>
/* Phase 5 radial presentation polish */
.radial-presentation-active .chart-title {
  font-size: 21px !important;
  line-height: 1.18 !important;
  margin-bottom: 4px !important;
}
.radial-presentation-active .chart-subtitle {
  font-size: 12px !important;
  line-height: 1.25 !important;
}
.radial-presentation-active .canvas-card {
  overflow: visible !important;
}

/* Phase 5 advanced application */
.dashboard-empty {
  grid-column: 1 / -1;
  padding: 18px;
  border: 1px dashed rgba(255,255,255,.16);
  border-radius: 12px;
  text-align: center;
  color: #9ca3af;
  background: rgba(15,23,42,.45);
  font-size: 13px;
}

.stats-panel {
  margin-top: 8px;
  border: 1px solid rgba(255,255,255,.08);
  border-radius: 10px;
  overflow: hidden;
}
.stat-row {
  display: flex;
  justify-content: space-between;
  gap: 12px;
  padding: 7px 9px;
  font-size: 12px;
  border-bottom: 1px solid rgba(255,255,255,.06);
}
.stat-row:last-child { border-bottom: 0; }
.stat-row strong { font-variant-numeric: tabular-nums; }
.dashboard-canvas {
  width: 100%;
  max-width: 100%;
  min-width: 0;
  display: grid;
  gap: 14px;
  padding: 16px 0 24px;
  box-sizing: border-box;
}
.dashboard-card {
  min-width: 0;
  overflow: hidden;
  border: 1px solid rgba(255,255,255,.08);
  border-radius: 14px;
  background: rgba(15,23,42,.82);
  box-sizing: border-box;
}
.dashboard-card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 8px;
  padding: 10px 12px;
}
.dashboard-card-actions {
  display: flex;
  gap: 4px;
  flex-wrap: wrap;
}
.mini-action {
  border: 0;
  border-radius: 6px;
  padding: 4px 7px;
  cursor: pointer;
  background: rgba(255,255,255,.09);
  color: inherit;
  font-size: 11px;
}
.dashboard-plot {
  width: 100%;
  min-width: 0;
  height: 360px;
}
.dashboard-list {
  display: grid;
  gap: 5px;
  margin-top: 8px;
}
.dashboard-list-row {
  display: flex;
  justify-content: space-between;
  gap: 8px;
  padding: 6px 8px;
  border-radius: 7px;
  background: rgba(255,255,255,.04);
  font-size: 11px;
}
.accessibility-mode {
  --focus-ring: #ffffff;
}
.accessibility-mode button:focus-visible,
.accessibility-mode input:focus-visible,
.accessibility-mode select:focus-visible,
.accessibility-mode textarea:focus-visible {
  outline: 3px solid var(--focus-ring);
  outline-offset: 2px;
}
.accessibility-mode button,
.accessibility-mode input,
.accessibility-mode select,
.accessibility-mode textarea {
  font-size: 14px;
}
@media (max-width: 1100px) {
  .dashboard-canvas { grid-template-columns: 1fr !important; }
}


/* Phase 4 layout fix v4: keep the native file picker fully inside Import Data. */
#dropZone {
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  box-sizing: border-box !important;
  overflow: hidden !important;
}
#csvFile {
  display: block !important;
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  box-sizing: border-box !important;
  overflow: hidden !important;
}
#pasteData {
  display: block !important;
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}
.sidebar input,
.sidebar select,
.sidebar button,
.sidebar textarea {
  max-width: 100%;
  box-sizing: border-box;
}

/* Phase 4 layout fix: keep the Plotly canvas strictly inside the right content column. */
.app-container {
  display: grid !important;
  grid-template-columns: minmax(360px, 390px) minmax(0, 1fr) !important;
  width: 100% !important;
  max-width: 100vw !important;
  min-width: 0 !important;
  overflow-x: hidden !important;
  box-sizing: border-box !important;
}
.sidebar {
  width: auto !important;
  min-width: 0 !important;
  max-width: 390px !important;
  box-sizing: border-box !important;
  position: relative !important;
  overflow-x: hidden !important;
  z-index: 2 !important;
}
.main-content {
  min-width: 0 !important;
  width: auto !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
  overflow: visible !important;
  position: relative !important;
  z-index: 1 !important;
  padding-bottom: 24px !important;
}
.canvas-card {
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  height: auto !important;
  min-height: 0 !important;
  box-sizing: border-box !important;
  overflow: visible !important;
}
.canvas-wrapper,
#visualizationCanvas {
  width: 100% !important;
  max-width: 100% !important;
  min-width: 0 !important;
  height: 650px;
  min-height: 0 !important;
  flex: 0 0 auto !important;
  box-sizing: border-box !important;
  overflow: visible !important;
}
#visualizationCanvas .js-plotly-plot,
#visualizationCanvas .plot-container,
#visualizationCanvas .svg-container {
  width: 100% !important;
  max-width: 100% !important;
  box-sizing: border-box !important;
}
@media (max-width: 900px) {
  .app-container {
    grid-template-columns: 320px minmax(0, 1fr) !important;
  }
}
 .mini-label{display:block;font-size:.62rem;color:#9ca3af;text-transform:uppercase;margin-bottom:3px}.check-row{display:flex;align-items:center;gap:5px;font-size:.68rem;color:#d1d5db;text-transform:none;letter-spacing:0}.check-row input{width:auto;padding:0}.status-text { font-size: 11px; opacity: .75; margin-top: 8px; word-break: break-word; }
.button-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 6px; margin-top: 6px; }
</style>`);
document.querySelector<HTMLDivElement>('#app')!.innerHTML = `
  <div class="app-container">
    <aside class="sidebar" style="width: 390px;">
      <div class="brand">
        <h1>VizNova</h1>
        <p style="font-size: 0.75rem; color: #6b7280; margin-top: 4px;">Ultimate Multi-Domain Master Engine</p>
      </div>

      <div class="control-group">
        <label>1. Import Data</label>
        <div id="dropZone" style="border:1px dashed rgba(96,165,250,0.55); border-radius:10px; padding:14px; text-align:center; background:rgba(59,130,246,0.05);">
          <div style="font-size:0.78rem; color:#d1d5db; margin-bottom:8px;">Drop Excel / CSV / TSV here</div>
          <input type="file" id="csvFile" accept=".csv,.tsv,.xlsx,.xls" style="width:100%;" />
          <button id="pasteData" type="button" style="margin-top:8px; width:100%; background:rgba(59,130,246,0.75);">Paste Data from Clipboard</button>
        </div>
        <div id="sheetSelectorWrap" style="display:none; margin-top:8px;">
          <span style="font-size:0.65rem; color:#9ca3af; text-transform:uppercase;">Excel Sheet</span>
          <select id="sheetSelect" style="width:100%; margin-top:3px;"></select>
        </div>
        <div id="dataStatus" style="font-size:0.72rem; color:#9ca3af; margin-top:8px;"></div>
      </div>

      <div id="columnMappingSection" class="control-group" style="display: none; background: rgba(255,255,255,0.03); padding: 12px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.08);">
        <label style="color: #60a5fa; font-weight: 700; margin-bottom: 6px; display: block;">2. Data Mapping</label>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 8px;">
          <div>
            <span style="font-size: 0.65rem; color: #9ca3af; text-transform: uppercase;">X-Axis / Category</span>
            <select id="colX" style="width:100%; font-size:0.75rem; padding:6px; margin-top:2px;"></select>
          </div>
          <div>
            <span style="font-size: 0.65rem; color: #9ca3af; text-transform: uppercase;">Y-Axis / Primary Metric</span>
            <select id="colY" style="width:100%; font-size:0.75rem; padding:6px; margin-top:2px;"></select>
          </div>
          <div>
            <span style="font-size: 0.65rem; color: #fbbf24; text-transform: uppercase;">Chart Value / Metric</span>
            <select id="chartMetric" style="width:100%; font-size:0.75rem; padding:6px; margin-top:2px;"></select>
          </div>
          <div>
            <span style="font-size: 0.65rem; color: #9ca3af; text-transform: uppercase;">Z-Axis / Matrix Value</span>
            <select id="colZ" style="width:100%; font-size:0.75rem; padding:6px; margin-top:2px;"></select>
          </div>
          <div>
            <span style="font-size: 0.65rem; color: #9ca3af; text-transform: uppercase;">Group / Series</span>
            <select id="colT" style="width:100%; font-size:0.75rem; padding:6px; margin-top:2px;"></select>
          </div>
        </div>
        <div id="mappingHint" style="font-size:0.68rem; color:#9ca3af; margin-top:8px; line-height:1.35;"></div>
      </div>

      <div id="dataPreviewSection" class="control-group" style="display:none; background:rgba(255,255,255,0.03); padding:12px; border-radius:8px; border:1px solid rgba(255,255,255,0.08);">
        <label style="color:#34d399; font-weight:700;">3. Data Preview & Validation</label>
        <div id="validationSummary" style="font-size:0.72rem; margin:6px 0; line-height:1.45;"></div>
        <div id="dataPreview" style="max-height:230px; overflow:auto; border:1px solid rgba(255,255,255,0.08); border-radius:6px;"></div>
      </div>

      <div class="control-group">
        <label>Exhaustive Master Plot Library</label>
        <select id="chartTypeSelect">
          <optgroup label="Corporate & Financial">
            <option value="bar">Vertical Corporate Column Bar</option>
            <option value="horizontalBar">Horizontal Matrix Bar</option>
            <option value="stackedBar">Stacked Multi-Tier Bar</option>
            <option value="waterfall">Financial Waterfall Variance</option>
            <option value="pie">Proportional Composition Pie</option>
            <option value="doughnut">Glowing Polar Donut</option>
          </optgroup>
          <optgroup label="Scientific & Statistical">
            <option value="heatmap">Advanced Scientific Matrix Heatmap</option>
            <option value="contour">Contour Elevation Map</option>
            <option value="boxplot">Statistical Box & Whisker Distribution</option>
            <option value="violin">Violin Probability Density</option>
            <option value="scatter">Cosmic Scatter Correlation</option>
            <option value="bubble">Multi-Dimensional Bubble Cluster</option>
          </optgroup>
          <optgroup label="Creative & Advanced">
            <option value="radar">Cyberpunk Multi-Axis Spider Web (Radar)</option>
            <option value="polarArea">Aurora Polar Area Matrix</option>
            <option value="annotatedRadial">Presentation Radial Bar • Callout Ring</option>
            <option value="funnel">Conversion Drop-off Funnel</option>
            <option value="surface3d">3D Volumetric Surface Space</option>
          </optgroup>
        </select>
      </div>

      <div class="control-group" id="radialPresentationControls" style="display:none;">
        <label style="color:#f59e0b;">Presentation Radial Controls</label>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
          <div>
            <span class="mini-label">Center Title</span>
            <input id="radialCenterTitle" value="Girls Scholarship" style="width:100%;box-sizing:border-box;" />
          </div>
          <div>
            <span class="mini-label">Detail Metric</span>
            <select id="radialDetailMetric" style="width:100%;box-sizing:border-box;">
              <option value="">None</option>
            </select>
          </div>
        </div>
        <div id="radialLabelEditorGroup" style="margin-top:10px;display:none;">
          <span class="mini-label">Editable Callout Labels</span>
          <div id="radialLabelEditor" style="display:grid;gap:6px;margin-top:6px;"></div>
          <button id="resetRadialLabels" type="button" style="margin-top:8px;width:100%;">
            Reset Callout Labels to Data
          </button>
        </div>
        <label class="check-row" style="margin-top:8px;">
          <input id="radialCallouts" type="checkbox" checked />
          <span>External callout labels</span>
        </label>
        <label class="check-row">
          <input id="radialConnectors" type="checkbox" checked />
          <span>Connector lines</span>
        </label>
        <div id="radialMetricStatus" class="status-text">The radial chart uses the selected Chart Value / Metric.</div>
        <div class="status-text">Designed for presentation-style radial comparisons. Labels stay outside the polar plot.</div>
      </div>

      <div class="control-group">
        <label>Color Combination & Scientific Scales</label>
        <select id="paletteSelect">
          <option value="Viridis">Scientific Viridis (Default)</option>
          <option value="Plasma">Scientific Plasma</option>
          <option value="Cividis">Scientific Cividis</option>
          <option value="YlGnBu">Yellow-Green-Blue Gradient</option>
          <option value="Jet">Classic Jet Spectrum</option>
          <option value="cyberpunk">Cyberpunk Neon (Custom Palette)</option>
          <option value="aurora">Aurora Borealis (Custom Palette)</option>
          <option value="custom">Custom Editable Hex Palette</option>
        </select>
      </div>

      <div class="control-group" id="customPaletteGroup" style="display: none;">
        <label>Custom Hex Colors (Comma Separated)</label>
        <input type="text" id="customHexInput" value="#06b6d4, #ec4899, #8b5cf6, #3b82f6, #10b981" />
      </div>

      <div class="control-group">
        <label>Built-in Data Presets</label>
        <select id="datasetSelect">
          <option value="tech">Tech Ecosystem Velocity</option>
          <option value="neural">Neural Network Weights</option>
          <option value="financial">Cosmic Market Flux</option>
        </select>
      </div>

      <div class="control-group" id="categoryLabelEditorGroup"
        style="display: none; background: rgba(255,255,255,0.03); padding: 12px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.08);">
        <label style="color: #60a5fa; font-weight: 700; margin-bottom: 6px; display: block;">
          Editable Category Labels
        </label>
        <p style="font-size: 0.7rem; color: #9ca3af; margin: 0 0 8px 0; line-height: 1.35;">
          Edit one category label per line. Use <b>|</b> to split a label into two lines.
        </p>
        <textarea id="categoryLabelInput"
          spellcheck="false"
          style="width:100%; min-height:180px; resize:vertical; box-sizing:border-box; line-height:1.45; font-size:0.78rem;"></textarea>
        <button id="applyCategoryLabels" style="margin-top:8px; width:100%;">
          Apply Category Labels
        </button>
        <button id="resetCategoryLabels"
          style="margin-top:6px; width:100%; background: rgba(75,85,99,0.8);">
          Reset to Data Labels
        </button>
      </div>

      <div class="control-group" id="appearanceControls" style="background:rgba(255,255,255,0.03); padding:12px; border-radius:8px; border:1px solid rgba(255,255,255,0.08);">
        <label style="color:#c084fc; font-weight:700;">4. Chart Appearance & Controls</label>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:7px;">
          <div><span class="mini-label">Title Size</span><input id="titleSize" type="number" min="12" max="60" value="24" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">Label Size</span><input id="labelSize" type="number" min="7" max="24" value="11" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">Plot Height</span><input id="plotHeight" type="number" min="300" max="1400" value="650" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">Opacity</span><input id="plotOpacity" type="number" min="0.1" max="1" step="0.05" value="0.9" style="width:100%;box-sizing:border-box;"></div>
        </div>
        <div style="margin-top:8px;display:grid;grid-template-columns:1fr 1fr;gap:6px;">
          <label class="check-row"><input id="showDataLabels" type="checkbox"> Data labels</label>
          <label class="check-row"><input id="showLegend" type="checkbox" checked> Legend</label>
          <label class="check-row"><input id="showGrid" type="checkbox" checked> Grid</label>
          <label class="check-row"><input id="reverseColors" type="checkbox"> Reverse colors</label>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:8px;">
          <div><span class="mini-label">X Axis Title</span><input id="xAxisTitle" type="text" placeholder="Auto" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">Y Axis Title</span><input id="yAxisTitle" type="text" placeholder="Auto" style="width:100%;box-sizing:border-box;"></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:8px;">
          <div><span class="mini-label">X Min</span><input id="xMin" type="number" placeholder="Auto" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">X Max</span><input id="xMax" type="number" placeholder="Auto" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">Y Min</span><input id="yMin" type="number" placeholder="Auto" style="width:100%;box-sizing:border-box;"></div>
          <div><span class="mini-label">Y Max</span><input id="yMax" type="number" placeholder="Auto" style="width:100%;box-sizing:border-box;"></div>
        </div>
        <button id="resetAppearance" type="button" style="margin-top:8px;width:100%;background:rgba(75,85,99,0.8);">Reset Appearance</button>
      </div>

      <hr style="border-color: rgba(255,255,255,0.08); margin: 4px 0;" />

      
      <div class="control-group" id="advancedControls"
        style="background:rgba(255,255,255,0.03); padding:12px; border-radius:8px; border:1px solid rgba(255,255,255,0.08);">
        <label style="color:#f59e0b; font-weight:700;">5. Advanced Application</label>

        <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;">
          <button id="addDashboardChart" type="button">Add to Dashboard</button>
          <button id="clearDashboard" type="button" style="background:rgba(75,85,99,.8);">Clear Dashboard</button>
        </div>

        <div style="margin-top:8px;">
          <span class="mini-label">Dashboard Columns</span>
          <select id="dashboardColumns" style="width:100%;box-sizing:border-box;">
            <option value="1">1 column</option>
            <option value="2" selected>2 columns</option>
            <option value="3">3 columns</option>
          </select>
        </div>

        <div style="margin-top:8px;">
          <span class="mini-label">Cross-Filter</span>
          <button id="clearDashboardFilter" type="button" style="width:100%;background:rgba(75,85,99,.8);">Clear Dashboard Filter</button>
          <div id="dashboardFilterStatus" class="status-text">No dashboard filter active.</div>
        </div>

        <div style="margin-top:8px;">
          <span class="mini-label">Statistics</span>
          <div id="statsPanel" class="stats-panel">
            <div class="stat-row"><span>Rows</span><strong id="statRows">0</strong></div>
            <div class="stat-row"><span>Numeric values</span><strong id="statNumeric">0</strong></div>
            <div class="stat-row"><span>Missing values</span><strong id="statMissing">0</strong></div>
            <div class="stat-row"><span>Mean (Y)</span><strong id="statMean">—</strong></div>
            <div class="stat-row"><span>Median (Y)</span><strong id="statMedian">—</strong></div>
            <div class="stat-row"><span>Min / Max (Y)</span><strong id="statRange">—</strong></div>
          </div>
        </div>

        <label class="check-row" style="margin-top:8px;">
          <input id="accessibilityMode" type="checkbox" />
          <span>Accessibility mode</span>
        </label>
      </div>

      <div class="control-group" id="projectControls"
        style="background:rgba(255,255,255,0.03); padding:12px; border-radius:8px; border:1px solid rgba(255,255,255,0.08);">
        <label style="color:#34d399; font-weight:700;">6. Project</label>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;">
          <button id="newProject" type="button" style="background:rgba(75,85,99,.8);">New</button>
          <button id="saveProject" type="button">Save Project</button>
          <button id="loadProject" type="button">Load Project</button>
          <button id="saveTemplate" type="button">Save Template</button>
          <button id="loadTemplate" type="button">Load Template</button>
        </div>
        <div style="margin-top:8px;">
          <span class="mini-label">Recent Projects</span>
          <select id="recentProjects" style="width:100%;box-sizing:border-box;">
            <option value="">Recent projects...</option>
          </select>
        </div>
        <label class="check-row" style="margin-top:8px;">
          <input id="autoSaveProject" type="checkbox" checked />
          <span>Autosave project locally</span>
        </label>
        <div id="projectStatus" class="status-text">Project: Untitled</div>
      </div>

<div class="control-group" id="exportControls" style="background:rgba(255,255,255,0.03); padding:12px; border-radius:8px; border:1px solid rgba(255,255,255,0.08);">
        <label style="color:#60a5fa; font-weight:700;">7. Professional Export</label>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:7px;">
          <div><span class="mini-label">Export Preset</span><select id="exportPreset" style="width:100%;box-sizing:border-box;">
            <option value="custom">Custom</option>
            <option value="word">Word / Report</option>
            <option value="presentation">Presentation</option>
            <option value="web">Web / Transparent</option>
          </select></div>
          <div><span class="mini-label">Background</span><select id="exportBackground" style="width:100%;box-sizing:border-box;">
            <option value="transparent">Transparent</option>
            <option value="white">White</option>
            <option value="dark">Dark</option>
            <option value="custom">Custom</option>
          </select></div>
          <div><span class="mini-label">Custom BG</span><input id="exportCustomBg" type="color" value="#ffffff" style="width:100%;height:34px;padding:2px;box-sizing:border-box;"></div>
          <div><span class="mini-label">Raster Scale</span><select id="exportScale" style="width:100%;box-sizing:border-box;">
            <option value="2">2×</option><option value="4" selected>4×</option><option value="6">6×</option>
          </select></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:8px;">
          <button id="exportPng">PNG</button>
          <button id="exportJpeg">JPEG</button>
          <button id="exportSvg">SVG</button>
          <button id="exportPdf">PDF</button>
          <button id="exportWebp">WebP</button>
          <button id="copyChart">Copy Chart</button>
        </div>
        <div id="exportStatus" style="font-size:.68rem;color:#9ca3af;margin-top:7px;min-height:16px;"></div>
      </div>
    </aside>

    <main class="main-content">
      <div id="dashboardCanvas" class="dashboard-canvas" aria-label="VizNova dashboard"></div>
      <div class="canvas-card" id="renderCard" style="display: flex; flex-direction: column;">
        <div class="chart-title" contenteditable="true" spellcheck="false">Universal Master Analytics Workspace</div>
        <div class="chart-subtitle" contenteditable="true" spellcheck="false">Double-click to edit title • Powered by Plotly Scientific Engine</div>
        
        <div class="canvas-wrapper" style="width: 100%;" id="visualizationCanvas"></div>
      </div>
    </main>
  </div>
`;

// Palettes & Presets

const projectFileInput = document.createElement('input');
projectFileInput.type = 'file';
projectFileInput.accept = '.aetherviz,.json';
projectFileInput.style.display = 'none';
document.body.appendChild(projectFileInput);

const templateFileInput = document.createElement('input');
templateFileInput.type = 'file';
templateFileInput.accept = '.aethertemplate,.json';
templateFileInput.style.display = 'none';
document.body.appendChild(templateFileInput);

const palettes: Record<string, string[]> = {
  cyberpunk: ['#06b6d4', '#ec4899', '#8b5cf6', '#3b82f6', '#10b981', '#f59e0b', '#f43f5e', '#14b8a6'],
  aurora: ['#10b981', '#06b6d4', '#6366f1', '#a855f7', '#ec4899', '#14b8a6', '#84cc16', '#3b82f6']
};

const defaultDatasets: Record<string, { labels: string[], data: number[], label: string }> = {
  tech: { labels: ['Alpha', 'Beta', 'Gamma', 'Delta', 'Omega'], data: [85, 92, 78, 95, 88], label: 'Velocity Index' },
  neural: { labels: ['Layer 1', 'Layer 2', 'Layer 3', 'Layer 4', 'Layer 5'], data: [64, 79, 91, 98, 84], label: 'Synapse Activation' },
  financial: { labels: ['Sector A', 'Sector B', 'Sector C', 'Sector D', 'Sector E'], data: [120, 210, 150, 300, 280], label: 'Capital Flow ($B)' }
};

let rawParsedData: any[] = [];
let availableColumns: string[] = [];
let workbook: XLSX.WorkBook | null = null;
let currentFileName = '';
let currentLabels: string[] = defaultDatasets.tech.labels;
let currentDisplayLabels: string[] = [...currentLabels];
let radialPresentationLabels: string[] = [...currentDisplayLabels];
let currentValues: number[] = defaultDatasets.tech.data;
let currentLabelName: string = defaultDatasets.tech.label;
let chartMetricKey = '';

function getChartMetricKey(): string {
  const metric = (document.getElementById('chartMetric') as HTMLSelectElement | null)?.value;
  return metric || (document.getElementById('colY') as HTMLSelectElement | null)?.value || '';
}

function metricLooksLikeRate(name: string): boolean {
  return /%|percent|percentage|rate|ratio|share|proportion|female|girl/i.test(name);
}


function getActiveColors(): any {
  const paletteKey = (document.getElementById('paletteSelect') as HTMLSelectElement).value;
  if (paletteKey === 'custom') {
    const customText = (document.getElementById('customHexInput') as HTMLInputElement).value;
    return customText.split(',').map(c => c.trim()).filter(c => c.length > 0);
  }
  if (palettes[paletteKey]) {
    return palettes[paletteKey];
  }
  return paletteKey; // Returns Plotly scientific color scale name (e.g. 'Viridis', 'Plasma')
}

function isNumericColumn(column: string): boolean {
  if (!rawParsedData.length) return false;
  const values = rawParsedData.map(row => row[column]).filter(v => v !== '' && v !== null && v !== undefined);
  if (!values.length) return false;
  const numeric = values.filter(v => Number.isFinite(Number(v))).length;
  return numeric / values.length >= 0.8;
}


function chooseBestNumericMetric(numericColumns: string[], preferredRate = false): string {
  if (!numericColumns.length) return '';
  const scored = numericColumns.map(col => {
    const name = String(col).toLowerCase();
    let score = 0;
    if (/percentage|percent|%/.test(name)) score += 100;
    if (/female|girls?|women/.test(name)) score += 70;
    if (/rate|ratio|share|proportion/.test(name)) score += 45;
    if (/scholarship/.test(name)) score += 20;
    if (preferredRate && score > 0) score += 50;
    return { col, score };
  });
  scored.sort((a, b) => b.score - a.score);
  return scored[0]?.col || numericColumns[0];
}

function markImportedDatasetActive(): void {
  const ds = document.getElementById('datasetSelect') as HTMLSelectElement | null;
  if (!ds) return;
  let imported = Array.from(ds.options).find(o => o.value === '__imported__');
  if (!imported) {
    imported = new Option('Imported Data (current file)', '__imported__');
    ds.insertBefore(imported, ds.firstChild);
  }
  ds.value = '__imported__';
}

function populateMappingControls(): void {
  const selects = ['colX', 'colY', 'colZ', 'colT'].map(id => document.getElementById(id) as HTMLSelectElement | null);
  if (selects.some(s => !s)) return;
  const [xSelect, ySelect, zSelect, tSelect] = selects as HTMLSelectElement[];

  [xSelect, ySelect, zSelect, tSelect].forEach(sel => {
    sel.innerHTML = '';
    if (sel === zSelect || sel === tSelect) sel.add(new Option('— None —', ''));
  });

  availableColumns.forEach((col) => {
    xSelect.add(new Option(col, col));
    ySelect.add(new Option(col, col));
    zSelect.add(new Option(col, col));
    tSelect.add(new Option(col, col));
  });

  // Robust automatic mapping: first non-numeric column = category;
  // first numeric = primary metric; second/third numeric = additional metrics.
  const nonNumeric = availableColumns.find(col => !isNumericColumn(col));
  const numericColumns = availableColumns.filter(isNumericColumn);
  xSelect.value = nonNumeric || availableColumns[0] || '';
  ySelect.value = numericColumns[0] || availableColumns[1] || '';
  if (numericColumns[1]) zSelect.value = numericColumns[1];

  const metricSelect = document.getElementById('chartMetric') as HTMLSelectElement | null;
  if (metricSelect) {
    metricSelect.innerHTML = '';
    numericColumns.forEach(col => metricSelect.add(new Option(col, col)));
    chartMetricKey = chooseBestNumericMetric(numericColumns, (document.getElementById('chartTypeSelect') as HTMLSelectElement | null)?.value === 'annotatedRadial') || ySelect.value || '';
    metricSelect.value = chartMetricKey;
  }

  document.getElementById('columnMappingSection')!.style.display = 'block';
  document.getElementById('dataPreviewSection')!.style.display = 'block';
  updateMappingHint();
}

function updateMappingHint(): void {
  const chartType = (document.getElementById('chartTypeSelect') as HTMLSelectElement)?.value;
  const hint = document.getElementById('mappingHint');
  if (!hint) return;
  if (chartType === 'heatmap' || chartType === 'contour') {
    hint.textContent = 'Use X = category, Y = second category, and Z = numeric value for a true matrix. Group is optional.';
  } else if (['scatter', 'bubble', 'stackedBar'].includes(chartType)) {
    hint.textContent = 'Use Group / Series to split the data into multiple traces. Bubble charts also use Y as the primary size/value source.';
  } else {
    hint.textContent = 'X = category. Chart Value / Metric controls the numeric field plotted. Y remains the primary mapping; Z is available for a second numeric field or matrix charts.';
  }
}

function validateData(): { missingX: number; invalidY: number; duplicateX: number; rows: number } {
  const xKey = (document.getElementById('colX') as HTMLSelectElement)?.value;
  const yKey = (document.getElementById('colY') as HTMLSelectElement)?.value;
  const missingX = xKey ? rawParsedData.filter(row => row[xKey] === '' || row[xKey] === null || row[xKey] === undefined).length : 0;
  const invalidY = yKey ? rawParsedData.filter(row => row[yKey] === '' || row[yKey] === null || row[yKey] === undefined || !Number.isFinite(Number(row[yKey]))).length : 0;
  const labels = xKey ? rawParsedData.map(row => String(row[xKey] ?? '')).filter(Boolean) : [];
  const duplicateX = labels.length - new Set(labels).size;
  return { missingX, invalidY, duplicateX, rows: rawParsedData.length };
}

function updateValidationAndPreview(): void {
  const summary = document.getElementById('validationSummary');
  const preview = document.getElementById('dataPreview');
  if (!summary || !preview || !rawParsedData.length) return;

  const v = validateData();
  const valid = Math.max(0, v.rows - v.missingX - v.invalidY);
  const warnings: string[] = [];
  if (v.missingX) warnings.push(`${v.missingX} missing X values`);
  if (v.invalidY) warnings.push(`${v.invalidY} non-numeric/missing Y values`);
  if (v.duplicateX) warnings.push(`${v.duplicateX} repeated X categories`);

  summary.innerHTML = `<span style="color:#34d399">✓ ${v.rows.toLocaleString()} rows detected</span> · ${availableColumns.length} columns · ${valid.toLocaleString()} chartable rows` +
    (warnings.length ? `<br><span style="color:#fbbf24">⚠ ${warnings.join(' · ')}</span>` : `<br><span style="color:#34d399">✓ No mapping warnings detected</span>`);

  const cols = availableColumns;
  const rows = rawParsedData.slice(0, 12);
  let html = '<table style="border-collapse:collapse; width:100%; font-size:0.68rem;"><thead><tr>';
  cols.forEach(c => { html += `<th style="position:sticky;top:0;background:#111827;color:#9ca3af;text-align:left;padding:5px;border-bottom:1px solid rgba(255,255,255,.1);white-space:nowrap;">${escapePlotlyLabel(c)}</th>`; });
  html += '</tr></thead><tbody>';
  rows.forEach(row => {
    html += '<tr>';
    cols.forEach(c => { html += `<td style="padding:5px;border-bottom:1px solid rgba(255,255,255,.06);white-space:nowrap;">${escapePlotlyLabel(String(row[c] ?? ''))}</td>`; });
    html += '</tr>';
  });
  html += '</tbody></table>';
  preview.innerHTML = html;
}

function getMappedRows(): any[] {
  const xKey = (document.getElementById('colX') as HTMLSelectElement)?.value;
  const yKey = getChartMetricKey();
  if (!rawParsedData.length || !xKey || !yKey) return [];
  return rawParsedData
    .map(row => ({
      x: row[xKey],
      y: Number(row[yKey]),
      z: (document.getElementById('colZ') as HTMLSelectElement)?.value ? Number(row[(document.getElementById('colZ') as HTMLSelectElement).value]) : NaN,
      group: (document.getElementById('colT') as HTMLSelectElement)?.value ? String(row[(document.getElementById('colT') as HTMLSelectElement).value]) : '',
      raw: row
    }))
    .filter(r => r.x !== '' && r.x !== null && r.x !== undefined && Number.isFinite(r.y));
}

function processMappingAndRender() {
  if (rawParsedData.length > 0) {
    const xSelect = document.getElementById('colX') as HTMLSelectElement;
    const ySelect = document.getElementById('colY') as HTMLSelectElement;
    if (xSelect?.value && ySelect?.value) {
      const rows = getMappedRows();
      currentLabels = rows.map(row => String(row.x));
      currentDisplayLabels = [...currentLabels];
      radialPresentationLabels = [...currentDisplayLabels];
      currentValues = rows.map(row => row.y);
      currentLabelName = getChartMetricKey() || ySelect.value;
      syncCategoryLabelEditor();
      syncRadialPresentationLabelEditor();

    }
    updateValidationAndPreview();
  }
  renderVisualization();
}

function escapePlotlyLabel(value: string): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function formatPolarLabel(value: string): string {
  // A pipe lets the user manually control a two-line category label.
  const parts = String(value).split('|').map(part => part.trim()).filter(Boolean);

  if (parts.length >= 2) {
    return `${escapePlotlyLabel(parts[0])}<br>${escapePlotlyLabel(parts.slice(1).join(' | '))}`;
  }

  return escapePlotlyLabel(value);
}

function syncCategoryLabelEditor(): void {
  const group = document.getElementById('categoryLabelEditorGroup') as HTMLElement | null;
  const input = document.getElementById('categoryLabelInput') as HTMLTextAreaElement | null;

  if (!group || !input) return;

  // Show the editor whenever the loaded data has categories.
  group.style.display = currentLabels.length > 0 ? 'flex' : 'none';

  input.value = currentDisplayLabels.join('\n');
}

function syncRadialPresentationLabelEditor(): void {
  const group = document.getElementById('radialLabelEditorGroup') as HTMLElement | null;
  const editor = document.getElementById('radialLabelEditor') as HTMLElement | null;
  if (!group || !editor) return;

  const active =
    (document.getElementById('chartTypeSelect') as HTMLSelectElement | null)?.value === 'annotatedRadial';

  group.style.display = active && currentDisplayLabels.length ? 'block' : 'none';

  if (!active) return;

  if (radialPresentationLabels.length !== currentDisplayLabels.length) {
    radialPresentationLabels = [...currentDisplayLabels];
  }

  editor.innerHTML = '';

  radialPresentationLabels.forEach((label, index) => {
    const wrap = document.createElement('div');
    wrap.style.display = 'grid';
    wrap.style.gridTemplateColumns = '28px 1fr';
    wrap.style.gap = '6px';
    wrap.style.alignItems = 'center';

    const number = document.createElement('span');
    number.textContent = String(index + 1);
    number.style.opacity = '0.65';
    number.style.fontSize = '12px';

    const input = document.createElement('input');
    input.type = 'text';
    input.value = label;
    input.setAttribute('aria-label', `Radial callout label ${index + 1}`);
    input.style.width = '100%';
    input.style.boxSizing = 'border-box';

    input.addEventListener('input', () => {
      radialPresentationLabels[index] = input.value;
      renderVisualization();
    });

    wrap.append(number, input);
    editor.appendChild(wrap);
  });
}

function resetRadialPresentationLabels(): void {
  radialPresentationLabels = [...currentDisplayLabels];
  syncRadialPresentationLabelEditor();
  renderVisualization();
}

function resetDisplayLabels(): void {
  currentDisplayLabels = [...currentLabels];
  radialPresentationLabels = [...currentDisplayLabels];
  syncCategoryLabelEditor();
  syncRadialPresentationLabelEditor();
  renderVisualization();
}

function applyCategoryLabels(): void {
  const input = document.getElementById('categoryLabelInput') as HTMLTextAreaElement | null;
  if (!input) return;

  const edited = input.value
    .split(/\r?\n/)
    .map(label => label.trim())
    .filter(label => label.length > 0);

  // Keep the chart/data alignment intact. If fewer labels are supplied,
  // preserve the original data labels for the remaining categories.
  currentDisplayLabels = currentLabels.map((original, index) => edited[index] || String(original));
  radialPresentationLabels = [...currentDisplayLabels];

  input.value = currentDisplayLabels.join('\n');
  syncRadialPresentationLabelEditor();
  renderVisualization();
}

function colorAt(colors: any, index: number, fallback = '#06b6d4'): string {
  return Array.isArray(colors) && colors.length ? colors[index % colors.length] : fallback;
}

function buildGroupedTraces(chartType: string, colorInput: any): any[] {
  const rows = getMappedRows();
  const groupKey = (document.getElementById('colT') as HTMLSelectElement)?.value;
  if (!groupKey) return [];
  const groups = Array.from(new Set(rows.map(r => r.group || 'Series 1')));
  return groups.map((group, index) => {
    const data = rows.filter(r => (r.group || 'Series 1') === group);
    const base: any = {
      name: group,
      x: data.map(r => String(r.x)),
      y: data.map(r => r.y),
      marker: { color: colorAt(colorInput, index) }
    };
    if (chartType === 'scatter' || chartType === 'bubble') {
      base.type = 'scatter';
      base.mode = 'markers';
      base.marker.size = chartType === 'bubble'
        ? data.map(r => Math.max(8, (Math.abs(r.y) / Math.max(...rows.map(x => Math.abs(x.y)), 1)) * 36))
        : 12;
      return base;
    }
    base.type = 'bar';
    if (chartType === 'stackedBar') return base;
    base.type = 'scatter';
    base.mode = chartType === 'area' ? 'lines' : 'lines+markers';
    base.fill = chartType === 'area' ? 'tozeroy' : 'none';
    return base;
  });
}

function getControlValue(id: string, fallback = ''): string {
  const el = document.getElementById(id) as HTMLInputElement | null;
  return el ? el.value : fallback;
}

function getOptionalNumber(id: string): number | undefined {
  const value = getControlValue(id).trim();
  if (value === '') return undefined;
  const n = Number(value);
  return Number.isFinite(n) ? n : undefined;
}

function getAppearanceSettings() {
  const reverse = (document.getElementById('reverseColors') as HTMLInputElement | null)?.checked ?? false;
  const colors = getActiveColors();
  const palette = Array.isArray(colors) ? [...colors] : colors;
  if (Array.isArray(palette) && reverse) palette.reverse();
  return {
    titleSize: Number(getControlValue('titleSize','24')) || 24,
    labelSize: Number(getControlValue('labelSize','11')) || 11,
    plotHeight: Number(getControlValue('plotHeight','650')) || 650,
    opacity: Number(getControlValue('plotOpacity','0.9')) || 0.9,
    showDataLabels: (document.getElementById('showDataLabels') as HTMLInputElement | null)?.checked ?? false,
    showLegend: (document.getElementById('showLegend') as HTMLInputElement | null)?.checked ?? true,
    showGrid: (document.getElementById('showGrid') as HTMLInputElement | null)?.checked ?? true,
    colors: palette,
    xTitle: getControlValue('xAxisTitle'),
    yTitle: getControlValue('yAxisTitle'),
    xMin: getOptionalNumber('xMin'), xMax: getOptionalNumber('xMax'),
    yMin: getOptionalNumber('yMin'), yMax: getOptionalNumber('yMax')
  };
}

function applyAppearanceToPage(settings: ReturnType<typeof getAppearanceSettings>): void {
  const title = document.querySelector('.chart-title') as HTMLElement | null;
  const subtitle = document.querySelector('.chart-subtitle') as HTMLElement | null;
  const wrapper = document.getElementById('visualizationCanvas') as HTMLElement | null;
  const radialActive = document.body.classList.contains('radial-presentation-active');
  if (title) title.style.fontSize = `${radialActive ? Math.min(settings.titleSize, 21) : settings.titleSize}px`;
  if (wrapper) { wrapper.style.height = `${settings.plotHeight}px`; wrapper.style.minHeight = '0'; wrapper.style.flex = '0 0 auto'; wrapper.style.overflow = 'visible'; }
  if (subtitle) subtitle.style.fontSize = `${Math.max(10, settings.labelSize - 1)}px`;
}

function makeDataLabels(values: number[]): string[] {
  return values.map(v => Number.isFinite(v) ? String(Number(v.toFixed(2))) : '');
}


type ProjectSnapshot = {
  version: 4;
  name: string;
  createdAt: string;
  updatedAt: string;
  source: {
    kind: string;
    fileName: string;
    rows: Record<string, unknown>[];
    sheetName: string;
  };
  mapping: {
    xKey: string;
    yKey: string;
    zKey: string;
    groupKey: string;
  };
  chart: {
    type: string;
    palette: string;
    customPalette: string;
    categoryLabels: string[];
    title: string;
    subtitle: string;
    metricKey?: string;
    radialPresentationLabels?: string[];
  };
  appearance: Record<string, unknown>;
  dashboardItems?: DashboardItem[];
  dashboardFilterValue?: string | null;
};

type TemplateSnapshot = {
  version: 1;
  name: string;
  createdAt: string;
  chart: ProjectSnapshot["chart"];
  appearance: ProjectSnapshot["appearance"];
};

const PROJECT_STORAGE_KEY = 'aetherviz.projects.v1';
const AUTOSAVE_KEY = 'aetherviz.autosave.v1';
const TEMPLATE_STORAGE_KEY = 'aetherviz.templates.v1';
let currentProjectName = 'Untitled';
let currentSourceFileName = '';
let currentSourceKind = '';
let currentSheetName = '';

function projectStatus(message?: string) {
  const el = document.querySelector<HTMLElement>('#projectStatus');
  if (el) el.textContent = message || `Project: ${currentProjectName}`;
}

function safeJson<T>(value: string | null, fallback: T): T {
  if (!value) return fallback;
  try { return JSON.parse(value) as T; } catch { return fallback; }
}

function getRecentProjects(): ProjectSnapshot[] {
  return safeJson<ProjectSnapshot[]>(localStorage.getItem(PROJECT_STORAGE_KEY), []);
}

function setRecentProjects(items: ProjectSnapshot[]) {
  localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(items.slice(0, 10)));
  refreshRecentProjects();
}

function refreshRecentProjects() {
  const select = document.querySelector<HTMLSelectElement>('#recentProjects');
  if (!select) return;
  const items = getRecentProjects();
  select.innerHTML = '<option value="">Recent projects...</option>';
  items.forEach((item, index) => {
    const opt = document.createElement('option');
    opt.value = String(index);
    opt.textContent = `${item.name} — ${new Date(item.updatedAt).toLocaleString()}`;
    select.appendChild(opt);
  });
}

function getCurrentProjectSnapshot(name = currentProjectName): ProjectSnapshot {
  const settings = getAppearanceSettings();
  return {
    version: 4,
    name,
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    source: {
      kind: currentSourceKind,
      fileName: currentSourceFileName,
      rows: rawParsedData.map(r => ({ ...r })),
      sheetName: currentSheetName,
    },
    mapping: {
      xKey: (document.getElementById('colX') as HTMLSelectElement | null)?.value || '',
      yKey: (document.getElementById('colY') as HTMLSelectElement | null)?.value || '',
      zKey: (document.getElementById('colZ') as HTMLSelectElement | null)?.value || '',
      groupKey: (document.getElementById('colT') as HTMLSelectElement | null)?.value || '',
    },
    chart: {
      type: (document.getElementById('chartTypeSelect') as HTMLSelectElement | null)?.value || 'bar',
      palette: (document.getElementById('paletteSelect') as HTMLSelectElement | null)?.value || 'Viridis',
      customPalette: (document.getElementById('customHexInput') as HTMLInputElement | null)?.value || '',
      categoryLabels: [...currentDisplayLabels],
      radialPresentationLabels: [...radialPresentationLabels],
      title: (document.querySelector('.chart-title') as HTMLElement | null)?.textContent?.trim() || '',
      subtitle: (document.querySelector('.chart-subtitle') as HTMLElement | null)?.textContent?.trim() || '',
      metricKey: getChartMetricKey(),
    },
    appearance: { ...settings },
    dashboardItems: dashboardItems.map(item => ({ ...item, rows: item.rows.map(r => ({ ...r })) })),
    dashboardFilterValue,
  } as ProjectSnapshot & { dashboardItems?: DashboardItem[]; dashboardFilterValue?: string | null };
}

function applyAppearanceSnapshot(appearance: Record<string, unknown>) {
  const ids = [
    'titleSize', 'labelSize', 'plotHeight', 'plotOpacity',
    'showDataLabels', 'showLegend', 'showGrid', 'reverseColors',
    'xAxisTitle', 'yAxisTitle', 'xMin', 'xMax', 'yMin', 'yMax'
  ];
  ids.forEach(id => {
    const el = document.querySelector<HTMLInputElement>(`#${id}`);
    if (!el || !(id in appearance)) return;
    const value = appearance[id];
    if (el.type === 'checkbox') el.checked = Boolean(value);
    else el.value = value == null ? '' : String(value);
  });
}

function applyProjectSnapshot(snapshot: ProjectSnapshot) {
  currentProjectName = snapshot.name || 'Untitled';
  currentSourceFileName = snapshot.source?.fileName || '';
  currentSourceKind = snapshot.source?.kind || '';
  currentSheetName = snapshot.source?.sheetName || '';

  rawParsedData = Array.isArray(snapshot.source?.rows) ? snapshot.source.rows.map(r => ({ ...r })) : [];
  availableColumns = rawParsedData.length ? Object.keys(rawParsedData[0]) : [];

  const chartTypeEl = document.getElementById('chartTypeSelect') as HTMLSelectElement | null;
  const paletteEl = document.getElementById('paletteSelect') as HTMLSelectElement | null;
  const customHexEl = document.getElementById('customHexInput') as HTMLInputElement | null;
  const titleEl = document.querySelector('.chart-title') as HTMLElement | null;
  const subtitleEl = document.querySelector('.chart-subtitle') as HTMLElement | null;

  if (chartTypeEl) chartTypeEl.value = snapshot.chart?.type || chartTypeEl.value;
  if (paletteEl) paletteEl.value = snapshot.chart?.palette || paletteEl.value;
  if (customHexEl) customHexEl.value = snapshot.chart?.customPalette || customHexEl.value;
  const customGroup = document.getElementById('customPaletteGroup') as HTMLElement | null;
  if (customGroup) customGroup.style.display = paletteEl?.value === 'custom' ? 'flex' : 'none';
  if (titleEl) titleEl.textContent = snapshot.chart?.title || '';
  if (subtitleEl) subtitleEl.textContent = snapshot.chart?.subtitle || '';

  applyAppearanceSnapshot(snapshot.appearance || {});

  if (rawParsedData.length) {
    populateMappingControls();
    const setSelect = (id: string, value: string) => {
      const el = document.getElementById(id) as HTMLSelectElement | null;
      if (el && value && Array.from(el.options).some(o => o.value === value)) el.value = value;
    };
    setSelect('colX', snapshot.mapping?.xKey || '');
    setSelect('colY', snapshot.mapping?.yKey || '');
    setSelect('colZ', snapshot.mapping?.zKey || '');
    setSelect('colT', snapshot.mapping?.groupKey || '');

    const numericColumns = availableColumns.filter(isNumericColumn);
    const metricEl = document.getElementById('chartMetric') as HTMLSelectElement | null;
    if (metricEl) {
      metricEl.innerHTML = '';
      numericColumns.forEach(col => metricEl.add(new Option(col, col)));
      const savedMetric = snapshot.chart?.metricKey || chooseBestNumericMetric(numericColumns);
      chartMetricKey = savedMetric || '';
      if (savedMetric && Array.from(metricEl.options).some(o => o.value === savedMetric)) {
        metricEl.value = savedMetric;
      }
    }
    markImportedDatasetActive();
    updateValidationAndPreview();
  } else {
    document.getElementById('columnMappingSection')?.setAttribute('style', 'display:none;');
    document.getElementById('dataPreviewSection')?.setAttribute('style', 'display:none;');
  }

  const xKey = (document.getElementById('colX') as HTMLSelectElement | null)?.value || '';
  currentLabels = rawParsedData.length
    ? rawParsedData.map(row => String(row[xKey] ?? ''))
    : (Array.isArray(snapshot.chart?.categoryLabels) ? [...snapshot.chart.categoryLabels] : []);
  currentDisplayLabels = Array.isArray(snapshot.chart?.categoryLabels) && snapshot.chart.categoryLabels.length
    ? [...snapshot.chart.categoryLabels]
    : [...currentLabels];
  radialPresentationLabels = Array.isArray(snapshot.chart?.radialPresentationLabels) &&
    snapshot.chart.radialPresentationLabels.length === currentDisplayLabels.length
    ? [...snapshot.chart.radialPresentationLabels]
    : [...currentDisplayLabels];
  currentValues = getMappedRows().map(row => row.y);
  currentLabelName = getChartMetricKey() || 'Value';

  const projectExtras = snapshot as ProjectSnapshot & { dashboardItems?: DashboardItem[]; dashboardFilterValue?: string | null };
  dashboardItems = Array.isArray(projectExtras.dashboardItems) ? projectExtras.dashboardItems : [];
  dashboardFilterValue = projectExtras.dashboardFilterValue || null;

  syncCategoryLabelEditor();
  updateStatistics();
  enableLargeDatasetMode();
  updateDashboardFilterStatus();
  applyAppearanceToPage(getAppearanceSettings());
  updateRadialPresentationVisibility();
  updateRadialDetailOptions();
  renderVisualization();
  renderDashboard();
  updateDashboardFilterStatus();
  projectStatus(`Project: ${currentProjectName}`);
}

function downloadJson(filename: string, payload: unknown) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function openJsonFile(input: HTMLInputElement, handler: (value: any) => void) {
  input.value = '';
  input.onchange = async () => {
    const file = input.files?.[0];
    if (!file) return;
    try {
      const value = JSON.parse(await file.text());
      handler(value);
    } catch (err) {
      console.error(err);
      projectStatus('Could not read the selected project/template file.');
    }
  };
  input.click();
}

function saveProjectToBrowser() {
  const snapshot = getCurrentProjectSnapshot();
  const items = getRecentProjects().filter(p => p.name !== snapshot.name);
  setRecentProjects([snapshot, ...items]);
  localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(snapshot));
  projectStatus(`Saved: ${snapshot.name}`);
}

function saveProjectToFile() {
  const name = currentProjectName === 'Untitled' ? 'VizNova-Project' : currentProjectName;
  const snapshot = getCurrentProjectSnapshot(name);
  currentProjectName = name;
  downloadJson(`${name.replace(/[^a-z0-9_-]+/gi, '_')}.aetherviz`, snapshot);
  saveProjectToBrowser();
}

function saveTemplateToBrowser() {
  const template: TemplateSnapshot = {
    version: 1,
    name: currentProjectName === 'Untitled' ? 'My Template' : currentProjectName,
    createdAt: new Date().toISOString(),
    chart: getCurrentProjectSnapshot().chart,
    appearance: getCurrentProjectSnapshot().appearance,
  };
  const templates = safeJson<TemplateSnapshot[]>(localStorage.getItem(TEMPLATE_STORAGE_KEY), []);
  localStorage.setItem(TEMPLATE_STORAGE_KEY, JSON.stringify([template, ...templates.filter(t => t.name !== template.name)].slice(0, 20)));
  downloadJson(`${template.name.replace(/[^a-z0-9_-]+/gi, '_')}.aethertemplate`, template);
  projectStatus(`Template saved: ${template.name}`);
}

function loadTemplateSnapshot(template: TemplateSnapshot) {
  const chartTypeEl = document.getElementById('chartTypeSelect') as HTMLSelectElement | null;
  const paletteEl = document.getElementById('paletteSelect') as HTMLSelectElement | null;
  const customHexEl = document.getElementById('customHexInput') as HTMLInputElement | null;
  const titleEl = document.querySelector('.chart-title') as HTMLElement | null;
  const subtitleEl = document.querySelector('.chart-subtitle') as HTMLElement | null;

  if (chartTypeEl) chartTypeEl.value = template.chart.type || chartTypeEl.value;
  if (paletteEl) paletteEl.value = template.chart.palette || paletteEl.value;
  if (customHexEl) customHexEl.value = template.chart.customPalette || customHexEl.value;
  const customGroup = document.getElementById('customPaletteGroup') as HTMLElement | null;
  if (customGroup) customGroup.style.display = paletteEl?.value === 'custom' ? 'flex' : 'none';
  if (titleEl) titleEl.textContent = template.chart.title || '';
  if (subtitleEl) subtitleEl.textContent = template.chart.subtitle || '';

  applyAppearanceSnapshot(template.appearance || {});
  currentDisplayLabels = Array.isArray(template.chart.categoryLabels) ? [...template.chart.categoryLabels] : [...currentLabels];
  syncCategoryLabelEditor();
  updateRadialPresentationVisibility();
  updateRadialDetailOptions();
  applyAppearanceToPage(getAppearanceSettings());
  renderVisualization();
  projectStatus(`Template loaded: ${template.name}`);
}


type DashboardItem = {
  id: string;
  title: string;
  chartType: string;
  palette: string;
  customPalette: string;
  rows: Record<string, unknown>[];
  mapping: { xKey: string; yKey: string; zKey: string; groupKey: string };
  metricKey: string;
  labels: string[];
  appearance: Record<string, unknown>;
};

let dashboardItems: DashboardItem[] = [];
let dashboardCounter = 0;
let dashboardFilterValue: string | null = null;

function filteredDashboardRows(item: DashboardItem) {
  if (!dashboardFilterValue) return item.rows;
  const key = item.mapping.xKey;
  return item.rows.filter(row => String(row[key] ?? '') === dashboardFilterValue);
}

function updateDashboardFilterStatus() {
  const el = document.querySelector<HTMLElement>('#dashboardFilterStatus');
  if (el) el.textContent = dashboardFilterValue
    ? `Active filter: ${dashboardFilterValue}`
    : 'No dashboard filter active.';
}

function applyDashboardFilter(value: string) {
  dashboardFilterValue = value;
  updateDashboardFilterStatus();
  renderDashboard();
}

function cloneDashboardItem(): DashboardItem {
  const settings = getAppearanceSettings();
  const getValue = (id: string) =>
    (document.getElementById(id) as HTMLInputElement | HTMLSelectElement | null)?.value || '';
  const title =
    (document.querySelector('.chart-title') as HTMLElement | null)?.textContent?.trim() ||
    `Chart ${dashboardCounter + 1}`;

  return {
    id: `chart-${Date.now()}-${dashboardCounter++}`,
    title,
    chartType: getValue('chartTypeSelect'),
    palette: getValue('paletteSelect'),
    customPalette: getValue('customHexInput'),
    rows: rawParsedData.map(r => ({ ...r })),
    mapping: {
      xKey: getValue('colX'),
      yKey: getValue('colY'),
      zKey: getValue('colZ'),
      groupKey: getValue('colT'),
    },
    metricKey: getChartMetricKey(),
    labels: [...currentDisplayLabels],
    appearance: { ...settings },
  };
}

function dashboardColumnCount() {
  return Math.max(1, Math.min(3, Number(document.querySelector<HTMLSelectElement>('#dashboardColumns')?.value || 2)));
}

function renderDashboard() {
  const canvas = document.querySelector<HTMLElement>('#dashboardCanvas');
  const list = document.querySelector<HTMLElement>('#dashboardList');
  const empty = document.querySelector<HTMLElement>('#dashboardEmpty');
  if (!canvas) {
    console.warn('VizNova: #dashboardCanvas is missing; dashboard cannot render.');
    return;
  }

  canvas.style.gridTemplateColumns = `repeat(${dashboardColumnCount()}, minmax(0, 1fr))`;
  canvas.innerHTML = '';

  if (!dashboardItems.length) {
    canvas.innerHTML = '<div class="dashboard-empty">Dashboard is empty. Click <strong>Add to Dashboard</strong> to add the current chart.</div>';
  }

  dashboardItems.forEach((item, index) => {
    const card = document.createElement('div');
    card.className = 'dashboard-card';
    card.dataset.dashboardId = item.id;

    const header = document.createElement('div');
    header.className = 'dashboard-card-header';
    header.innerHTML = `<strong>${escapeHtml(item.title)}</strong>`;

    const actions = document.createElement('div');
    actions.className = 'dashboard-card-actions';

    const remove = document.createElement('button');
    remove.className = 'mini-action';
    remove.textContent = 'Remove';
    remove.addEventListener('click', () => {
      dashboardItems = dashboardItems.filter(x => x.id !== item.id);
      renderDashboard();
    });

    const moveUp = document.createElement('button');
    moveUp.className = 'mini-action';
    moveUp.textContent = '↑';
    moveUp.title = 'Move chart up';
    moveUp.addEventListener('click', () => {
      if (index <= 0) return;
      [dashboardItems[index - 1], dashboardItems[index]] = [dashboardItems[index], dashboardItems[index - 1]];
      renderDashboard();
    });

    const moveDown = document.createElement('button');
    moveDown.className = 'mini-action';
    moveDown.textContent = '↓';
    moveDown.title = 'Move chart down';
    moveDown.addEventListener('click', () => {
      if (index >= dashboardItems.length - 1) return;
      [dashboardItems[index + 1], dashboardItems[index]] = [dashboardItems[index], dashboardItems[index + 1]];
      renderDashboard();
    });

    actions.append(moveUp, moveDown, remove);
    header.appendChild(actions);
    card.appendChild(header);

    const plot = document.createElement('div');
    plot.className = 'dashboard-plot';
    plot.id = `dashboard-plot-${item.id}`;
    card.appendChild(plot);
    canvas.appendChild(card);

    renderDashboardItem(plot, item);
  });

  if (list) {
    list.innerHTML = dashboardItems.length
      ? dashboardItems.map((item, i) => `<div class="dashboard-list-row"><span>${i + 1}. ${escapeHtml(item.title)}</span><span>${escapeHtml(item.chartType)}</span></div>`).join('')
      : '';
  }

  if (empty) empty.style.display = dashboardItems.length ? 'none' : 'block';
}

function escapeHtml(value: string) {
  return value.replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch] || ch));
}

function renderDashboardItem(container: HTMLElement, item: DashboardItem) {
  const itemRows = filteredDashboardRows(item);
  const xKey = item.mapping.xKey;
  const yKey = item.metricKey || item.mapping.yKey;
  const labels = itemRows.map(r => String(r[xKey] ?? ''));

  const values = itemRows.map(r => Number(r[yKey])).filter(Number.isFinite);
  const traceColor = item.palette === 'custom' && item.customPalette
    ? item.customPalette.split(',').map(s => s.trim()).filter(Boolean)[0]
    : undefined;

  const common = {
    paper_bgcolor: 'rgba(0,0,0,0)',
    plot_bgcolor: 'rgba(0,0,0,0)',
    font: { color: '#e5e7eb', size: 11 },
    margin: { l: 55, r: 20, t: 55, b: 50 },
    height: 360,
    showlegend: Boolean(item.appearance.showLegend),
  };

  let traces: any[] = [];
  if (item.chartType === 'pie' || item.chartType === 'doughnut') {
    traces = [{
      type: 'pie',
      labels,
      values,
      hole: item.chartType === 'doughnut' ? 0.55 : 0,
      textinfo: item.appearance.showDataLabels ? 'label+percent' : 'none',
      marker: traceColor ? { colors: [traceColor] } : undefined,
    }];
  } else if (item.chartType === 'scatter' || item.chartType === 'bubble') {
    traces = [{
      type: 'scatter',
      mode: item.appearance.showDataLabels ? 'lines+markers+text' : 'lines+markers',
      x: labels,
      y: values,
      text: item.appearance.showDataLabels ? labels : undefined,
      textposition: 'top center',
      name: 'Series',
      marker: traceColor ? { color: traceColor } : undefined,
    }];
  } else if (item.chartType === 'horizontalBar') {
    traces = [{
      type: 'bar',
      orientation: 'h',
      x: values,
      y: labels,
      name: 'Series',
      text: item.appearance.showDataLabels ? values.map(String) : undefined,
      textposition: 'auto',
      marker: traceColor ? { color: traceColor } : undefined,
    }];
  } else {
    traces = [{
      type: 'bar',
      x: labels,
      y: values,
      name: 'Series',
      text: item.appearance.showDataLabels ? values.map(String) : undefined,
      textposition: 'auto',
      marker: traceColor ? { color: traceColor } : undefined,
    }];
  }

  const layout: any = {
    ...common,
    title: { text: item.title, font: { size: Number(item.appearance.titleSize) || 18 } },
    xaxis: {
      title: item.appearance.xTitle || undefined,
      showgrid: Boolean(item.appearance.showGrid),
    },
    yaxis: {
      title: item.appearance.yTitle || undefined,
      showgrid: Boolean(item.appearance.showGrid),
    },
  };

  Plotly.newPlot(container, traces, layout, {
    responsive: true,
    displaylogo: false,
    modeBarButtonsToRemove: ['lasso2d', 'select2d'],
  }).then(() => {
    (container as any).on('plotly_click', (event: any) => {
      const point = event?.points?.[0];
      if (!point) return;
      const clicked = point.label ?? point.x ?? point.y;
      if (clicked !== undefined && clicked !== null) applyDashboardFilter(String(clicked));
    });
  });
}

function updateStatistics() {
  const rowCount = rawParsedData.length;
  let numeric = 0;
  let missing = 0;
  const ys: number[] = [];

  rawParsedData.forEach(row => {
    Object.values(row).forEach(value => {
      if (value === '' || value === null || value === undefined) missing++;
      else if (typeof value === 'number' && Number.isFinite(value)) numeric++;
      else if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value))) numeric++;
    });
    const y = Number(row[getChartMetricKey()]);
    if (Number.isFinite(y)) ys.push(y);
  });

  ys.sort((a, b) => a - b);
  const mean = ys.length ? ys.reduce((a, b) => a + b, 0) / ys.length : NaN;
  const median = ys.length ? (ys.length % 2 ? ys[(ys.length - 1) / 2] : (ys[ys.length / 2 - 1] + ys[ys.length / 2]) / 2) : NaN;

  const set = (id: string, value: string) => {
    const el = document.querySelector<HTMLElement>(`#${id}`);
    if (el) el.textContent = value;
  };
  set('statRows', String(rowCount));
  set('statNumeric', String(numeric));
  set('statMissing', String(missing));
  set('statMean', Number.isFinite(mean) ? mean.toFixed(2) : '—');
  set('statMedian', Number.isFinite(median) ? median.toFixed(2) : '—');
  set('statRange', ys.length ? `${ys[0].toFixed(2)} / ${ys[ys.length - 1].toFixed(2)}` : '—');
}

function enableLargeDatasetMode() {
  // Keep raw data intact, but cap category labels shown in heavy dashboard cards.
  // The main visualization continues to use the complete dataset.
  const large = rawParsedData.length > 5000;
  document.body.classList.toggle('large-dataset', large);
  if (large) {
    projectStatus(`Large dataset detected: ${rawParsedData.length.toLocaleString()} rows. Rendering controls remain enabled.`);
  }
}

function setupPhase5AdvancedApplication() {
  updateStatistics();
  enableLargeDatasetMode();
  renderDashboard();
  updateDashboardFilterStatus();

  const addDashboardButton = document.querySelector<HTMLButtonElement>('#addDashboardChart');
  const clearDashboardButton = document.querySelector<HTMLButtonElement>('#clearDashboard');

  if (!addDashboardButton || !clearDashboardButton) {
    console.error('VizNova: dashboard controls were not found in the DOM.');
    projectStatus('Dashboard controls are not connected.');
  }

  addDashboardButton?.addEventListener('click', (event) => {
    event.preventDefault();
    if (!rawParsedData.length) {
      projectStatus('Import data before adding a dashboard chart.');
      return;
    }
    dashboardItems.push(cloneDashboardItem());
    renderDashboard();
    projectStatus(`Dashboard: ${dashboardItems.length} chart${dashboardItems.length === 1 ? '' : 's'}`);
  });

  clearDashboardButton?.addEventListener('click', (event) => {
    event.preventDefault();
    dashboardItems = [];
    dashboardFilterValue = null;
    renderDashboard();
    updateDashboardFilterStatus();
    projectStatus('Dashboard cleared.');
  });

  document.querySelector<HTMLSelectElement>('#dashboardColumns')?.addEventListener('change', renderDashboard);

  document.querySelector<HTMLButtonElement>('#clearDashboardFilter')?.addEventListener('click', (event) => {
    event.preventDefault();
    dashboardFilterValue = null;
    updateDashboardFilterStatus();
    renderDashboard();
  });

  document.querySelector<HTMLInputElement>('#accessibilityMode')?.addEventListener('change', (event) => {
    document.body.classList.toggle('accessibility-mode', (event.target as HTMLInputElement).checked);
  });

  // Keep statistics current when mappings or data change.
  ['colX', 'colY', 'colZ', 'colT'].forEach(id => {
    document.getElementById(id)?.addEventListener('change', updateStatistics);
    document.getElementById(id)?.addEventListener('input', updateStatistics);
  });
}

function updateRadialPresentationVisibility(): void {
  const type = (document.getElementById('chartTypeSelect') as HTMLSelectElement | null)?.value;
  const controls = document.getElementById('radialPresentationControls') as HTMLElement | null;
  if (controls) controls.style.display = type === 'annotatedRadial' ? 'block' : 'none';
}

function updateRadialDetailOptions(): void {
  const select = document.getElementById('radialDetailMetric') as HTMLSelectElement | null;
  if (!select) return;
  const previous = select.value;
  select.innerHTML = '<option value="">None</option>';
  availableColumns.forEach(key => {
    const option = document.createElement('option');
    option.value = key;
    option.textContent = key;
    select.appendChild(option);
  });
  if (availableColumns.includes(previous)) select.value = previous;
}

function formatRadialMetric(value: number, metricKey: string): string {
  if (!Number.isFinite(value)) return '';
  if (metricLooksLikeRate(metricKey)) return `${value.toFixed(2)}%`;
  return Number.isInteger(value) ? value.toLocaleString() : value.toFixed(2);
}

function radialDetailText(row: any, metricKey: string): string {
  if (!metricKey) return '';
  const source = row?.raw ?? row;
  const value = source?.[metricKey];
  if (value === undefined || value === null || value === '') return '';
  const numeric = Number(value);
  return Number.isFinite(numeric) ? `${numeric.toLocaleString()} ${metricKey}` : `${String(value)} ${metricKey}`;
}

function buildRadialCallouts(
  labels: string[],
  values: number[],
  metricKey: string,
  rowsForChart: any[],
  showCallouts: boolean,
  showConnectors: boolean,
  labelSize: number
): any[] {
  if (!showCallouts || !labels.length) return [];

  const annotations: any[] = [];
  const n = labels.length;
  const radius = 0.43;
  const labelRadius = 0.48;
  const displayLabels = radialPresentationLabels.length === labels.length
    ? radialPresentationLabels
    : labels;

  // IMPORTANT: barpolar uses rotation: 90° and direction: clockwise.
  // Therefore theta=0 is at the top, theta increases clockwise, and the
  // paper-coordinate conversion must use x = sin(theta), y = cos(theta).
  // This keeps each external callout aligned with the same bar index.
  labels.forEach((label, i) => {
    const theta = (2 * Math.PI * i / n);
    const px = 0.5 + radius * Math.sin(theta);
    const py = 0.5 + radius * Math.cos(theta);
    const lx = 0.5 + labelRadius * Math.sin(theta);
    const ly = 0.5 + labelRadius * Math.cos(theta);

    const anchor = lx < 0.45 ? 'right' : lx > 0.55 ? 'left' : 'center';
    const detail = radialDetailText(rowsForChart[i], (document.getElementById('radialDetailMetric') as HTMLSelectElement | null)?.value || '');
    const title = escapePlotlyLabel(String(displayLabels[i] ?? label));
    const valueText = `<b>${formatRadialMetric(values[i], metricKey)}</b>`;
    const detailText = detail ? `<br><span style="font-size:${Math.max(9, labelSize - 1)}px">${escapePlotlyLabel(detail)}</span>` : '';

    annotations.push({
      x: px,
      y: py,
      xref: 'paper',
      yref: 'paper',
      text: `${title}<br>${valueText}${detailText}`,
      showarrow: showConnectors,
      arrowhead: 2,
      arrowsize: 0.7,
      arrowwidth: 1.5,
      arrowcolor: colorAt(getAppearanceSettings().colors, i, '#06b6d4'),
      ax: (lx - px) * 650,
      ay: (py - ly) * 650,
      xanchor: anchor,
      yanchor: 'middle',
      align: anchor === 'center' ? 'center' : anchor,
      font: {
        family: 'Inter, sans-serif',
        size: Math.max(10, Math.min(labelSize, 12)),
        color: '#f3f4f6'
      },
      bgcolor: 'rgba(15,23,42,0.84)',
      bordercolor: colorAt(getAppearanceSettings().colors, i, '#06b6d4'),
      borderwidth: 1,
      borderpad: 6,
      opacity: 0.98
    });
  });

  return annotations;
}

function addRadialCenterAnnotation(layout: any, centerTitle: string): void {
  layout.shapes = layout.shapes || [];
  layout.shapes.push({
    type: 'circle',
    xref: 'paper',
    yref: 'paper',
    x0: 0.43,
    x1: 0.57,
    y0: 0.43,
    y1: 0.57,
    fillcolor: 'rgba(255,255,255,0.97)',
    line: { color: '#4c1d95', width: 2 },
    layer: 'above'
  });
  layout.annotations = layout.annotations || [];
  layout.annotations.push({
    x: 0.5,
    y: 0.5,
    xref: 'paper',
    yref: 'paper',
    text: `<b>${escapePlotlyLabel(centerTitle || 'Girls Scholarship')}</b>`,
    showarrow: false,
    font: { family: 'Inter, sans-serif', size: 14, color: '#111827' },
    bgcolor: 'rgba(255,255,255,0.97)',
    bordercolor: '#4c1d95',
    borderwidth: 2,
    borderpad: 10,
    opacity: 0.98
  });
}


function getColumnDisplayName(key: string): string {
  const option = document.querySelector<HTMLOptionElement>(`#chartMetric option[value="${CSS.escape(key)}"]`);
  return option?.textContent?.trim() || key;
}

function isDefaultWorkspaceTitle(value: string): boolean {
  const v = value.trim().toLowerCase();
  return !v || v === 'universal master analytics workspace' || v === 'comparison';
}

function syncRadialPresentationTitle(): void {
  const chartTypeEl = document.getElementById('chartTypeSelect') as HTMLSelectElement | null;
  const titleEl = document.querySelector<HTMLElement>('.chart-title');
  const subtitleEl = document.querySelector<HTMLElement>('.chart-subtitle');
  const centerEl = document.getElementById('radialCenterTitle') as HTMLInputElement | null;
  const active = chartTypeEl?.value === 'annotatedRadial';

  document.body.classList.toggle('radial-presentation-active', active);
  if (!active) return;

  const metricKey = getChartMetricKey();
  const metricName = getColumnDisplayName(metricKey);
  const looksFemaleRate = metricLooksLikeRate(metricKey) &&
    /female|girls?|women/i.test(metricName);

  const title = looksFemaleRate
    ? 'Percentage of Girls Getting Scholarship (2024–25)'
    : `${metricName} — Faculty / Department Wise`;

  // Presentation radial owns its presentation title so it cannot fall back
  // to the generic workspace title.
  if (titleEl) titleEl.textContent = title;
  if (subtitleEl) subtitleEl.textContent = 'Faculty / Department Wise';

  if (centerEl && (!centerEl.value.trim() || isDefaultWorkspaceTitle(centerEl.value))) {
    centerEl.value = looksFemaleRate ? 'Girls Scholarship' : 'Comparison';
  }
}

function configureRadialDetailMetric(): void {
  const select = document.getElementById('radialDetailMetric') as HTMLSelectElement | null;
  if (!select) return;

  const previous = select.value;
  const currentMetric = getChartMetricKey();
  select.innerHTML = '<option value="">None</option>';

  const options = Array.from(document.querySelectorAll<HTMLOptionElement>('#chartMetric option'));
  options.forEach(opt => {
    if (!opt.value || opt.value === currentMetric) return;
    const values = rawParsedData.map(r => Number(r[opt.value])).filter(Number.isFinite);
    if (!values.length) return;
    const option = document.createElement('option');
    option.value = opt.value;
    option.textContent = opt.textContent || opt.value;
    select.appendChild(option);
  });

  // Keep an existing numeric choice only; never allow the category column here.
  if (Array.from(select.options).some(o => o.value === previous)) {
    select.value = previous;
  } else {
    select.value = '';
  }
}

function syncMetricBasedTitle(): void {
  const chartType = (document.getElementById('chartTypeSelect') as HTMLSelectElement | null)?.value || '';
  const titleEl = document.querySelector<HTMLElement>('.chart-title');
  const subtitleEl = document.querySelector<HTMLElement>('.chart-subtitle');
  if (!titleEl || !subtitleEl) return;

  const currentTitle = titleEl.textContent?.trim() || '';
  const metricKey = getChartMetricKey();
  const metricName = getColumnDisplayName(metricKey) || metricKey;
  const isDefaultTitle = isDefaultWorkspaceTitle(currentTitle);

  if (!metricName || !isDefaultTitle) return;

  const normalized = metricName.replace(/\s+/g, ' ').trim();
  const lower = normalized.toLowerCase();
  const isGirlsScholarshipRate =
    /percentage|percent|%/.test(lower) &&
    /female|girls?|women/.test(lower) &&
    /scholarship/.test(lower);

  if (isGirlsScholarshipRate) {
    titleEl.textContent = 'Percentage of Girls Getting Scholarship (2024–25)';
    subtitleEl.textContent = 'Faculty / Department Wise';
  } else if (chartType) {
    titleEl.textContent = `${normalized} — Category Wise`;
    subtitleEl.textContent = 'Faculty / Department Wise';
  }
}

function renderVisualization() {
  syncMetricBasedTitle();
  const chartType = (document.getElementById('chartTypeSelect') as HTMLSelectElement).value;
  if (chartType === 'annotatedRadial') {
    configureRadialDetailMetric();
    syncRadialPresentationTitle();
  } else {
    document.body.classList.remove('radial-presentation-active');
  }
  const metricKey = getChartMetricKey();
  const settings = getAppearanceSettings();
  const radialStatus = document.getElementById('radialMetricStatus');
  if (radialStatus && chartType === 'annotatedRadial') {
    const metricName = getColumnDisplayName(metricKey);
    radialStatus.textContent = `Using metric: ${metricName || 'None'} · ${rawParsedData.length ? rawParsedData.length.toLocaleString() + ' imported rows' : 'built-in data'}`;
  }
  const rateWarning = (chartType === 'pie' || chartType === 'doughnut') && metricLooksLikeRate(metricKey);
  const status = document.getElementById('mappingHint');
  if (status && rateWarning) {
    status.innerHTML = '⚠ The selected metric looks like a percentage/rate. Pie/doughnut shows composition; use Bar, Radar, or Polar for faculty-wise percentage rates.';
  } else if (status) { updateMappingHint(); }
  const colorInput = settings.colors;
  applyAppearanceToPage(settings);
  const container = document.getElementById('visualizationCanvas') as HTMLElement;

  let trace: any = {};
  let layout: any = {
    paper_bgcolor: 'rgba(17, 24, 39, 0)',
    plot_bgcolor: 'rgba(17, 24, 39, 0)',
    font: { family: 'Inter, sans-serif', color: '#f3f4f6', size: 12 },
    margin: { l: 70, r: 70, t: 70, b: 90 },
    autosize: true,
    xaxis: {
      title: settings.xTitle || undefined,
      gridcolor: settings.showGrid ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0,0,0,0)',
      zerolinecolor: settings.showGrid ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0,0,0,0)',
      range: settings.xMin !== undefined && settings.xMax !== undefined ? [settings.xMin, settings.xMax] : undefined,
      tickfont: { size: settings.labelSize }
    },
    yaxis: {
      title: settings.yTitle || currentLabelName,
      gridcolor: settings.showGrid ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0,0,0,0)',
      zerolinecolor: settings.showGrid ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0,0,0,0)',
      range: settings.yMin !== undefined && settings.yMax !== undefined ? [settings.yMin, settings.yMax] : undefined,
      tickfont: { size: settings.labelSize }
    },
    legend: { font: { color: '#9ca3af', size: settings.labelSize }, orientation: 'h' },
    showlegend: settings.showLegend
  };

  // Configure Traces based on exhaustive chart selections
  const groupKey = (document.getElementById('colT') as HTMLSelectElement)?.value;
  if (groupKey && ['bar', 'stackedBar', 'scatter', 'bubble', 'area'].includes(chartType)) {
    trace = buildGroupedTraces(chartType, colorInput);
    if (chartType === 'stackedBar') layout.barmode = 'stack';
  } else if (chartType === 'bar' || chartType === 'horizontalBar' || chartType === 'waterfall') {
    trace = {
      x: chartType === 'horizontalBar' ? currentValues : currentLabels,
      y: chartType === 'horizontalBar' ? currentLabels : currentValues,
      type: 'bar',
      orientation: chartType === 'horizontalBar' ? 'h' : 'v',
      marker: { color: Array.isArray(colorInput) ? colorInput : '#06b6d4', opacity: settings.opacity },
      text: settings.showDataLabels ? makeDataLabels(currentValues) : undefined,
      textposition: settings.showDataLabels ? 'auto' : undefined,
      textfont: { size: settings.labelSize }
    };
  } else if (chartType === 'stackedBar') {
    trace = {
      x: currentLabels,
      y: currentValues,
      type: 'bar',
      name: currentLabelName,
      marker: { color: Array.isArray(colorInput) ? colorInput[0] : '#06b6d4' }
    };
    layout.barmode = 'stack';
  } else if (chartType === 'pie' || chartType === 'doughnut') {
    trace = {
      labels: currentLabels,
      values: currentValues,
      type: 'pie',
      hole: chartType === 'doughnut' ? 0.5 : 0,
      marker: { colors: Array.isArray(colorInput) ? colorInput : undefined, opacity: settings.opacity },
      textinfo: settings.showDataLabels ? 'label+percent' : 'percent',
      textfont: { size: settings.labelSize }
    };
  } else if (chartType === 'heatmap' || chartType === 'contour') {
    const xKey = (document.getElementById('colX') as HTMLSelectElement)?.value;
    const yKey = (document.getElementById('colY') as HTMLSelectElement)?.value;
    const zKey = (document.getElementById('colZ') as HTMLSelectElement)?.value;
    const matrixRows = rawParsedData.filter(row => xKey && yKey && zKey && row[xKey] !== '' && row[yKey] !== '' && Number.isFinite(Number(row[zKey])));
    const xs = Array.from(new Set(matrixRows.map(row => String(row[xKey]))));
    const ys = Array.from(new Set(matrixRows.map(row => String(row[yKey]))));
    const z = ys.map(y => xs.map(x => {
      const matches = matrixRows.filter(row => String(row[xKey]) === x && String(row[yKey]) === y).map(row => Number(row[zKey]));
      return matches.length ? matches.reduce((a,b) => a+b, 0) / matches.length : null;
    }));
    trace = {
      z: xs.length && ys.length && zKey ? z : [currentValues],
      x: xs.length && zKey ? xs : currentLabels,
      y: ys.length && zKey ? ys : ['Metric'],
      type: chartType,
      colorscale: typeof colorInput === 'string' ? colorInput : 'Viridis'
    };
  } else if (chartType === 'boxplot' || chartType === 'violin') {
    trace = {
      y: currentValues,
      type: chartType === 'boxplot' ? 'box' : 'violin',
      name: currentLabelName,
      boxpoints: 'all',
      marker: { color: Array.isArray(colorInput) ? colorInput[0] : '#8b5cf6' }
    };
  } else if (chartType === 'scatter' || chartType === 'bubble') {
    trace = {
      x: currentLabels,
      y: currentValues,
      mode: 'markers',
      type: 'scatter',
      marker: {
        size: chartType === 'bubble' ? currentValues.map(v => Math.max(8, (v / Math.max(...currentValues)) * 30)) : 12,
        color: currentValues,
        colorscale: typeof colorInput === 'string' ? colorInput : 'Viridis',
        showscale: true,
        opacity: settings.opacity
      },
      text: settings.showDataLabels ? makeDataLabels(currentValues) : undefined,
      textposition: settings.showDataLabels ? 'top center' : undefined,
      textfont: { size: settings.labelSize }
    };
  } else if (chartType === 'annotatedRadial') {
    // Presentation radial must always use the currently selected imported metric,
    // never a stale built-in preset value.
    const metricSelect = document.getElementById('chartMetric') as HTMLSelectElement | null;
    const metricKeyForRadial = metricSelect?.value || getChartMetricKey();
    const xKeyForRadial = (document.getElementById('colX') as HTMLSelectElement | null)?.value || '';

    const mappedRows = rawParsedData.length && xKeyForRadial && metricKeyForRadial
      ? rawParsedData
          .map(row => ({
            x: row[xKeyForRadial],
            y: Number(row[metricKeyForRadial]),
            raw: row
          }))
          .filter(r =>
            r.x !== '' &&
            r.x !== null &&
            r.x !== undefined &&
            Number.isFinite(r.y)
          )
      : [];

    const labels = mappedRows.length
      ? mappedRows.map((r, i) => String(currentDisplayLabels[i] ?? r.x))
      : [...currentDisplayLabels];

    const values = mappedRows.length
      ? mappedRows.map(r => r.y)
      : [...currentValues];

    const centerTitle = getControlValue('radialCenterTitle', 'Girls Scholarship');
    const showCallouts = (document.getElementById('radialCallouts') as HTMLInputElement | null)?.checked ?? true;
    const showConnectors = (document.getElementById('radialConnectors') as HTMLInputElement | null)?.checked ?? true;
    const colors = Array.isArray(colorInput) ? colorInput : ['#3b82f6'];

    trace = {
      type: 'barpolar',
      r: values,
      theta: labels.map((_, i) => i * (360 / Math.max(labels.length, 1))),
      width: labels.map(() => Math.max(8, (360 / Math.max(labels.length, 1)) * 0.72)),
      marker: {
        color: values.map((_, i) => colorAt(colors, i, '#3b82f6')),
        opacity: settings.opacity,
        line: { color: 'rgba(255,255,255,0.82)', width: 1.5 }
      },
      text: settings.showDataLabels ? values.map(v => formatRadialMetric(v, metricKeyForRadial)) : undefined,
      textposition: settings.showDataLabels ? 'outside' : undefined,
      hovertext: labels.map((label, i) => `${label}<br>${formatRadialMetric(values[i], metricKeyForRadial)}`),
      hoverinfo: 'text'
    };

    const finiteValues = values.filter(Number.isFinite);
    const maxValue = Math.max(...finiteValues, 1);
    const rateMetric = metricLooksLikeRate(metricKeyForRadial);
    const radialMax = rateMetric ? 100 : maxValue * 1.18;

    layout.margin = { l: 145, r: 145, t: 48, b: 48 };
    layout.showlegend = false;
    layout.polar = {
      bgcolor: 'rgba(0,0,0,0)',
      radialaxis: {
        visible: settings.showGrid,
        range: [0, radialMax],
        tickfont: { size: Math.min(settings.labelSize, 11), color: '#d1d5db' },
        gridcolor: settings.showGrid ? 'rgba(156,163,175,0.22)' : 'rgba(0,0,0,0)',
        linecolor: 'rgba(156,163,175,0.28)',
        ticksuffix: rateMetric ? '%' : ''
      },
      angularaxis: {
        visible: false,
        direction: 'clockwise',
        rotation: 90,
        gridcolor: 'rgba(156,163,175,0.16)'
      }
    };

    layout.annotations = buildRadialCallouts(
      labels,
      values,
      metricKeyForRadial,
      mappedRows.length ? mappedRows : values.map((value, i) => ({ x: labels[i], y: value, raw: {} })),
      showCallouts,
      showConnectors,
      settings.labelSize
    );
    addRadialCenterAnnotation(layout, centerTitle);
  } else if (chartType === 'radar') {
    trace = {
      type: 'scatterpolar',
      r: currentValues,
      theta: currentLabels,
      fill: 'toself',
      opacity: settings.opacity,
      marker: { color: colorAt(colorInput, 0, '#ec4899'), opacity: settings.opacity },
      text: settings.showDataLabels ? makeDataLabels(currentValues) : undefined,
      textposition: settings.showDataLabels ? 'top center' : undefined
    };
    layout.polar = {
      radialaxis: { visible: true, range: [0, Math.max(...currentValues) * 1.2], gridcolor: 'rgba(255,255,255,0.1)' },
      bgcolor: 'rgba(0,0,0,0)'
    };
  } else if (chartType === 'polarArea') {
    trace = {
      r: currentValues,
      // Keep the original data categories as the angular positions.
      theta: currentLabels,
      type: 'barpolar',
      marker: { color: Array.isArray(colorInput) ? colorInput : '#3b82f6', opacity: settings.opacity },
      text: settings.showDataLabels ? makeDataLabels(currentValues) : undefined,
      textposition: settings.showDataLabels ? 'auto' : undefined
    };

    layout.polar = {
      bgcolor: 'rgba(0,0,0,0)',
      angularaxis: {
        tickmode: 'array',
        tickvals: currentLabels,
        ticktext: currentDisplayLabels.map(formatPolarLabel),
        tickfont: {
          family: 'Inter, sans-serif',
          color: '#f3f4f6',
          size: settings.labelSize
        },
        gridcolor: 'rgba(255,255,255,0.12)',
        linecolor: 'rgba(255,255,255,0.25)'
      },
      radialaxis: {
        visible: true,
        range: [0, 3],
        gridcolor: 'rgba(255,255,255,0.12)',
        linecolor: 'rgba(255,255,255,0.25)',
        tickfont: {
          family: 'Inter, sans-serif',
          color: '#f3f4f6',
          size: 11
        }
      }
    };
  } else if (chartType === 'funnel') {
    trace = {
      y: currentLabels,
      x: currentValues,
      type: 'funnel',
      marker: { color: Array.isArray(colorInput) ? colorInput : '#f59e0b' }
    };
  } else if (chartType === 'surface3d') {
    trace = {
      z: [
        currentValues,
        currentValues.map(v => v * 0.9),
        currentValues.map(v => v * 1.1),
        currentValues.map(v => v * 0.8)
      ],
      type: 'surface',
      colorscale: typeof colorInput === 'string' ? colorInput : 'Viridis'
    };
  } else {
    // Default Line / Area Flow
    trace = {
      x: currentLabels,
      y: currentValues,
      type: 'scatter',
      mode: 'lines+markers',
      fill: chartType === 'area' ? 'tozeroy' : 'none',
      line: { color: Array.isArray(colorInput) ? colorInput[0] : '#06b6d4', width: 3 },
      marker: { color: Array.isArray(colorInput) ? colorInput[0] : '#06b6d4' }
    };
  }

  layout.height = settings.plotHeight;
  layout.width = undefined;
  Plotly.newPlot(container, [trace], layout, { responsive: true, displaylogo: false })
    .then(() => {
      Plotly.Plots.resize(container);
    });
}

// Event Bindings
document.getElementById('chartTypeSelect')?.addEventListener('change', () => { updateMappingHint(); updateRadialPresentationVisibility(); updateRadialDetailOptions(); renderVisualization(); });
document.getElementById('paletteSelect')?.addEventListener('change', (e) => {
  const val = (e.target as HTMLSelectElement).value;
  document.getElementById('customPaletteGroup')!.style.display = val === 'custom' ? 'flex' : 'none';
  renderVisualization();
});
document.getElementById('customHexInput')?.addEventListener('input', renderVisualization);
document.getElementById('radialCenterTitle')?.addEventListener('input', renderVisualization);
document.getElementById('radialDetailMetric')?.addEventListener('change', renderVisualization);
document.getElementById('radialCallouts')?.addEventListener('change', renderVisualization);
document.getElementById('radialConnectors')?.addEventListener('change', renderVisualization);

document.getElementById('colX')?.addEventListener('change', processMappingAndRender);
document.getElementById('colY')?.addEventListener('change', () => {
  const y = (document.getElementById('colY') as HTMLSelectElement).value;
  const metric = document.getElementById('chartMetric') as HTMLSelectElement | null;
  if (metric) { chartMetricKey = y; metric.value = y; }
  processMappingAndRender();
});
document.getElementById('colZ')?.addEventListener('change', processMappingAndRender);
document.getElementById('chartMetric')?.addEventListener('change', () => {
  chartMetricKey = (document.getElementById('chartMetric') as HTMLSelectElement).value;
  processMappingAndRender();
});
document.getElementById('colT')?.addEventListener('change', processMappingAndRender);

document.getElementById('applyCategoryLabels')?.addEventListener('click', applyCategoryLabels);
document.getElementById('resetCategoryLabels')?.addEventListener('click', resetDisplayLabels);

['titleSize','labelSize','plotHeight','plotOpacity','xAxisTitle','yAxisTitle','xMin','xMax','yMin','yMax','showDataLabels','showLegend','showGrid','reverseColors'].forEach(id => {
  const el = document.getElementById(id);
  el?.addEventListener('input', renderVisualization);
  el?.addEventListener('change', renderVisualization);
});

document.getElementById('resetAppearance')?.addEventListener('click', () => {
  const defaults: Record<string,string|boolean> = { titleSize:'24', labelSize:'11', plotHeight:'650', plotOpacity:'0.9', xAxisTitle:'', yAxisTitle:'', xMin:'', xMax:'', yMin:'', yMax:'', showDataLabels:false, showLegend:true, showGrid:true, reverseColors:false };
  Object.entries(defaults).forEach(([id,value]) => {
    const el = document.getElementById(id) as HTMLInputElement | null;
    if (!el) return;
    if (el.type === 'checkbox') el.checked = Boolean(value); else el.value = String(value);
  });
  renderVisualization();
});

document.getElementById('datasetSelect')?.addEventListener('change', (e) => {
  const key = (e.target as HTMLSelectElement).value;
  if (key === '__imported__') return;
  const ds = defaultDatasets[key];
  currentLabels = ds.labels;
  currentDisplayLabels = [...currentLabels];
  currentValues = ds.data;
  currentLabelName = ds.label;
  rawParsedData = [];
  availableColumns = [];
  workbook = null;
  currentSourceFileName = '';
  currentSourceKind = 'preset';
  currentSheetName = '';
  document.getElementById('columnMappingSection')!.style.display = 'none';
  document.getElementById('dataPreviewSection')!.style.display = 'none';
  document.getElementById('sheetSelectorWrap')!.style.display = 'none';
  syncCategoryLabelEditor();
  renderVisualization();
});

async function loadRows(rows: any[], sourceName = 'Dataset'): Promise<void> {
  rawParsedData = rows.filter(row => row && typeof row === 'object');
  currentFileName = sourceName;
  currentSourceFileName = sourceName;
  currentSourceKind = 'import';
  if (!rawParsedData.length) throw new Error('No data rows were found.');
  availableColumns = Object.keys(rawParsedData[0]);
  if (!availableColumns.length) throw new Error('No columns were found.');
  populateMappingControls();
  const xSelect = document.getElementById('colX') as HTMLSelectElement;
  const ySelect = document.getElementById('colY') as HTMLSelectElement;
  if (xSelect && ySelect) {
    const nonNumeric = availableColumns.find(col => !isNumericColumn(col));
    const numericColumns = availableColumns.filter(isNumericColumn);
    xSelect.value = nonNumeric || availableColumns[0] || '';
    const numeric = numericColumns[0];
    if (numeric) ySelect.value = numeric;

    const metricSelect = document.getElementById('chartMetric') as HTMLSelectElement | null;
    if (metricSelect) {
      metricSelect.innerHTML = '';
      numericColumns.forEach(col => metricSelect.add(new Option(col, col)));
      chartMetricKey = chooseBestNumericMetric(
        numericColumns,
        (document.getElementById('chartTypeSelect') as HTMLSelectElement | null)?.value === 'annotatedRadial'
      ) || numeric || '';
      metricSelect.value = chartMetricKey;
    }
  }
  markImportedDatasetActive();
  updateValidationAndPreview();
  processMappingAndRender();
  const status = document.getElementById('dataStatus');
  if (status) status.innerHTML = `<span style="color:#34d399">Loaded:</span> ${escapePlotlyLabel(sourceName)} · ${rawParsedData.length.toLocaleString()} rows · ${availableColumns.length} columns`;
}

async function parseTextData(text: string, sourceName: string): Promise<void> {
  const delimiter = text.includes('\t') ? '\t' : undefined;
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      delimiter,
      complete: async (results: any) => {
        try { await loadRows(results.data as any[], sourceName); resolve(); }
        catch (e) { reject(e); }
      },
      error: (error: any) => reject(error)
    });
  });
}

async function parseFile(file: File): Promise<void> {
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    const buffer = await file.arrayBuffer();
    workbook = XLSX.read(buffer, { type: 'array' });
    const wrap = document.getElementById('sheetSelectorWrap')!;
    const sheetSelect = document.getElementById('sheetSelect') as HTMLSelectElement;
    sheetSelect.innerHTML = '';
    workbook.SheetNames.forEach(name => sheetSelect.add(new Option(name, name)));
    wrap.style.display = workbook.SheetNames.length > 1 ? 'block' : 'none';
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    await loadRows(rows, `${file.name} / ${sheetName}`);
    return;
  }
  const text = await file.text();
  await parseTextData(text, file.name);
}

document.getElementById('csvFile')?.addEventListener('change', async (e) => {
  const file = (e.target as HTMLInputElement).files?.[0];
  if (!file) return;
  try { await parseFile(file); }
  catch (error) {
    const status = document.getElementById('dataStatus');
    if (status) status.innerHTML = `<span style="color:#f87171">Import error:</span> ${escapePlotlyLabel(String(error))}`;
  }
});

document.getElementById('sheetSelect')?.addEventListener('change', async (e) => {
  if (!workbook) return;
  const sheetName = (e.target as HTMLSelectElement).value;
  try {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    await loadRows(rows, `${currentFileName.split(' / ')[0]} / ${sheetName}`);
  } catch (error) {
    const status = document.getElementById('dataStatus');
    if (status) status.innerHTML = `<span style="color:#f87171">Sheet error:</span> ${escapePlotlyLabel(String(error))}`;
  }
});

document.getElementById('pasteData')?.addEventListener('click', async () => {
  try {
    const text = await navigator.clipboard.readText();
    if (!text.trim()) throw new Error('Clipboard is empty.');
    await parseTextData(text, 'Clipboard Data');
  } catch (error) {
    const status = document.getElementById('dataStatus');
    if (status) status.innerHTML = `<span style="color:#f87171">Paste error:</span> ${escapePlotlyLabel(String(error))}`;
  }
});

const dropZone = document.getElementById('dropZone');
dropZone?.addEventListener('dragover', (event) => {
  event.preventDefault();
  (dropZone as HTMLElement).style.borderColor = '#60a5fa';
});
dropZone?.addEventListener('dragleave', () => {
  (dropZone as HTMLElement).style.borderColor = 'rgba(96,165,250,0.55)';
});
dropZone?.addEventListener('drop', async (event) => {
  event.preventDefault();
  (dropZone as HTMLElement).style.borderColor = 'rgba(96,165,250,0.55)';
  const file = event.dataTransfer?.files?.[0];
  if (!file) return;
  try { await parseFile(file); }
  catch (error) {
    const status = document.getElementById('dataStatus');
    if (status) status.innerHTML = `<span style="color:#f87171">Import error:</span> ${escapePlotlyLabel(String(error))}`;
  }
});

type ExportBackground = 'transparent' | 'white' | 'dark' | 'custom';

type ExportSettings = {
  background: ExportBackground;
  customBackground: string;
  scale: number;
};

function getExportSettings(): ExportSettings {
  const background = (getControlValue('exportBackground', 'transparent') || 'transparent') as ExportBackground;
  const customBackground = getControlValue('exportCustomBg', '#ffffff') || '#ffffff';
  const scale = Math.max(1, Math.min(8, Number(getControlValue('exportScale', '4')) || 4));
  return { background, customBackground, scale };
}

function backgroundColor(settings: ExportSettings): string | null {
  if (settings.background === 'transparent') return null;
  if (settings.background === 'white') return '#ffffff';
  if (settings.background === 'dark') return '#090d16';
  return settings.customBackground;
}

function setExportStatus(message: string): void {
  const el = document.getElementById('exportStatus');
  if (el) el.textContent = message;
}

function applyExportPreset(preset: string): void {
  const bg = document.getElementById('exportBackground') as HTMLSelectElement | null;
  const scale = document.getElementById('exportScale') as HTMLSelectElement | null;
  const custom = document.getElementById('exportCustomBg') as HTMLInputElement | null;
  if (!bg || !scale || !custom) return;
  if (preset === 'word') { bg.value = 'transparent'; scale.value = '4'; }
  else if (preset === 'presentation') { bg.value = 'white'; scale.value = '4'; }
  else if (preset === 'web') { bg.value = 'transparent'; scale.value = '2'; }
  if (preset !== 'custom') setExportStatus(`${preset[0].toUpperCase()}${preset.slice(1)} export preset selected.`);
}

async function withExportAppearance<T>(
  settings: ExportSettings,
  work: () => Promise<T>
): Promise<T> {
  const card = document.getElementById('renderCard') as HTMLElement;
  const title = card.querySelector('.chart-title') as HTMLElement | null;
  const subtitle = card.querySelector('.chart-subtitle') as HTMLElement | null;
  const plot = document.getElementById('visualizationCanvas') as HTMLElement;
  const bg = backgroundColor(settings);
  const originalBackground = card.style.background;
  const originalBoxShadow = card.style.boxShadow;
  const originalBorder = card.style.border;
  const originalTitleColor = title?.style.color || '';
  const originalSubtitleColor = subtitle?.style.color || '';

  card.style.boxShadow = 'none';
  card.style.border = 'none';
  card.style.background = bg ?? 'transparent';
  if (bg === null || bg === '#ffffff') {
    if (title) title.style.color = '#111827';
    if (subtitle) subtitle.style.color = '#4b5563';
  }

  try {
    await Plotly.relayout(plot, {
      paper_bgcolor: bg ?? 'rgba(0,0,0,0)',
      plot_bgcolor: bg ?? 'rgba(0,0,0,0)',
      
      'xaxis.tickfont.color': bg === null || bg === '#ffffff' ? '#111827' : '#f3f4f6',
      'yaxis.tickfont.color': bg === null || bg === '#ffffff' ? '#111827' : '#f3f4f6',
      'xaxis.gridcolor': bg === null || bg === '#ffffff' ? 'rgba(17,24,39,0.14)' : 'rgba(255,255,255,0.08)',
      'yaxis.gridcolor': bg === null || bg === '#ffffff' ? 'rgba(17,24,39,0.14)' : 'rgba(255,255,255,0.08)',
      'polar.angularaxis.tickfont.color': bg === null || bg === '#ffffff' ? '#111827' : '#f3f4f6',
      'polar.radialaxis.tickfont.color': bg === null || bg === '#ffffff' ? '#111827' : '#f3f4f6'
    } as any);
    return await work();
  } finally {
    card.style.background = originalBackground;
    card.style.boxShadow = originalBoxShadow;
    card.style.border = originalBorder;
    if (title) title.style.color = originalTitleColor;
    if (subtitle) subtitle.style.color = originalSubtitleColor;
    renderVisualization();
  }
}

async function captureExportCanvas(settings: ExportSettings): Promise<HTMLCanvasElement> {
  const card = document.getElementById('renderCard') as HTMLElement;
  const bg = backgroundColor(settings);
  return await withExportAppearance(settings, async () => {
    return await html2canvas(card, {
      scale: settings.scale,
      backgroundColor: bg,
      logging: false,
      useCORS: true,
      imageTimeout: 15000
    });
  });
}

async function downloadCanvas(canvas: HTMLCanvasElement, format: 'png' | 'jpeg' | 'webp', quality = 1): Promise<void> {
  const mime = `image/${format}`;
  const blob = await new Promise<Blob | null>(resolve => canvas.toBlob(resolve, mime, quality));
  if (!blob) throw new Error(`Could not create ${format.toUpperCase()} image.`);
  const link = document.createElement('a');
  link.download = `VizNova_Master_${Date.now()}.${format}`;
  link.href = URL.createObjectURL(blob);
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1500);
}

async function exportImage(format: 'png' | 'jpeg' | 'svg' | 'webp' | 'pdf'): Promise<void> {
  try {
    setExportStatus(`Preparing ${format.toUpperCase()} export…`);
    const settings = getExportSettings();
    if (format === 'svg') {
      const plot = document.getElementById('visualizationCanvas') as HTMLElement;
      const bg = backgroundColor(settings);
      const dataUrl = await withExportAppearance(settings, async () => Plotly.toImage(plot, {
        format: 'svg',
        width: Math.max(1400, plot.clientWidth * 2),
        height: Math.max(800, plot.clientHeight * 2)
      }));
      const response = await fetch(dataUrl);
      const blob = await response.blob();
      const link = document.createElement('a');
      link.download = `VizNova_Master_${Date.now()}.svg`;
      link.href = URL.createObjectURL(blob);
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 1500);
      setExportStatus(`SVG exported${bg ? ` on ${bg} background` : ' with transparency'}.`);
      return;
    }

    const canvas = await captureExportCanvas(settings);
    if (format === 'pdf') {
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4', compress: true });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 8;
      const maxW = pageWidth - margin * 2;
      const maxH = pageHeight - margin * 2;
      const ratio = Math.min(maxW / canvas.width, maxH / canvas.height);
      const w = canvas.width * ratio;
      const h = canvas.height * ratio;
      const x = (pageWidth - w) / 2;
      const y = (pageHeight - h) / 2;
      const imgData = canvas.toDataURL('image/jpeg', 0.98);
      pdf.addImage(imgData, 'JPEG', x, y, w, h, undefined, 'FAST');
      pdf.save(`VizNova_Master_${Date.now()}.pdf`);
      setExportStatus('PDF exported as an A4 landscape report page.');
      return;
    }

    await downloadCanvas(canvas, format as 'png' | 'jpeg' | 'webp', format === 'jpeg' ? 0.98 : 1);
    setExportStatus(`${format.toUpperCase()} exported successfully.`);
  } catch (error) {
    console.error(error);
    setExportStatus(`Export failed: ${String(error)}`);
  }
}

async function copyChartToClipboard(): Promise<void> {
  if (!navigator.clipboard || !('ClipboardItem' in window)) {
    alert('Clipboard image copy is not supported by this environment. Use PNG export instead.');
    return;
  }
  try {
    const settings = getExportSettings();
    const canvas = await captureExportCanvas({ ...settings, background: 'white' });
    canvas.toBlob(async blob => {
      if (!blob) return;
      try {
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
        const button = document.getElementById('copyChart');
        if (button) { const old = button.textContent; button.textContent = 'Copied ✓'; setTimeout(() => button.textContent = old, 1200); }
        setExportStatus('Chart copied as a white-background PNG.');
      } catch { alert('Could not copy the chart image. Use PNG export instead.'); }
    }, 'image/png');
  } catch (error) {
    setExportStatus(`Copy failed: ${String(error)}`);
  }
}


let phase4ProjectSystemInitialized = false;

function setupPhase4ProjectSystem(): void {
  if (phase4ProjectSystemInitialized) return;
  phase4ProjectSystemInitialized = true;

  refreshRecentProjects();

  document.getElementById('newProject')?.addEventListener('click', () => {
    if (!confirm('Start a new project? Unsaved changes will remain only if already saved.')) return;
    currentProjectName = 'Untitled';
    currentSourceFileName = '';
    currentSourceKind = '';
    currentSheetName = '';
    rawParsedData = [];
    availableColumns = [];
    workbook = null;
    currentLabels = [];
    currentDisplayLabels = [];
    currentValues = [];
    currentLabelName = '';
    chartMetricKey = '';
    dashboardItems = [];
    dashboardFilterValue = null;
    document.getElementById('columnMappingSection')?.setAttribute('style', 'display:none;');
    document.getElementById('dataPreviewSection')?.setAttribute('style', 'display:none;');
    document.getElementById('sheetSelectorWrap')?.setAttribute('style', 'display:none;');
    syncCategoryLabelEditor();
    renderDashboard();
    updateDashboardFilterStatus();
    updateStatistics();
    renderVisualization();
    projectStatus();
  });

  document.getElementById('saveProject')?.addEventListener('click', saveProjectToFile);

  document.getElementById('loadProject')?.addEventListener('click', () => {
    openJsonFile(projectFileInput, (snapshot) => {
      if (!snapshot || snapshot.version !== 4) {
        projectStatus('Unsupported VizNova project file.');
        return;
      }
      applyProjectSnapshot(snapshot as ProjectSnapshot);
      saveProjectToBrowser();
    });
  });

  document.getElementById('saveTemplate')?.addEventListener('click', saveTemplateToBrowser);

  document.getElementById('loadTemplate')?.addEventListener('click', () => {
    openJsonFile(templateFileInput, (template) => {
      if (!template || template.version !== 1) {
        projectStatus('Unsupported VizNova template file.');
        return;
      }
      loadTemplateSnapshot(template as TemplateSnapshot);
    });
  });

  document.getElementById('recentProjects')?.addEventListener('change', (event) => {
    const index = Number((event.target as HTMLSelectElement).value);
    if (!Number.isFinite(index)) return;
    const snapshot = getRecentProjects()[index];
    if (snapshot) applyProjectSnapshot(snapshot);
  });

  const autosave = document.getElementById('autoSaveProject') as HTMLInputElement | null;
  autosave?.addEventListener('change', () => {
    if (autosave.checked) {
      localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(getCurrentProjectSnapshot()));
      projectStatus('Autosave enabled');
    } else {
      projectStatus('Autosave disabled');
    }
  });

  const autosaved = safeJson<ProjectSnapshot | null>(localStorage.getItem(AUTOSAVE_KEY), null);
  if (autosaved && autosaved.version === 4) {
    try {
      const shouldRestore = confirm(`Restore autosaved project "${autosaved.name}"?`);
      if (shouldRestore) applyProjectSnapshot(autosaved);
    } catch (error) {
      console.error('VizNova autosave restore skipped:', error);
      localStorage.removeItem(AUTOSAVE_KEY);
      projectStatus('Previous autosave was incompatible and has been reset.');
    }
  }

  window.setInterval(() => {
    if (autosave?.checked) {
      const snapshot = getCurrentProjectSnapshot();
      localStorage.setItem(AUTOSAVE_KEY, JSON.stringify(snapshot));
      projectStatus(`Autosaved: ${currentProjectName}`);
    }
  }, 30000);
}

document.getElementById('exportPng')?.addEventListener('click', () => exportImage('png'));
document.getElementById('exportJpeg')?.addEventListener('click', () => exportImage('jpeg'));
document.getElementById('exportSvg')?.addEventListener('click', () => exportImage('svg'));
document.getElementById('exportPdf')?.addEventListener('click', () => exportImage('pdf'));
document.getElementById('exportWebp')?.addEventListener('click', () => exportImage('webp'));
document.getElementById('copyChart')?.addEventListener('click', copyChartToClipboard);
document.getElementById('exportPreset')?.addEventListener('change', (event) => {
  applyExportPreset((event.target as HTMLSelectElement).value);
});
document.getElementById('exportBackground')?.addEventListener('change', () => {
  const preset = document.getElementById('exportPreset') as HTMLSelectElement | null;
  if (preset) preset.value = 'custom';
});
document.getElementById('exportCustomBg')?.addEventListener('input', () => {
  const preset = document.getElementById('exportPreset') as HTMLSelectElement | null;
  if (preset) preset.value = 'custom';
});
document.getElementById('exportScale')?.addEventListener('change', () => {
  const preset = document.getElementById('exportPreset') as HTMLSelectElement | null;
  if (preset) preset.value = 'custom';
});

setupPhase4ProjectSystem();

function setupRadialPresentationControls(): void {
  const chartTypeEl = document.getElementById('chartTypeSelect') as HTMLSelectElement | null;
  const metricEl = document.getElementById('chartMetric') as HTMLSelectElement | null;
  const radialDetail = document.getElementById('radialDetailMetric') as HTMLSelectElement | null;
  const radialCenter = document.getElementById('radialCenterTitle') as HTMLInputElement | null;
  const radialReset =
    document.getElementById('resetRadialLabels') as HTMLButtonElement | null;

  const refresh = () => {
    if (chartTypeEl?.value === 'annotatedRadial') {
      configureRadialDetailMetric();
      syncRadialPresentationTitle();
      syncRadialPresentationLabelEditor();
    } else {
      document.body.classList.remove('radial-presentation-active');
      syncRadialPresentationLabelEditor();
    }
  };

  chartTypeEl?.addEventListener('change', refresh);
  metricEl?.addEventListener('change', refresh);
  radialDetail?.addEventListener('change', () => renderVisualization());
  radialCenter?.addEventListener('input', () => renderVisualization());
  radialReset?.addEventListener('click', resetRadialPresentationLabels);
  refresh();
}
setupRadialPresentationControls();
setupPhase5AdvancedApplication();

syncCategoryLabelEditor();
syncRadialPresentationLabelEditor();
updateRadialDetailOptions();
updateRadialPresentationVisibility();
renderVisualization();