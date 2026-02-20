/* ============================================================
   GraphMaker - Beautiful Graphs for Presentations
   Pure client-side: ECharts + SheetJS, zero dependencies
   ============================================================ */

// ── State ──────────────────────────────────────────────────
const state = {
    rawData: [],          // Array of row objects from the spreadsheet
    columns: [],          // { name, type: 'text'|'number'|'date' }
    pivot: { rows: [], values: [], columns: [], filters: [] },
    filterSelections: {},  // { colName: Set of selected values }
    aggregatedData: null,
    chartType: 'bar',
    chartInstance: null,
    currentBg: '#ffffff',
    currentPalette: 0,
};

// ── Color Palettes (PPT-friendly, high-contrast) ──────────
const PALETTES = [
    { name: 'Indigo',   colors: ['#6366f1','#8b5cf6','#ec4899','#f59e0b','#10b981','#3b82f6','#ef4444','#14b8a6','#f97316','#84cc16'] },
    { name: 'Business', colors: ['#1e3a5f','#2563eb','#0891b2','#059669','#d97706','#dc2626','#7c3aed','#db2777','#ea580c','#4f46e5'] },
    { name: 'Pastel',   colors: ['#93c5fd','#a5b4fc','#f9a8d4','#fcd34d','#6ee7b7','#fdba74','#fca5a5','#67e8f9','#d8b4fe','#86efac'] },
    { name: 'Bold',     colors: ['#ef4444','#f59e0b','#22c55e','#3b82f6','#a855f7','#ec4899','#14b8a6','#f97316','#6366f1','#eab308'] },
    { name: 'Earth',    colors: ['#92400e','#b45309','#a16207','#4d7c0f','#047857','#0e7490','#1e40af','#6b21a8','#be123c','#78716c'] },
    { name: 'Ocean',    colors: ['#0c4a6e','#0369a1','#0284c7','#0891b2','#06b6d4','#22d3ee','#67e8f9','#a5f3fc','#155e75','#164e63'] },
];

// ── Chart type definitions ──────────────────────────────────
const CHART_TYPES = [
    { id: 'bar',         icon: '📊', label: 'Bar' },
    { id: 'bar_h',       icon: '📊', label: 'H-Bar' },
    { id: 'stacked_bar', icon: '📊', label: 'Stacked' },
    { id: 'grouped_bar', icon: '📊', label: 'Grouped' },
    { id: 'line',        icon: '📈', label: 'Line' },
    { id: 'area',        icon: '📈', label: 'Area' },
    { id: 'pie',         icon: '🥧', label: 'Pie' },
    { id: 'donut',       icon: '🍩', label: 'Donut' },
    { id: 'scatter',     icon: '⚬',  label: 'Scatter' },
    { id: 'radar',       icon: '🕸', label: 'Radar' },
    { id: 'treemap',     icon: '▦',  label: 'Treemap' },
    { id: 'funnel',      icon: '▽',  label: 'Funnel' },
    { id: 'heatmap',     icon: '🟧', label: 'Heatmap' },
    { id: 'bubble',      icon: '◉',  label: 'Bubble' },
    { id: 'waterfall',   icon: '▥',  label: 'Waterfall' },
    { id: 'gantt',       icon: '▬',  label: 'Gantt' },
];

// ── Initialization ─────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    initUpload();
    initChartTypeGrid();
    initPaletteGrid();
    initDesignControls();
    initExportControls();
    initNavigation();
});

// ── Navigation ─────────────────────────────────────────────
function goToStep(n) {
    for (let i = 1; i <= 4; i++) {
        document.getElementById('step' + i).style.display = i === n ? 'flex' : 'none';
        const ind = document.querySelector(`.step-indicator[data-step="${i}"]`);
        ind.classList.toggle('active', i === n);
        ind.classList.toggle('done', i < n);
    }
    if (n === 3) renderChart();
    if (n === 4) generateExportPreview();
}

function initNavigation() {
    document.getElementById('btnToDesign').addEventListener('click', () => goToStep(3));
    document.getElementById('btnToExport').addEventListener('click', () => goToStep(4));
    document.getElementById('btnBackToPrepare').addEventListener('click', () => goToStep(2));
    document.getElementById('btnBackToDesign').addEventListener('click', () => goToStep(3));
    document.getElementById('btnStartOver').addEventListener('click', () => {
        if (state.chartInstance) { state.chartInstance.dispose(); state.chartInstance = null; }
        state.rawData = []; state.columns = [];
        state.pivot = { rows: [], values: [], columns: [], filters: [] };
        state.filterSelections = {};
        goToStep(1);
    });
    document.getElementById('btnMakeAnother').addEventListener('click', () => goToStep(2));
}

// ── Step 1: File Upload ────────────────────────────────────
function initUpload() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', e => {
        e.preventDefault(); dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', e => { if (e.target.files.length) handleFile(e.target.files[0]); });

    document.getElementById('btnDemoData').addEventListener('click', loadDemoData);
}

function handleFile(file) {
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(ws, { raw: false, dateNF: 'yyyy-mm-dd' });
            if (!json.length) { alert('The file appears to be empty.'); return; }
            processData(json);
        } catch (err) {
            alert('Could not read this file. Please try a different .xlsx, .xls, or .csv file.\n\n' + err.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function loadDemoData() {
    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const regions = ['North','South','East','West'];
    const products = ['Laptops','Phones','Tablets','Accessories'];
    const data = [];
    for (const region of regions) {
        for (const product of products) {
            for (const month of months) {
                data.push({
                    Month: month,
                    Region: region,
                    Product: product,
                    Revenue: Math.round(10000 + Math.random() * 90000),
                    Units: Math.round(50 + Math.random() * 450),
                    Profit: Math.round(2000 + Math.random() * 30000),
                });
            }
        }
    }
    processData(data);
}

function processData(json) {
    state.rawData = json;
    state.columns = detectColumns(json);
    state.pivot = { rows: [], values: [], columns: [], filters: [] };
    state.filterSelections = {};
    buildColumnChips();
    renderDataTable(json);
    goToStep(2);
}

function detectColumns(data) {
    const cols = Object.keys(data[0]);
    return cols.map(name => {
        const samples = data.slice(0, 100).map(r => r[name]).filter(v => v != null && v !== '');
        let type = 'text';
        if (samples.length) {
            const numCount = samples.filter(v => !isNaN(parseFloat(v)) && isFinite(v)).length;
            if (numCount / samples.length > 0.8) type = 'number';
            else {
                const dateCount = samples.filter(v => !isNaN(Date.parse(v))).length;
                if (dateCount / samples.length > 0.8) type = 'date';
            }
        }
        return { name, type };
    });
}

// ── Step 2: Pivot / Data Preparation ────────────────────────

function buildColumnChips() {
    const list = document.getElementById('columnList');
    list.innerHTML = '';
    for (const col of state.columns) {
        const chip = document.createElement('div');
        chip.className = 'column-chip type-' + col.type;
        chip.draggable = true;
        chip.dataset.col = col.name;
        const iconMap = { text: 'Aa', number: '#', date: '📅' };
        chip.innerHTML = `<span class="chip-icon">${iconMap[col.type]}</span> ${col.name}`;
        chip.addEventListener('dragstart', e => {
            e.dataTransfer.setData('text/plain', col.name);
            e.dataTransfer.effectAllowed = 'copy';
        });
        list.appendChild(chip);
    }

    // Init drop zones
    for (const zone of document.querySelectorAll('.pivot-drop')) {
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', e => {
            e.preventDefault(); zone.classList.remove('dragover');
            const colName = e.dataTransfer.getData('text/plain');
            addToPivot(zone.dataset.role, colName);
        });
    }
}

function addToPivot(role, colName) {
    // Prevent duplicates in same role
    if (state.pivot[role].find(c => c.name === colName)) return;

    const col = state.columns.find(c => c.name === colName);
    if (!col) return;

    const entry = { name: colName, agg: 'sum' };
    if (role === 'values' && col.type !== 'number') entry.agg = 'count';
    state.pivot[role].push(entry);

    if (role === 'filters') buildFilterUI(colName);
    refreshPivotZones();
    updatePreview();
}

function removeFromPivot(role, colName) {
    state.pivot[role] = state.pivot[role].filter(c => c.name !== colName);
    if (role === 'filters') delete state.filterSelections[colName];
    refreshPivotZones();
    updatePreview();
}

function refreshPivotZones() {
    for (const role of ['rows','values','columns','filters']) {
        const zone = document.getElementById('pivot' + role.charAt(0).toUpperCase() + role.slice(1));
        zone.innerHTML = '';
        if (!state.pivot[role].length) {
            zone.innerHTML = '<span class="pivot-placeholder">Drop a column here</span>';
            continue;
        }
        for (const entry of state.pivot[role]) {
            const chip = document.createElement('span');
            chip.className = 'pivot-chip';
            let inner = entry.name;
            if (role === 'values') {
                inner += ` <select onchange="changeAgg('${entry.name}', this.value)">`;
                for (const a of ['sum','avg','count','min','max']) {
                    inner += `<option value="${a}" ${entry.agg===a?'selected':''}>${a.toUpperCase()}</option>`;
                }
                inner += '</select>';
            }
            inner += ` <span class="remove-chip" onclick="removeFromPivot('${role}','${entry.name}')">&times;</span>`;
            chip.innerHTML = inner;
            zone.appendChild(chip);
        }
    }
    buildFilterControls();
}

function changeAgg(colName, agg) {
    const entry = state.pivot.values.find(c => c.name === colName);
    if (entry) entry.agg = agg;
    updatePreview();
}

function buildFilterUI(colName) {
    const vals = [...new Set(state.rawData.map(r => String(r[colName] ?? '')))].sort();
    state.filterSelections[colName] = new Set(vals);
}

function buildFilterControls() {
    const wrap = document.getElementById('filterControls');
    wrap.innerHTML = '';
    for (const entry of state.pivot.filters) {
        const colName = entry.name;
        const vals = [...new Set(state.rawData.map(r => String(r[colName] ?? '')))].sort();
        const group = document.createElement('div');
        group.className = 'filter-group';
        group.innerHTML = `<label>Filter: ${colName}</label>`;
        const checks = document.createElement('div');
        checks.className = 'filter-checks';
        for (const v of vals) {
            const id = `f_${colName}_${v}`.replace(/\W/g,'_');
            const checked = !state.filterSelections[colName] || state.filterSelections[colName].has(v);
            const lbl = document.createElement('label');
            lbl.innerHTML = `<input type="checkbox" id="${id}" ${checked?'checked':''} value="${v}"> ${v || '(empty)'}`;
            lbl.querySelector('input').addEventListener('change', function() {
                if (!state.filterSelections[colName]) state.filterSelections[colName] = new Set(vals);
                if (this.checked) state.filterSelections[colName].add(v);
                else state.filterSelections[colName].delete(v);
                updatePreview();
            });
            checks.appendChild(lbl);
        }
        group.appendChild(checks);
        wrap.appendChild(group);
    }
}

function getFilteredData() {
    let data = state.rawData;
    for (const entry of state.pivot.filters) {
        const sel = state.filterSelections[entry.name];
        if (sel) data = data.filter(r => sel.has(String(r[entry.name] ?? '')));
    }
    return data;
}

function aggregateData() {
    const data = getFilteredData();
    const { rows, values, columns } = state.pivot;

    if (!rows.length || !values.length) {
        state.aggregatedData = null;
        return null;
    }

    // Group data
    const groups = {};
    for (const row of data) {
        const rowKey = rows.map(r => String(row[r.name] ?? '')).join('|||');
        const colKey = columns.length ? columns.map(c => String(row[c.name] ?? '')).join('|||') : '__all__';
        const key = rowKey + '###' + colKey;
        if (!groups[key]) groups[key] = { rowKey, colKey, rows: rows.map(r => row[r.name]), colVals: columns.map(c => row[c.name]), items: [] };
        groups[key].items.push(row);
    }

    // Aggregate
    const result = [];
    for (const g of Object.values(groups)) {
        const entry = { _rowKey: g.rowKey, _colKey: g.colKey, _rowLabels: g.rows, _colLabels: g.colVals };
        for (const v of values) {
            const nums = g.items.map(r => parseFloat(r[v.name])).filter(n => !isNaN(n));
            switch (v.agg) {
                case 'sum': entry[v.name] = nums.reduce((a,b) => a+b, 0); break;
                case 'avg': entry[v.name] = nums.length ? nums.reduce((a,b) => a+b, 0) / nums.length : 0; break;
                case 'count': entry[v.name] = g.items.length; break;
                case 'min': entry[v.name] = nums.length ? Math.min(...nums) : 0; break;
                case 'max': entry[v.name] = nums.length ? Math.max(...nums) : 0; break;
            }
        }
        result.push(entry);
    }

    state.aggregatedData = result;
    return result;
}

function updatePreview() {
    const aggData = aggregateData();
    if (!aggData) {
        renderDataTable(getFilteredData().slice(0, 200));
        return;
    }

    // Build preview table from aggregated data
    const { rows, values, columns } = state.pivot;
    const table = document.getElementById('dataTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    // Unique column keys
    const colKeys = [...new Set(aggData.map(r => r._colKey))].sort();
    const hasMultiCols = columns.length > 0 && colKeys.length > 1;

    // Header
    let headerHTML = '<tr>';
    for (const r of rows) headerHTML += `<th>${r.name}</th>`;
    if (hasMultiCols) {
        for (const ck of colKeys) {
            for (const v of values) {
                headerHTML += `<th>${ck === '__all__' ? '' : ck.replace(/\|\|\|/g,' / ')} ${v.name} (${v.agg})</th>`;
            }
        }
    } else {
        for (const v of values) headerHTML += `<th>${v.name} (${v.agg})</th>`;
    }
    headerHTML += '</tr>';
    thead.innerHTML = headerHTML;

    // Group by row key for pivot table display
    const rowGroups = {};
    for (const r of aggData) {
        if (!rowGroups[r._rowKey]) rowGroups[r._rowKey] = { labels: r._rowLabels, byCol: {} };
        rowGroups[r._rowKey].byCol[r._colKey] = r;
    }

    let bodyHTML = '';
    for (const rg of Object.values(rowGroups)) {
        bodyHTML += '<tr>';
        for (const l of rg.labels) bodyHTML += `<td>${l ?? ''}</td>`;
        if (hasMultiCols) {
            for (const ck of colKeys) {
                for (const v of values) {
                    const val = rg.byCol[ck] ? rg.byCol[ck][v.name] : 0;
                    bodyHTML += `<td>${formatNumber(val)}</td>`;
                }
            }
        } else {
            for (const v of values) {
                const val = rg.byCol['__all__'] ? rg.byCol['__all__'][v.name] : 0;
                bodyHTML += `<td>${formatNumber(val)}</td>`;
            }
        }
        bodyHTML += '</tr>';
    }
    tbody.innerHTML = bodyHTML;
    document.getElementById('rowCount').textContent = `${Object.keys(rowGroups).length} groups from ${getFilteredData().length} rows`;
}

function renderDataTable(data) {
    const table = document.getElementById('dataTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    if (!data.length) { thead.innerHTML = ''; tbody.innerHTML = ''; return; }

    const cols = Object.keys(data[0]);
    thead.innerHTML = '<tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr>';

    const rows = data.slice(0, 200);
    tbody.innerHTML = rows.map(r =>
        '<tr>' + cols.map(c => `<td>${r[c] ?? ''}</td>`).join('') + '</tr>'
    ).join('');

    document.getElementById('rowCount').textContent = `${state.rawData.length} rows, ${cols.length} columns` +
        (data.length > 200 ? ' (showing first 200)' : '');
}

function formatNumber(n) {
    if (n == null || isNaN(n)) return '';
    if (Number.isInteger(n)) return n.toLocaleString();
    return n.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

// ── Step 3: Chart Design ────────────────────────────────────

function initChartTypeGrid() {
    const grid = document.getElementById('chartTypeGrid');
    for (const ct of CHART_TYPES) {
        const btn = document.createElement('button');
        btn.className = 'chart-type-btn' + (ct.id === state.chartType ? ' active' : '');
        btn.dataset.type = ct.id;
        btn.innerHTML = `<span class="ct-icon">${ct.icon}</span>${ct.label}`;
        btn.addEventListener('click', () => {
            document.querySelectorAll('.chart-type-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.chartType = ct.id;
            document.getElementById('ganttOptions').style.display = ct.id === 'gantt' ? 'block' : 'none';
            renderChart();
        });
        grid.appendChild(btn);
    }
}

function initPaletteGrid() {
    const grid = document.getElementById('paletteGrid');
    PALETTES.forEach((p, i) => {
        const btn = document.createElement('button');
        btn.className = 'palette-btn' + (i === 0 ? ' active' : '');
        btn.title = p.name;
        btn.innerHTML = '<div class="palette-colors">' +
            p.colors.slice(0,5).map(c => `<span style="background:${c}"></span>`).join('') +
            '</div>';
        btn.addEventListener('click', () => {
            document.querySelectorAll('.palette-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.currentPalette = i;
            renderChart();
        });
        grid.appendChild(btn);
    });
}

function initDesignControls() {
    const controls = ['chartTitle','chartSubtitle','xAxisLabel','yAxisLabel','fontSize',
                      'showLegend','showDataLabels','showGrid','legendPosition'];
    for (const id of controls) {
        const el = document.getElementById(id);
        el.addEventListener(el.type === 'range' ? 'input' : 'change', () => {
            if (id === 'fontSize') document.getElementById('fontSizeVal').textContent = el.value + 'px';
            renderChart();
        });
    }

    // Background buttons
    document.querySelectorAll('.bg-opt').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.bg-opt').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.currentBg = btn.dataset.bg;
            renderChart();
        });
    });

    // Gantt column selectors
    for (const id of ['ganttTask','ganttStart','ganttEnd','ganttCategory','ganttProgress']) {
        document.getElementById(id).addEventListener('change', () => renderChart());
    }
}

function populateGanttSelectors() {
    const cols = state.columns.map(c => c.name);
    for (const id of ['ganttTask','ganttStart','ganttEnd','ganttCategory','ganttProgress']) {
        const sel = document.getElementById(id);
        const prev = sel.value;
        const hasNone = id === 'ganttCategory' || id === 'ganttProgress';
        sel.innerHTML = (hasNone ? '<option value="">None</option>' : '') +
            cols.map(c => `<option value="${c}">${c}</option>`).join('');
        if (prev && cols.includes(prev)) sel.value = prev;
    }
}

// ── Chart Rendering ─────────────────────────────────────────

function renderChart() {
    const container = document.getElementById('chartDiv');
    if (state.chartInstance) state.chartInstance.dispose();
    state.chartInstance = echarts.init(container, null, { renderer: 'svg' });

    if (state.chartType === 'gantt') {
        populateGanttSelectors();
        renderGanttChart();
        return;
    }

    const aggData = aggregateData();
    if (!aggData || !aggData.length) {
        state.chartInstance.setOption({ title: { text: 'Set up your data in Step 2', left: 'center', top: 'center', textStyle: { color: '#94a3b8', fontSize: 18 } } });
        return;
    }

    const option = buildChartOption(aggData);
    state.chartInstance.setOption(option);
    window.addEventListener('resize', () => state.chartInstance && state.chartInstance.resize());
}

function buildChartOption(aggData) {
    const { rows, values, columns } = state.pivot;
    const colors = PALETTES[state.currentPalette].colors;
    const fs = parseInt(document.getElementById('fontSize').value);
    const bg = state.currentBg;
    const isDark = bg === '#1e293b';
    const textColor = isDark ? '#e2e8f0' : '#334155';
    const subColor = isDark ? '#94a3b8' : '#64748b';
    const gridColor = isDark ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)';
    const showLegend = document.getElementById('showLegend').checked;
    const showLabels = document.getElementById('showDataLabels').checked;
    const showGrid = document.getElementById('showGrid').checked;
    const legendPos = document.getElementById('legendPosition').value;

    // Get unique row keys as categories
    const rowKeys = [...new Map(aggData.map(r => [r._rowKey, r._rowLabels])).entries()];
    const categories = rowKeys.map(([k, labels]) => labels.join(' / '));

    // Get unique col keys as series
    const colKeys = [...new Set(aggData.map(r => r._colKey))].sort();
    const hasSeries = columns.length > 0 && colKeys.length > 1;

    // Map data by row/col
    const dataMap = {};
    for (const r of aggData) dataMap[r._rowKey + '___' + r._colKey] = r;

    const title = document.getElementById('chartTitle').value || '';
    const subtitle = document.getElementById('chartSubtitle').value || '';
    const xLabel = document.getElementById('xAxisLabel').value || (rows.length ? rows.map(r=>r.name).join(', ') : '');
    const yLabel = document.getElementById('yAxisLabel').value || (values.length ? values.map(v=>`${v.name} (${v.agg})`).join(', ') : '');

    // Base options for PPT-quality text
    const baseTextStyle = {
        fontFamily: 'Inter, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif',
        fontSize: fs,
        fontWeight: 500,
    };

    const option = {
        backgroundColor: bg === 'transparent' ? 'transparent' : bg,
        color: colors,
        textStyle: { ...baseTextStyle, color: textColor },
        title: {
            text: title,
            subtext: subtitle,
            left: 'center',
            top: 8,
            textStyle: { ...baseTextStyle, fontSize: fs + 6, fontWeight: 700, color: textColor },
            subtextStyle: { ...baseTextStyle, fontSize: fs - 1, color: subColor },
        },
        tooltip: {
            trigger: isPieLike() ? 'item' : 'axis',
            textStyle: { fontSize: fs - 2 },
        },
        animation: true,
        animationDuration: 600,
    };

    // Legend
    if (showLegend) {
        const legendConfig = {
            show: true,
            textStyle: { ...baseTextStyle, fontSize: fs - 2, color: textColor },
            itemGap: 16,
            itemWidth: 14, itemHeight: 10,
        };
        if (legendPos === 'bottom') { legendConfig.bottom = 10; legendConfig.left = 'center'; }
        else if (legendPos === 'top') { legendConfig.top = title ? 60 : 10; legendConfig.left = 'center'; }
        else if (legendPos === 'left') { legendConfig.left = 10; legendConfig.top = 'middle'; legendConfig.orient = 'vertical'; }
        else if (legendPos === 'right') { legendConfig.right = 10; legendConfig.top = 'middle'; legendConfig.orient = 'vertical'; }
        option.legend = legendConfig;
    }

    // Grid (for cartesian charts) - containLabel ensures axis labels never overlap
    const gridPadding = { top: 80, right: 80, bottom: showLegend && legendPos === 'bottom' ? 80 : 50, left: 30, containLabel: true };
    if (showLegend && legendPos === 'left') gridPadding.left = 140;
    if (showLegend && legendPos === 'right') gridPadding.right = 160;

    // Build per chart type
    const type = state.chartType;

    if (type === 'pie' || type === 'donut') {
        return buildPieOption(option, aggData, categories, values, colors, fs, showLabels, textColor, type === 'donut');
    }
    if (type === 'treemap') {
        return buildTreemapOption(option, aggData, categories, values, colors, fs, showLabels, textColor);
    }
    if (type === 'funnel') {
        return buildFunnelOption(option, aggData, categories, values, colors, fs, showLabels, textColor);
    }
    if (type === 'radar') {
        return buildRadarOption(option, aggData, categories, values, colKeys, hasSeries, dataMap, rowKeys, colors, fs, showLabels, textColor, gridColor);
    }
    if (type === 'heatmap') {
        return buildHeatmapOption(option, aggData, categories, values, colKeys, hasSeries, dataMap, rowKeys, colors, fs, showLabels, textColor, gridColor, gridPadding, isDark);
    }

    // Cartesian charts (bar, line, area, scatter, bubble, waterfall)
    const isHorizontal = type === 'bar_h';
    const isStacked = type === 'stacked_bar';
    const isGrouped = type === 'grouped_bar';

    const catAxis = {
        type: 'category',
        data: categories,
        axisLabel: {
            ...baseTextStyle,
            fontSize: fs - 2,
            color: subColor,
            rotate: categories.length > 10 ? 45 : 0,
            overflow: 'truncate',
            width: 120,
        },
        axisLine: { lineStyle: { color: gridColor } },
        axisTick: { show: false },
        name: isHorizontal ? yLabel : xLabel,
        nameTextStyle: { ...baseTextStyle, fontSize: fs - 1, color: subColor },
        nameLocation: 'end',
        nameGap: 15,
    };

    const valAxis = {
        type: 'value',
        axisLabel: { ...baseTextStyle, fontSize: fs - 2, color: subColor, formatter: v => abbreviate(v) },
        splitLine: { show: showGrid, lineStyle: { color: gridColor } },
        axisLine: { show: false },
        name: isHorizontal ? xLabel : yLabel,
        nameTextStyle: { ...baseTextStyle, fontSize: fs - 1, color: subColor, padding: isHorizontal ? [10,0,0,0] : [0,0,0,0] },
        nameLocation: 'end',
        nameGap: 15,
    };

    option.grid = gridPadding;
    option.xAxis = isHorizontal ? valAxis : catAxis;
    option.yAxis = isHorizontal ? catAxis : valAxis;

    // Build series
    const series = [];
    const valField = values[0].name;

    if (type === 'waterfall') {
        return buildWaterfallOption(option, aggData, categories, values, rowKeys, dataMap, colors, fs, showLabels, textColor, baseTextStyle);
    }

    if (type === 'scatter' || type === 'bubble') {
        return buildScatterOption(option, aggData, values, hasSeries, colKeys, dataMap, rowKeys, colors, fs, showLabels, textColor, type === 'bubble');
    }

    if (hasSeries) {
        for (const ck of colKeys) {
            const sData = rowKeys.map(([rk]) => {
                const d = dataMap[rk + '___' + ck];
                return d ? d[valField] : 0;
            });
            const s = {
                name: ck.replace(/\|\|\|/g, ' / '),
                type: (type === 'line' || type === 'area') ? 'line' : 'bar',
                data: sData,
                label: { show: showLabels, ...baseTextStyle, fontSize: fs - 3, formatter: p => abbreviate(p.value) },
            };
            if (type === 'area') s.areaStyle = { opacity: 0.3 };
            if (isStacked) s.stack = 'total';
            if (isGrouped) s.barGap = '10%';
            series.push(s);
        }
    } else {
        // One series per value field
        for (const v of values) {
            const sData = rowKeys.map(([rk]) => {
                const d = dataMap[rk + '___' + '__all__'];
                return d ? d[v.name] : 0;
            });
            const s = {
                name: `${v.name} (${v.agg})`,
                type: (type === 'line' || type === 'area') ? 'line' : 'bar',
                data: sData,
                label: { show: showLabels, ...baseTextStyle, fontSize: fs - 3, formatter: p => abbreviate(p.value) },
            };
            if (type === 'area') s.areaStyle = { opacity: 0.3 };
            if (isStacked) s.stack = 'total';
            series.push(s);
        }
    }

    option.series = series;
    return option;
}

function isPieLike() {
    return ['pie','donut','treemap','funnel'].includes(state.chartType);
}

function abbreviate(n) {
    if (n == null || isNaN(n)) return '';
    const abs = Math.abs(n);
    if (abs >= 1e9) return (n/1e9).toFixed(1) + 'B';
    if (abs >= 1e6) return (n/1e6).toFixed(1) + 'M';
    if (abs >= 1e3) return (n/1e3).toFixed(1) + 'K';
    return Number.isInteger(n) ? n.toString() : n.toFixed(1);
}

// ── Pie / Donut ──────────────────────────────────────────────
function buildPieOption(option, aggData, categories, values, colors, fs, showLabels, textColor, isDonut) {
    const valField = values[0].name;
    const rowKeys = [...new Map(aggData.map(r => [r._rowKey, r._rowLabels])).entries()];
    const data = rowKeys.map(([rk, labels], i) => {
        const d = aggData.find(r => r._rowKey === rk);
        return { name: labels.join(' / '), value: d ? d[valField] : 0 };
    });

    option.series = [{
        type: 'pie',
        radius: isDonut ? ['40%','72%'] : ['0%','72%'],
        center: ['50%','55%'],
        data,
        label: {
            show: true,
            formatter: showLabels ? '{b}: {d}%' : '{b}',
            fontSize: fs - 2,
            fontWeight: 500,
            color: textColor,
            fontFamily: 'Inter, sans-serif',
        },
        emphasis: { itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0,0,0,0.15)' } },
        itemStyle: { borderRadius: isDonut ? 6 : 4, borderColor: option.backgroundColor || '#fff', borderWidth: 2 },
    }];
    return option;
}

// ── Treemap ──────────────────────────────────────────────────
function buildTreemapOption(option, aggData, categories, values, colors, fs, showLabels, textColor) {
    const valField = values[0].name;
    const rowKeys = [...new Map(aggData.map(r => [r._rowKey, r._rowLabels])).entries()];
    const data = rowKeys.map(([rk, labels]) => {
        const d = aggData.find(r => r._rowKey === rk);
        return { name: labels.join(' / '), value: d ? d[valField] : 0 };
    });

    option.series = [{
        type: 'treemap',
        data,
        label: { show: true, formatter: '{b}\n{c}', fontSize: fs - 2, fontWeight: 500, color: '#fff', fontFamily: 'Inter, sans-serif' },
        breadcrumb: { show: false },
        itemStyle: { borderWidth: 2, borderColor: option.backgroundColor || '#fff', gapWidth: 2 },
    }];
    return option;
}

// ── Funnel ───────────────────────────────────────────────────
function buildFunnelOption(option, aggData, categories, values, colors, fs, showLabels, textColor) {
    const valField = values[0].name;
    const rowKeys = [...new Map(aggData.map(r => [r._rowKey, r._rowLabels])).entries()];
    const data = rowKeys.map(([rk, labels]) => {
        const d = aggData.find(r => r._rowKey === rk);
        return { name: labels.join(' / '), value: d ? d[valField] : 0 };
    }).sort((a,b) => b.value - a.value);

    option.series = [{
        type: 'funnel',
        left: '15%', right: '15%', top: 80, bottom: 40,
        data,
        label: { show: true, formatter: showLabels ? '{b}: {c}' : '{b}', fontSize: fs - 1, fontWeight: 500, color: textColor, fontFamily: 'Inter, sans-serif' },
        itemStyle: { borderWidth: 1, borderColor: option.backgroundColor || '#fff' },
    }];
    return option;
}

// ── Radar ────────────────────────────────────────────────────
function buildRadarOption(option, aggData, categories, values, colKeys, hasSeries, dataMap, rowKeys, colors, fs, showLabels, textColor, gridColor) {
    const valField = values[0].name;
    const indicator = categories.map(c => {
        const maxVal = Math.max(...aggData.map(r => r[valField] || 0));
        return { name: c, max: maxVal * 1.2 };
    });

    option.radar = {
        indicator,
        shape: 'polygon',
        axisName: { color: textColor, fontSize: fs - 3, fontFamily: 'Inter, sans-serif' },
        splitLine: { lineStyle: { color: gridColor } },
        splitArea: { show: false },
    };

    if (hasSeries) {
        option.series = [{
            type: 'radar',
            data: colKeys.map(ck => ({
                name: ck.replace(/\|\|\|/g, ' / '),
                value: rowKeys.map(([rk]) => { const d = dataMap[rk + '___' + ck]; return d ? d[valField] : 0; }),
            })),
            label: { show: showLabels, fontSize: fs - 3 },
        }];
    } else {
        option.series = [{
            type: 'radar',
            data: [{
                name: `${valField}`,
                value: rowKeys.map(([rk]) => { const d = dataMap[rk + '___' + '__all__']; return d ? d[valField] : 0; }),
            }],
            label: { show: showLabels, fontSize: fs - 3 },
        }];
    }
    return option;
}

// ── Heatmap ──────────────────────────────────────────────────
function buildHeatmapOption(option, aggData, categories, values, colKeys, hasSeries, dataMap, rowKeys, colors, fs, showLabels, textColor, gridColor, gridPadding, isDark) {
    if (!hasSeries || colKeys.length < 2) {
        // Need split-by for heatmap
        option.title.subtext = 'Heatmap requires a "Split By" column. Please add one in Step 2.';
        return option;
    }

    const valField = values[0].name;
    const yLabels = colKeys.map(ck => ck.replace(/\|\|\|/g, ' / '));
    const allVals = aggData.map(r => r[valField]).filter(v => !isNaN(v));
    const hData = [];
    rowKeys.forEach(([rk], xi) => {
        colKeys.forEach((ck, yi) => {
            const d = dataMap[rk + '___' + ck];
            hData.push([xi, yi, d ? d[valField] : 0]);
        });
    });

    option.grid = { ...gridPadding, left: 30, containLabel: true };
    option.xAxis = { type: 'category', data: categories, axisLabel: { fontSize: fs - 3, color: textColor, rotate: categories.length > 8 ? 45 : 0, fontFamily: 'Inter, sans-serif' }, splitArea: { show: true } };
    option.yAxis = { type: 'category', data: yLabels, axisLabel: { fontSize: fs - 3, color: textColor, fontFamily: 'Inter, sans-serif' }, splitArea: { show: true } };
    option.visualMap = {
        min: allVals.length ? Math.min(...allVals) : 0,
        max: allVals.length ? Math.max(...allVals) : 100,
        calculable: true,
        orient: 'horizontal',
        left: 'center', bottom: 10,
        inRange: { color: [isDark ? '#1e293b' : '#eef2ff', colors[0]] },
        textStyle: { color: textColor, fontSize: fs - 3 },
    };
    option.series = [{
        type: 'heatmap',
        data: hData,
        label: { show: showLabels, fontSize: fs - 3, fontWeight: 500, color: textColor, fontFamily: 'Inter, sans-serif', formatter: p => abbreviate(p.value[2]) },
        emphasis: { itemStyle: { shadowBlur: 10 } },
    }];
    delete option.legend;
    return option;
}

// ── Scatter / Bubble ─────────────────────────────────────────
function buildScatterOption(option, aggData, values, hasSeries, colKeys, dataMap, rowKeys, colors, fs, showLabels, textColor, isBubble) {
    if (values.length < 2) {
        option.title.subtext = 'Scatter/Bubble needs at least 2 value columns.';
        return option;
    }

    const xField = values[0].name;
    const yField = values[1].name;
    const sizeField = isBubble && values.length >= 3 ? values[2].name : null;

    option.xAxis = { type: 'value', name: xField, nameLocation: 'middle', nameGap: 30, nameTextStyle: { fontSize: fs - 1, color: textColor, fontFamily: 'Inter, sans-serif' }, axisLabel: { fontSize: fs - 2, color: textColor, formatter: v => abbreviate(v) }, splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } } };
    option.yAxis = { type: 'value', name: yField, nameLocation: 'end', nameGap: 15, nameTextStyle: { fontSize: fs - 1, color: textColor, fontFamily: 'Inter, sans-serif' }, axisLabel: { fontSize: fs - 2, color: textColor, formatter: v => abbreviate(v) }, splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } } };
    option.grid = { top: 80, right: 50, bottom: 60, left: 30, containLabel: true };

    if (hasSeries) {
        option.series = colKeys.map(ck => ({
            name: ck.replace(/\|\|\|/g, ' / '),
            type: 'scatter',
            data: rowKeys.map(([rk]) => {
                const d = dataMap[rk + '___' + ck];
                if (!d) return [0,0];
                const pt = [d[xField], d[yField]];
                if (sizeField) pt.push(d[sizeField]);
                return pt;
            }),
            symbolSize: isBubble ? (v => Math.sqrt(v[2]) / 3 + 5) : 10,
            label: { show: showLabels, fontSize: fs - 3, formatter: p => p.data[1] },
        }));
    } else {
        option.series = [{
            type: 'scatter',
            data: rowKeys.map(([rk]) => {
                const d = dataMap[rk + '___' + '__all__'];
                if (!d) return [0,0];
                const pt = [d[xField], d[yField]];
                if (sizeField) pt.push(d[sizeField]);
                return pt;
            }),
            symbolSize: isBubble ? (v => Math.sqrt(v[2]) / 3 + 5) : 10,
            label: { show: showLabels, fontSize: fs - 3 },
        }];
    }
    return option;
}

// ── Waterfall ────────────────────────────────────────────────
function buildWaterfallOption(option, aggData, categories, values, rowKeys, dataMap, colors, fs, showLabels, textColor, baseTextStyle) {
    const valField = values[0].name;
    const rawValues = rowKeys.map(([rk]) => {
        const d = dataMap[rk + '___' + '__all__'] || dataMap[rk + '___' + Object.keys(dataMap).find(k => k.startsWith(rk))?.split('___')[1]];
        return d ? d[valField] : 0;
    });

    // Calculate waterfall base and visible parts
    const base = []; const pos = []; const neg = [];
    let cumulative = 0;
    for (const v of rawValues) {
        if (v >= 0) {
            base.push(cumulative);
            pos.push(v);
            neg.push(0);
        } else {
            base.push(cumulative + v);
            pos.push(0);
            neg.push(-v);
        }
        cumulative += v;
    }

    option.series = [
        { type: 'bar', stack: 'wf', data: base, itemStyle: { color: 'transparent' }, emphasis: { itemStyle: { color: 'transparent' } } },
        { name: 'Increase', type: 'bar', stack: 'wf', data: pos, itemStyle: { color: colors[0] },
          label: { show: showLabels, position: 'top', ...baseTextStyle, fontSize: fs - 3, formatter: p => p.value ? abbreviate(p.value) : '' } },
        { name: 'Decrease', type: 'bar', stack: 'wf', data: neg, itemStyle: { color: colors[4] || '#ef4444' },
          label: { show: showLabels, position: 'bottom', ...baseTextStyle, fontSize: fs - 3, formatter: p => p.value ? '-' + abbreviate(p.value) : '' } },
    ];
    return option;
}

// ── Gantt Chart ──────────────────────────────────────────────
function renderGanttChart() {
    const data = getFilteredData();
    const taskCol = document.getElementById('ganttTask').value;
    const startCol = document.getElementById('ganttStart').value;
    const endCol = document.getElementById('ganttEnd').value;
    const catCol = document.getElementById('ganttCategory').value;
    const progCol = document.getElementById('ganttProgress').value;

    if (!taskCol || !startCol || !endCol) {
        state.chartInstance.setOption({
            title: { text: 'Select Task, Start Date, and End Date columns', left: 'center', top: 'center',
                     textStyle: { color: '#94a3b8', fontSize: 16, fontFamily: 'Inter, sans-serif' } }
        });
        return;
    }

    const colors = PALETTES[state.currentPalette].colors;
    const fs = parseInt(document.getElementById('fontSize').value);
    const bg = state.currentBg;
    const isDark = bg === '#1e293b';
    const textColor = isDark ? '#e2e8f0' : '#334155';

    // Parse tasks
    const tasks = data.map((row, i) => {
        const start = new Date(row[startCol]);
        const end = new Date(row[endCol]);
        return {
            name: String(row[taskCol] || 'Task ' + (i+1)),
            start: isNaN(start) ? new Date() : start,
            end: isNaN(end) ? new Date() : end,
            category: catCol ? String(row[catCol] || '') : '',
            progress: progCol ? parseFloat(row[progCol]) || 0 : -1,
        };
    }).filter(t => !isNaN(t.start.getTime()) && !isNaN(t.end.getTime()));

    if (!tasks.length) {
        state.chartInstance.setOption({
            title: { text: 'No valid date data found', left: 'center', top: 'center',
                     textStyle: { color: '#94a3b8', fontSize: 16 } }
        });
        return;
    }

    // Sort by start date
    tasks.sort((a,b) => a.start - b.start);

    const taskNames = tasks.map(t => t.name);
    const categories = [...new Set(tasks.map(t => t.category))];
    const catColorMap = {};
    categories.forEach((c, i) => catColorMap[c] = colors[i % colors.length]);

    // Date range
    const allDates = tasks.flatMap(t => [t.start, t.end]);
    const minDate = new Date(Math.min(...allDates));
    const maxDate = new Date(Math.max(...allDates));

    // Build bars
    const seriesData = tasks.map((t, i) => ({
        name: t.name,
        value: [i, t.start.getTime(), t.end.getTime(), t.progress, t.category],
        itemStyle: { color: catColorMap[t.category] || colors[0] },
    }));

    // Progress overlay
    const progressData = tasks.filter(t => t.progress >= 0).map((t, i) => {
        const dur = t.end.getTime() - t.start.getTime();
        const progEnd = t.start.getTime() + dur * (t.progress / 100);
        return {
            value: [i, t.start.getTime(), progEnd, t.progress, t.category],
            itemStyle: { color: 'rgba(255,255,255,0.35)' },
        };
    });

    const title = document.getElementById('chartTitle').value || '';
    const subtitle = document.getElementById('chartSubtitle').value || '';
    const showLegend = document.getElementById('showLegend').checked;

    const option = {
        backgroundColor: bg === 'transparent' ? 'transparent' : bg,
        color: colors,
        textStyle: { fontFamily: 'Inter, sans-serif', fontSize: fs, color: textColor },
        title: {
            text: title, subtext: subtitle, left: 'center', top: 8,
            textStyle: { fontSize: fs + 6, fontWeight: 700, color: textColor, fontFamily: 'Inter, sans-serif' },
            subtextStyle: { fontSize: fs - 1, color: isDark ? '#94a3b8' : '#64748b', fontFamily: 'Inter, sans-serif' },
        },
        tooltip: {
            formatter: p => {
                const v = p.value;
                const s = new Date(v[1]).toLocaleDateString();
                const e = new Date(v[2]).toLocaleDateString();
                let tip = `<b>${p.name}</b><br/>Start: ${s}<br/>End: ${e}`;
                if (v[3] >= 0) tip += `<br/>Progress: ${v[3]}%`;
                if (v[4]) tip += `<br/>Category: ${v[4]}`;
                return tip;
            },
        },
        grid: { top: 80, right: 40, bottom: 50, left: Math.max(120, Math.max(...taskNames.map(n => n.length)) * 7) },
        xAxis: {
            type: 'time',
            min: minDate.getTime(),
            max: maxDate.getTime(),
            axisLabel: { fontSize: fs - 2, color: textColor, fontFamily: 'Inter, sans-serif' },
            splitLine: { show: true, lineStyle: { color: isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.04)' } },
        },
        yAxis: {
            type: 'category',
            data: taskNames,
            inverse: true,
            axisLabel: { fontSize: fs - 2, color: textColor, fontFamily: 'Inter, sans-serif', width: 150, overflow: 'truncate' },
            axisTick: { show: false },
            axisLine: { show: false },
        },
        series: [
            {
                type: 'custom',
                renderItem: (params, api) => {
                    const catIdx = api.value(0);
                    const start = api.coord([api.value(1), catIdx]);
                    const end = api.coord([api.value(2), catIdx]);
                    const height = api.size([0, 1])[1] * 0.6;
                    return {
                        type: 'rect',
                        shape: { x: start[0], y: start[1] - height/2, width: Math.max(end[0] - start[0], 2), height },
                        style: { ...api.style(), fill: api.visual('color') },
                        styleEmphasis: api.styleEmphasis(),
                    };
                },
                encode: { x: [1, 2], y: 0 },
                data: seriesData,
            },
            ...(progressData.length ? [{
                type: 'custom',
                renderItem: (params, api) => {
                    const catIdx = api.value(0);
                    const start = api.coord([api.value(1), catIdx]);
                    const end = api.coord([api.value(2), catIdx]);
                    const height = api.size([0, 1])[1] * 0.6;
                    return {
                        type: 'rect',
                        shape: { x: start[0], y: start[1] - height/2, width: Math.max(end[0] - start[0], 2), height },
                        style: { fill: 'rgba(255,255,255,0.35)' },
                    };
                },
                encode: { x: [1, 2], y: 0 },
                data: progressData,
                silent: true,
            }] : []),
        ],
    };

    if (showLegend && categories.length > 1) {
        option.legend = {
            data: categories.filter(c => c),
            bottom: 10, left: 'center',
            textStyle: { fontSize: fs - 2, color: textColor, fontFamily: 'Inter, sans-serif' },
        };
        // Add invisible series for legend
        for (const cat of categories.filter(c => c)) {
            option.series.push({ type: 'bar', name: cat, data: [], itemStyle: { color: catColorMap[cat] } });
        }
    }

    state.chartInstance.setOption(option);
}

// ── Step 4: Export ───────────────────────────────────────────

function initExportControls() {
    const presetSel = document.getElementById('exportPreset');
    presetSel.addEventListener('change', () => {
        const custom = presetSel.value === 'custom';
        document.getElementById('customSizeFields').style.display = custom ? 'block' : 'none';
        generateExportPreview();
    });

    document.getElementById('exportFormat').addEventListener('change', generateExportPreview);
    document.getElementById('exportScale').addEventListener('change', generateExportPreview);
    document.getElementById('exportWidth').addEventListener('change', generateExportPreview);
    document.getElementById('exportHeight').addEventListener('change', generateExportPreview);

    document.getElementById('btnDownload').addEventListener('click', downloadImage);
    document.getElementById('btnCopyClipboard').addEventListener('click', copyToClipboard);
}

function getExportDimensions() {
    const preset = document.getElementById('exportPreset').value;
    if (preset === 'custom') {
        return {
            width: parseInt(document.getElementById('exportWidth').value) || 1920,
            height: parseInt(document.getElementById('exportHeight').value) || 1080,
        };
    }
    const [w, h] = preset.split('x').map(Number);
    return { width: w, height: h };
}

function generateExportPreview() {
    const { width, height } = getExportDimensions();
    const scale = parseInt(document.getElementById('exportScale').value) || 2;
    const format = document.getElementById('exportFormat').value;
    const preview = document.getElementById('exportPreview');

    // Create offscreen chart
    const offDiv = document.createElement('div');
    offDiv.style.width = width + 'px';
    offDiv.style.height = height + 'px';
    offDiv.style.position = 'absolute';
    offDiv.style.left = '-9999px';
    document.body.appendChild(offDiv);

    const offChart = echarts.init(offDiv, null, { renderer: 'svg', width, height });
    const currentOption = state.chartInstance ? state.chartInstance.getOption() : {};
    offChart.setOption(currentOption);

    setTimeout(() => {
        if (format === 'svg') {
            const svgStr = offChart.renderToSVGString();
            preview.innerHTML = `<img src="data:image/svg+xml;charset=utf-8,${encodeURIComponent(svgStr)}" style="max-width:100%;height:auto">`;
        } else {
            const url = offChart.getDataURL({ type: 'png', pixelRatio: scale, backgroundColor: state.currentBg === 'transparent' ? 'transparent' : (state.currentBg || '#ffffff') });
            preview.innerHTML = `<img src="${url}" style="max-width:100%;height:auto">`;
        }
        offChart.dispose();
        offDiv.remove();
    }, 300);
}

function downloadImage() {
    const { width, height } = getExportDimensions();
    const scale = parseInt(document.getElementById('exportScale').value) || 2;
    const format = document.getElementById('exportFormat').value;
    const status = document.getElementById('exportStatus');

    const offDiv = document.createElement('div');
    offDiv.style.width = width + 'px';
    offDiv.style.height = height + 'px';
    offDiv.style.position = 'absolute';
    offDiv.style.left = '-9999px';
    document.body.appendChild(offDiv);

    const offChart = echarts.init(offDiv, null, { renderer: 'svg', width, height });
    offChart.setOption(state.chartInstance.getOption());

    setTimeout(() => {
        let dataUrl, filename;
        if (format === 'svg') {
            const svgStr = offChart.renderToSVGString();
            dataUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgStr);
            filename = 'chart.svg';
        } else {
            dataUrl = offChart.getDataURL({ type: 'png', pixelRatio: scale, backgroundColor: state.currentBg === 'transparent' ? 'transparent' : (state.currentBg || '#ffffff') });
            filename = `chart_${width}x${height}@${scale}x.png`;
        }

        const a = document.createElement('a');
        a.href = dataUrl;
        a.download = filename;
        a.click();

        offChart.dispose();
        offDiv.remove();
        status.textContent = `Downloaded ${filename} (${width*scale}×${height*scale} pixels)`;
    }, 300);
}

async function copyToClipboard() {
    const { width, height } = getExportDimensions();
    const scale = parseInt(document.getElementById('exportScale').value) || 2;
    const status = document.getElementById('exportStatus');

    const offDiv = document.createElement('div');
    offDiv.style.width = width + 'px';
    offDiv.style.height = height + 'px';
    offDiv.style.position = 'absolute';
    offDiv.style.left = '-9999px';
    document.body.appendChild(offDiv);

    const offChart = echarts.init(offDiv, null, { renderer: 'canvas', width, height, devicePixelRatio: scale });
    offChart.setOption(state.chartInstance.getOption());

    setTimeout(async () => {
        try {
            const canvas = offDiv.querySelector('canvas');
            const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
            await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
            status.textContent = 'Copied to clipboard! Paste directly into PowerPoint.';
        } catch (e) {
            status.textContent = 'Clipboard copy failed. Please use Download instead.';
        }
        offChart.dispose();
        offDiv.remove();
    }, 300);
}

// Make functions available globally for inline event handlers
window.removeFromPivot = removeFromPivot;
window.changeAgg = changeAgg;
