/**
 * Data Analysis Dashboard Logic Engine
 * 100% Client-side processing using Chart.js, Plotly.js, jStat, XLSX
 */

// Global State
let rawDataset = [];
let columnNames = [];
let numericColumns = [];
let categoricalColumns = [];

// Chart Instances
let timeSeriesChart = null;
let regressionChart = null;

// Initialize on DOM Loaded
document.addEventListener('DOMContentLoaded', () => {
  setupEventListeners();
  loadDefaultData();
});

function setupEventListeners() {
  // Tab Switching
  const navTabs = document.querySelectorAll('.nav-tab');
  navTabs.forEach(tab => {
    tab.addEventListener('click', () => {
      navTabs.forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      
      tab.classList.add('active');
      const target = tab.getAttribute('data-tab');
      document.getElementById(target).classList.add('active');

      // Trigger redraw/update if needed
      handleTabSwitch(target);
    });
  });

  // File Upload
  const fileInput = document.getElementById('excelFileInput');
  const uploadBtn = document.getElementById('uploadFileBtn');
  const dropZone = document.getElementById('dropZone');

  uploadBtn.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleFileUpload);

  // Drag & Drop
  ['dragenter', 'dragover'].forEach(eventName => {
    window.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('hidden');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (eventName === 'dragleave' && e.target === dropZone) {
        dropZone.classList.add('hidden');
      }
    });
  });

  dropZone.addEventListener('drop', (e) => {
    dropZone.classList.add('hidden');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      processExcelFile(files[0]);
    }
  });

  // Reset to default
  document.getElementById('resetDataBtn').addEventListener('click', loadDefaultData);

  // Variable Select Change Handlers
  document.getElementById('tsColSelect').addEventListener('change', renderTimeSeriesSection);
  document.getElementById('tsAnomalyMethod').addEventListener('change', renderTimeSeriesSection);
  document.getElementById('regXSelect').addEventListener('change', renderRegressionSection);
  document.getElementById('regYSelect').addEventListener('change', renderRegressionSection);
  document.getElementById('distColSelect').addEventListener('change', renderDistributionSection);
  document.getElementById('groupValColSelect').addEventListener('change', renderHypothesisSection);
  document.getElementById('groupCategoryColSelect').addEventListener('change', renderHypothesisSection);
}

// 1. Data Ingestion
function loadDefaultData() {
  if (typeof DEFAULT_SENSOR_DATA !== 'undefined') {
    processDataArray(DEFAULT_SENSOR_DATA, 'sensor_data_dummy.xlsx (내장 데이터)');
  }
}

function handleFileUpload(e) {
  const file = e.target.files[0];
  if (!file) return;
  processExcelFile(file);
}

function processExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    
    if (jsonData.length === 0) {
      alert('파일에 유효한 데이터가 없습니다.');
      return;
    }
    processDataArray(jsonData, file.name);
  };
  reader.readAsArrayBuffer(file);
}

function processDataArray(data, sourceName) {
  rawDataset = data;
  document.getElementById('currentFileName').textContent = sourceName;
  document.getElementById('totalRowCount').textContent = data.length.toLocaleString();

  // Extract columns
  columnNames = Object.keys(data[0] || {});
  document.getElementById('totalColCount').textContent = columnNames.length;

  // Identify numeric vs categorical
  numericColumns = [];
  categoricalColumns = [];

  columnNames.forEach(col => {
    const sampleValues = data.map(r => r[col]).filter(v => v !== null && v !== undefined && v !== '');
    const numValid = sampleValues.filter(v => typeof v === 'number' || (!isNaN(Number(v)) && typeof v !== 'boolean')).length;
    if (numValid / sampleValues.length > 0.8) {
      numericColumns.push(col);
    } else {
      categoricalColumns.push(col);
    }
  });

  // Populate Select Boxes
  populateSelectOptions();

  // Initial Overview & Stats Table Render
  renderDataPreviewTable();
  renderDescriptiveStats();

  // Trigger default renders for all tabs
  renderTimeSeriesSection();
  renderCorrelationSection();
  renderRegressionSection();
  renderDistributionSection();
  renderHypothesisSection();
}

function populateSelectOptions() {
  const populate = (selectId, options, defaultIdx = 0) => {
    const select = document.getElementById(selectId);
    if (!select) return;
    select.innerHTML = '';
    options.forEach((opt, idx) => {
      const optionEl = document.createElement('option');
      optionEl.value = opt;
      optionEl.textContent = opt;
      if (idx === defaultIdx) optionEl.selected = true;
      select.appendChild(optionEl);
    });
  };

  populate('tsColSelect', numericColumns, 0);
  populate('regXSelect', numericColumns, 0);
  populate('regYSelect', numericColumns, Math.min(1, numericColumns.length - 1));
  populate('distColSelect', numericColumns, 0);
  populate('groupValColSelect', numericColumns, 0);
  populate('groupCategoryColSelect', categoricalColumns.length > 0 ? categoricalColumns : columnNames, 0);
}

function handleTabSwitch(tabId) {
  setTimeout(() => {
    if (tabId === 'correlation') renderCorrelationSection();
    if (tabId === 'timeseries' && timeSeriesChart) timeSeriesChart.resize();
    if (tabId === 'regression' && regressionChart) regressionChart.resize();
    if (tabId === 'distribution') renderDistributionSection();
    if (tabId === 'hypothesis') renderHypothesisSection();
  }, 50);
}

// 2. Data Preview Table
function renderDataPreviewTable() {
  const tableHead = document.getElementById('previewTableHead');
  const tableBody = document.getElementById('previewTableBody');
  if (!tableHead || !tableBody) return;

  // Header
  tableHead.innerHTML = `<tr>${columnNames.map(c => `<th class="p-3 border-b border-gray-700 text-xs font-semibold text-gray-400 uppercase tracking-wider">${c}</th>`).join('')}</tr>`;

  // Rows (first 10)
  const previewRows = rawDataset.slice(0, 10);
  tableBody.innerHTML = previewRows.map(row => {
    return `<tr class="hover:bg-gray-800/40 transition">
      ${columnNames.map(col => {
        let val = row[col];
        if (col === '이벤트') {
          let badgeClass = val === '정상' ? 'badge-normal' : (val === '주의' ? 'badge-warning' : 'badge-danger');
          return `<td class="p-3 border-b border-gray-800 text-xs"><span class="stat-badge ${badgeClass}">${val}</span></td>`;
        }
        if (typeof val === 'number') {
          val = Number.isInteger(val) ? val : val.toFixed(4);
        }
        return `<td class="p-3 border-b border-gray-800 text-xs text-gray-300">${val ?? '-'}</td>`;
      }).join('')}
    </tr>`;
  }).join('');
}

// 3. Descriptive Statistics
function getColumnStats(colName) {
  const vals = rawDataset.map(r => Number(r[colName])).filter(v => !isNaN(v) && v !== null);
  if (vals.length === 0) return null;

  const n = vals.length;
  const mean = jStat.mean(vals);
  const std = jStat.stdev(vals, true); // sample stdev
  const variance = jStat.variance(vals, true);
  const min = jStat.min(vals);
  const max = jStat.max(vals);
  const median = jStat.median(vals);
  const q1 = jStat.percentile(vals, 0.25);
  const q3 = jStat.percentile(vals, 0.75);
  const iqr = q3 - q1;
  const skewness = jStat.skewness(vals);
  const kurtosis = jStat.kurtosis(vals);

  return { n, mean, std, variance, min, max, median, q1, q3, iqr, skewness, kurtosis };
}

function renderDescriptiveStats() {
  const container = document.getElementById('descriptiveStatsCards');
  const tableContainer = document.getElementById('descriptiveStatsTable');
  if (!container || !tableContainer) return;

  let tableHtml = `
    <table class="data-table">
      <thead>
        <tr>
          <th>변수명</th>
          <th>데이터수 (N)</th>
          <th>평균 (Mean)</th>
          <th>표준편차 (Std)</th>
          <th>중앙값 (Median)</th>
          <th>최솟값 (Min)</th>
          <th>최댓값 (Max)</th>
          <th>IQR</th>
          <th>왜도 (Skew)</th>
          <th>첨도 (Kurt)</th>
        </tr>
      </thead>
      <tbody>
  `;

  numericColumns.forEach(col => {
    const s = getColumnStats(col);
    if (!s) return;
    tableHtml += `
      <tr>
        <td class="font-semibold text-blue-400">${col}</td>
        <td>${s.n}</td>
        <td>${s.mean.toFixed(4)}</td>
        <td>${s.std.toFixed(4)}</td>
        <td>${s.median.toFixed(4)}</td>
        <td>${s.min.toFixed(4)}</td>
        <td>${s.max.toFixed(4)}</td>
        <td>${s.iqr.toFixed(4)}</td>
        <td>${s.skewness.toFixed(3)}</td>
        <td>${s.kurtosis.toFixed(3)}</td>
      </tr>
    `;
  });

  tableHtml += `</tbody></table>`;
  tableContainer.innerHTML = tableHtml;

  // Render Top Quick Summary Cards for Key Sensor variables
  container.innerHTML = numericColumns.slice(0, 4).map(col => {
    const s = getColumnStats(col);
    if (!s) return '';
    return `
      <div class="glass-card p-4 flex flex-col justify-between">
        <div class="flex items-center justify-between mb-2">
          <span class="text-xs font-semibold text-gray-400 truncate max-w-[140px]" title="${col}">${col}</span>
          <span class="stat-badge bg-blue-500/10 text-blue-400 border border-blue-500/20">Mean: ${s.mean.toFixed(2)}</span>
        </div>
        <div class="flex items-baseline gap-2 mb-1">
          <span class="text-2xl font-bold text-white">${s.median.toFixed(3)}</span>
          <span class="text-xs text-gray-400">중앙값</span>
        </div>
        <div class="flex justify-between text-xs text-gray-400 border-t border-gray-800 pt-2 mt-2">
          <span>범위: ${s.min.toFixed(2)} ~ ${s.max.toFixed(2)}</span>
          <span>표준편차: ${s.std.toFixed(2)}</span>
        </div>
      </div>
    `;
  }).join('');
}

// 4. Time Series & Anomaly Detection
function renderTimeSeriesSection() {
  const col = document.getElementById('tsColSelect').value;
  const method = document.getElementById('tsAnomalyMethod').value;
  if (!col) return;

  const timeCol = columnNames.find(c => c.includes('시각') || c.includes('시간') || c.includes('time')) || columnNames[0];
  const labels = rawDataset.map((r, i) => r[timeCol] !== undefined ? String(r[timeCol]) : `#${i + 1}`);
  const values = rawDataset.map(r => Number(r[col]) || 0);

  // Stats for Anomaly Detection
  const stats = getColumnStats(col);
  let upperLimit = 0, lowerLimit = 0;
  let anomalies = [];

  if (method === '3sigma') {
    upperLimit = stats.mean + 2.5 * stats.std;
    lowerLimit = Math.max(0, stats.mean - 2.5 * stats.std);
    anomalies = values.map((v, i) => (v > upperLimit || v < lowerLimit) ? { x: labels[i], y: v, index: i } : null).filter(Boolean);
  } else if (method === 'iqr') {
    upperLimit = stats.q3 + 1.5 * stats.iqr;
    lowerLimit = Math.max(0, stats.q1 - 1.5 * stats.iqr);
    anomalies = values.map((v, i) => (v > upperLimit || v < lowerLimit) ? { x: labels[i], y: v, index: i } : null).filter(Boolean);
  } else {
    // 이동평균 3구간
    const windowSize = 3;
    const ma = values.map((v, idx, arr) => {
      const sub = arr.slice(Math.max(0, idx - windowSize + 1), idx + 1);
      return sub.reduce((a, b) => a + b, 0) / sub.length;
    });
    upperLimit = stats.mean + 2 * stats.std;
    lowerLimit = stats.mean - 2 * stats.std;
  }

  // Render Time Series Chart
  const ctx = document.getElementById('timeSeriesCanvas').getContext('2d');
  if (timeSeriesChart) timeSeriesChart.destroy();

  const anomalyPointsData = values.map(v => (v > upperLimit || v < lowerLimit) ? v : null);

  timeSeriesChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: labels,
      datasets: [
        {
          label: `${col} 측정값`,
          data: values,
          borderColor: '#3b82f6',
          backgroundColor: 'rgba(59, 130, 246, 0.1)',
          fill: true,
          tension: 0.3,
          pointRadius: 3,
          pointHoverRadius: 6,
          borderWidth: 2
        },
        {
          label: '이상치(Anomaly)',
          data: anomalyPointsData,
          borderColor: 'transparent',
          backgroundColor: '#ef4444',
          pointRadius: 6,
          pointHoverRadius: 8,
          pointStyle: 'circle',
          showLine: false
        },
        {
          label: `상한 임계선 (${upperLimit.toFixed(3)})`,
          data: labels.map(() => upperLimit),
          borderColor: 'rgba(239, 68, 68, 0.7)',
          borderDash: [5, 5],
          pointRadius: 0,
          fill: false
        },
        {
          label: `하한 임계선 (${lowerLimit.toFixed(3)})`,
          data: labels.map(() => lowerLimit),
          borderColor: 'rgba(245, 158, 11, 0.7)',
          borderDash: [5, 5],
          pointRadius: 0,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: {
        mode: 'index',
        intersect: false,
      },
      plugins: {
        legend: {
          labels: { color: '#9ca3af', font: { family: 'Pretendard' } }
        },
        tooltip: {
          backgroundColor: '#1e293b',
          titleColor: '#f8fafc',
          bodyColor: '#cbd5e1',
          borderColor: 'rgba(255,255,255,0.1)',
          borderWidth: 1
        }
      },
      scales: {
        x: {
          grid: { color: 'rgba(255,255,255,0.05)' },
          ticks: { color: '#9ca3af', maxRotation: 45 }
        },
        y: {
          grid: { color: 'rgba(255,255,255,0.05)' },
          ticks: { color: '#9ca3af' }
        }
      }
    }
  });

  // Update Anomaly Insights
  const count = anomalies.length;
  document.getElementById('anomalyCount').textContent = `${count}건 감지됨`;
  document.getElementById('anomalyInsightText').innerHTML = `
    <strong>[${col}]</strong> 변수 대상 <em>${method === '3sigma' ? '3-Sigma' : (method === 'iqr' ? 'IQR 1.5x' : '이동평균')}</em> 기법 적용 결과:
    총 <strong>${values.length}</strong>개 데이터 중 <strong>${count}</strong>개의 이상치/경고치 포인트가 감지되었습니다.
    <ul class="list-disc pl-5 mt-2 space-y-1 text-gray-300">
      <li>정상 예상 범위: [${lowerLimit.toFixed(4)} ~ ${upperLimit.toFixed(4)}]</li>
      <li>이상치 감지 시각/인덱스: ${count > 0 ? anomalies.map(a => `${a.x} (${a.y.toFixed(3)})`).slice(0, 5).join(', ') + (count > 5 ? ' 외' : '') : '이상 없음'}</li>
    </ul>
  `;
}

// 5. Correlation Matrix & Heatmap
function renderCorrelationSection() {
  const cols = numericColumns;
  if (cols.length === 0) return;

  const matrix = [];
  for (let i = 0; i < cols.length; i++) {
    const row = [];
    const valsA = rawDataset.map(r => Number(r[cols[i]]) || 0);
    for (let j = 0; j < cols.length; j++) {
      const valsB = rawDataset.map(r => Number(r[cols[j]]) || 0);
      const r = jStat.corrcoeff(valsA, valsB);
      row.push(isNaN(r) ? 0 : Number(r.toFixed(3)));
    }
    matrix.push(row);
  }

  // Plotly Heatmap
  const data = [{
    z: matrix,
    x: cols,
    y: cols,
    type: 'heatmap',
    colorscale: [
      [0, '#3b82f6'],
      [0.5, '#1e293b'],
      [1, '#ef4444']
    ],
    zmin: -1,
    zmax: 1,
    hoverongaps: false
  }];

  const layout = {
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: '#9ca3af', family: 'Pretendard' },
    margin: { t: 30, r: 30, b: 80, l: 120 },
    xaxis: { tickangle: -45, gridcolor: 'transparent' },
    yaxis: { autorange: 'reversed', gridcolor: 'transparent' }
  };

  Plotly.newPlot('correlationHeatmapDiv', data, layout, { responsive: true, displayModeBar: false });

  // Find Top Strongest Correlations
  let pairs = [];
  for (let i = 0; i < cols.length; i++) {
    for (let j = i + 1; j < cols.length; j++) {
      pairs.push({
        pair: `${cols[i]} ↔ ${cols[j]}`,
        r: matrix[i][j],
        absR: Math.abs(matrix[i][j])
      });
    }
  }
  pairs.sort((a, b) => b.absR - a.absR);

  const topPairs = pairs.slice(0, 5);
  const pairHtml = topPairs.map(p => {
    const strength = p.absR > 0.7 ? '강한 상관관계' : (p.absR > 0.4 ? '보통 상관관계' : '약한/무상관');
    const color = p.r > 0 ? 'text-rose-400' : 'text-blue-400';
    return `
      <div class="flex items-center justify-between p-2.5 rounded bg-gray-800/40 border border-gray-700/50">
        <span class="text-xs text-gray-300 font-medium">${p.pair}</span>
        <div class="flex items-center gap-3">
          <span class="text-xs font-mono font-bold ${color}">${p.r > 0 ? '+' : ''}${p.r.toFixed(3)}</span>
          <span class="text-[11px] text-gray-400">${strength}</span>
        </div>
      </div>
    `;
  }).join('');

  document.getElementById('topCorrelationsList').innerHTML = pairHtml || '<p class="text-gray-500 text-xs">상관 데이터가 충분하지 않습니다.</p>';
}

// 6. Linear Regression
function renderRegressionSection() {
  const xCol = document.getElementById('regXSelect').value;
  const yCol = document.getElementById('regYSelect').value;
  if (!xCol || !yCol) return;

  const points = rawDataset
    .map(r => ({ x: Number(r[xCol]), y: Number(r[yCol]) }))
    .filter(p => !isNaN(p.x) && !isNaN(p.y));

  if (points.length < 2) return;

  const xVals = points.map(p => p.x);
  const yVals = points.map(p => p.y);

  // Linear Regression Calculation
  const xMean = jStat.mean(xVals);
  const yMean = jStat.mean(yVals);
  
  let num = 0, den = 0;
  for (let i = 0; i < points.length; i++) {
    num += (points[i].x - xMean) * (points[i].y - yMean);
    den += Math.pow(points[i].x - xMean, 2);
  }
  
  const slope = den !== 0 ? num / den : 0;
  const intercept = yMean - slope * xMean;
  const corr = jStat.corrcoeff(xVals, yVals);
  const r2 = Math.pow(isNaN(corr) ? 0 : corr, 2);

  // Residual standard error & F-test p-value approximation
  const n = points.length;
  const residuals = points.map(p => p.y - (slope * p.x + intercept));
  const ssRes = residuals.reduce((sum, r) => sum + r * r, 0);
  const ssTot = yVals.reduce((sum, y) => sum + Math.pow(y - yMean, 2), 0);
  const dfRes = n - 2;
  const fStat = (dfRes > 0 && ssRes > 0) ? ((ssTot - ssRes) / 1) / (ssRes / dfRes) : 0;
  const pVal = dfRes > 0 ? (1 - jStat.centralF.cdf(fStat, 1, dfRes)) : 1;

  // UI Stats Update
  document.getElementById('regEquation').textContent = `y = ${slope.toFixed(4)}x ${intercept >= 0 ? '+' : '-'} ${Math.abs(intercept).toFixed(4)}`;
  document.getElementById('regR2').textContent = r2.toFixed(4);
  document.getElementById('regPVal').textContent = pVal < 0.001 ? '< 0.001 (매우 유의)' : pVal.toFixed(4);

  // Line Points for plotting
  const minX = Math.min(...xVals);
  const maxX = Math.max(...xVals);
  const lineData = [
    { x: minX, y: slope * minX + intercept },
    { x: maxX, y: slope * maxX + intercept }
  ];

  const ctx = document.getElementById('regressionCanvas').getContext('2d');
  if (regressionChart) regressionChart.destroy();

  regressionChart = new Chart(ctx, {
    type: 'scatter',
    data: {
      datasets: [
        {
          label: '실제 관측치',
          data: points,
          backgroundColor: '#3b82f6',
          borderColor: '#60a5fa',
          pointRadius: 5,
          pointHoverRadius: 7
        },
        {
          label: `회귀선 (R²=${r2.toFixed(3)})`,
          data: lineData,
          type: 'line',
          borderColor: '#ef4444',
          borderWidth: 2,
          pointRadius: 0,
          fill: false
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#9ca3af' } },
        tooltip: {
          backgroundColor: '#1e293b',
          callbacks: {
            label: (ctx) => `(${ctx.parsed.x.toFixed(3)}, ${ctx.parsed.y.toFixed(3)})`
          }
        }
      },
      scales: {
        x: {
          title: { display: true, text: xCol, color: '#9ca3af' },
          grid: { color: 'rgba(255,255,255,0.05)' },
          ticks: { color: '#9ca3af' }
        },
        y: {
          title: { display: true, text: yCol, color: '#9ca3af' },
          grid: { color: 'rgba(255,255,255,0.05)' },
          ticks: { color: '#9ca3af' }
        }
      }
    }
  });

  // Interpretations
  document.getElementById('regInsightText').innerHTML = `
    <strong>[회귀분석 결과 요약]</strong><br>
    독립변수(X: <code>${xCol}</code>)가 1단위 증가할 때 종속변수(Y: <code>${yCol}</code>)는 평균적으로 <strong>${slope.toFixed(4)}</strong>만큼 ${slope >= 0 ? '증가' : '감소'}합니다.<br>
    설명력(결정계수 $R^2$)은 <strong>${(r2 * 100).toFixed(2)}%</strong>이며, 모델의 통계적 유의확률 p-value는 <strong>${pVal < 0.001 ? '< 0.001' : pVal.toFixed(4)}</strong>로 ${pVal < 0.05 ? '<span class="text-emerald-400 font-semibold">통계적으로 유의미합니다 (p < 0.05).</span>' : '<span class="text-amber-400 font-semibold">유의수준 5%에서 기각할 수 없습니다.</span>'}
  `;
}

// 7. Distribution & Normality Test
function renderDistributionSection() {
  const col = document.getElementById('distColSelect').value;
  if (!col) return;

  const vals = rawDataset.map(r => Number(r[col])).filter(v => !isNaN(v));
  if (vals.length === 0) return;

  // 1) Histogram & Boxplot via Plotly
  const histData = [
    {
      x: vals,
      type: 'histogram',
      name: '빈도 분포',
      marker: { color: 'rgba(59, 130, 246, 0.7)' },
      opacity: 0.8
    }
  ];

  const histLayout = {
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: '#9ca3af', family: 'Pretendard' },
    margin: { t: 20, r: 20, b: 40, l: 40 },
    xaxis: { gridcolor: 'rgba(255,255,255,0.05)', title: col },
    yaxis: { gridcolor: 'rgba(255,255,255,0.05)', title: '빈도(Count)' },
    bargap: 0.08
  };

  Plotly.newPlot('distributionPlotDiv', histData, histLayout, { responsive: true, displayModeBar: false });

  // Boxplot
  const boxData = [{
    y: vals,
    type: 'box',
    name: col,
    marker: { color: '#8b5cf6' },
    boxpoints: 'all',
    jitter: 0.3,
    pointpos: -1.8
  }];
  const boxLayout = {
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: '#9ca3af', family: 'Pretendard' },
    margin: { t: 20, r: 20, b: 40, l: 50 },
    yaxis: { gridcolor: 'rgba(255,255,255,0.05)' }
  };
  Plotly.newPlot('boxPlotDiv', boxData, boxLayout, { responsive: true, displayModeBar: false });

  // 2) Q-Q Plot (Normal Quantiles vs Sample Quantiles)
  const sorted = [...vals].sort((a, b) => a - b);
  const n = sorted.length;
  const mean = jStat.mean(sorted);
  const std = jStat.stdev(sorted, true);

  const theoreticalQ = [];
  for (let i = 1; i <= n; i++) {
    const p = (i - 0.5) / n;
    theoreticalQ.push(jStat.normal.inv(p, mean, std));
  }

  const qqData = [
    {
      x: theoreticalQ,
      y: sorted,
      mode: 'markers',
      type: 'scatter',
      name: '샘플 분위수',
      marker: { color: '#06b6d4', size: 7 }
    },
    {
      x: [Math.min(...theoreticalQ), Math.max(...theoreticalQ)],
      y: [Math.min(...theoreticalQ), Math.max(...theoreticalQ)],
      mode: 'lines',
      name: '정규 기준선',
      line: { color: '#ef4444', dash: 'dash' }
    }
  ];

  const qqLayout = {
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: '#9ca3af', family: 'Pretendard' },
    margin: { t: 20, r: 20, b: 40, l: 50 },
    xaxis: { gridcolor: 'rgba(255,255,255,0.05)', title: '이론적 정규분포 분위수' },
    yaxis: { gridcolor: 'rgba(255,255,255,0.05)', title: '샘플 분위수' }
  };
  Plotly.newPlot('qqPlotDiv', qqData, qqLayout, { responsive: true, displayModeBar: false });

  // Normality Assessment (Skewness & Kurtosis)
  const skew = jStat.skewness(vals);
  const kurt = jStat.kurtosis(vals);
  const isNormalSkew = Math.abs(skew) < 1.0;
  const isNormalKurt = Math.abs(kurt) < 2.0;

  document.getElementById('normalityInsightText').innerHTML = `
    <strong>[${col} 정규성 진단]</strong><br>
    - 왜도 (Skewness): <strong>${skew.toFixed(3)}</strong> (${Math.abs(skew) < 0.5 ? '대칭적' : (skew > 0 ? '우측 꼬리가 긴 분포' : '좌측 꼬리가 긴 분포')})<br>
    - 첨도 (Kurtosis): <strong>${kurt.toFixed(3)}</strong> (${Math.abs(kurt) < 0.5 ? '정규분포 수준의 완만한 높이' : (kurt > 0 ? '중앙에 뾰족하게 집중' : '완만하게 평평한 분포')})<br>
    - 종합 판단: ${isNormalSkew && isNormalKurt 
      ? '<span class="text-emerald-400 font-semibold">왜도와 첨도가 기준 범위(-1 ~ +1, -2 ~ +2) 내에 있어 대체로 정규성을 만족합니다.</span>' 
      : '<span class="text-amber-400 font-semibold">정규분포 기준에서 일부 이탈이 감지됩니다 (비모수 분석 기법 고려 권장).</span>'}
  `;
}

// 8. Hypothesis Testing & Group Analysis
function renderHypothesisSection() {
  const valCol = document.getElementById('groupValColSelect').value;
  const catCol = document.getElementById('groupCategoryColSelect').value;
  if (!valCol || !catCol) return;

  // Grouping
  const groups = {};
  rawDataset.forEach(r => {
    const cat = r[catCol] !== undefined ? String(r[catCol]) : '미분류';
    const val = Number(r[valCol]);
    if (!isNaN(val)) {
      if (!groups[cat]) groups[cat] = [];
      groups[cat].push(val);
    }
  });

  const groupKeys = Object.keys(groups);
  if (groupKeys.length < 2) {
    document.getElementById('hypothesisResultDiv').innerHTML = `
      <div class="p-4 bg-amber-500/10 border border-amber-500/20 text-amber-400 rounded-lg text-sm">
        선택한 범주 변수에 최소 2개 이상의 서로 다른 그룹이 필요합니다.
      </div>
    `;
    return;
  }

  // Group Boxplot
  const plotData = groupKeys.map(k => ({
    y: groups[k],
    type: 'box',
    name: `${k} (n=${groups[k].length})`,
    boxmean: true
  }));

  const layout = {
    paper_bgcolor: 'transparent',
    plot_bgcolor: 'transparent',
    font: { color: '#9ca3af', family: 'Pretendard' },
    margin: { t: 20, r: 20, b: 40, l: 50 },
    yaxis: { gridcolor: 'rgba(255,255,255,0.05)', title: valCol }
  };
  Plotly.newPlot('groupComparisonPlotDiv', plotData, layout, { responsive: true, displayModeBar: false });

  // Statistical Test: 2 groups -> Two-sample t-test, 3+ groups -> ANOVA
  let testName = '';
  let testStat = 0;
  let pVal = 0;
  let df = 0;

  if (groupKeys.length === 2) {
    testName = '독립표본 t-검정 (Independent Two-Sample t-Test)';
    const g1 = groups[groupKeys[0]];
    const g2 = groups[groupKeys[1]];
    const t = jStat.ttest(g1, g2);
    testStat = Math.abs(t);
    df = g1.length + g2.length - 2;
    pVal = (1 - jStat.studentt.cdf(testStat, df)) * 2;
  } else {
    testName = '일원배치 분산분석 (One-Way ANOVA)';
    const sampleArrays = groupKeys.map(k => groups[k]);
    const f = jStat.anovafscore(...sampleArrays);
    testStat = isNaN(f) ? 0 : f;
    const k = groupKeys.length;
    const N = sampleArrays.reduce((sum, arr) => sum + arr.length, 0);
    const dfBetween = k - 1;
    const dfWithin = N - k;
    pVal = (dfWithin > 0 && testStat > 0) ? (1 - jStat.centralF.cdf(testStat, dfBetween, dfWithin)) : 1;
  }

  const isSignificant = pVal < 0.05;

  document.getElementById('hypothesisResultDiv').innerHTML = `
    <div class="glass-card p-5">
      <div class="flex items-center justify-between border-b border-gray-800 pb-3 mb-4">
        <div>
          <h4 class="text-sm font-bold text-white">${testName}</h4>
          <p class="text-xs text-gray-400 mt-0.5">비교 변수: <strong>${valCol}</strong> / 그룹 기준: <strong>${catCol}</strong></p>
        </div>
        <span class="stat-badge ${isSignificant ? 'badge-danger' : 'badge-normal'}">
          ${isSignificant ? '통계적 유의차 있음 (p < 0.05)' : '유의차 없음 (p ≥ 0.05)'}
        </span>
      </div>

      <div class="grid grid-cols-3 gap-4 mb-4">
        <div class="p-3 bg-gray-800/40 rounded-lg border border-gray-700/50">
          <span class="text-xs text-gray-400 block mb-1">검정통계량 (${groupKeys.length === 2 ? 't' : 'F'})</span>
          <span class="text-xl font-bold font-mono text-white">${testStat.toFixed(4)}</span>
        </div>
        <div class="p-3 bg-gray-800/40 rounded-lg border border-gray-700/50">
          <span class="text-xs text-gray-400 block mb-1">유의확률 (p-value)</span>
          <span class="text-xl font-bold font-mono ${isSignificant ? 'text-rose-400' : 'text-blue-400'}">
            ${pVal < 0.001 ? '< 0.001' : pVal.toFixed(4)}
          </span>
        </div>
        <div class="p-3 bg-gray-800/40 rounded-lg border border-gray-700/50">
          <span class="text-xs text-gray-400 block mb-1">비교 그룹 수</span>
          <span class="text-xl font-bold font-mono text-emerald-400">${groupKeys.length}개 그룹</span>
        </div>
      </div>

      <div class="insight-box text-xs text-gray-300">
        <strong>[검정 결과 해석]</strong><br>
        귀무가설($H_0$): 그룹 간 <code>${valCol}</code>의 평균 차이가 없다.<br>
        대립가설($H_1$): 그룹 간 <code>${valCol}</code>의 평균에 유의미한 차이가 있다.<br><br>
        유의확률 p-value는 <strong>${pVal < 0.001 ? '< 0.001' : pVal.toFixed(4)}</strong>로, 
        ${isSignificant 
          ? '<span class="text-rose-400 font-semibold">유의수준 5% 하에서 귀무가설을 기각합니다. 즉, 그룹 간에 통계적으로 유의미한 차이가 존재합니다.</span>'
          : '<span class="text-blue-400 font-semibold">유의수준 5% 하에서 귀무가설을 기각할 수 없습니다. 즉, 그룹 간 평균 차이가 통계적으로 유의하지 않습니다.</span>'}
      </div>
    </div>
  `;
}
