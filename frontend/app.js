/**
 * PDF & LaTeX to Word Converter — Frontend Application Logic (Unified DrawingML, OMML & PDF Decompiler Engine)
 * Communicates with FastAPI backend via REST API
 */

// ─── Configuration ───────────────────────────────────────────────
const API_BASE = window.location.origin;

// ─── State ───────────────────────────────────────────────────────
const state = {
  mode: 2,
  dpi: 400,
  strict: false,
  qaPdf: true,
  selectedFile: null,
  currentJobId: null,
  theme: localStorage.getItem('latex2word-theme') || 'dark',
};

// ─── Initialization ──────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  applyTheme(state.theme);
  selectMode(2);
  checkSystemStatus();
  loadDefaultTemplate();

  const dropzone = document.getElementById('dropzone');
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dropzone--dragover');
  });
  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dropzone--dragover');
  });
  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dropzone--dragover');
    if (e.dataTransfer.files.length) {
      const input = document.getElementById('file-input');
      input.files = e.dataTransfer.files;
      handleFileSelect(input);
    }
  });
});

// ─── Theme ───────────────────────────────────────────────────────
function toggleTheme() {
  state.theme = state.theme === 'dark' ? 'light' : 'dark';
  applyTheme(state.theme);
  localStorage.setItem('latex2word-theme', state.theme);
}

function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
}

// ─── Tab Switching ───────────────────────────────────────────────
function switchTab(tab) {
  const tabs = ['paste', 'upload'];
  tabs.forEach(t => {
    document.getElementById(`tab-${t}`).classList.toggle('tab-panel--active', t === tab);
    document.getElementById(`tab-${t}-btn`).classList.toggle('tab-btn--active', t === tab);
  });
}

// ─── Mode Selection ──────────────────────────────────────────────
function selectMode(mode) {
  state.mode = mode;
  document.getElementById('mode1-card').classList.toggle('radio-card--active', mode === 1);
  document.getElementById('mode2-card').classList.toggle('radio-card--active', mode === 2);
}

// ─── DPI Slider ──────────────────────────────────────────────────
function updateDPI(value) {
  state.dpi = parseInt(value);
  document.getElementById('dpi-value').textContent = `${value} DPI`;
}

// ─── Toggle Options ──────────────────────────────────────────────
function toggleOption(option) {
  if (option === 'strict') {
    state.strict = !state.strict;
    document.getElementById('toggle-strict').classList.toggle('toggle--active', state.strict);
  } else if (option === 'qa') {
    state.qaPdf = !state.qaPdf;
    document.getElementById('toggle-qa').classList.toggle('toggle--active', state.qaPdf);
  }
}

// ─── File Selection ──────────────────────────────────────────────
function handleFileSelect(input) {
  const file = input.files[0];
  if (!file) return;
  state.selectedFile = file;
  const display = document.getElementById('filename-display');
  display.textContent = `📎 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
  display.style.display = 'inline-flex';
  document.getElementById('btn-convert-upload').disabled = false;
}

// ─── Default Template ────────────────────────────────────────────
function loadDefaultTemplate() {
  const editor = document.getElementById('latex-editor');
  editor.value = `\\documentclass[12pt,a4paper]{article}
\\usepackage{amsmath,amssymb}
\\usepackage{tikz,pgfplots,tkz-tab}
\\usepackage{booktabs}
\\pgfplotsset{compat=1.18}

\\title{Mẫu Kiểm Thử LaTeX sang Word (Native DrawingML Shapes)}
\\author{Unified Core Engine}
\\date{\\today}

\\begin{document}
\\maketitle

\\textbf{Câu 1.} Cho hàm số bậc ba $y = f(x)$ có đồ thị như hình vẽ bên. Hỏi phương trình $f(f(x)) = 0$ có tất cả bao nhiêu nghiệm thực phân biệt?
\\immini{
Khảo sát hàm số hợp $g(x) = f(f(x))$:
\\[
x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}, \\quad \\binom{n}{k} = \\frac{n!}{k!(n-k)!}
\\]
\\choiceB{$3$.}{$7$.}{$9$.}{$5$.}
}{
\\begin{tikzpicture}
\\begin{axis}[
    axis lines=middle,
    xlabel={$x$}, ylabel={$y$},
    width=6.5cm, height=4.2cm,
    xmin=-2, xmax=2, ymin=-3.5, ymax=2.5
]
\\addplot[blue, thick, domain=-1.65:1.65] {x^3 - 3*x};
\\end{axis}
\\end{tikzpicture}
}

\\textbf{Câu 2.} Cho hàm số $y = f(x)$ có bảng biến thiên như sau:
\\begin{center}
\\begin{tikzpicture}
    \\tkzTabInit[lgt=1.2, espcl=1.6, deltacl=0.6]
      {$x$ / 0.8, $f'(x)$ / 0.8, $f(x)$ / 2}
      {$-\\infty$, $-1$, $0$, $1$, $+\\infty$}
    \\tkzTabLine{ , -, 0, +, 0, -, 0, +, }
    \\tkzTabVar{+/ $+\\infty$, -/ $-2$, +/ $3$, -/ $-2$, +/ $+\\infty$}
\\end{tikzpicture}
\\end{center}
\\choiceC{$4$}{$6$}{$7$}{$5$}

\\end{document}`;
}

// ─── System Status ───────────────────────────────────────────────
async function checkSystemStatus() {
  try {
    const res = await fetch(`${API_BASE}/api/status`);
    const data = await res.json();
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');

    if (data.status === 'ready') {
      dot.className = 'status-dot status-dot--ready';
      text.textContent = 'Hệ thống sẵn sàng';
    } else {
      dot.className = 'status-dot status-dot--error';
      text.textContent = 'Thiếu công cụ';
    }

    window._systemTools = data.tools;
  } catch {
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');
    dot.className = 'status-dot status-dot--error';
    text.textContent = 'API không phản hồi';
  }
}

function showSystemStatus() {
  const tools = window._systemTools || {};
  const container = document.getElementById('modal-tools');
  const toolLabels = {
    'pdf_decompiler': 'PDF to LaTeX/OMML Decompiler (.pdf / .zip)',
    'drawingml_engine': 'Native Word Shapes (DrawingML Engine)',
    'math_tables_engine': 'Math Tables (tkz-tab, Choice, Tabular)',
    'omml_math_engine': 'Native OMML Math Engine (Binom, Cases, Matrix)',
    'latex2word_py': 'latex2word.py (Unified Core)',
    'libreoffice': 'MS Word / LibreOffice (QA PDF)',
  };

  container.innerHTML = Object.entries(toolLabels).map(([key, label]) => {
    const ok = tools[key];
    return `
      <div class="tool-check">
        <span class="tool-check__name">${label}</span>
        <span class="tool-check__status ${ok ? 'tool-check__status--ok' : 'tool-check__status--missing'}">
          ${ok ? '✅ Sẵn sàng' : '⚠️ Tùy chọn'}
        </span>
      </div>`;
  }).join('');

  document.getElementById('modal-overlay').classList.add('modal-overlay--open');
}

function closeModal() {
  document.getElementById('modal-overlay').classList.remove('modal-overlay--open');
}

// ─── Progress Overlay ────────────────────────────────────────────
function showProgress(title, step) {
  document.getElementById('progress-title').textContent = title || 'Đang chuyển đổi...';
  document.getElementById('progress-step').textContent = step || 'Đang khởi động...';
  document.getElementById('progress-bar').style.width = '5%';
  document.getElementById('progress-overlay').classList.add('progress-overlay--active');
  startProgressTimer();
}

function hideProgress() {
  if (window._progressInterval) clearInterval(window._progressInterval);
  document.getElementById('progress-overlay').classList.remove('progress-overlay--active');
}

function startProgressTimer() {
  const bar = document.getElementById('progress-bar');
  const step = document.getElementById('progress-step');
  let seconds = 0;
  if (window._progressInterval) clearInterval(window._progressInterval);

  window._progressInterval = setInterval(() => {
    seconds++;
    const pct = Math.min(95, Math.round(100 * (1 - Math.exp(-seconds / 4))));
    bar.style.width = pct + '%';

    if (seconds < 2) {
      step.textContent = `1️⃣ Phân tích cấu trúc tài liệu PDF / LaTeX & công thức OMML (⏱️ ${seconds}s)...`;
    } else if (seconds < 5) {
      step.textContent = `2️⃣ Dựng hình học & đồ thị sang Word DrawingML Shapes (⏱️ ${seconds}s)...`;
    } else {
      step.textContent = `3️⃣ Dựng Bảng biến thiên tkz-tab, Bảng trắc nghiệm & Đóng gói .docx (⏱️ ${seconds}s)...`;
    }
  }, 1000);
}

function finishProgress() {
  if (window._progressInterval) clearInterval(window._progressInterval);
  document.getElementById('progress-bar').style.width = '100%';
  setTimeout(hideProgress, 350);
}

// ─── Toast Notifications ─────────────────────────────────────────
function dismissToast() {
  const toast = document.getElementById('toast');
  toast.classList.remove('toast--visible');
  if (window._toastTimeout) clearTimeout(window._toastTimeout);
}

function showToast(message, type = 'success') {
  const toast = document.getElementById('toast');
  const icon = document.getElementById('toast-icon');
  const text = document.getElementById('toast-text');
  icon.textContent = type === 'error' ? '❌' : '✅';
  text.textContent = message;
  toast.className = `toast toast--visible toast--${type}`;
  if (window._toastTimeout) clearTimeout(window._toastTimeout);
  window._toastTimeout = setTimeout(() => {
    toast.classList.remove('toast--visible');
  }, type === 'error' ? 8000 : 4000);
}

function _cleanDocxFilename(rawName) {
  if (!rawName) return 'document.docx';
  const base = rawName.replace(/\.(pdf|tex|zip)$/i, '');
  return `${base}.docx`;
}

// ─── Conversion: Paste ───────────────────────────────────────────
async function convertPaste() {
  const source = document.getElementById('latex-editor').value.trim();
  if (!source) {
    showToast('Vui lòng nhập mã LaTeX trước khi chuyển đổi.', 'error');
    return;
  }

  showProgress('Đang chuyển đổi LaTeX → Word...', 'Đang khởi tạo...');

  const formData = new FormData();
  formData.append('latex_source', source);
  formData.append('mode', state.mode);
  formData.append('dpi', state.dpi);
  formData.append('strict', state.strict);
  formData.append('qa_pdf', state.qaPdf);

  try {
    const res = await fetch(`${API_BASE}/api/convert/paste`, { method: 'POST', body: formData });
    const data = await res.json();
    finishProgress();

    if (res.ok && data.status === 'completed') {
      state.currentJobId = data.job_id;
      showResults(data);
      showToast('🎉 Chuyển đổi thành công! Đang tự động tải file Word...');

      const dlLink = document.createElement('a');
      dlLink.href = `${API_BASE}/api/download/${data.job_id}/docx`;
      dlLink.download = _cleanDocxFilename(data.filename);
      document.body.appendChild(dlLink);
      dlLink.click();
      document.body.removeChild(dlLink);
    } else {
      showToast(data.error || data.detail || 'Chuyển đổi thất bại.', 'error');
    }
  } catch (err) {
    finishProgress();
    showToast('Lỗi kết nối tới API server: ' + err.message, 'error');
  }
}

// ─── Conversion: Upload ──────────────────────────────────────────
async function convertUpload() {
  if (!state.selectedFile) {
    showToast('Vui lòng chọn file (.pdf, .tex hoặc .zip) trước khi chuyển đổi.', 'error');
    return;
  }

  showProgress('Đang chuyển đổi File → Word...', 'Tải lên & phân tích file...');

  const formData = new FormData();
  formData.append('file', state.selectedFile);
  formData.append('main_tex', document.getElementById('main-tex-input').value);
  formData.append('mode', state.mode);
  formData.append('dpi', state.dpi);
  formData.append('strict', state.strict);
  formData.append('qa_pdf', state.qaPdf);

  try {
    const res = await fetch(`${API_BASE}/api/convert`, { method: 'POST', body: formData });
    const data = await res.json();
    finishProgress();

    if (res.ok && data.status === 'completed') {
      state.currentJobId = data.job_id;
      showResults(data);
      showToast('🎉 Chuyển đổi thành công! Đang tự động tải file Word...');

      const dlLink = document.createElement('a');
      dlLink.href = `${API_BASE}/api/download/${data.job_id}/docx`;
      dlLink.download = _cleanDocxFilename(data.filename);
      document.body.appendChild(dlLink);
      dlLink.click();
      document.body.removeChild(dlLink);
    } else {
      showToast(data.error || data.detail || 'Chuyển đổi thất bại.', 'error');
    }
  } catch (err) {
    finishProgress();
    showToast('Lỗi kết nối tới API server: ' + err.message, 'error');
  }
}

// ─── Display Results ─────────────────────────────────────────────
function showResults(data) {
  const section = document.getElementById('results-section');
  section.classList.add('results--visible');

  setTimeout(() => section.scrollIntoView({ behavior: 'smooth', block: 'start' }), 200);

  const header = document.getElementById('result-header');
  const icon = document.getElementById('result-icon');
  const title = document.getElementById('result-title');
  const desc = document.getElementById('result-desc');

  header.className = 'result-header';
  if (data.overall === 'passed') {
    header.classList.add('result-header--passed');
    icon.textContent = '✅';
    title.textContent = 'Chuyển đổi thành công — Tất cả kiểm thử đạt yêu cầu';
    desc.textContent = `Job ID: ${data.job_id} • Mode ${data.mode} • ${data.filename}`;
  } else {
    header.classList.add('result-header--review');
    icon.textContent = '⚠️';
    title.textContent = 'Chuyển đổi hoàn tất — Cần xem xét một số cảnh báo';
    desc.textContent = `Job ID: ${data.job_id} • Mode ${data.mode} • ${data.filename}`;
  }

  const ooxmlBanner = document.getElementById('ooxml-banner');
  if (data.ooxml_stats && data.ooxml_stats.tables_processed > 0) {
    ooxmlBanner.classList.add('ooxml-banner--visible');
    document.getElementById('ooxml-text').textContent =
      `Unified Math Engine: Đã dựng ${data.ooxml_stats.variation_tables || 0} bảng biến thiên, ` +
      `${data.ooxml_stats.choice_tables || 0} bảng trắc nghiệm, ` +
      `${data.ooxml_stats.data_tables || 0} bảng dữ liệu (${data.ooxml_stats.cells_formatted} ô).`;
  } else {
    ooxmlBanner.classList.remove('ooxml-banner--visible');
  }

  const jobId = data.job_id;
  document.getElementById('dl-docx').href = `${API_BASE}/api/download/${jobId}/docx`;
  document.getElementById('dl-report').href = `${API_BASE}/api/download/${jobId}/report`;
  const pdfBtn = document.getElementById('dl-pdf');
  if (data.engine_report && data.engine_report.qa_rendered_pdf) {
    pdfBtn.href = `${API_BASE}/api/download/${jobId}/pdf`;
    pdfBtn.style.display = 'inline-flex';
  } else {
    pdfBtn.style.display = 'none';
  }

  renderLayers(data.layers || []);
  renderEngineStats(data.engine_report || {});
}

// ─── Render 5 Validation Layers ──────────────────────────────────
function renderLayers(layers) {
  const grid = document.getElementById('layers-grid');
  grid.innerHTML = layers.map((layer, idx) => {
    const statusClass = (layer.status || 'skip').toLowerCase();
    const statusLabel = {
      'PASS': '✅ PASS', 'WARN': '⚠️ WARN',
      'FAIL': '❌ FAIL', 'SKIP': '⏭️ SKIP',
    }[layer.status] || layer.status;

    return `
      <div class="layer-card" onclick="toggleLayerDetails(${idx})">
        <span class="layer-badge layer-badge--${statusClass}">${statusLabel}</span>
        <div class="layer-info">
          <div class="layer-name">Layer ${layer.layer} — ${layer.name}</div>
          <div class="layer-summary">${layer.summary}</div>
          <div class="layer-details" id="layer-details-${idx}">
            <pre>${JSON.stringify(layer.details, null, 2)}</pre>
          </div>
        </div>
      </div>`;
  }).join('');
}

function toggleLayerDetails(idx) {
  const el = document.getElementById(`layer-details-${idx}`);
  el.classList.toggle('layer-details--open');
}

// ─── Render Engine Stats ─────────────────────────────────────────
function renderEngineStats(report) {
  const card = document.getElementById('stats-card');
  const grid = document.getElementById('report-grid');

  const fields = [
    { key: 'drawingml_groups', label: 'Nhóm Hình DrawingML' },
    { key: 'drawingml_shapes', label: 'Nét Vẽ Vector Word' },
    { key: 'tkz_tab_tables', label: 'Bảng Biến Thiên' },
    { key: 'choice_tables', label: 'Bảng Trắc Nghiệm' },
    { key: 'word_tables', label: 'Tổng Bảng Word' },
    { key: 'equation_nodes', label: 'Công Thức OMML' },
  ];

  const hasData = fields.some(f => report[f.key] !== undefined);
  if (!hasData) {
    card.style.display = 'none';
    return;
  }

  card.style.display = 'block';
  grid.innerHTML = fields
    .filter(f => report[f.key] !== undefined)
    .map(f => `
      <div class="report-stat">
        <div class="report-stat__value">${report[f.key]}</div>
        <div class="report-stat__label">${f.label}</div>
      </div>`)
    .join('');
}
