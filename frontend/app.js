/**
 * PDF & LaTeX to Word Converter — Frontend (pipeline chung: TeX/ZIP/PDF → Word, báo cáo trung thực)
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
  answerMode: 'source',
  showSolutions: true,
  figurePolicy: 'auto',
  selectedFile: null,
  currentJobId: null,
  theme: (() => { try { return localStorage.getItem('latex2word-theme') || 'dark'; } catch { return 'dark'; } })(),
};

// Chống chèn HTML: mọi chuỗi từ máy chủ (tên tệp, tóm tắt, mã LaTeX trong báo cáo) phải escape trước khi vào innerHTML.
function escapeHtml(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

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
  try { localStorage.setItem('latex2word-theme', state.theme); } catch { /* bỏ qua */ }
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
  } else if (option === 'solutions') {
    state.showSolutions = !state.showSolutions;
    document.getElementById('toggle-solutions').classList.toggle('toggle--active', state.showSolutions);
  }
}

function updateOption(key, value) {
  state[key] = value;
  if (key === 'answerMode' && value === 'student' && state.showSolutions) {
    toggleOption('solutions');
  }
}

function appendCommonOptions(formData) {
  formData.append('mode', state.mode);
  formData.append('dpi', state.dpi);
  formData.append('strict', state.strict);
  formData.append('qa_pdf', state.qaPdf);
  formData.append('answer_mode', state.answerMode);
  formData.append('show_solutions', state.showSolutions);
  formData.append('figure_policy', state.figurePolicy);
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
      text.textContent = 'Thiếu thành phần — xem chi tiết';
    }

    window._systemTools = data.tools;
    window._capabilities = data.capabilities || null;
  } catch {
    const dot = document.getElementById('status-dot');
    const text = document.getElementById('status-text');
    dot.className = 'status-dot status-dot--error';
    text.textContent = 'API không phản hồi';
  }
}

function showSystemStatus() {
  const container = document.getElementById('modal-tools');
  const caps = window._capabilities;
  let rows = [];
  if (caps && caps.derived) {
    const labels = {
      tex_upload: 'Engine dựng Word từ TeX (.tex / .zip)',
      paste_smart: 'Tiền xử lý thông minh (latex_preprocessor)',
      pdf_reader: 'Bộ đọc PDF (pdfplumber + pypdfium2 + pikepdf)',
      pdf_vector_crop: 'Cắt vùng PDF dạng vector (pdftocairo)',
      ocr: 'OCR trang scan (tesseract)',
      math_ocr: 'Nhận dạng công thức từ ảnh',
      visual_qa: 'Render QA Word → PDF (LibreOffice / Word)',
      tex_reference_compile: 'Biên dịch PDF tham chiếu từ TeX',
    };
    rows = Object.entries(labels).map(([key, label]) => {
      const item = caps.derived[key] || { available: false, detail: 'không có thông tin' };
      return `
        <div class="tool-check">
          <span class="tool-check__name">${escapeHtml(label)}<br><small>${escapeHtml(item.detail)}</small></span>
          <span class="tool-check__status ${item.available ? 'tool-check__status--ok' : 'tool-check__status--missing'}">
            ${item.available ? '✅ Có' : '❌ Không có'}
          </span>
        </div>`;
    });
  } else {
    const tools = window._systemTools || {};
    rows = Object.entries(tools).map(([key, ok]) => `
        <div class="tool-check">
          <span class="tool-check__name">${escapeHtml(key)}</span>
          <span class="tool-check__status ${ok ? 'tool-check__status--ok' : 'tool-check__status--missing'}">${ok ? '✅ Có' : '❌ Không có'}</span>
        </div>`);
  }
  container.innerHTML = rows.join('');
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
      step.textContent = `1️⃣ Đọc nguồn (PDF / TeX / ZIP) và xác định khối nội dung (⏱️ ${seconds}s)...`;
    } else if (seconds < 5) {
      step.textContent = `2️⃣ Dựng Word: chữ, OMML, bảng, hình / ảnh vùng gốc (⏱️ ${seconds}s)...`;
    } else {
      step.textContent = `3️⃣ Đóng gói, render đối chiếu và kiểm định 5 nhóm (⏱️ ${seconds}s)...`;
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
  appendCommonOptions(formData);

  try {
    const res = await fetch(`${API_BASE}/api/convert/paste`, { method: 'POST', body: formData });
    const data = await res.json();
    finishProgress();

    if (data && data.job_id && (data.status === 'completed' || data.strict_violation)) {
      state.currentJobId = data.job_id;
      showResults(data);
      if (data.strict_violation) {
        showToast(data.error || 'Chế độ nghiêm ngặt: kết quả chưa đạt.', 'error');
        return;
      }
      showToast(data.quality_status === 'passed'
        ? '🎉 Đã tạo Word — đạt các kiểm tra tự động. Đang tải xuống...'
        : '⚠️ Đã tạo Word — cần xem lại một số mục. Đang tải xuống...');

      const dlLink = document.createElement('a');
      dlLink.href = `${API_BASE}/api/download/${data.job_id}/docx`;
      dlLink.download = _cleanDocxFilename(data.filename);
      document.body.appendChild(dlLink);
      dlLink.click();
      document.body.removeChild(dlLink);
    } else {
      showError(data);
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
  appendCommonOptions(formData);

  try {
    const res = await fetch(`${API_BASE}/api/convert`, { method: 'POST', body: formData });
    const data = await res.json();
    finishProgress();

    if (data && data.job_id && (data.status === 'completed' || data.strict_violation)) {
      state.currentJobId = data.job_id;
      showResults(data);
      if (data.strict_violation) {
        showToast(data.error || 'Chế độ nghiêm ngặt: kết quả chưa đạt.', 'error');
        return;
      }
      showToast(data.quality_status === 'passed'
        ? '🎉 Đã tạo Word — đạt các kiểm tra tự động. Đang tải xuống...'
        : '⚠️ Đã tạo Word — cần xem lại một số mục. Đang tải xuống...');

      const dlLink = document.createElement('a');
      dlLink.href = `${API_BASE}/api/download/${data.job_id}/docx`;
      dlLink.download = _cleanDocxFilename(data.filename);
      document.body.appendChild(dlLink);
      dlLink.click();
      document.body.removeChild(dlLink);
    } else {
      showError(data);
    }
  } catch (err) {
    finishProgress();
    showToast('Lỗi kết nối tới API server: ' + err.message, 'error');
  }
}

function showError(data) {
  let msg = (data && (data.error || data.detail)) || 'Chuyển đổi thất bại.';
  if (data && Array.isArray(data.candidates) && data.candidates.length) {
    const names = data.candidates.map(c => (typeof c === 'string' ? c : c.path)).filter(Boolean);
    msg += ' Ứng viên: ' + names.join(', ');
    const input = document.getElementById('main-tex-input');
    if (input && names.length) input.placeholder = names[0];
  }
  showToast(msg, 'error');
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
  const quality = data.quality_status || data.overall;
  desc.textContent = `Job ID: ${data.job_id} • Mode ${data.mode} • ${data.filename} • tạo tệp: ${data.conversion_completed ? 'có' : 'không'}`;
  if (!data.conversion_completed && data.status !== 'completed') {
    header.classList.add('result-header--failed');
    icon.textContent = '❌';
    title.textContent = 'Không tạo được gói Word hợp lệ';
  } else if (quality === 'passed') {
    header.classList.add('result-header--passed');
    icon.textContent = '✅';
    title.textContent = 'Đã tạo Word — đạt các tiêu chí tự động đã nêu (không có nghĩa đúng 100%)';
  } else if (quality === 'failed') {
    header.classList.add('result-header--failed');
    icon.textContent = '❌';
    title.textContent = 'Đã tạo tệp nhưng kiểm định thất bại — xem lý do';
  } else {
    header.classList.add('result-header--review');
    icon.textContent = '⚠️';
    title.textContent = 'Đã tạo Word — cần người dùng kiểm tra (review_required)';
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

  renderQuality(data);
  renderLayers(data.layers || []);
  renderEngineStats(data.engine_report || {});
}

function renderQuality(data) {
  const grid = document.getElementById('quality-grid');
  const notes = document.getElementById('quality-notes');
  if (!grid || !notes) return;
  const bs = data.blocks_summary || {};
  const q = data.quality || {};
  const cells = [
    ['Khối dựng native', bs.native || 0],
    ['Khối giữ bằng ảnh', bs.as_image || 0],
    ['Khối unresolved', bs.unresolved || 0],
    ['Tổng khối', bs.total || 0],
  ];
  grid.innerHTML = cells.map(([label, value]) => `
      <div class="report-stat">
        <div class="report-stat__value">${escapeHtml(value)}</div>
        <div class="report-stat__label">${escapeHtml(label)}</div>
      </div>`).join('');
  let html = '<div><small>Mẫu số là tổng khối trong manifest của job; đây là thống kê phương thức dựng, '
    + 'không phải độ chính xác thị giác đã đo.</small></div>';
  if ((q.not_performed || []).length) {
    html += '<div><strong>Kiểm tra bắt buộc chưa thực hiện:</strong><ul>'
      + q.not_performed.map(x => `<li>${escapeHtml(x)}</li>`).join('') + '</ul></div>';
  }
  if ((q.reasons || []).length) {
    html += '<div><strong>Lý do trạng thái:</strong><ul>'
      + q.reasons.slice(0, 15).map(x => `<li>${escapeHtml(x)}</li>`).join('') + '</ul></div>';
  }
  if (data.source && data.source.main) {
    html += `<div><strong>Tệp chính:</strong> ${escapeHtml(data.source.main)} (${escapeHtml(data.source.selection || '')})</div>`;
  }
  notes.innerHTML = html;
}

// ─── Render 5 Validation Layers ──────────────────────────────────
function renderLayers(layers) {
  const grid = document.getElementById('layers-grid');
  grid.innerHTML = layers.map((layer, idx) => {
    const statusClass = String(layer.status || 'skip').toLowerCase().replace(/[^a-z_]/g, '');
    const statusLabel = {
      'PASS': '✅ PASS', 'WARN': '⚠️ WARN',
      'FAIL': '❌ FAIL', 'SKIP': '⏭️ SKIP', 'NOT_MEASURED': '⏸️ CHƯA ĐO',
    }[layer.status] || escapeHtml(layer.status);

    return `
      <div class="layer-card" onclick="toggleLayerDetails(${idx})">
        <span class="layer-badge layer-badge--${statusClass}">${statusLabel}</span>
        <div class="layer-info">
          <div class="layer-name">Layer ${escapeHtml(layer.layer)} — ${escapeHtml(layer.name)}</div>
          <div class="layer-summary">${escapeHtml(layer.summary)}</div>
          <div class="layer-details" id="layer-details-${idx}">
            <pre>${escapeHtml(JSON.stringify(layer.details, null, 2))}</pre>
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
    { key: 'drawingml_groups', label: 'Nhóm Hình DrawingML (tệp cuối)' },
    { key: 'drawingml_shapes', label: 'Nét Vẽ Vector Word' },
    { key: 'tkz_tab_tables', label: 'Bảng Biến Thiên' },
    { key: 'choice_tables', label: 'Bảng Trắc Nghiệm' },
    { key: 'word_tables', label: 'Tổng Bảng Word' },
    { key: 'equation_nodes', label: 'Công Thức OMML (tệp cuối)' },
    { key: 'embedded_media_files', label: 'Ảnh logic (SVG+PNG = 1)' },
    { key: 'smart_paste_status', label: 'Tiền Xử Lý Thông Minh (dán mã)' },
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
        <div class="report-stat__value">${escapeHtml(report[f.key])}</div>
        <div class="report-stat__label">${escapeHtml(f.label)}</div>
      </div>`)
    .join('');
}
