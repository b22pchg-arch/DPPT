/* SCADA Load Forecast Offline PWA LV9 - Workflow UI
   Mục tiêu: chỉ điều khiển hiển thị theo luồng, không khóa nút và không thay đổi lõi tính toán. */
(function(){
  'use strict';
  const VERSION = 'LV9';
  const STORAGE_KEY = 'scadaLoadForecast.lv9.workflowState';

  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function titleOf(card){ const h = card.querySelector('h2'); return h ? h.textContent.trim() : ''; }
  function getCards(){ return Array.from(document.querySelectorAll('main .grid > .card')).filter(c => c.id !== 'lv9WorkflowShell'); }
  function findCard(pattern){ return getCards().find(c => pattern.test(titleOf(c))); }
  function cardIndexMap(){
    const cards = getCards();
    const by = {};
    cards.forEach(c => {
      const t = titleOf(c);
      if (/0\) Chế độ/.test(t)) by.config = c;
      else if (/Dashboard LV8\.9/.test(t)) by.dashboard = c;
      else if (/Nhật ký thao tác LV8\.9/.test(t)) by.oplog = c;
      else if (/Hướng dẫn luồng/.test(t)) by.guide = c;
      else if (/1\) Nạp dữ liệu/.test(t)) by.load = c;
      else if (/2\) Ánh xạ/.test(t)) by.map = c;
      else if (/3\) Quy tắc chỉ danh/.test(t)) by.designation = c;
      else if (/4\) Hiệu chỉnh/.test(t)) by.editor = c;
      else if (/5\) Kiểm tra chất lượng/.test(t)) by.quality = c;
      else if (/6\) Nội suy/.test(t)) by.interpolate = c;
      else if (/7\) Phân tích sự kiện/.test(t)) by.operation = c;
      else if (/8\) Huấn luyện/.test(t)) by.train = c;
      else if (/9\) So sánh mô hình/.test(t)) by.compare = c;
      else if (/10\) Dùng model/.test(t)) by.scada = c;
      else if (/Kết quả/.test(t)) by.results = c;
      else if (/Bảng dữ liệu chuẩn hóa/.test(t)) by.table = c;
    });
    return by;
  }

  const FLOWS = {
    A: {
      name: 'Luồng A - Tạo model ngoài SCADA',
      badge: 'Mạng ngoài',
      hint: 'Chuẩn hóa dữ liệu lịch sử, xử lý chất lượng/chuyển tải, huấn luyện, so sánh chiến lược và xuất model vận hành.',
      steps: [
        {id:'A0', label:'Tổng quan', cards:['config','dashboard','guide'], primary:null, help:'Chọn chế độ Mạng ngoài, kiểm tra trạng thái dữ liệu/model và đọc luồng thao tác.'},
        {id:'A1', label:'Nạp dữ liệu', cards:['load'], primary:'csvFile', help:'Chọn file dữ liệu lịch sử Excel/CSV/TXT/JSON. Nếu là Excel, chọn đúng sheet sau khi nạp.'},
        {id:'A2', label:'Ánh xạ cột', cards:['map','load'], primary:'applyMapBtn', help:'Ánh xạ thời gian, P, chỉ danh trạm/lộ và các cờ vận hành. Sau đó bấm Áp dụng ánh xạ.'},
        {id:'A3', label:'Tách chỉ danh', cards:['designation','table'], primary:'parseDesignationBtn', help:'Tách Đơn vị/Trạm/Lộ/Lộ nối vòng theo quy tắc LV6 để tạo khóa dự báo chuẩn.'},
        {id:'A4', label:'Hiệu chỉnh dữ liệu', cards:['editor'], primary:'applyEditsBtn', help:'Lọc theo ngày, sửa bảng, điền nhanh nhiệt độ/ngày lễ/cờ vận hành rồi lưu thay đổi vào dữ liệu.'},
        {id:'A5', label:'Kiểm tra chất lượng', cards:['quality','editor'], primary:'runQualityBtn', help:'Kiểm tra toàn bộ RAM. Khi có lỗi, bảng hiệu chỉnh sẽ chỉ hiện dòng lỗi để rà soát.'},
        {id:'A6', label:'Nội suy / mốc thiếu', cards:['interpolate','editor'], primary:'interpolateBtn', help:'Bổ sung mốc thiếu, xử lý P trống/P thấp/P lỗi. Với mất mốc, xem dòng liền trước và liền sau mốc thiếu.'},
        {id:'A7', label:'Phân tích vận hành', cards:['operation','editor'], primary:'analyzeOperationBtn', help:'Phân tích P=0, chuyển tải, lộ giảm/lộ nhận tải. Nếu khôi phục P nền chuyển tải, app điều chỉnh cả lộ nhận tải.'},
        {id:'A8', label:'Huấn luyện model', cards:['train','results'], primary:'trainAllStationsBtn', help:'Huấn luyện GBDT cho từng chỉ danh/trạm/lộ. Đây là bước tạo model nội bộ trước khi xuất.'},
        {id:'A9', label:'So sánh chiến lược', cards:['compare','results'], primary:'trainSelectBestAllBtn', help:'So sánh GBDT, Similar Day, Hybrid, Auto Blend và áp dụng chiến lược tốt nhất vào model.'},
        {id:'A10', label:'Xuất model', cards:['train','dashboard'], primary:'exportModelBtn', help:'Sau khi huấn luyện và áp dụng chiến lược, xuất model vận hành để đưa vào SCADA.'}
      ]
    },
    B: {
      name: 'Luồng B - Dự báo trong mạng SCADA',
      badge: 'SCADA',
      hint: 'Nạp dữ liệu vận hành mới, nạp model vận hành đã xuất từ luồng A, dự báo đa cấp, cảnh báo và xuất báo cáo.',
      steps: [
        {id:'B0', label:'Tổng quan SCADA', cards:['config','dashboard'], primary:null, help:'Chọn chế độ SCADA. Có thể nạp cấu hình đã xuất từ luồng A để giữ ánh xạ/ngưỡng.'},
        {id:'B1', label:'Nạp dữ liệu mới', cards:['load','map'], primary:'applyMapBtn', help:'Dữ liệu mới được phép khác file huấn luyện, nhưng phải cùng cấu trúc cột, đơn vị MW và quy tắc chỉ danh.'},
        {id:'B2', label:'Tách chỉ danh', cards:['designation','table'], primary:'parseDesignationBtn', help:'Tách chỉ danh nếu file mới chưa có cột LV6. Nếu có chỉ danh mới chưa nằm trong model, nên huấn luyện bổ sung ở luồng A.'},
        {id:'B3', label:'Nạp model vận hành', cards:['scada'], primary:'modelFile', help:'Chỉ nạp model xuất từ Mục 8. Model dùng được cho dữ liệu mới cùng phạm vi chỉ danh đã học.'},
        {id:'B4', label:'Dự báo nhanh', cards:['scada','results'], primary:'forecastAllBtn', help:'Chọn trạm/lộ hoặc tất cả, số bước dự báo và chiến lược; bấm dự báo.'},
        {id:'B5', label:'Dự báo đa cấp', cards:['scada','results'], primary:'lv8QuickForecastBtn', help:'Chọn 15/30/60 phút, 24h, 48h hoặc 7 ngày; chọn điểm bắt đầu dự báo; chạy dự báo LV8.5.'},
        {id:'B6', label:'Báo cáo & cảnh báo', cards:['scada','results','table'], primary:'lv8SummaryBtn', help:'Tạo báo cáo đa cấp, xem Pmax, giờ Pmax, MWh, phụ tải theo ca và cảnh báo theo ngưỡng.'},
        {id:'B7', label:'Cập nhật RAM', cards:['scada','table'], primary:'lv82AppendForecastBtn', help:'Chèn forecast vào dữ liệu RAM để dự báo nối tiếp. Dòng forecast được đánh dấu không dùng huấn luyện.'}
      ]
    },
    C: {
      name: 'Luồng C - Đánh giá sai số và hiệu chỉnh model',
      badge: 'Đánh giá',
      hint: 'Sau vận hành, nạp dữ liệu thực tế và file dự báo để đánh giá sai số, tạo hiệu chỉnh và xuất model đã hiệu chỉnh.',
      steps: [
        {id:'C0', label:'Tổng quan đánh giá', cards:['dashboard','guide'], primary:null, help:'Dùng sau khi đã có dữ liệu thực tế tương ứng với forecast đã xuất.'},
        {id:'C1', label:'Nạp file đánh giá', cards:['scada'], primary:'lv84EvalFilesBtn', help:'Chọn file dữ liệu thực tế và file dữ liệu dự báo. App ghép theo thời gian + chỉ danh/trạm/lộ.'},
        {id:'C2', label:'Đánh giá sai số', cards:['scada','results'], primary:'lv84EvalRamBtn', help:'Tính MAE, MAPE, RMSE, BIAS từng mốc và tổng hợp theo chỉ danh/ngày.'},
        {id:'C3', label:'Tạo hiệu chỉnh', cards:['scada'], primary:'lv85BuildCalibrationBtn', help:'Tạo bảng bias MW/% theo toàn hệ thống, chỉ danh, giờ, loại ngày hoặc chỉ danh+giờ.'},
        {id:'C4', label:'Áp dụng hiệu chỉnh', cards:['scada'], primary:'lv85ApplyCalibrationBtn', help:'Áp dụng bảng hiệu chỉnh vào model đang nạp. Dự báo sau sẽ có cột forecast_before_calibration_mw.'},
        {id:'C5', label:'So sánh trước/sau', cards:['scada','results'], primary:'lv89CompareCalibrationBtn', help:'So sánh sai số trước/sau hiệu chỉnh để quyết định có nên dùng model đã hiệu chỉnh không.'},
        {id:'C6', label:'Xuất hồ sơ', cards:['scada','dashboard','oplog'], primary:'lv85ExportModelBtn', help:'Xuất calibration, model đã hiệu chỉnh, báo cáo sai số, nhật ký thao tác và trạng thái dashboard.'}
      ]
    }
  };

  let state = { flow:'A', step:0, expert:false, side:false };

  function saveState(){ try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch(_){} }
  function loadState(){
    try { const raw = localStorage.getItem(STORAGE_KEY); if (raw) state = Object.assign(state, JSON.parse(raw)); } catch(_){}
    if (!FLOWS[state.flow]) state.flow = 'A';
    state.step = Math.max(0, Math.min(state.step || 0, FLOWS[state.flow].steps.length - 1));
  }

  function createShell(){
    const grid = document.querySelector('main .grid');
    if (!grid || $('lv9WorkflowShell')) return;
    const shell = document.createElement('div');
    shell.className = 'card span12 lv9-shell';
    shell.id = 'lv9WorkflowShell';
    shell.innerHTML = `
      <div class="lv9-topline">
        <div>
          <h2>LV9) Giao diện Workflow / Workbench</h2>
          <div class="compactNote">Chọn luồng A/B/C để chỉ hiển thị bước đang thao tác. <b>Không khóa nút</b>, không thay đổi lõi tính toán; có thể quay lại Chế độ chuyên gia để xem toàn bộ mục dọc.</div>
        </div>
        <div class="lv9-mode-actions">
          <button class="secondary" id="lv9ExpertBtn">Chế độ chuyên gia</button>
          <button class="secondary" id="lv9SideBtn">Dashboard / Nhật ký</button>
        </div>
      </div>
      <div class="lv9-flow-row" id="lv9FlowButtons"></div>
      <div class="lv9-workbench">
        <aside class="lv9-steps" id="lv9StepNav"></aside>
        <div class="lv9-current">
          <div class="lv9-status" id="lv9StatusBox"></div>
          <div class="lv9-help" id="lv9HelpBox"></div>
          <div class="lv9-navrow">
            <button class="secondary" id="lv9PrevBtn">← Quay lại</button>
            <button class="good" id="lv9PrimaryBtn">Thực hiện chính</button>
            <button class="secondary" id="lv9NextBtn">Tiếp tục →</button>
          </div>
        </div>
      </div>
      <div class="lv9-sidepanel" id="lv9SidePanel" style="display:none">
        <div class="statusBox"><span class="pill modeBadge">Ngăn phụ LV9</span><span class="pill">Chỉ để xem nhanh, không can thiệp chức năng</span></div>
        <div class="row" style="margin-top:8px">
          <button class="secondary" data-lv9-jump="dashboard">Mở Dashboard</button>
          <button class="secondary" data-lv9-jump="oplog">Mở Nhật ký</button>
          <button class="secondary" data-lv9-jump="results">Mở Kết quả</button>
          <button class="secondary" data-lv9-jump="table">Mở Bảng dữ liệu</button>
        </div>
      </div>`;
    grid.insertBefore(shell, grid.firstElementChild);
  }

  function renderFlowButtons(){
    const box = $('lv9FlowButtons'); if (!box) return;
    box.innerHTML = Object.entries(FLOWS).map(([key, flow]) =>
      `<button class="${state.flow===key && !state.expert ? 'good' : 'secondary'}" data-flow="${key}"><b>${key}</b> ${esc(flow.name.replace(/^Luồng [ABC] - /,''))}</button>`
    ).join('') + `<button class="${state.expert ? 'good' : 'secondary'}" data-expert="1">Hiển thị tất cả mục</button>`;
    box.querySelectorAll('[data-flow]').forEach(btn => btn.addEventListener('click', () => {
      state.flow = btn.dataset.flow; state.step = 0; state.expert = false; saveState(); applyWorkflow();
    }));
    box.querySelector('[data-expert]')?.addEventListener('click', () => { state.expert = !state.expert; saveState(); applyWorkflow(); });
  }

  function renderSteps(){
    const nav = $('lv9StepNav'); if (!nav) return;
    const flow = FLOWS[state.flow];
    nav.innerHTML = flow.steps.map((s, i) => `<button class="${i===state.step ? 'active' : ''}" data-step="${i}"><span>${esc(s.id)}</span>${esc(s.label)}</button>`).join('');
    nav.querySelectorAll('[data-step]').forEach(btn => btn.addEventListener('click', () => {
      state.step = Number(btn.dataset.step); state.expert = false; saveState(); applyWorkflow();
    }));
  }

  function clickPrimary(){
    const flow = FLOWS[state.flow]; const step = flow.steps[state.step];
    const id = step.primary;
    if (!id) { setStatus('Bước này không có nút thao tác chính. Hãy thao tác trong khung đang hiển thị.', 'warn'); return; }
    const el = $(id);
    if (!el) { setStatus(`Không tìm thấy nút/ô thao tác chính: ${id}`, 'bad'); return; }
    if (el.tagName === 'INPUT' && el.type === 'file') { el.click(); setStatus(`Đã mở hộp chọn file cho ${id}.`, 'ok'); return; }
    if (el.disabled) { setStatus(`Nút chính đang bị khóa theo trạng thái nội bộ của chức năng gốc: ${id}. Kiểm tra dữ liệu/model yêu cầu trước.`, 'warn'); return; }
    el.click();
    setStatus(`Đã kích hoạt thao tác chính: ${id}`, 'ok');
  }

  function setStatus(msg, level){
    const box = $('lv9StatusBox'); if (!box) return;
    const cls = level === 'bad' ? 'bad' : level === 'ok' ? 'ok' : level === 'warn' ? 'warn' : '';
    box.innerHTML = `<span class="pill ${cls}">${esc(msg)}</span><span class="pill">${new Date().toLocaleString('vi-VN')}</span>`;
  }

  function showCardsForCurrentStep(){
    const cards = getCards();
    cards.forEach(c => { c.classList.remove('lv9-hidden','lv9-current-card'); });
    document.documentElement.classList.toggle('lv9-workflow-mode', !state.expert);
    document.documentElement.classList.toggle('lv9-expert-mode', !!state.expert);
    if (state.expert) return;
    const map = cardIndexMap();
    const step = FLOWS[state.flow].steps[state.step];
    const targets = new Set(['lv9WorkflowShell']);
    step.cards.forEach(k => { if (map[k]) targets.add(map[k].id || assignId(map[k], k)); });
    cards.forEach(c => {
      const id = c.id || assignId(c, 'card');
      if (!targets.has(id)) c.classList.add('lv9-hidden');
      else c.classList.add('lv9-current-card');
    });
  }

  function assignId(el, prefix){
    if (!el.id) el.id = `lv9_${prefix}_${Math.random().toString(36).slice(2,8)}`;
    return el.id;
  }

  function updateInfo(){
    const flow = FLOWS[state.flow]; const step = flow.steps[state.step];
    const status = $('lv9StatusBox');
    if (status) status.innerHTML = `<span class="pill modeBadge">${esc(flow.badge)}</span><span class="pill ok">${esc(step.id)} - ${esc(step.label)}</span><span class="pill">Bước ${state.step+1}/${flow.steps.length}</span>`;
    const help = $('lv9HelpBox');
    if (help) help.innerHTML = `<b>${esc(flow.name)}</b><br>${esc(flow.hint)}<hr style="border:0;border-top:1px solid var(--line);margin:8px 0"><b>Việc cần làm:</b> ${esc(step.help || '')}`;
    const prev = $('lv9PrevBtn'), next = $('lv9NextBtn'), pri = $('lv9PrimaryBtn');
    if (prev) prev.disabled = state.step <= 0;
    if (next) next.disabled = state.step >= flow.steps.length - 1;
    if (pri) pri.textContent = step.primary ? `Thực hiện chính (${step.primary})` : 'Bước hướng dẫn / kiểm tra';
  }

  function applyWorkflow(){
    renderFlowButtons();
    renderSteps();
    showCardsForCurrentStep();
    updateInfo();
    const side = $('lv9SidePanel'); if (side) side.style.display = state.side ? 'block' : 'none';
    setTimeout(() => {
      const shell = $('lv9WorkflowShell'); if (shell && !state.expert) shell.scrollIntoView({block:'start', behavior:'smooth'});
    }, 20);
  }

  function wire(){
    $('lv9PrevBtn')?.addEventListener('click', () => { if (state.step > 0) { state.step--; saveState(); applyWorkflow(); } });
    $('lv9NextBtn')?.addEventListener('click', () => { const max = FLOWS[state.flow].steps.length - 1; if (state.step < max) { state.step++; saveState(); applyWorkflow(); } });
    $('lv9PrimaryBtn')?.addEventListener('click', clickPrimary);
    $('lv9ExpertBtn')?.addEventListener('click', () => { state.expert = !state.expert; saveState(); applyWorkflow(); });
    $('lv9SideBtn')?.addEventListener('click', () => { state.side = !state.side; saveState(); applyWorkflow(); });
    document.querySelectorAll('[data-lv9-jump]').forEach(btn => btn.addEventListener('click', () => {
      const key = btn.getAttribute('data-lv9-jump');
      state.expert = false;
      const flow = FLOWS[state.flow];
      const idx = flow.steps.findIndex(s => s.cards.includes(key));
      if (idx >= 0) state.step = idx;
      else {
        const fallback = {dashboard:['A',0], oplog:['C',6], results:['B',6], table:['B',7]}[key];
        if (fallback) { state.flow = fallback[0]; state.step = fallback[1]; }
      }
      saveState(); applyWorkflow();
    }));
  }

  function markVersion(){
    const h1 = document.querySelector('h1'); if (h1) h1.textContent = 'SCADA Load Forecast Offline PWA LV9';
    document.title = 'SCADA Load Forecast Offline PWA LV9';
    const vi = $('versionInfo'); if (vi) vi.innerHTML = '<span class="pill modeBadge">LV9</span><span class="pill ok">Workflow UI + Workbench, không khóa nút</span>';
    const cfg = $('configName'); if (cfg && /LV8/i.test(cfg.value)) cfg.value = 'SCADA_LOAD_FORECAST_LV9';
  }

  function init(){
    createShell();
    loadState();
    wire();
    markVersion();
    applyWorkflow();
    try { if (window.log) window.log('Sẵn sàng LV9: giao diện Workflow/Workbench chỉ điều khiển hiển thị, không khóa nút và không thay đổi lõi LV8.9.'); } catch(_){}
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
