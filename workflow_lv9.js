/* SCADA Load Forecast Offline PWA LV9.8.1 - B2 Forecast Buttons Fix / Workbench
   Chỉ điều khiển HIỂN THỊ, không khóa nút, không thay đổi lõi tính toán LV8.9. */
(function(){
  'use strict';
  const VERSION = 'LV9.8.1';
  const STORAGE_KEY = 'scadaLoadForecast.lv9_7.workflowState';

  function $(id){ return document.getElementById(id); }
  function esc(s){ return String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function titleOf(card){ const h = card.querySelector('h2'); return h ? h.textContent.trim() : ''; }
  const CARD_REGISTRY = [];
  const CARD_HOME = new Map();

  function ensureCardRegistry(){
    if (CARD_REGISTRY.length) return;
    const grid = document.querySelector('main .grid');
    if (!grid) return;
    Array.from(grid.children).forEach((c, idx) => {
      if (!(c instanceof HTMLElement)) return;
      if (!c.classList.contains('card')) return;
      if (c.id === 'lv9WorkflowShell') return;
      const marker = document.createComment(`lv9-card-placeholder-${idx}`);
      c.parentNode.insertBefore(marker, c);
      CARD_REGISTRY.push(c);
      CARD_HOME.set(c, { marker });
    });
  }

  function getCards(){ ensureCardRegistry(); return CARD_REGISTRY.slice(); }
  function assignId(el, prefix){ if (!el.id) el.id = `lv92_${prefix}_${Math.random().toString(36).slice(2,8)}`; return el.id; }

  function undockAllCards(){
    ensureCardRegistry();
    CARD_REGISTRY.forEach(card => {
      card.classList.remove('lv9-hidden','lv9-current-card','lv9-docked-card');
      const home = CARD_HOME.get(card);
      if (home && home.marker && home.marker.parentNode && card.previousSibling !== home.marker) {
        home.marker.parentNode.insertBefore(card, home.marker.nextSibling);
      }
    });
    const slot = $('lv9ActiveCardSlot');
    if (slot) slot.innerHTML = '';
  }

  function dockCardsIntoWorkbench(cards){
    const slot = $('lv9ActiveCardSlot');
    if (!slot) return;
    slot.innerHTML = '';
    if (!cards.length) {
      slot.innerHTML = '<div class="compactNote">Bước này chưa có khung chức năng tương ứng. Hãy dùng nút điều hướng hoặc bật Chế độ chuyên gia nếu cần.</div>';
      return;
    }
    cards.forEach(card => {
      card.classList.remove('lv9-hidden');
      card.classList.add('lv9-current-card','lv9-docked-card');
      slot.appendChild(card);
    });
  }

  function ensureStyle(){
    if ($('lv92WorkflowStyle')) return;
    const st = document.createElement('style');
    st.id = 'lv92WorkflowStyle';
    st.textContent = `
      html.lv9-workflow-mode main{max-width:none!important;width:100%!important;padding-left:8px!important;padding-right:8px!important;}
      html.lv9-workflow-mode main .grid{display:block!important;}
      html.lv9-workflow-mode #lv9WorkflowShell{width:100%!important;max-width:none!important;margin-bottom:12px!important;}
      html.lv9-workflow-mode .lv9-hidden{display:none!important;}
      html.lv9-workflow-mode .lv9-current-card{display:block!important;min-width:0!important;width:100%!important;max-width:none!important;margin:0 0 12px 0!important;}
      html.lv9-workflow-mode .lv9-docked-card{box-shadow:none!important;}
      html.lv9-workflow-mode .lv9-current-card .tableWrap,
      html.lv9-workflow-mode .lv9-current-card .tablebox,
      html.lv9-workflow-mode .lv9-current-card .scrollBox{max-width:100%!important;width:100%!important;}
      html.lv9-workflow-mode .lv9-current-card table{min-width:100%;}
      .lv9-shell{border:1px solid rgba(56,189,248,.45)!important;background:linear-gradient(180deg,rgba(15,23,42,.98),rgba(15,23,42,.94))!important;}
      .lv9-topline{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;}
      .lv9-mode-actions{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end;}
      .lv9-flow-row{display:flex;flex-wrap:wrap;gap:8px;margin:10px 0 8px;}
      .lv9-flow-row button{border-radius:12px!important;min-height:44px;}
      .lv9-workbench{display:grid;grid-template-columns:230px minmax(0,1fr);gap:12px;align-items:start;margin-top:8px;}
      .lv9-steps{border:1px solid var(--line);border-radius:14px;background:rgba(2,6,23,.38);padding:8px;max-height:calc(100vh - 260px);overflow:auto;position:sticky;top:8px;}
      .lv9-steps button{display:flex!important;width:100%;gap:8px;align-items:center;justify-content:flex-start;margin:0 0 6px 0;background:rgba(51,65,85,.92);border-radius:10px;line-height:1.25;text-align:left;}
      .lv9-steps button span{font-weight:800;color:#93c5fd;min-width:34px;}
      .lv9-steps button.active{background:#16a34a!important;color:#fff!important;box-shadow:inset 4px 0 0 #bbf7d0;}
      .lv9-steps button.active span{color:#fff;}
      .lv9-current{border:1px solid var(--line);border-radius:14px;background:rgba(2,6,23,.38);padding:10px;min-width:0;min-height:calc(100vh - 245px);}
      .lv9-active-slot{margin-top:10px;padding-top:10px;border-top:1px solid var(--line);min-height:360px;}
      .lv9-active-slot > .card{border-color:rgba(56,189,248,.35)!important;background:rgba(15,23,42,.96)!important;}
      .lv9-active-slot > .card .tableWrap{max-height:62vh;}
      .lv9-status,.lv9-help{margin-bottom:8px;}
      .lv9-help{font-size:14px;color:var(--muted);line-height:1.45;}
      .lv9-navrow{display:grid;grid-template-columns:160px 1fr 160px;gap:8px;align-items:center;}
      .lv9-navrow button{min-height:42px;}
      .lv9-sidepanel{margin-top:10px;border:1px dashed rgba(148,163,184,.4);border-radius:14px;padding:10px;background:rgba(2,6,23,.28);}
      .lv9-homecards{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;margin-top:10px;}
      .lv9-homecards button{white-space:normal;min-height:74px;text-align:left;line-height:1.35;}
      .lv93-compactbar{display:none;align-items:center;gap:6px;grid-template-columns:auto minmax(150px,210px) minmax(220px,1fr) auto auto auto auto auto minmax(120px,210px);position:sticky;top:0;z-index:20;background:rgba(15,23,42,.98);padding:6px;border:1px solid rgba(56,189,248,.28);border-radius:12px;margin:4px 0 6px;box-shadow:0 8px 18px rgba(0,0,0,.18);}
      .lv93-compactbar select,.lv93-compactbar button{min-height:32px!important;padding:5px 8px!important;border-radius:8px!important;font-size:12px!important;}
      .lv93-compactbar .lv93-title{font-weight:900;color:#dbeafe;white-space:nowrap;}
      .lv93-compactbar .lv93-status{font-size:12px;color:#9fb0c6;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;min-width:120px;}
      html.lv93-compact-mode header{padding:6px 14px!important;}
      html.lv93-compact-mode header h1{font-size:16px!important;margin:0!important;}
      html.lv93-compact-mode header .sub{display:none!important;}
      html.lv93-compact-mode #lv9WorkflowShell{padding:8px!important;margin-bottom:6px!important;border-radius:12px!important;}
      html.lv93-compact-mode #lv9WorkflowShell .lv9-topline,
      html.lv93-compact-mode #lv9WorkflowShell .lv9-flow-row,
      html.lv93-compact-mode #lv9WorkflowShell .lv9-homecards,
      html.lv93-compact-mode #lv9WorkflowShell #lv9StatusBox,
      html.lv93-compact-mode #lv9WorkflowShell #lv9HelpBox,
      html.lv93-compact-mode #lv9WorkflowShell #lv92TrueNote,
      html.lv93-compact-mode #lv9WorkflowShell .lv9-navrow{display:none!important;}
      html.lv93-compact-mode .lv93-compactbar{display:grid!important;}
      html.lv93-compact-mode .lv9-workbench{grid-template-columns:1fr!important;gap:0!important;margin-top:4px!important;}
      html.lv93-compact-mode .lv9-steps{display:none!important;}
      html.lv93-compact-mode.lv93-show-steps .lv9-workbench{grid-template-columns:96px minmax(0,1fr)!important;gap:8px!important;}
      html.lv93-compact-mode.lv93-show-steps .lv9-steps{display:block!important;max-height:calc(100vh - 138px)!important;padding:5px!important;}
      html.lv93-compact-mode.lv93-show-steps .lv9-steps button{justify-content:center!important;text-align:center!important;padding:7px 4px!important;min-height:32px!important;font-size:11px!important;}
      html.lv93-compact-mode.lv93-show-steps .lv9-steps button .lv9-step-label{display:none!important;}
      html.lv93-compact-mode.lv93-show-steps .lv9-steps button span{min-width:0!important;}
      html.lv93-compact-mode .lv9-current{padding:4px!important;min-height:calc(100vh - 142px)!important;border-radius:10px!important;}
      html.lv93-compact-mode .lv9-active-slot{margin-top:0!important;padding-top:0!important;border-top:0!important;min-height:calc(100vh - 150px)!important;}
      html.lv93-compact-mode .lv9-active-slot > .card{padding:10px!important;margin-bottom:8px!important;border-radius:12px!important;}
      html.lv93-compact-mode .lv9-active-slot > .card h2{margin-bottom:8px!important;font-size:16px!important;}
      html.lv93-compact-mode .lv9-active-slot > .card .tablewrap{max-height:calc(100vh - 235px)!important;}
      html.lv93-compact-mode main{padding:6px!important;gap:6px!important;}
      html.lv93-compact-mode main .grid{gap:6px!important;}
      html.lv93-compact-mode .card{padding:10px;}
      html.lv93-compact-mode .toolbar{gap:6px!important;}
      html.lv93-compact-mode .row{gap:6px!important;}
      html.lv93-compact-mode label{margin:5px 0 2px!important;}
      html.lv93-compact-mode input,html.lv93-compact-mode select,html.lv93-compact-mode textarea{padding:7px 8px!important;}
      html.lv93-compact-mode button{min-height:34px;padding:7px 9px;}
      .lv92-true-note{border:1px solid rgba(34,197,94,.35);background:rgba(22,163,74,.08);border-radius:12px;padding:8px 10px;margin-top:8px;color:#d1fae5;}
      html.lv9-expert-mode .lv9-hidden{display:block!important;}
      html.lv9-expert-mode #lv9WorkflowShell{position:static!important;}
      @media (max-width:900px){
        html.lv9-workflow-mode main .grid{grid-template-columns:1fr!important;}
        .lv9-workbench{grid-template-columns:1fr;}
        .lv9-steps{display:flex;gap:6px;overflow-x:auto;max-height:none;}
        .lv9-steps button{width:auto;min-width:150px;margin:0;}
        .lv9-navrow{grid-template-columns:1fr;}
        .lv9-homecards{grid-template-columns:1fr;}
        .lv93-compactbar{grid-template-columns:1fr 1fr;position:static;}
        .lv93-compactbar .lv93-title,.lv93-compactbar .lv93-status{grid-column:1/-1;}
        html.lv93-compact-mode.lv93-show-steps .lv9-workbench{grid-template-columns:1fr!important;}
        html.lv93-compact-mode.lv93-show-steps .lv9-steps{display:flex!important;overflow-x:auto;max-height:none!important;}
      }
    `;
    document.head.appendChild(st);
  }

  function cardIndexMap(){
    const cards = getCards();
    const by = {};
    cards.forEach(c => {
      const t = titleOf(c);
      if (/0\) Chế độ/.test(t)) by.config = c;
      else if (/0\.3\).*Dashboard LV8\.9/.test(t) || /Dashboard LV8\.9/.test(t)) by.dashboard = c;
      else if (/0\.4\).*Nhật ký thao tác LV8\.9/.test(t) || /Nhật ký thao tác LV8\.9/.test(t)) by.oplog = c;
      else if (/0\.1\).*Hướng dẫn luồng/.test(t) || /Hướng dẫn luồng/.test(t)) by.guide = c;
      else if (/0\.2\).*Điều khiển thu gọn/.test(t)) by.collapse = c;
      else if (/1\) Nạp dữ liệu/.test(t)) by.load = c;
      else if (/2\) Ánh xạ/.test(t)) by.map = c;
      else if (/3\) Quy tắc chỉ danh/.test(t)) by.designation = c;
      else if (/4\) Hiệu chỉnh/.test(t)) by.editor = c;
      else if (/5\) Kiểm tra chất lượng/.test(t)) by.quality = c;
      else if (/6\) Nội suy/.test(t)) by.interpolate = c;
      else if (/7\) Phân tích sự kiện/.test(t)) by.operation = c;
      else if (/8\) Huấn luyện/.test(t)) by.train = c;
      else if (/9\) So sánh mô hình/.test(t)) by.compare = c;
      else if (/10A\) Nạp model vận hành/.test(t)) { by.modelImport = c; by.scada = c; }
      else if (/10B\).*Dự báo nhanh/.test(t)) { by.forecastBasic = c; by.scada = c; }
      else if (/10C1\) B2/.test(t)) { by.forecastRun = c; by.forecastMulti = c; }
      else if (/10C2\) B3/.test(t)) { by.forecastReview = c; }
      else if (/10C3\) B4/.test(t)) { by.forecastAppend = c; }
      else if (/10C4\) B5/.test(t)) { by.forecastExport = c; }
      else if (/10C\) Dự báo đa cấp/.test(t)) by.forecastMulti = c;
      else if (/10D\) Đánh giá sai số/.test(t)) by.eval = c;
      else if (/10E\) Hiệu chỉnh mô hình/.test(t)) by.calibration = c;
      else if (/10F\) So sánh trước\/sau/.test(t)) by.calCompare = c;
      else if (/10G\) Ngưỡng cảnh báo/.test(t)) by.thresholds = c;
      else if (/10\) Dùng model/.test(t)) by.scada = c;
      else if (/Kết quả/.test(t)) by.results = c;
      else if (/Bảng dữ liệu chuẩn hóa/.test(t)) by.table = c;
    });
    return by;
  }

  const FLOWS = {
    HOME: {
      name: 'Chọn luồng công việc', badge: 'WORKFLOW', hint: 'LV9.8.1 sửa B1/B2 không bị ẩn Mục 10B: B2 hiển thị cả khung Dự báo nhanh 10B và khung Dự báo đa cấp 10C1. Chọn A tạo model, S chuẩn bị chung, B dự báo, C đánh giá/hiệu chỉnh.',
      steps: [{id:'HOME', label:'Chọn luồng', cards:['dashboard'], primary:null, help:'Chọn Luồng A, S, B hoặc C. Khuyến nghị: A tạo model; S chuẩn bị dữ liệu/model; B dự báo; C đánh giá sai số và hiệu chỉnh.'}]
    },
    A: {
      name: 'Luồng A - Tạo model ngoài SCADA', badge: 'Mạng ngoài', hint: 'Chuẩn hóa dữ liệu lịch sử, xử lý chất lượng/chuyển tải, huấn luyện, so sánh chiến lược và xuất model vận hành.',
      steps: [
        {id:'A0', label:'Tổng quan tạo model', cards:['config','dashboard','guide'], primary:null, help:'Chọn chế độ Mạng ngoài, kiểm tra trạng thái dữ liệu/model và đọc luồng thao tác.'},
        {id:'A1', label:'Nạp dữ liệu lịch sử', cards:['load'], primary:'csvFile', help:'Chọn file dữ liệu lịch sử Excel/CSV/TXT/JSON. Nếu là Excel, chọn đúng sheet sau khi nạp.'},
        {id:'A2', label:'Ánh xạ cột', cards:['map','load'], primary:'applyMapBtn', help:'Ánh xạ thời gian, P, chỉ danh trạm/lộ và các cờ vận hành. Sau đó bấm Áp dụng ánh xạ.'},
        {id:'A3', label:'Tách chỉ danh', cards:['designation','table'], primary:'parseDesignationBtn', help:'Tách Đơn vị/Trạm/Lộ/Lộ nối vòng theo quy tắc LV6 để tạo khóa dự báo chuẩn.'},
        {id:'A4', label:'Hiệu chỉnh dữ liệu', cards:['editor'], primary:'applyEditsBtn', help:'Lọc theo ngày, sửa bảng, điền nhanh nhiệt độ/ngày lễ/cờ vận hành rồi lưu thay đổi vào dữ liệu.'},
        {id:'A5', label:'Kiểm tra chất lượng', cards:['quality','editor'], primary:'runQualityBtn', help:'Kiểm tra toàn bộ RAM. Khi có lỗi, bảng hiệu chỉnh chỉ hiện dòng lỗi để rà soát.'},
        {id:'A6', label:'Nội suy / mốc thiếu', cards:['interpolate','editor'], primary:'interpolateBtn', help:'Bổ sung mốc thiếu, xử lý P trống/P thấp/P lỗi. Với mất mốc, xem dòng liền trước và liền sau mốc thiếu.'},
        {id:'A7', label:'Phân tích vận hành', cards:['operation','editor'], primary:'analyzeOperationBtn', help:'Phân tích P=0, chuyển tải, lộ giảm/lộ nhận tải. Nếu khôi phục P nền chuyển tải, app điều chỉnh cả lộ nhận tải.'},
        {id:'A8', label:'Huấn luyện model', cards:['train','results'], primary:'trainAllStationsBtn', help:'Huấn luyện GBDT cho từng chỉ danh/trạm/lộ. Đây là bước tạo model nội bộ trước khi xuất.'},
        {id:'A9', label:'So sánh chiến lược', cards:['compare','results'], primary:'trainSelectBestAllBtn', help:'So sánh GBDT, Similar Day, Hybrid, Auto Blend và áp dụng chiến lược tốt nhất vào model.'},
        {id:'A10', label:'Xuất model vận hành', cards:['train','dashboard'], primary:'exportModelBtn', help:'Sau khi huấn luyện và áp dụng chiến lược, xuất model vận hành để đưa vào SCADA.'}
      ]
    },
    S: {
      name: 'Luồng S - Chuẩn bị chung cho B/C', badge: 'Dùng chung', hint: 'Dùng một lần để nạp dữ liệu vận hành/thực tế, ánh xạ, tách chỉ danh và nạp model. Sau đó chuyển sang B để dự báo hoặc C để đánh giá/hiệu chỉnh.',
      steps: [
        {id:'S0', label:'Tổng quan dùng chung', cards:['config','dashboard'], primary:null, help:'Luồng S gom các bước trùng giữa B và C. Nếu đã chuẩn bị xong dữ liệu/model, có thể chuyển thẳng sang B hoặc C.'},
        {id:'S1', label:'Nạp dữ liệu', cards:['load'], primary:'csvFile', help:'Nạp dữ liệu vận hành mới, dữ liệu thực tế sau vận hành hoặc dữ liệu cần dự báo. Dữ liệu được phép khác file đã huấn luyện nhưng cần cùng cấu trúc và đơn vị.'},
        {id:'S2', label:'Ánh xạ cột', cards:['map','load'], primary:'applyMapBtn', help:'Ánh xạ thời gian, P, chỉ danh, nhiệt độ và các cờ vận hành. Sau khi bấm áp dụng, dữ liệu sẽ vào RAM.'},
        {id:'S3', label:'Tách chỉ danh', cards:['designation','table'], primary:'parseDesignationBtn', help:'Tách chỉ danh Đơn vị/Trạm/Lộ/Lộ nối vòng. Nếu file đã có cột LV6 thì vẫn có thể kiểm tra lại.'},
        {id:'S4', label:'Nạp model vận hành', cards:['modelImport','dashboard'], primary:'modelFile', help:'Nạp model vận hành xuất từ Luồng A. Model dùng được cho dữ liệu mới nếu cùng phạm vi chỉ danh đã học.'},
        {id:'S5', label:'Kiểm tra sẵn sàng', cards:['dashboard','table'], primary:null, help:'Kiểm tra nhanh dữ liệu RAM, số chỉ danh, model và calibration. Sau đó chọn Luồng B để dự báo hoặc Luồng C để đánh giá/hiệu chỉnh.'}
      ]
    },
    B: {
      name: 'Luồng B - Dự báo vận hành', badge: 'Dự báo', hint: 'Chỉ tập trung dự báo. Các bước nạp dữ liệu, ánh xạ, tách chỉ danh và nạp model đã chuyển sang Luồng S dùng chung.',
      steps: [
        {id:'B0', label:'Sẵn sàng dự báo', cards:['dashboard','modelImport'], primary:null, help:'Dùng sau khi đã hoàn thành Luồng S. Kiểm tra model, dữ liệu RAM và phạm vi chỉ danh trước khi dự báo.'},
        {id:'B1', label:'Dự báo nhanh', cards:['forecastBasic','results','table'], primary:'forecastAllBtn', help:'Chọn trạm/lộ hoặc tất cả, số bước dự báo và chiến lược; bấm dự báo nhanh.'},
        {id:'B2', label:'Chạy dự báo / tạo báo cáo LV8.5', cards:['forecastBasic','forecastRun','results','table'], primary:'lv8QuickForecastBtn', help:'B1/B2 hiển thị đúng 10B và 10C1: có thể dùng Dự báo trạm/lộ đang chọn, Dự báo tất cả trạm/lộ, hoặc Dự báo/Tạo báo cáo đa cấp LV8.5. Bảng kết quả và biểu đồ sẽ nằm ngay trong bước này.'},
        {id:'B3', label:'Xem kết quả, Pmax, cảnh báo', cards:['forecastReview','thresholds','results'], primary:null, help:'B3 chỉ dùng để xem kết quả sau B2: Pmax, giờ Pmax, MWh, phụ tải theo ca và cảnh báo theo ngưỡng. Không cập nhật RAM ở bước này.'},
        {id:'B4', label:'Cập nhật RAM và dự báo tiếp', cards:['forecastAppend','forecastRun','table'], primary:'lv82AppendForecastBtn', help:'B4 cập nhật forecast vào dữ liệu RAM. Sau khi cập nhật xong, dùng lại các nút B2 để dự báo tiếp từ dữ liệu RAM đã bổ sung.'},
        {id:'B5', label:'Xuất forecast_summary', cards:['forecastExport','forecastReview','dashboard'], primary:'lv8ExportSummaryBtn', help:'B5 chỉ dùng để xuất forecast_summary_lv8_5.csv và kiểm tra tổng hợp cuối cùng. Nếu cần đánh giá sau vận hành, chuyển sang Luồng C.'}
      ]
    },
    C: {
      name: 'Luồng C - Đánh giá sai số và hiệu chỉnh model', badge: 'Đánh giá', hint: 'Chỉ tập trung đánh giá forecast và hiệu chỉnh model. Nếu chưa nạp dữ liệu/model, làm Luồng S trước.',
      steps: [
        {id:'C0', label:'Sẵn sàng đánh giá', cards:['dashboard','modelImport'], primary:null, help:'Dùng sau khi có dữ liệu thực tế tương ứng với forecast đã xuất. Có thể đánh giá bằng 2 file hoặc RAM + forecast hiện có.'},
        {id:'C1', label:'Nạp 2 file đánh giá', cards:['eval'], primary:'lv84EvalFilesBtn', help:'Chọn file dữ liệu thực tế và file dữ liệu dự báo. App ghép theo thời gian + chỉ danh/trạm/lộ.'},
        {id:'C2', label:'Đánh giá sai số', cards:['eval','results'], primary:'lv84EvalRamBtn', help:'Tính MAE, MAPE, RMSE, BIAS từng mốc và tổng hợp theo chỉ danh/ngày.'},
        {id:'C3', label:'Tạo hiệu chỉnh', cards:['calibration'], primary:'lv85BuildCalibrationBtn', help:'Tạo bảng bias MW/% theo toàn hệ thống, chỉ danh, giờ, loại ngày hoặc chỉ danh+giờ.'},
        {id:'C4', label:'Áp dụng hiệu chỉnh', cards:['calibration'], primary:'lv85ApplyCalibrationBtn', help:'Áp dụng bảng hiệu chỉnh vào model đang nạp. Dự báo sau sẽ có cột forecast_before_calibration_mw.'},
        {id:'C5', label:'So sánh trước/sau', cards:['calCompare','results'], primary:'lv89CompareCalibrationBtn', help:'So sánh sai số trước/sau hiệu chỉnh để quyết định có nên dùng model đã hiệu chỉnh không.'},
        {id:'C6', label:'Xuất model/hồ sơ hiệu chỉnh', cards:['calibration','dashboard','oplog'], primary:'lv85ExportModelBtn', help:'Xuất calibration, model đã hiệu chỉnh, báo cáo sai số, nhật ký thao tác và trạng thái dashboard.'}
      ]
    }
  };

  let state = { flow:'A', step:0, expert:false, side:false, compact:true, rail:false };
  function saveState(){ try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch(_){} }
  function loadState(){
    try { const raw = localStorage.getItem(STORAGE_KEY); if (raw) state = Object.assign(state, JSON.parse(raw)); } catch(_){}
    if (!FLOWS[state.flow]) state.flow = 'HOME';
    state.step = Math.max(0, Math.min(state.step || 0, FLOWS[state.flow].steps.length - 1));
  }

  function createShell(){
    const grid = document.querySelector('main .grid'); if (!grid || $('lv9WorkflowShell')) return;
    const shell = document.createElement('div');
    shell.className = 'card span12 lv9-shell'; shell.id = 'lv9WorkflowShell';
    shell.innerHTML = `
      <div class="lv9-topline">
        <div>
          <h2>LV9.8.1) Workflow / Workbench - sửa B1/B1 hiển thị 10B, B2 vẫn có 10B</h2>
          <div class="compactNote">Mặc định <b>chỉ hiện một bước đang thao tác</b>. Các mục không thuộc bước hiện tại sẽ được ẩn hoàn toàn để tăng không gian làm việc. Không khóa nút, không thay đổi lõi tính toán.</div>
        </div>
        <div class="lv9-mode-actions">
          <button class="secondary" id="lv9ExpertBtn">Chế độ chuyên gia / xem tất cả</button>
          <button class="secondary" id="lv9SideBtn">Ngăn phụ Dashboard / Nhật ký</button>
        </div>
      </div>
      <div class="lv93-compactbar" id="lv93CompactBar">
        <span class="lv93-title">LV9.8.1</span>
        <select id="lv93FlowSelect" title="Chọn luồng"></select>
        <select id="lv93StepSelect" title="Chọn bước"></select>
        <button class="secondary" id="lv93PrevBtn">←</button>
        <button class="good" id="lv93PrimaryBtn">Thực hiện</button>
        <button class="secondary" id="lv93NextBtn">→</button>
        <button class="secondary" id="lv93RailBtn">Bước</button>
        <button class="secondary" id="lv93CompactBtn">Đầy đủ</button>
        <span class="lv93-status" id="lv93StatusLine"></span>
      </div>
      <div class="lv9-flow-row" id="lv9FlowButtons"></div>
      <div class="lv9-homecards" id="lv9HomeCards"></div>
      <div class="lv9-workbench">
        <aside class="lv9-steps" id="lv9StepNav"></aside>
        <div class="lv9-current">
          <div class="lv9-status" id="lv9StatusBox"></div>
          <div class="lv9-help" id="lv9HelpBox"></div>
          <div class="lv92-true-note" id="lv92TrueNote">LV9.8.1: B1 hiển thị 10B, B2 vẫn có 10B luôn hiện cả biểu đồ và bảng kết quả; dự báo tất cả trạm/lộ có bảng tổng hợp đầy đủ.</div>
          <div class="lv9-navrow" style="margin-top:8px">
            <button class="secondary" id="lv9PrevBtn">← Quay lại</button>
            <button class="good" id="lv9PrimaryBtn">Thực hiện chính</button>
            <button class="secondary" id="lv9NextBtn">Tiếp tục →</button>
          </div>
          <div class="lv9-active-slot" id="lv9ActiveCardSlot"></div>
        </div>
      </div>
      <div class="lv9-sidepanel" id="lv9SidePanel" style="display:none">
        <div class="statusBox"><span class="pill modeBadge">Ngăn phụ LV9.8.1</span><span class="pill">Chỉ để xem nhanh, không can thiệp chức năng</span></div>
        <div class="row" style="margin-top:8px">
          <button class="secondary" data-lv9-jump="dashboard">Mở Dashboard</button>
          <button class="secondary" data-lv9-jump="oplog">Mở Nhật ký</button>
          <button class="secondary" data-lv9-jump="results">Mở Kết quả</button>
          <button class="secondary" data-lv9-jump="table">Mở Bảng dữ liệu</button>
        </div>
      </div>`;
    grid.insertBefore(shell, grid.firstElementChild);
  }

  function renderCompactBar(){
    const flowSel = $('lv93FlowSelect'), stepSel = $('lv93StepSelect');
    if (!flowSel || !stepSel) return;
    const flows = ['A','S','B','C'];
    flowSel.innerHTML = flows.map(k => `<option value="${k}" ${state.flow===k?'selected':''}>${k} - ${esc(FLOWS[k].name.replace(/^Luồng [ASBC] - /,''))}</option>`).join('');
    const flow = FLOWS[state.flow] || FLOWS.A;
    stepSel.innerHTML = flow.steps.map((s,i) => `<option value="${i}" ${i===state.step?'selected':''}>${esc(s.id)} - ${esc(s.label)}</option>`).join('');
    const line = $('lv93StatusLine');
    const step = flow.steps[state.step];
    if (line) line.textContent = `${step.id} - ${step.label} • ${state.step+1}/${flow.steps.length}`;
    const prev=$('lv93PrevBtn'), next=$('lv93NextBtn'), pri=$('lv93PrimaryBtn'), rail=$('lv93RailBtn'), compact=$('lv93CompactBtn');
    if (prev) prev.disabled = state.step <= 0;
    if (next) next.disabled = state.step >= flow.steps.length - 1;
    if (pri) pri.textContent = step.primary ? `▶ ${step.label}` : '✓ Xem bước';
    if (rail) rail.textContent = state.rail ? 'Ẩn bước' : 'Bước';
    if (compact) compact.textContent = state.compact ? 'Đầy đủ' : 'Tối giản';
    if (!flowSel.dataset.wired) {
      flowSel.addEventListener('change', () => { state.flow = flowSel.value; state.step = 0; state.expert = false; saveState(); applyWorkflow(); });
      stepSel.addEventListener('change', () => { state.step = Number(stepSel.value) || 0; state.expert = false; saveState(); applyWorkflow(); });
      prev?.addEventListener('click', () => { if (state.step > 0) { state.step--; saveState(); applyWorkflow(); } });
      next?.addEventListener('click', () => { const max=(FLOWS[state.flow]||FLOWS.A).steps.length-1; if (state.step < max) { state.step++; saveState(); applyWorkflow(); } });
      pri?.addEventListener('click', clickPrimary);
      rail?.addEventListener('click', () => { state.rail = !state.rail; saveState(); applyWorkflow(); });
      compact?.addEventListener('click', () => { state.compact = !state.compact; saveState(); applyWorkflow(); });
      flowSel.dataset.wired = '1';
    }
  }

  function renderFlowButtons(){
    const box = $('lv9FlowButtons'); if (!box) return;
    box.innerHTML = ['A','S','B','C'].map(key => {
      const flow = FLOWS[key];
      return `<button class="${state.flow===key && !state.expert ? 'good' : 'secondary'}" data-flow="${key}"><b>${key}</b> ${esc(flow.name.replace(/^Luồng [ASBC] - /,''))}</button>`;
    }).join('') + `<button class="${state.expert ? 'good' : 'secondary'}" data-expert="1">Hiển thị tất cả mục</button>`;
    box.querySelectorAll('[data-flow]').forEach(btn => btn.addEventListener('click', () => { state.flow = btn.dataset.flow; state.step = 0; state.expert = false; saveState(); applyWorkflow(); }));
    box.querySelector('[data-expert]')?.addEventListener('click', () => { state.expert = !state.expert; saveState(); applyWorkflow(); });
  }

  function renderHomeCards(){
    const home = $('lv9HomeCards'); if (!home) return;
    if (state.flow !== 'HOME' || state.expert) { home.style.display = 'none'; home.innerHTML = ''; return; }
    home.style.display = 'grid';
    home.innerHTML = ['A','S','B','C'].map(key => {
      const f = FLOWS[key];
      return `<button class="secondary" data-home-flow="${key}"><b>${esc(f.name)}</b><br><span style="font-weight:500;color:var(--muted)">${esc(f.hint)}</span></button>`;
    }).join('');
    home.querySelectorAll('[data-home-flow]').forEach(btn => btn.addEventListener('click', () => { state.flow = btn.dataset.homeFlow; state.step = 0; state.expert = false; saveState(); applyWorkflow(); }));
  }

  function renderSteps(){
    const nav = $('lv9StepNav'); if (!nav) return;
    const flow = FLOWS[state.flow];
    nav.innerHTML = flow.steps.map((s, i) => `<button class="${i===state.step ? 'active' : ''}" data-step="${i}" title="${esc(s.id)} - ${esc(s.label)}"><span>${esc(s.id)}</span><em class="lv9-step-label">${esc(s.label)}</em></button>`).join('');
    nav.querySelectorAll('[data-step]').forEach(btn => btn.addEventListener('click', () => { state.step = Number(btn.dataset.step); state.expert = false; saveState(); applyWorkflow(); }));
  }

  function clickPrimary(){
    const step = FLOWS[state.flow].steps[state.step]; const id = step.primary;
    if (!id) { setStatus('Bước này là bước hướng dẫn/kiểm tra. Hãy thao tác trong khung đang hiển thị hoặc bấm Tiếp tục.', 'warn'); return; }
    const el = $(id);
    if (!el) { setStatus(`Không tìm thấy nút/ô thao tác chính: ${id}`, 'bad'); return; }
    if (el.tagName === 'INPUT' && el.type === 'file') { el.click(); setStatus(`Đã mở hộp chọn file: ${id}`, 'ok'); return; }
    if (el.disabled) { setStatus(`Nút chính đang bị khóa bởi logic chức năng gốc: ${id}. Kiểm tra điều kiện dữ liệu/model của bước này.`, 'warn'); return; }
    el.click(); setStatus(`Đã kích hoạt thao tác chính: ${id}`, 'ok');
  }

  function setStatus(msg, level){
    const box = $('lv9StatusBox'); if (!box) return;
    const cls = level === 'bad' ? 'bad' : level === 'ok' ? 'ok' : level === 'warn' ? 'warn' : '';
    box.innerHTML = `<span class="pill ${cls}">${esc(msg)}</span><span class="pill">${new Date().toLocaleString('vi-VN')}</span>`;
  }

  function showCardsForCurrentStep(){
    ensureCardRegistry();
    undockAllCards();
    const cards = getCards();
    document.documentElement.classList.toggle('lv9-workflow-mode', !state.expert);
    document.documentElement.classList.toggle('lv9-expert-mode', !!state.expert);
    if (state.expert) {
      cards.forEach(c => c.classList.remove('lv9-hidden','lv9-current-card','lv9-docked-card'));
      return;
    }
    const map = cardIndexMap(); const step = FLOWS[state.flow].steps[state.step];
    const targetCards = [];
    step.cards.forEach(k => { if (map[k] && !targetCards.includes(map[k])) targetCards.push(map[k]); });
    const targetSet = new Set(targetCards);
    cards.forEach(c => { if (!targetSet.has(c)) c.classList.add('lv9-hidden'); });
    dockCardsIntoWorkbench(targetCards);
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
    if (pri) pri.textContent = step.primary ? `Thực hiện chính: ${step.label}` : 'Bước hướng dẫn / kiểm tra';
  }

  function applyWorkflow(){
    ensureStyle(); renderFlowButtons(); renderHomeCards(); renderSteps(); renderCompactBar(); showCardsForCurrentStep(); updateInfo();
    document.documentElement.classList.toggle('lv93-compact-mode', !!state.compact && !state.expert);
    document.documentElement.classList.toggle('lv93-show-steps', !!state.rail && !state.expert);
    const side = $('lv9SidePanel'); if (side) side.style.display = state.side ? 'block' : 'none';
  }

  function wire(){
    $('lv9PrevBtn')?.addEventListener('click', () => { if (state.step > 0) { state.step--; saveState(); applyWorkflow(); } });
    $('lv9NextBtn')?.addEventListener('click', () => { const max = FLOWS[state.flow].steps.length - 1; if (state.step < max) { state.step++; saveState(); applyWorkflow(); } });
    $('lv9PrimaryBtn')?.addEventListener('click', clickPrimary);
    $('lv9ExpertBtn')?.addEventListener('click', () => { state.expert = !state.expert; saveState(); applyWorkflow(); });
    $('lv9SideBtn')?.addEventListener('click', () => { state.side = !state.side; saveState(); applyWorkflow(); });
    document.querySelectorAll('[data-lv9-jump]').forEach(btn => btn.addEventListener('click', () => {
      const key = btn.getAttribute('data-lv9-jump'); state.expert = false;
      const flow = FLOWS[state.flow]; const idx = flow.steps.findIndex(s => s.cards.includes(key));
      if (idx >= 0) state.step = idx; else { const fb = {dashboard:['A',0], oplog:['C',6], results:['B',1], table:['B',1]}[key]; if (fb) { state.flow = fb[0]; state.step = fb[1]; } }
      saveState(); applyWorkflow();
    }));
  }

  function markVersion(){
    const h1 = document.querySelector('h1'); if (h1) h1.textContent = 'SCADA Load Forecast Offline PWA LV9.8.1';
    document.title = 'SCADA Load Forecast Offline PWA LV9.8.1';
    const vi = $('versionInfo'); if (vi) vi.innerHTML = '<span class="pill modeBadge">LV9.8.1</span><span class="pill ok">B1/B2 không ẩn 10B: đủ nút dự báo</span>';
    const cfg = $('configName'); if (cfg && /LV[0-9]/i.test(cfg.value)) cfg.value = 'SCADA_LOAD_FORECAST_LV9_7';
  }

  function init(){
    ensureStyle(); createShell(); loadState(); wire(); markVersion(); applyWorkflow();
    try { if (window.log) window.log('Sẵn sàng LV9.8.1: B1 đã hiển thị lại 10B cùng 10C1, không còn mất nút dự báo trạm/lộ và dự báo tất cả trạm/lộ.'); } catch(_){}
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
