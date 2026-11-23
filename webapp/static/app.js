/* CodeArts · pywebview 完整版前端 */
"use strict";

const api = () => window.pywebview.api;
const $ = (id) => document.getElementById(id);

const state = {
  files: [],
  selected: new Set(),
  schemes: null,
  settings: {},
  previewFile: null,
  printing: false,
  paused: false,
};

/* ================= 工具 ================= */
function basename(p) { return (p || "").split(/[\\/]/).pop(); }
function now() {
  const d = new Date();
  return [d.getHours(), d.getMinutes(), d.getSeconds()]
    .map((x) => String(x).padStart(2, "0")).join(":");
}
function setStatus(main, sub, cls) {
  $("status-text").textContent = main;
  $("status-sub").textContent = sub || "";
  $("status-text").className = "status" + (cls ? " " + cls : "");
}
function showModal(id) { $(id).style.display = "flex"; }
function hideModal(id) { $(id).style.display = "none"; }
window.hideModal = hideModal;

/* ================= 后端事件 ================= */
window.__pushLog = function (stamp, text, error) {
  const line = document.createElement("div");
  line.className = "log-line" + (error ? " err" : "");
  const t = document.createElement("span");
  t.className = "t"; t.textContent = `[${stamp}]`;
  line.appendChild(t);
  line.appendChild(document.createTextNode(text));
  const body = $("log-body");
  body.appendChild(line);
  while (body.childElementCount > 3000) body.removeChild(body.firstChild);
  body.scrollTop = body.scrollHeight;
  $("log-summary").textContent = `[${stamp}] ${text}`;
};

window.__event = function (kind, payload) {
  switch (kind) {
    case "log":
      window.__pushLog(payload.stamp, payload.text, payload.error);
      break;
    case "rows":
      for (const [iid, status] of payload) {
        const f = state.files.find((x) => x.iid === iid);
        if (f) f.status = status;
      }
      renderFiles();
      break;
    case "file-update": {
      const idx = state.files.findIndex((x) => x.iid === payload.iid);
      if (idx >= 0) state.files[idx] = payload;
      else state.files.push(payload);
      renderFiles();
      refreshMetrics();
      break;
    }
    case "status":
      setStatus(payload, "", "running");
      break;
    case "progress":
      $("progress-track").classList.add("active");
      $("progress-fill").style.width =
        `${Math.round(100 * payload.index / payload.total)}%`;
      $("status-sub").textContent = payload.text;
      break;
    case "print-done":
      state.printing = false;
      state.paused = false;
      $("btn-print").disabled = false;
      $("btn-pause").disabled = true;
      $("btn-cancel").disabled = true;
      $("progress-track").classList.remove("active");
      if (payload.completed) {
        setStatus("完成", "全部打印阶段已完成", "");
      } else if (payload.cancelled) {
        setStatus("已取消", "任务已取消", "error");
      } else {
        setStatus("失败", payload.error, "error");
      }
      break;
    case "flip-modal":
      $("flip-title").textContent = payload.title;
      $("flip-body").innerHTML =
        payload.detail.map((l) => `<p>${l}</p>`).join("") +
        `<p class="muted small">${payload.sample}</p>`;
      showModal("modal-flip");
      break;
    case "phase-confirm":
      $("phase-body").innerHTML =
        `<p><b>${payload.text}</b> 的完成状态无法确认：</p>` +
        `<p class="muted">${payload.reason}</p>` +
        `<p>请检查打印机出纸状态，若已完成请点击继续，否则取消。</p>`;
      showModal("modal-phase");
      break;
  }
};

window.flipAnswer = (proceed) => { hideModal("modal-flip"); api().confirm_flip(proceed); };
window.phaseAnswer = (proceed) => { hideModal("modal-phase"); api().confirm_phase(proceed); };

/* ================= 文件列表 ================= */
function renderFiles() {
  const list = $("file-list");
  const empty = $("empty-state");
  list.querySelectorAll(".file-row").forEach((n) => n.remove());
  if (!state.files.length) {
    empty.style.display = "flex";
    $("file-count").textContent = "0 个文件";
    return;
  }
  empty.style.display = "none";
  $("file-count").textContent = `${state.files.length} 个文件`;

  for (const f of state.files) {
    const row = document.createElement("div");
    row.className = "file-row" + (state.selected.has(f.iid) ? " selected" : "");
    row.dataset.iid = f.iid;

    const check = document.createElement("input");
    check.type = "checkbox";
    check.checked = state.selected.has(f.iid);
    check.onclick = (e) => { e.stopPropagation(); toggleSelect(f.iid); };

    const name = document.createElement("span");
    name.className = "name";
    name.textContent = basename(f.source || f.path);
    name.title = f.source || f.path;

    const pages = document.createElement("span");
    pages.className = "pages";
    const rt = (f.range_text || "").trim();
    pages.textContent = `${rt || "全部"} / ${f.pages || "?"} 页 × ${f.copies || 1}`;

    const st = document.createElement("span");
    st.className = "st" + (f.status === "失败" ? " err" : "") +
                   (f.status === "转换中" ? " conv" : "");
    st.textContent = f.status;

    row.append(check, name, pages, st);
    row.onclick = () => toggleSelect(f.iid);
    row.ondblclick = () => { selectPreviewFile(f); openFileOpt(f.iid); };
    list.appendChild(row);
  }
}

function toggleSelect(iid) {
  if (state.selected.has(iid)) state.selected.delete(iid);
  else state.selected.add(iid);
  renderFiles();
}

async function doAdd(paths) {
  const res = await api().add_files(paths);
  if (res.ok) {
    for (const rec of res.added) state.files.push(rec);
    renderFiles();
    refreshMetrics();
    const last = state.files[state.files.length - 1];
    if (last && last.pages) selectPreviewFile(last);
  } else {
    window.__pushLog(now(), res.message, true);
  }
}

async function addFiles() {
  const input = document.createElement("input");
  input.type = "file";
  input.multiple = true;
  input.accept = ".pdf,.doc,.docx,.wps,.png,.jpg,.jpeg,.bmp,.tif,.tiff,.webp,.gif";
  input.onchange = async () => {
    const paths = Array.from(input.files).map((f) => f.path);
    if (paths.length) await doAdd(paths);
  };
  input.click();
}

async function addFolder() {
  // 用文件选择器选多个文件近似（WebView 无目录选择 API 时退化为文件选择）
  const input = document.createElement("input");
  input.type = "file";
  input.multiple = true;
  input.webkitdirectory = true;
  input.onchange = async () => {
    const paths = Array.from(input.files).map((f) => f.path);
    if (paths.length) await doAdd(paths);
  };
  input.click();
}

/* ================= 单文件设置 ================= */
let fileoptIid = null;
function openFileOpt(iid) {
  const f = state.files.find((x) => x.iid === iid);
  if (!f) return;
  fileoptIid = iid;
  $("fileopt-title").textContent = basename(f.source || f.path);
  $("fileopt-range").value = f.range_text || "";
  $("fileopt-copies").value = f.copies || 1;
  showModal("modal-fileopt");
}

async function applyFileOpt() {
  const range = $("fileopt-range").value.trim();
  const copies = parseInt($("fileopt-copies").value || "1", 10);
  await api().set_file_options(fileoptIid, range, copies);
  const f = state.files.find((x) => x.iid === fileoptIid);
  if (f) { f.range_text = range; f.copies = copies; }
  hideModal("modal-fileopt");
  renderFiles();
  refreshMetrics();
}
window.applyFileOpt = applyFileOpt;

/* ================= 指标 ================= */
async function refreshMetrics() {
  const opts = collectOptions();
  const total = state.files.reduce((s, f) => s + (f.pages || 0), 0);
  $("m-files").textContent = state.files.length;
  $("m-pages").textContent = total;
  if (state.files.length) {
    const res = await api().compute_plan(opts);
    if (res.ok) {
      const s = res.summary;
      $("m-sheets").textContent = s.sheets ?? s.sheet_count ?? "—";
      $("m-flips").textContent = s.flip_count ?? "—";
    }
  } else {
    $("m-sheets").textContent = "0";
    $("m-flips").textContent = "0";
  }
}

function collectOptions() {
  const duplex = document.querySelector('input[name="duplex"]:checked').value;
  const flow = document.querySelector('input[name="flow"]:checked').value;
  return {
    printer: $("printer-select").value,
    copies: parseInt($("copies").value || "1", 10),
    manual_duplex: duplex === "manual",
    print_flow: flow,
    ...state.advanced,
  };
}

/* ================= 预览 ================= */
async function selectPreviewFile(f) {
  state.previewFile = f;
  $("preview-total").textContent = `/ ${f.pages || "—"}`;
  $("preview-page").value = 1;
  await renderPreview(1);
}

async function renderPreview(pageNo) {
  if (!state.previewFile || !state.previewFile.pages) return;
  const res = await api().preview_page(state.previewFile.iid, pageNo, collectOptions());
  const stage = $("preview-stage");
  if (res.ok) {
    stage.innerHTML = "";
    const img = document.createElement("img");
    img.src = res.dataurl;
    stage.appendChild(img);
    $("preview-page").value = res.page;
    $("preview-total").textContent = `/ ${res.total}`;
  }
}

async function showPlanDetail() {
  if (!state.previewFile) { window.__pushLog(now(), "请先选择文件", true); return; }
  const res = await api().plan_description(state.previewFile.iid, collectOptions());
  if (res.ok) {
    $("plan-text").textContent = res.text + "\n\n" + res.sample;
    showModal("modal-plan");
  } else {
    window.__pushLog(now(), res.message, true);
  }
}

/* ================= 打印机 ================= */
async function loadPrinters() {
  const res = await api().list_printers();
  if (!res.ok) { $("printer-name").textContent = "打印机列表获取失败"; return; }
  const sel = $("printer-select");
  sel.innerHTML = "";
  for (const name of res.printers) {
    const opt = document.createElement("option");
    opt.value = name; opt.textContent = name;
    if (name === res.default) opt.selected = true;
    sel.appendChild(opt);
  }
  $("printer-name").textContent = res.default || res.printers[0] || "无打印机";
  updatePrinterStatus(sel.value);
}

async function updatePrinterStatus(name) {
  if (!name) return;
  const res = await api().printer_status(name);
  const dot = $("printer-dot");
  if (res.ok && res.ready) {
    dot.className = "dot ok";
    $("printer-name").textContent = name;
  } else {
    dot.className = "dot bad";
    $("printer-name").textContent = `${name}（异常）`;
  }
}

/* ================= 打印 ================= */
async function startPrint() {
  if (!state.files.length) { window.__pushLog(now(), "请先添加文件", true); return; }
  if (state.files.some((f) => f.status === "转换中")) {
    window.__pushLog(now(), "部分文件还在转换，请稍候", true); return;
  }
  const opts = collectOptions();
  const res = await api().start_print(opts);
  if (res.ok) {
    state.printing = true;
    setStatus("打印中", "正在排版并发送", "running");
    $("btn-print").disabled = true;
    $("btn-pause").disabled = false;
    $("btn-cancel").disabled = false;
  } else {
    window.__pushLog(now(), `无法开始打印：${res.message}`, true);
    setStatus("错误", res.message, "error");
  }
}

async function pauseToggle() {
  const res = await api().pause_toggle();
  if (res.ok) {
    state.paused = res.paused;
    $("btn-pause").textContent = state.paused ? "继续" : "暂停";
    setStatus(state.paused ? "已暂停" : "打印中", "", state.paused ? "" : "running");
  }
}

async function cancelPrint() {
  await api().cancel_print();
  window.__pushLog(now(), "已请求取消…");
}

/* ================= 高级设置 ================= */
state.advanced = {
  paper_name: "A4", orientation: "auto", scale_mode: "shrink",
  scale_percent: 100, nup: 1, nup_border: false, page_order: "row",
  booklet: false, auto_rotate: true, annots: true, skip_blank: false,
  margin_left_mm: 10, margin_top_mm: 10, margin_right_mm: 10, margin_bottom_mm: 8,
  crop_left_mm: 0, crop_top_mm: 0, crop_right_mm: 0, crop_bottom_mm: 0,
  color_mode: "color", print_quality: "normal", dpi: 300, ink_saver: false,
  wait_mode: "queue", delay: 20,
  source_name: "自动选择", collate: "no_collate", reverse: false,
  copy_exec_on: false, keep_output: false, auto_open: false, output_dir: "",
};

function buildAdvancedForm() {
  const schemes = state.schemes;
  const a = state.advanced;
  const body = $("advanced-body");
  const opt = (v, cur, l) => `<option value="${v}" ${v === cur ? "selected" : ""}>${l}</option>`;
  const sel = (id, list, cur) =>
    `<select id="adv-${id}" class="input">${list.map((x) => opt(x.value, cur, x.label)).join("")}</select>`;
  const num = (id, val, min = 0, max = 9999) =>
    `<input type="number" id="adv-${id}" class="input num" value="${val}" min="${min}" max="${max}">`;
  const chk = (id, label, cur) =>
    `<label class="checkbox-row"><input type="checkbox" id="adv-${id}" ${cur ? "checked" : ""}>${label}</label>`;

  body.innerHTML = `
  <div class="adv-section"><div class="adv-section-title">纸张与方向</div>
    <div class="adv-grid">
      <label class="form-row"><span>纸张</span>${sel("paper_name", schemes.paper_names.map((p) => ({value: p, label: p})), a.paper_name)}</label>
      <label class="form-row"><span>方向</span>${sel("orientation", schemes.orientations, a.orientation)}</label>
      <label class="form-row"><span>缩放</span>${sel("scale_mode", schemes.scale_modes, a.scale_mode)}</label>
      <label class="form-row"><span>缩放比例 %</span>${num("scale_percent", a.scale_percent, 1, 400)}</label>
      <label class="form-row"><span>纸盒</span><input id="adv-source_name" class="input" value="${a.source_name}"></label>
      ${chk("auto_rotate", "自动旋转以适合纸张", a.auto_rotate)}
    </div>
  </div>
  <div class="adv-section"><div class="adv-section-title">拼版与装订</div>
    <div class="adv-grid">
      <label class="form-row"><span>每面页数 N-up</span>${sel("nup", schemes.nup, a.nup)}</label>
      <label class="form-row"><span>页序</span>${sel("page_order", schemes.page_orders, a.page_order)}</label>
      ${chk("nup_border", "拼版加边框", a.nup_border)}
      ${chk("booklet", "小册子装订", a.booklet)}
    </div>
  </div>
  <div class="adv-section"><div class="adv-section-title">页边距（mm）</div>
    <div class="adv-grid">
      <label class="form-row"><span>左</span>${num("margin_left_mm", a.margin_left_mm)}</label>
      <label class="form-row"><span>上</span>${num("margin_top_mm", a.margin_top_mm)}</label>
      <label class="form-row"><span>右</span>${num("margin_right_mm", a.margin_right_mm)}</label>
      <label class="form-row"><span>下</span>${num("margin_bottom_mm", a.margin_bottom_mm)}</label>
    </div>
  </div>
  <div class="adv-section"><div class="adv-section-title">内容裁剪（mm，0=不裁）</div>
    <div class="adv-grid">
      <label class="form-row"><span>左</span>${num("crop_left_mm", a.crop_left_mm)}</label>
      <label class="form-row"><span>上</span>${num("crop_top_mm", a.crop_top_mm)}</label>
      <label class="form-row"><span>右</span>${num("crop_right_mm", a.crop_right_mm)}</label>
      <label class="form-row"><span>下</span>${num("crop_bottom_mm", a.crop_bottom_mm)}</label>
    </div>
  </div>
  <div class="adv-section"><div class="adv-section-title">打印质量</div>
    <div class="adv-grid">
      <label class="form-row"><span>色彩</span>${sel("color_mode", schemes.color_modes, a.color_mode)}</label>
      <label class="form-row"><span>质量</span>${sel("print_quality", schemes.qualities, a.print_quality)}</label>
      <label class="form-row"><span>DPI</span>${num("dpi", a.dpi, 72, 1200)}</label>
      ${chk("ink_saver", "省墨模式", a.ink_saver)}
      ${chk("annots", "打印批注/注释", a.annots)}
      ${chk("skip_blank", "跳过空白页", a.skip_blank)}
    </div>
  </div>
  <div class="adv-section"><div class="adv-section-title">流程与份数</div>
    <div class="adv-grid">
      <label class="form-row"><span>完成等待</span>${sel("wait_mode", schemes.wait_modes, a.wait_mode)}</label>
      <label class="form-row"><span>固定延时（秒）</span>${num("delay", a.delay, 0, 600)}</label>
      ${chk("reverse", "逆序打印", a.reverse)}
      ${chk("copy_exec_on", "按份逐份执行", a.copy_exec_on)}
      ${chk("keep_output", "保留排版输出 PDF", a.keep_output)}
      ${chk("auto_open", "完成后打开输出目录", a.auto_open)}
      <label class="form-row"><span>输出目录</span><input id="adv-output_dir" class="input" value="${a.output_dir}"></label>
    </div>
  </div>`;
}

function readAdvancedForm() {
  const a = state.advanced;
  const ids = ["paper_name", "orientation", "scale_mode", "nup", "page_order",
               "color_mode", "print_quality", "wait_mode", "source_name", "output_dir"];
  for (const k of ids) {
    const el = $(`adv-${k}`);
    if (el) a[k] = k === "nup" ? parseInt(el.value, 10) : el.value;
  }
  const nums = ["scale_percent", "dpi", "delay", "margin_left_mm", "margin_top_mm",
                "margin_right_mm", "margin_bottom_mm", "crop_left_mm", "crop_top_mm",
                "crop_right_mm", "crop_bottom_mm"];
  for (const k of nums) {
    const el = $(`adv-${k}`);
    if (el) a[k] = parseFloat(el.value);
  }
  const chks = ["nup_border", "booklet", "auto_rotate", "ink_saver", "annots",
                "skip_blank", "reverse", "copy_exec_on", "keep_output", "auto_open"];
  for (const k of chks) {
    const el = $(`adv-${k}`);
    if (el) a[k] = el.checked;
  }
}

window.applyAdvanced = function () {
  readAdvancedForm();
  hideModal("modal-advanced");
  refreshMetrics();
};
window.resetAdvanced = function () {
  state.advanced = {
    paper_name: "A4", orientation: "auto", scale_mode: "shrink", scale_percent: 100,
    nup: 1, nup_border: false, page_order: "row", booklet: false, auto_rotate: true,
    annots: true, skip_blank: false, margin_left_mm: 10, margin_top_mm: 10,
    margin_right_mm: 10, margin_bottom_mm: 8, crop_left_mm: 0, crop_top_mm: 0,
    crop_right_mm: 0, crop_bottom_mm: 0, color_mode: "color", print_quality: "normal",
    dpi: 300, ink_saver: false, wait_mode: "queue", delay: 20,
    source_name: "自动选择", collate: "no_collate", reverse: false,
    copy_exec_on: false, keep_output: false, auto_open: false, output_dir: "",
  };
  buildAdvancedForm();
};

/* ================= 日志抽屉 ================= */
function initSash() {
  const sash = $("sash");
  let dragging = false;
  sash.addEventListener("pointerdown", (e) => {
    dragging = true;
    sash.setPointerCapture(e.pointerId);
    document.body.style.cursor = "sb_v_double_arrow";
  });
  sash.addEventListener("pointermove", (e) => {
    if (!dragging) return;
    const h = window.innerHeight - e.clientY;
    const clamped = Math.max(96, Math.min(h, Math.floor(window.innerHeight / 2)));
    document.documentElement.style.setProperty("--log-h", clamped + "px");
  });
  sash.addEventListener("pointerup", (e) => {
    dragging = false;
    sash.releasePointerCapture(e.pointerId);
    document.body.style.cursor = "";
  });
}

function exportLog() {
  const blob = new Blob([$("log-body").innerText], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  const d = new Date();
  a.download = `PDF打印日志_${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}_${now().replace(/:/g, "")}.txt`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ================= 初始化 ================= */
function onBridgeReady(fn) {
  if (window.pywebview && window.pywebview.api) { fn(); return; }
  const t = setInterval(() => {
    if (window.pywebview && window.pywebview.api) { clearInterval(t); fn(); }
  }, 100);
}

onBridgeReady(async () => {
  await api().cleanup_stale();
  const st = await api().get_settings();
  if (st.ok) {
    state.schemes = st.schemes;
    state.settings = st.settings;
  }
  await loadPrinters();
  initSash();

  $("btn-add").onclick = addFiles;
  $("btn-add-dir").onclick = addFolder;
  $("btn-remove").onclick = async () => {
    const iids = [...state.selected];
    if (!iids.length) return;
    await api().remove_files(iids);
    state.files = state.files.filter((f) => !state.selected.has(f.iid));
    state.selected.clear();
    renderFiles();
    refreshMetrics();
  };
  $("btn-clear").onclick = async () => {
    await api().clear_files();
    state.files = [];
    state.selected.clear();
    renderFiles();
    refreshMetrics();
    window.__pushLog(now(), "待打印列表已清空。");
  };
  $("btn-add-recent").onclick = async () => {
    const p = $("recent-input").value.trim();
    if (p) { const res = await api().add_recent(p); if (res.ok) { for (const r of res.added) state.files.push(r); renderFiles(); refreshMetrics(); } }
  };
  $("btn-refresh-printers").onclick = loadPrinters;
  $("printer-select").onchange = (e) => updatePrinterStatus(e.target.value);
  $("btn-print").onclick = startPrint;
  $("btn-pause").onclick = pauseToggle;
  $("btn-cancel").onclick = cancelPrint;
  $("btn-export-log").onclick = exportLog;
  $("btn-clear-log").onclick = () => { $("log-body").innerHTML = ""; $("log-summary").textContent = "暂无日志"; };
  $("btn-open-dir").onclick = () => api().open_log_dir();
  $("btn-advanced").onclick = () => { buildAdvancedForm(); showModal("modal-advanced"); };
  $("btn-plan-detail").onclick = showPlanDetail;
  $("copies").onchange = refreshMetrics;
  document.querySelectorAll('input[name="duplex"],input[name="flow"]').forEach((r) => r.onchange = refreshMetrics);
  $("preview-page").onchange = (e) => renderPreview(parseInt(e.target.value || "1", 10));

  // 拖放
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    const paths = Array.from(e.dataTransfer.files).map((f) => f.path);
    if (paths.length) doAdd(paths);
  });

  // 快捷键：Ctrl+P 打印 / Esc 取消 / Ctrl+A 全选 / Del 移除选中
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.key === "p") { e.preventDefault(); startPrint(); }
    else if (e.key === "Escape") {
      if (state.printing) cancelPrint();
      else document.querySelectorAll(".modal-mask").forEach((m) => m.style.display = "none");
    } else if (e.ctrlKey && e.key === "a") {
      e.preventDefault();
      state.selected = new Set(state.files.map((f) => f.iid));
      renderFiles();
    } else if (e.key === "Delete" && state.selected.size) {
      $("btn-remove").click();
    }
  });

  window.__pushLog(now(), "就绪，可以添加文件。");
});
