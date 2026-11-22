/* CodeArts · pywebview 版前端逻辑 */
"use strict";

const api = () => window.pywebview.api;
const $ = (id) => document.getElementById(id);

const state = {
  files: [],            // {iid, path, pages, status, range_text, copies, source}
  selected: new Set(),
  logLines: [],
};

/* ============ 工具 ============ */
function basename(p) { return (p || "").split(/[\\/]/).pop(); }

function setStatus(main, sub, cls) {
  $("status-text").textContent = main;
  $("status-sub").textContent = sub || "";
  $("status-text").className = "status" + (cls ? " " + cls : "");
}

function pushLog(stamp, text, error) {
  const line = document.createElement("div");
  line.className = "log-line" + (error ? " err" : "");
  const t = document.createElement("span");
  t.className = "t"; t.textContent = `[${stamp}]`;
  line.appendChild(t);
  line.appendChild(document.createTextNode(text));
  const body = $("log-body");
  body.appendChild(line);
  while (body.childElementCount > 2000) body.removeChild(body.firstChild);
  body.scrollTop = body.scrollHeight;
  const summary = $("log-summary");
  summary.textContent = `[${stamp}] ${text}`;
}

window.__pushLog = pushLog;

window.__fileConverted = function (rec) {
  state.files.push(rec);
  renderFiles();
  refreshMetrics();
};

window.__printDone = function () {
  setStatus("待开始", "就绪", "");
  $("btn-print").disabled = false;
  $("btn-cancel").disabled = true;
};

/* ============ 文件列表 ============ */
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
    pages.textContent = `${f.pages} 页`;

    const st = document.createElement("span");
    st.className = "st" + (f.status === "失败" ? " err" : "") +
                   (f.status === "转换中" ? " conv" : "");
    st.textContent = f.status;

    row.append(check, name, pages, st);
    row.onclick = () => toggleSelect(f.iid);
    row.ondblclick = () => selectPreviewFile(f);
    list.appendChild(row);
  }
}

function toggleSelect(iid) {
  if (state.selected.has(iid)) state.selected.delete(iid);
  else state.selected.add(iid);
  renderFiles();
}

async function addFiles() {
  // pywebview 无原生多选文件对话框 API 时，用 create_file_input 兜底
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

async function doAdd(paths) {
  const res = await api().add_files(paths);
  if (res.ok) {
    for (const rec of res.added) state.files.push(rec);
    renderFiles();
    refreshMetrics();
    if (res.converting) {
      pushLog(now(), `正在转换 ${res.converting} 个 Office/图片文件…`);
      for (const p of paths) {
        if (!res.added.some((a) => a.source === p || a.path === p)) {
          const f = state.files.find((x) => x.source === p);
          if (!f) {
            state.files.push({ iid: `tmp-${Math.random()}`, path: p, pages: 0,
                               status: "转换中", range_text: "", copies: 1, source: p });
          }
        }
      }
      renderFiles();
    }
    // 自动选中最新文件并预览
    const last = state.files[state.files.length - 1];
    if (last && last.pages) selectPreviewFile(last);
  } else {
    pushLog(now(), res.message, true);
  }
}

function now() {
  const d = new Date();
  return [d.getHours(), d.getMinutes(), d.getSeconds()]
    .map((x) => String(x).padStart(2, "0")).join(":");
}

/* ============ 指标 ============ */
async function refreshMetrics() {
  const opts = collectOptions();
  const total = state.files.reduce((s, f) => s + (f.pages || 0), 0);
  $("m-files").textContent = state.files.length;
  $("m-pages").textContent = total;
  if (state.files.length) {
    const res = await api().compute_plan(opts);
    if (res.ok) {
      const s = res.summary;
      $("m-sheets").textContent = s.sheet_count ?? "—";
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
    flow,
  };
}

/* ============ 预览 ============ */
let previewFile = null;

async function selectPreviewFile(f) {
  previewFile = f;
  $("preview-total").textContent = `/ ${f.pages || "—"}`;
  $("preview-page").value = 1;
  await renderPreview(1);
}

async function renderPreview(pageNo) {
  if (!previewFile || !previewFile.pages) return;
  const res = await api().preview_page(previewFile.iid, pageNo, {});
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

/* ============ 打印机 ============ */
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

/* ============ 打印 ============ */
async function startPrint() {
  if (!state.files.length) { pushLog(now(), "请先添加文件", true); return; }
  const opts = collectOptions();
  const res = await api().start_print(opts);
  if (res.ok) {
    setStatus("打印中", "正在排版并发送", "running");
    $("btn-print").disabled = true;
    $("btn-cancel").disabled = false;
  } else {
    pushLog(now(), `无法开始打印：${res.message}`, true);
    setStatus("错误", res.message, "error");
  }
}

async function cancelPrint() {
  await api().cancel_print();
  pushLog(now(), "已请求取消…");
}

/* ============ 日志抽屉（丝滑拖动：纯 CSS 变量 + pointer 事件） ============ */
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
    // 直接改 CSS 变量：浏览器 GPU 合成，无需 JS 布局计算，天然丝滑
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

/* ============ 导出日志 ============ */
function exportLog() {
  const text = state.logLines.join("\n") || $("log-body").innerText;
  const blob = new Blob([$("log-body").innerText], { type: "text/plain" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  const d = new Date();
  a.download = `PDF打印日志_${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}_${now().replace(/:/g, "")}.txt`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/* ============ 初始化 ============ */
// pywebviewready 可能早于脚本加载发出：轮询兜底
function onBridgeReady(fn) {
  if (window.pywebview && window.pywebview.api) {
    fn();
    return;
  }
  const t = setInterval(() => {
    if (window.pywebview && window.pywebview.api) {
      clearInterval(t);
      fn();
    }
  }, 100);
}

onBridgeReady(async () => {
  await api().cleanup_stale();
  await loadPrinters();
  initSash();

  $("btn-add").onclick = addFiles;
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
    pushLog(now(), "待打印列表已清空。");
  };
  $("btn-refresh-printers").onclick = loadPrinters;
  $("printer-select").onchange = (e) => updatePrinterStatus(e.target.value);
  $("btn-print").onclick = startPrint;
  $("btn-cancel").onclick = cancelPrint;
  $("btn-export-log").onclick = exportLog;
  $("btn-clear-log").onclick = () => { $("log-body").innerHTML = ""; $("log-summary").textContent = "暂无日志"; };
  $("btn-open-dir").onclick = () => api().open_log_dir();
  $("copies").onchange = refreshMetrics;
  document.querySelectorAll('input[name="duplex"]').forEach((r) => r.onchange = refreshMetrics);
  $("preview-page").onchange = (e) => renderPreview(parseInt(e.target.value || "1", 10));

  // 拖放文件（WebView2 支持 DataTransfer 文件路径）
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    const paths = Array.from(e.dataTransfer.files).map((f) => f.path);
    if (paths.length) doAdd(paths);
  });

  // 窗口尺寸变化时刷新指标（打印边距等可能变化）
  window.addEventListener("resize", () => {
    // GPU 合成下 resize 无卡顿，无需任何补偿
  });

  pushLog(now(), "就绪，可以添加文件。");
});
