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
  previewTotal: 0,
  printing: false,
  paused: false,
  // 用户输入的“指定范围”原文：非 custom/recent 模式输入框只显示提示文案，
  // 原文本单独保留，切换模式回来时不会丢失。
  rangeText: "1-5,8,10-12",
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

/* 应用内 Promise 式确认框：替代系统原生确认弹窗 提示，避免打断操作。 */
let confirmResolve = null;
function confirmAction({ title = "请确认", lines = [], confirmLabel = "确定", danger = false } = {}) {
  return new Promise((resolve) => {
    confirmResolve = resolve;
    $("confirm-title").textContent = title;
    fillTextLines($("confirm-body"), lines);
    const ok = $("btn-confirm-ok");
    ok.textContent = confirmLabel;
    ok.className = danger ? "btn ghost danger" : "btn accent";
    showModal("modal-confirm");
    $("btn-confirm-cancel").focus();
  });
}
window.answerConfirm = function (value) {
  hideModal("modal-confirm");
  const resolve = confirmResolve;
  confirmResolve = null;   // 防止重复确认/Esc 二次 resolve
  if (resolve) resolve(!!value);
};
function fillTextLines(element, lines) {
  element.replaceChildren();
  for (const text of lines) {
    const p = document.createElement("p");
    p.textContent = text;
    element.appendChild(p);
  }
}
/* 用户可见错误统一走这里：去掉 Python 异常类名前缀，空值给可操作的兜底提示。 */
function errorText(err, fallback = "请稍后重试。") {
  const raw = err && err.message ? String(err.message) : (err == null ? "" : String(err));
  const cleaned = raw
    .replace(/^[A-Za-z_][A-Za-z0-9_.]*(?:Error|Exception|Warning):\s*/, "")
    .trim();
  return cleaned || fallback;
}

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
    case "files-replaced":
      replaceFiles(payload);
      break;
    case "recent-updated":
      renderRecent(payload);
      break;
    case "rows":
      for (const [iid, status] of payload) {
        const f = state.files.find((x) => x.iid === iid);
        if (f) f.status = status;
      }
      renderFiles();
      break;
    case "file-update": {
      // 转换完成只替换记录对象：已选中的 iid 保持选中，正在预览的文件
      // 直接更新对象并按当前页码重绘，避免跳回第 1 面或丢失选中态。
      const idx = state.files.findIndex((x) => x.iid === payload.iid);
      if (idx >= 0) state.files[idx] = payload;
      else state.files.push(payload);
      const isPreviewing = state.previewFile && state.previewFile.iid === payload.iid;
      if (isPreviewing) state.previewFile = payload;
      renderFiles();
      refreshMetrics();
      if (isPreviewing) {
        if (payload.pages) {
          const wanted = Math.max(1, parseInt($("preview-page").value || "1", 10));
          renderPreview(Math.min(wanted, payload.pages));
        } else {
          selectPreviewFile(payload);
        }
      } else if (!state.previewFile && payload.pages) {
        selectPreviewFile(payload);
      }
      break;
    }
    case "files-added": {
      // 原生拖放通道入队的文件：与“添加文件”一致，新增文件默认全部选中。
      const added = Array.isArray(payload.added) ? payload.added : [];
      for (const rec of added) {
        if (!state.files.some((x) => x.iid === rec.iid)) state.files.push(rec);
        state.selected.add(rec.iid);
      }
      renderFiles();
      refreshMetrics();
      const last = added[added.length - 1];
      if (last) selectPreviewFile(last);
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
      updateControlStates();
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
      $("btn-flip-confirm").textContent = payload.proceed_label || "已翻面，继续打印反面";
      fillTextLines($("flip-body"), [...payload.detail, payload.sample]);
      showModal("modal-flip");
      break;
    case "phase-confirm":
      fillTextLines($("phase-body"), [
        `${payload.text} 的完成状态暂时无法确认：`, payload.reason,
        "请检查打印机出纸状态，若已完成请点击继续，否则取消。"]);
      showModal("modal-phase");
      break;
    case "summary":
      // 任务结束（完成/取消/失败）后的最终汇总，由用户手动关闭。
      $("summary-title").textContent = payload.title || "打印结果";
      fillTextLines($("summary-body"), String(payload.message || "").split("\n"));
      showModal("modal-summary");
      break;
  }
};

window.flipAnswer = (proceed) => { hideModal("modal-flip"); api().confirm_flip(proceed); };
window.phaseAnswer = (proceed) => { hideModal("modal-phase"); api().confirm_phase(proceed); };

/* ================= 文件列表 ================= */
const RANGE_TEXT_TITLE = "全局范围仅对未单独设置的文件生效；各文件可独立设置范围与份数";
/* 非自定义模式下 #range-text 的只读说明文案：切换模式瞬间不能被当成用户原文。 */
const RANGE_DISPLAY_TEXTS = new Set(["全部", "各文件独立", "仅奇数页", "仅偶数页", "当前页"]);

/* 是否存在逐文件自定义范围：有值时全局“全部”实际只作用于未单独设置的文件。 */
function hasPerFileRange() {
  return state.files.some((f) => (f.range_text || "").trim().length > 0);
}

/* 非 custom/recent 模式下 #range-text 显示的只读说明。 */
function rangeModeText(mode) {
  if (mode === "odd") return "仅奇数页";
  if (mode === "even") return "仅偶数页";
  if (mode === "current") return "当前页";
  return "全部";
}

/* 统一刷新 #range-text：syncInputStates / refreshMetrics / renderFiles 之后都必须调用。
   custom/recent 恢复用户原文；其余模式只显示说明，输入框保持 disabled。 */
function updateRangeTextDisplay() {
  const input = $("range-text");
  if (!input) return;
  const mode = $("range-mode").value;
  input.title = RANGE_TEXT_TITLE;
  if (mode === "custom" || mode === "recent") {
    input.value = state.rangeText;
  } else if (mode === "all") {
    input.value = hasPerFileRange() ? "各文件独立" : "全部";
  } else {
    input.value = rangeModeText(mode);
  }
}

function renderFiles() {
  updateRangeTextDisplay();
  const list = $("file-list");
  const empty = $("empty-state");
  const selectionBar = $("selection-bar");
  list.querySelectorAll(".file-row").forEach((n) => n.remove());
  if (!state.files.length) {
    empty.style.display = "flex";
    // 空状态隐藏选中操作条：否则它在 1024×680 下会盖住圆形加号/文件图标。
    selectionBar.hidden = true;
    $("file-count").textContent = "0 个文件";
    clearPreview();
    updateControlStates();
    return;
  }
  empty.style.display = "none";
  selectionBar.hidden = false;
  $("file-count").textContent = `${state.files.length} 个文件`;

  for (const f of state.files) {
    const row = document.createElement("div");
    row.className = "file-row" + (state.selected.has(f.iid) ? " selected" : "") +
      (state.previewFile && state.previewFile.iid === f.iid ? " previewing" : "");
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
    // range_error 由 compute_plan 逐文件回传：该文件范围在当前页数下没有有效页。
    // 行里必须显示“范围无效”，而不是沿用旧的有效范围/页数。
    const rangeError = f.range_error || "";
    // effective_range 由 compute_plan 按当前范围模式写入；逐文件自定义范围优先。
    const range = (f.effective_range || f.range_text || "全部").trim();
    const selectedPages = f.selected_pages ?? f.pages ?? "?";
    const plan = f.plan_text ? ` · ${f.plan_text}` : "";
    pages.textContent = `${selectedPages} 页 · ${range}${plan}`;
    pages.title = rangeError || `原始 ${f.pages || "?"} 页；范围 ${range}${plan}`;
    const st = document.createElement("span");
    st.className = "st" + (rangeError || f.status === "失败" ? " err" : "") +
                   (f.status === "转换中" ? " conv" : "");
    st.textContent = rangeError ? "范围无效" : f.status;
    if (rangeError) st.title = rangeError;
    row.append(check, name, pages, st);
    // 行单击只切换预览；选中状态仅由左侧复选框切换，已选中行再点不会取消。
    row.onclick = () => { selectPreviewFile(f); };
    // 双击仍然打开单文件设置；右键菜单行为保持不变。
    row.ondblclick = () => { selectPreviewFile(f); openFileOpt(f.iid, "range", "current"); };
    row.oncontextmenu = (e) => {
      e.preventDefault();
      if (!state.selected.has(f.iid)) { state.selected = new Set([f.iid]); renderFiles(); }
      openFileOpt(f.iid, "range", "current");
    };
    list.appendChild(row);
  }
  updateControlStates();
}

function toggleSelect(iid) {
  if (state.selected.has(iid)) state.selected.delete(iid);
  else state.selected.add(iid);
  renderFiles();
}

/* “更多”菜单与单文件设置的操作目标：优先当前预览文件，其次已选第一个。
   默认全选时不能再把整个选中集合当成目标，否则上移/下移会互相阻挡、
   资源管理器也只会定位到选中集合里的第一个文件。 */
function currentFileIid() {
  if (state.previewFile && state.files.some((f) => f.iid === state.previewFile.iid)) {
    return state.previewFile.iid;
  }
  const first = [...state.selected][0];
  return first || null;
}

function currentFileRecord() {
  const iid = currentFileIid();
  return iid ? state.files.find((f) => f.iid === iid) || null : null;
}

function updateControlStates() {
  const hasFiles = state.files.length > 0;
  const hasSelection = state.selected.size > 0;
  const hasCurrent = !!currentFileIid();
  $("btn-remove").disabled = !hasSelection || state.printing;
  $("btn-clear").disabled = !hasFiles || state.printing;
  // 上移/下移/定位只作用于当前文件，有当前文件即可用，不要求多选。
  for (const id of ["btn-move-up", "btn-move-down", "btn-show-explorer"]) {
    $(id).disabled = !hasCurrent || state.printing;
  }
  // 单文件设置同样以当前文件为目标；“应用到全部”只要有文件即可打开。
  for (const id of ["btn-file-range", "btn-file-copies", "btn-file-restore"]) {
    $(id).disabled = !hasCurrent || state.printing;
  }
  $("btn-file-apply-all").disabled = !hasFiles || state.printing;
  const ready = state.files.some((f) => f.pages && f.status !== "失败");
  // 任一文件范围无效（如 1 页文件设了第 2 页）时不得开始打印：否则该文件
  // 会被后端当作无有效页跳过，用户却以为整批都会打印。
  const hasRangeError = state.files.some((f) => f.range_error);
  $("btn-print").disabled = state.printing || !ready || hasRangeError;
  const info = $("selection-info");
  if (hasSelection) info.textContent = `已选 ${state.selected.size} 个文件`;
  else info.textContent = "选择文件后可单独设置范围与份数";
}

function replaceFiles(files) {
  const previousIid = state.previewFile && state.previewFile.iid;
  state.files = Array.isArray(files) ? files : [];
  state.selected = new Set([...state.selected].filter((iid) =>
    state.files.some((f) => f.iid === iid)));
  renderFiles();
  const current = state.files.find((f) => f.iid === previousIid);
  if (current) selectPreviewFile(current);
  else syncPreviewAfterFilesChanged();
  refreshMetrics();
}

function renderRecent(paths) {
  const select = $("recent-select");
  const previous = select.value;
  select.replaceChildren();
  const files = Array.isArray(paths) ? paths : [];
  if (!files.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "暂无最近文件";
    select.appendChild(option);
  } else {
    for (const path of files) {
      const option = document.createElement("option");
      option.value = path;
      option.textContent = basename(path);
      option.title = path;
      select.appendChild(option);
    }
    select.value = files.includes(previous) ? previous : files[0];
  }
  $("btn-add-recent").disabled = files.length === 0;
  // 清除记录按钮必须与列表状态保持一致：有记录才可用。
  $("btn-clear-recent").disabled = files.length === 0;
}

async function refreshRecent() {
  const res = await api().get_recent();
  if (res.ok) renderRecent(res.files);
}

async function doAdd(paths) {
  const res = await api().add_files(paths);
  if (res.ok) {
    const added = Array.isArray(res.added) ? res.added : [];
    for (const rec of added) {
      if (!state.files.some((x) => x.iid === rec.iid)) state.files.push(rec);
      // 新添加的文件默认选中，随即可直接设置范围/份数或移除。
      state.selected.add(rec.iid);
    }
    renderFiles();
    refreshMetrics();
    refreshRecent();
    const last = added[added.length - 1];
    if (last) selectPreviewFile(last);
  } else {
    window.__pushLog(now(), res.message, true);
  }
}

async function addFiles() {
  const res = await api().choose_files();
  if (!res.ok) {
    window.__pushLog(now(), `无法打开文件选择器：${res.message}`, true);
    return;
  }
  if (res.paths.length) await doAdd(res.paths);
}

async function addFolder() {
  // 使用系统文件夹对话框和本地路径，避免 WebView2 的“上传整个文件夹”安全提示。
  const res = await api().choose_folder();
  if (!res.ok) {
    window.__pushLog(now(), `无法选择文件夹：${res.message}`, true);
    return;
  }
  if (res.paths.length) await doAdd(res.paths);
}

/* ================= 单文件/多选设置 ================= */
/* 弹窗目标文件与“应用对象”。保存时按现场 scope 重新解析 iids：
   默认只改当前文件，绝不能因为列表默认全选而连带修改其它文件。 */
let fileoptTargetIid = null;
let fileoptScope = "current";
const FILEOPT_SCOPE_IDS = {
  current: "fileopt-scope-current",
  selected: "fileopt-scope-selected",
  all: "fileopt-scope-all",
};

/* 按当前 scope 解析实际目标 iid 列表。 */
function fileoptScopeIids() {
  if (fileoptScope === "selected") return [...state.selected];
  if (fileoptScope === "all") return state.files.map((f) => f.iid);
  return fileoptTargetIid ? [fileoptTargetIid] : [];
}

function fileoptScopeDescription(iids) {
  if (fileoptScope === "selected") return `已选的 ${iids.length} 个文件`;
  if (fileoptScope === "all") return `全部 ${iids.length} 个文件`;
  const f = state.files.find((x) => x.iid === fileoptTargetIid);
  const name = f ? basename(f.source || f.path) : "未选择";
  return `当前文件《${name}》`;
}

/* 标题与说明实时反映“将应用到”，切换 scope 立即刷新。 */
function updateFileOptScopeUI() {
  const iids = fileoptScopeIids();
  $("fileopt-scope-count").textContent = String(state.selected.size);
  for (const [scope, id] of Object.entries(FILEOPT_SCOPE_IDS)) {
    $(id).checked = fileoptScope === scope;
  }
  const desc = fileoptScopeDescription(iids);
  $("fileopt-title").textContent = `将应用到${desc}`;
  $("fileopt-preview").textContent = iids.length
    ? `将应用到${desc}；留空使用全部页面，范围支持 1-5,8,10-12。`
    : `将应用到${desc}；请先在左侧选择文件。`;
}

window.setFileOptScope = function (scope) {
  fileoptScope = FILEOPT_SCOPE_IDS[scope] ? scope : "current";
  updateFileOptScopeUI();
};

function openFileOpt(iid, focus = "range", defaultScope = "current") {
  const f = state.files.find((x) => x.iid === iid);
  if (!f) return;
  fileoptTargetIid = iid;
  fileoptScope = FILEOPT_SCOPE_IDS[defaultScope] ? defaultScope : "current";
  $("fileopt-range").value = f.range_text || "";
  $("fileopt-copies").value = f.copies || 1;
  updateFileOptScopeUI();
  showModal("modal-fileopt");
  if (focus === "copies") $("fileopt-copies").focus();
  else $("fileopt-range").focus();
}

window.pickFileRange = function (kind) {
  let value = kind;
  if (kind === "odd" || kind === "even") {
    // 预设只按当前作用对象的页数生成，避免把越界页码写进范围文本。
    const targets = fileoptScopeIids()
      .map((iid) => state.files.find((f) => f.iid === iid))
      .filter(Boolean);
    const maxPages = Math.max(1, ...targets.map((f) => f.pages || 1));
    const parity = kind === "odd" ? 1 : 0;
    value = Array.from({ length: maxPages }, (_, i) => i + 1)
      .filter((p) => p % 2 === parity).join(",");
  } else if (kind === "last") {
    value = (state.settings.recent_ranges || [])[0] || state.settings.last_range || "";
  }
  $("fileopt-range").value = value;
};

async function saveFileOptions({ inherit = false, applyAll = false } = {}) {
  const range = $("fileopt-range").value.trim();
  const copies = parseInt($("fileopt-copies").value || "1", 10);
  // 现场解析目标：弹窗打开期间选中集合/文件列表可能已变化。
  const iids = applyAll ? state.files.map((f) => f.iid) : fileoptScopeIids();
  if (!iids.length) {
    const message = "请先选择要设置的文件。";
    $("fileopt-preview").textContent = message;
    window.__pushLog(now(), message, true);
    return;
  }
  const res = await api().set_files_options(iids, range, copies, inherit, false);
  if (!res.ok) {
    $("fileopt-preview").textContent = res.message;
    window.__pushLog(now(), `无法保存文件设置：${errorText(res.message, "请稍后重试。")}`, true);
    return;
  }
  replaceFiles(res.files);
  state.advanced.per_file = !inherit;
  state.settings.per_file = !inherit;
  hideModal("modal-fileopt");
  // 立即按新范围重算：文件行显示的是 compute_plan 回传的 effective_range，
  // 否则保存后仍会看到上一次的“全部/奇数页”文案。
  refreshMetrics();
  window.__pushLog(now(), inherit
    ? `已将 ${iids.length} 个文件恢复为全局范围与份数。`
    : `已更新 ${iids.length} 个文件的范围与份数。`);
  persistSettings();
}

async function applyFileOpt() { await saveFileOptions(); }
window.applyFileOpt = applyFileOpt;

/* ================= 指标、设置与持久化 ================= */
let metricsGeneration = 0;
let settingsTimer = null;

async function refreshMetrics() {
  const generation = ++metricsGeneration;
  const opts = collectOptions();
  updateRangeTextDisplay();
  $("m-files").textContent = state.files.length;
  if (!state.files.length) {
    $("m-pages").textContent = "0";
    $("m-sheets").textContent = "0";
    $("m-flips").textContent = "0";
    updateControlStates();
    return;
  }
  const res = await api().compute_plan(opts);
  if (generation !== metricsGeneration) return;
  const rows = Array.isArray(res.rows) ? res.rows : [];
  const errors = Array.isArray(res.errors) ? res.errors : [];
  // 单个文件范围无效时后端仍回传全部 rows（无效行带 range_error + “范围无效”），
  // 必须先合并再重绘：否则一个无效文件会让整批行停在旧值。
  for (const row of rows) {
    const f = state.files.find((x) => x.iid === row.iid);
    if (!f) continue;
    f.effective_range = row.effective_range;
    f.selected_pages = row.selected_pages;
    f.effective_copies = row.effective_copies;
    f.plan_text = row.plan_text;
    // 有效文件回传 range_error=null，这里同时完成“修正后恢复”的清除动作。
    f.range_error = row.range_error || "";
  }
  if (res.ok) {
    $("m-pages").textContent = res.total_pages;
    $("m-sheets").textContent = res.summary.sheets ?? "—";
    $("m-flips").textContent = res.summary.flip_count ?? "—";
  } else if (errors.length) {
    // 全部文件范围无效：没有可选页，指标归零而不是保留上一批的旧数字。
    $("m-pages").textContent = "0";
    $("m-sheets").textContent = "—";
    $("m-flips").textContent = "—";
  }
  if (rows.length) renderFiles();
  if (errors.length) {
    // 只写一条可读日志说明哪些文件范围无效，不阻断其它行的刷新。
    const detail = errors.map((e) => `《${e.name}》${e.message}`).join("；");
    window.__pushLog(now(), `部分文件范围无效：${detail}`, true);
  } else if (!res.ok) {
    window.__pushLog(now(), `无法生成打印计划：${errorText(res.message, "请稍后重试。")}`, true);
  }
  updateControlStates();
}

function collectOptions() {
  const duplex = document.querySelector('input[name="duplex"]:checked').value;
  const flow = document.querySelector('input[name="flow"]:checked').value;
  const view = document.querySelector('input[name="preview-view"]:checked').value;
  const mode = $("range-mode").value;
  const customMode = mode === "custom" || mode === "recent";
  return {
    printer: $("printer-select").value,
    copies: parseInt($("copies").value || "1", 10),
    manual_duplex: duplex === "manual",
    print_flow: flow,
    range_mode: mode,
    // 非 custom/recent 时输入框显示的是提示文案（“全部/各文件独立”），
    // 真实范围文本必须取保留值，避免把提示文案持久化成用户输入。
    range_text: customMode ? $("range-text").value.trim() : state.rangeText.trim(),
    copy_exec_on: $("copy-exec").checked,
    preview_view: view,
    current_page: parseInt($("preview-page").value || "1", 10),
    ...state.advanced,
  };
}

function applySettingsToUI(settings) {
  state.settings = { ...settings };
  for (const key of Object.keys(state.advanced)) {
    if (key in settings) state.advanced[key] = settings[key];
  }
  state.advanced.delay = settings.delay_seconds ?? state.advanced.delay;
  $("copies").value = settings.copies ?? 1;
  const duplex = settings.manual_duplex ? "manual" : "simplex";
  document.querySelector(`input[name="duplex"][value="${duplex}"]`).checked = true;
  const flow = settings.print_flow || "batch_all";
  document.querySelector(`input[name="flow"][value="${flow}"]`).checked = true;
  $("range-mode").value = settings.range_mode || "all";
  // 保留用户原文（custom/recent 才显示）；其余模式输入框只显示说明文案。
  state.rangeText = settings.range_text || settings.last_range || "1-5,8,10-12";
  $("copy-exec").checked = !!settings.copy_exec_on;
  const view = settings.preview_view || "single";
  document.querySelector(`input[name="preview-view"][value="${view}"]`).checked = true;
  $("preview-scale-mode").value = state.advanced.scale_mode;
  $("preview-scale-percent").value = state.advanced.scale_percent;
  $("preview-nup").value = String(state.advanced.nup);
  $("preview-orientation").value = state.advanced.orientation;
  syncInputStates();
}

function syncInputStates() {
  const mode = $("range-mode").value;
  $("range-text").disabled = !(mode === "custom" || mode === "recent");
  $("preview-scale-percent").disabled = $("preview-scale-mode").value !== "custom";
  updateRangeTextDisplay();
}

function persistSettings() {
  const opts = collectOptions();
  const changes = { ...state.advanced };
  for (const key of ["printer", "copies", "manual_duplex", "print_flow", "range_mode",
                     "range_text", "copy_exec_on", "preview_view"]) changes[key] = opts[key];
  state.settings = { ...state.settings, ...changes };
  clearTimeout(settingsTimer);
  settingsTimer = setTimeout(() => api().save_settings(changes), 150);
}

function onSettingChange() {
  // custom/recent 下输入框才是用户原文；模式切换瞬间它可能还显示上一模式的
  // 说明文案（“全部/当前页”等），这类文案不能写回 state.rangeText。
  const mode = $("range-mode").value;
  if (mode === "custom" || mode === "recent") {
    const raw = $("range-text").value;
    if (!RANGE_DISPLAY_TEXTS.has(raw.trim())) state.rangeText = raw;
  }
  state.advanced.scale_mode = $("preview-scale-mode").value;
  state.advanced.scale_percent = parseFloat($("preview-scale-percent").value || "100");
  state.advanced.nup = parseInt($("preview-nup").value || "1", 10);
  state.advanced.orientation = $("preview-orientation").value;
  syncInputStates();
  persistSettings();
  refreshMetrics();
  if (state.previewFile) renderPreview(parseInt($("preview-page").value || "1", 10));
}

/* ================= 预览 ================= */
let previewGeneration = 0;

function showPreviewMessage(title, subtitle) {
  const stage = $("preview-stage");
  stage.replaceChildren();
  const box = document.createElement("div");
  box.className = "preview-empty";
  const heading = document.createElement("div");
  heading.className = "preview-empty-title";
  heading.textContent = title;
  const detail = document.createElement("div");
  detail.className = "empty-sub";
  detail.textContent = subtitle || "";
  box.append(heading, detail);
  stage.appendChild(box);
}

function clearPreview() {
  previewGeneration += 1;  // 使尚未返回的旧预览请求失效，防止清空后旧图回写。
  state.previewFile = null;
  state.previewTotal = 0;
  $("preview-page").value = 1;
  $("preview-page").disabled = true;
  $("preview-total").textContent = "/ —";
  $("preview-status").textContent = "等待预览";
  $("preview-stage").classList.remove("spread");
  showPreviewMessage("添加文件后显示实际打印效果", "在左侧添加 PDF、Word 或图片");
}

async function selectPreviewFile(f) {
  if (!f) { clearPreview(); return; }
  const generation = ++previewGeneration;
  state.previewFile = f;
  renderFiles();
  $("preview-page").value = 1;
  $("preview-total").textContent = `/ ${f.pages || "—"}`;
  if (!f.pages) {
    $("preview-page").disabled = true;
    const failed = f.status === "失败";
    showPreviewMessage(failed ? "文件无法预览" : "文件正在转换",
                       failed ? (f.error || "请检查文件后重试") : "转换完成后自动显示实际打印效果");
    return;
  }
  await renderPreview(1, generation);
}

async function renderPreview(pageNo, generation = ++previewGeneration) {
  const file = state.previewFile;
  if (!file || !file.pages) return;
  $("preview-page").disabled = true;
  showPreviewMessage("正在加载预览…", basename(file.source || file.path));
  const res = await api().preview_page(file.iid, pageNo, collectOptions());
  if (generation !== previewGeneration || state.previewFile?.iid !== file.iid) return;
  if (res.ok) {
    const spread = document.querySelector('input[name="preview-view"]:checked').value === "spread";
    const pages = [res];
    if (spread && res.page < res.total) {
      const next = await api().preview_page(file.iid, res.page + 1, collectOptions());
      if (next.ok) pages.push(next);
    }
    if (generation !== previewGeneration || state.previewFile?.iid !== file.iid) return;
    const stage = $("preview-stage");
    stage.replaceChildren();
    stage.classList.toggle("spread", spread);
    for (const item of pages) {
      const img = document.createElement("img");
      img.src = item.dataurl;
      img.alt = `第 ${item.page} 面`;
      stage.appendChild(img);
    }
    $("preview-page").disabled = false;
    $("preview-page").value = res.page;
    $("preview-total").textContent = `/ ${res.total}`;
    state.previewTotal = res.total;
    const endPage = pages[pages.length - 1].page;
    $("preview-status").textContent =
      `第 ${res.page}${endPage !== res.page ? `–${endPage}` : ""} 面 / 共 ${res.total} 面 · ` +
      `缩放 ${res.scale_percent || "—"}% · 纸张 ${res.paper}` +
      (res.nup > 1 ? ` · ${res.nup} 合 1` : "");
    $("preview-scale-chip").textContent = `实际 ${res.scale_percent || "—"}%`;
    $("detail-plan").textContent = res.plan || "";
  } else {
    showPreviewMessage("预览失败", res.message || "无法读取该文件");
  }
}

function syncPreviewAfterFilesChanged() {
  const current = state.previewFile &&
    state.files.find((f) => f.iid === state.previewFile.iid);
  if (current) return;
  const fallback = state.files[state.files.length - 1];
  if (fallback) selectPreviewFile(fallback);
  else clearPreview();
}

function changePreviewPage(delta) {
  if (!state.previewFile || !state.previewTotal) return;
  const spread = document.querySelector('input[name="preview-view"]:checked').value === "spread";
  const step = spread ? 2 : 1;
  const current = parseInt($("preview-page").value || "1", 10);
  const target = Math.max(1, Math.min(state.previewTotal, current + delta * step));
  renderPreview(target);
}

async function showPlanDetail() {
  if (!state.previewFile || !state.previewFile.pages) {
    window.__pushLog(now(), "请先选择可预览的文件", true); return;
  }
  const res = await api().plan_description(state.previewFile.iid, collectOptions());
  if (res.ok) {
    $("plan-text").textContent = res.text + "\n\n" + res.sample;
    showModal("modal-plan");
  } else {
    window.__pushLog(now(), `无法读取页序详情：${res.message}`, true);
  }
}

/* ================= 打印机 ================= */
async function loadPrinters() {
  const res = await api().list_printers();
  if (!res.ok) {
    $("printer-dot").className = "dot bad";
    $("printer-name").textContent = "无法获取打印机列表";
    $("printer-name").title = "无法获取打印机列表";
    window.__pushLog(now(), `无法获取打印机列表：${res.message}`, true);
    return;
  }
  const sel = $("printer-select");
  sel.innerHTML = "";
  for (const name of res.printers) {
    const opt = document.createElement("option");
    opt.value = name; opt.textContent = name;
    if (name === (state.settings.printer || res.default)) opt.selected = true;
    sel.appendChild(opt);
  }
  if (!sel.value && res.printers.length) sel.value = res.default || res.printers[0];
  state.settings.printer = sel.value;
  // 长打印机名在 220px 下拉框里省略显示，但 title 始终给全名。
  sel.title = sel.value || "";
  $("printer-name").textContent = res.default || res.printers[0] || "无打印机";
  $("printer-name").title = $("printer-name").textContent;
  updatePrinterStatus(sel.value);
}

async function updatePrinterStatus(name) {
  const dot = $("printer-dot");
  const label = $("printer-name");
  $("printer-select").title = name || "";
  if (!name) {
    dot.className = "dot bad";
    label.textContent = "无可用打印机";
    label.title = label.textContent;
    return;
  }
  dot.className = "dot pending";
  label.textContent = "检测打印机…";
  const res = await api().printer_status(name);
  if (res.ok && res.ready) {
    dot.className = "dot ok";
    label.textContent = name;
  } else {
    dot.className = "dot bad";
    // 后端统一把失败状态描述为“脱机”；详细原因来自 detail，并给出可操作建议。
    label.textContent = `${name}（脱机）`;
    const detail = (res && res.detail) || "请检查电源、USB/网络连接与纸张后重试。";
    // 后端 detail 已包含打印机名时不再重复前缀，日志保持简洁可读。
    const logText = detail.includes(name) ? detail : `打印机「${name}」脱机：${detail}`;
    window.__pushLog(now(), logText, true);
  }
  label.title = label.textContent;
}

async function loadPaperSources() {
  const printer = $("printer-select").value;
  if (!printer) return;
  const res = await api().paper_sources(printer);
  if (res.ok && res.sources.length) {
    state.advanced._paper_sources = res.sources;
    if ($("modal-advanced").style.display !== "none") buildAdvancedForm();
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
  if (!state.printing) return;
  const proceed = await confirmAction({
    title: "取消本次打印？",
    lines: ["已打印的纸张无法收回，临时文件会被清理。", "确定要取消本次整批打印吗？"],
    confirmLabel: "取消打印任务",
    danger: true,
  });
  if (!proceed) return;
  await api().cancel_print();
  window.__pushLog(now(), "已请求取消打印…");
}

/* ================= 高级设置 ================= */
state.advanced = {
  paper_name: "A4", paper_width_mm: 210, paper_height_mm: 297,
  orientation: "auto", scale_mode: "shrink", scale_percent: 100, nup: 1,
  nup_border: false, page_order: "row", booklet: false, auto_rotate: false,
  annots: false, skip_blank: false,
  margin_left_mm: 0, margin_top_mm: 0, margin_right_mm: 0, margin_bottom_mm: 0,
  crop_left_mm: 0, crop_top_mm: 0, crop_right_mm: 0, crop_bottom_mm: 0,
  color_mode: "color", print_quality: "normal", dpi: 300, ink_saver: false,
  wait_mode: "queue", delay: 15, source_name: "自动选择", collate: "collate",
  reverse: false, per_file: false, copy_exec_on: false,
  keep_output: false, auto_open: false, output_dir: "", _paper_sources: [["自动选择", 7]],
};
let advancedDraft = null;

function buildAdvancedForm() {
  const schemes = state.schemes;
  const a = advancedDraft || state.advanced;
  const body = $("advanced-body");
  const opt = (v, cur, l) => `<option value="${v}" ${v === cur ? "selected" : ""}>${l}</option>`;
  const sel = (id, list, cur) =>
    `<select id="adv-${id}" class="input">${list.map((x) => opt(x.value, cur, x.label)).join("")}</select>`;
  const num = (id, val, min = 0, max = 9999) =>
    `<input type="number" id="adv-${id}" class="input num" value="${val}" min="${min}" max="${max}">`;
  const chk = (id, label, cur) =>
    `<label class="checkbox-row"><input type="checkbox" id="adv-${id}" ${cur ? "checked" : ""}>${label}</label>`;

  body.innerHTML = `
  <div class="adv-section" id="adv-paper-section"><div class="adv-section-title">纸张与方向</div>
    <div class="adv-grid">
      <label class="form-row"><span>纸张</span>${sel("paper_name", schemes.paper_names.map((p) => ({value: p, label: p})), a.paper_name)}</label>
      <label class="form-row"><span>自定义宽（mm）</span>${num("paper_width_mm", a.paper_width_mm, 10, 2000)}</label>
      <label class="form-row"><span>自定义高（mm）</span>${num("paper_height_mm", a.paper_height_mm, 10, 2000)}</label>
      <label class="form-row"><span>纸盒</span><select id="adv-source_name" class="input">${(a._paper_sources || []).map((x) => opt(x[0], a.source_name, `${x[0]}（${x[1]}）`)).join("")}</select></label>
      <label class="form-row" title="自动：纸张方向跟随所选范围第一页；竖版/横版固定纸张方向。单页保持内容原方向缩放居中，拼接/小册子才按需旋转 90° 适配槽位"><span>纸张方向</span>${sel("orientation", schemes.orientations, a.orientation)}</label>
      <label class="form-row"><span>缩放</span>${sel("scale_mode", schemes.scale_modes, a.scale_mode)}</label>
      <label class="form-row"><span>缩放比例 %</span>${num("scale_percent", a.scale_percent, 10, 500)}</label>
      ${chk("auto_rotate", "方向不符时旋转 90° 适配槽位（默认关闭）", a.auto_rotate)}
      <button class="btn ghost sm" type="button" onclick="focusPageSetup()">页面设置…</button>
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
      <label class="form-row"><span>分页方式</span><select id="adv-collate" class="input"><option value="collate" ${a.collate === "collate" ? "selected" : ""}>自动分页（整份连续）</option><option value="uncollated" ${a.collate !== "collate" ? "selected" : ""}>逐页重复</option></select></label>
      ${chk("per_file", "逐文件单设范围/份数", a.per_file)}
      ${chk("reverse", "逆序打印", a.reverse)}
      ${chk("copy_exec_on", "按份逐份执行", a.copy_exec_on)}
      ${chk("keep_output", "保留排版输出 PDF", a.keep_output)}
      ${chk("auto_open", "完成后打开输出目录", a.auto_open)}
      <label class="form-row"><span>输出目录</span><span class="input-with-button"><input id="adv-output_dir" class="input" value="${a.output_dir}"><button class="btn ghost sm" type="button" onclick="browseOutputDir()">浏览…</button></span></label>
    </div>
  </div>`;
}

function readAdvancedForm() {
  const a = advancedDraft;
  const ids = ["paper_name", "orientation", "scale_mode", "nup", "page_order",
               "color_mode", "print_quality", "wait_mode", "source_name", "output_dir", "collate"];
  for (const k of ids) {
    const el = $(`adv-${k}`);
    if (el) a[k] = k === "nup" ? parseInt(el.value, 10) : el.value;
  }
  const nums = ["scale_percent", "dpi", "delay", "paper_width_mm", "paper_height_mm",
                "margin_left_mm", "margin_top_mm", "margin_right_mm", "margin_bottom_mm",
                "crop_left_mm", "crop_top_mm", "crop_right_mm", "crop_bottom_mm"];
  for (const k of nums) {
    const el = $(`adv-${k}`);
    if (el) a[k] = parseFloat(el.value);
  }
  const chks = ["nup_border", "booklet", "auto_rotate", "ink_saver", "annots", "per_file",
                "skip_blank", "reverse", "copy_exec_on", "keep_output", "auto_open"];
  for (const k of chks) {
    const el = $(`adv-${k}`);
    if (el) a[k] = el.checked;
  }
}

window.openAdvanced = function () {
  advancedDraft = { ...state.advanced };
  buildAdvancedForm();
  showModal("modal-advanced");
};
window.applyAdvanced = function () {
  readAdvancedForm();
  state.advanced = { ...advancedDraft };
  advancedDraft = null;
  $("copy-exec").checked = !!state.advanced.copy_exec_on;
  $("preview-scale-mode").value = state.advanced.scale_mode;
  $("preview-scale-percent").value = state.advanced.scale_percent;
  $("preview-nup").value = String(state.advanced.nup);
  $("preview-orientation").value = state.advanced.orientation;
  $("advanced-summary").textContent = "高级设置：自定义";
  hideModal("modal-advanced");
  onSettingChange();
};
// 仅重置高级设置草稿；已保存的打印设置不动，需再点“应用设置”才生效。
window.resetAdvanced = function (_acrobat = false) {
  advancedDraft = {
    paper_name: "A4", paper_width_mm: 210, paper_height_mm: 297,
    orientation: "auto", scale_mode: "shrink", scale_percent: 100, nup: 1,
    nup_border: false, page_order: "row", booklet: false, auto_rotate: !!_acrobat,
    annots: false, skip_blank: false, margin_left_mm: 0, margin_top_mm: 0,
    margin_right_mm: 0, margin_bottom_mm: 0, crop_left_mm: 0, crop_top_mm: 0,
    crop_right_mm: 0, crop_bottom_mm: 0, color_mode: "color", print_quality: "normal",
    dpi: 300, ink_saver: false, wait_mode: "queue", delay: 15,
    source_name: "自动选择", collate: "collate", reverse: false, per_file: false,
    copy_exec_on: false, keep_output: false, auto_open: false, output_dir: "",
    _paper_sources: state.advanced._paper_sources,
  };
  buildAdvancedForm();
};
// 恢复 Acrobat 等效默认：必须走后端 restore_defaults(acrobat=True)，
// 把单面/范围/份数/流程/高级参数整份回写到 UI，而不是只重置高级草稿。
window.restoreAcrobatDefaults = async function () {
  let res = null;
  try {
    res = await api().restore_defaults(true);
  } catch (err) {
    // 桥接层直接抛错（窗口未就绪等）：记日志并保持弹窗，绝不静默留在旧状态。
    window.__pushLog(now(), `无法恢复默认设置：${errorText(err, "请稍后重试。")}`, true);
    return;
  }
  if (!res || !res.ok) {
    window.__pushLog(now(), `无法恢复默认设置：${res && res.message ? res.message : "请稍后重试。"}`, true);
    return;  // 保持高级弹窗打开，便于用户重试或改用“恢复默认”。
  }
  applySettingsToUI(res.settings);
  advancedDraft = null;
  // 需求明确要求 Acrobat 基线为单面；再显式落一次，避免上游字段缺失时残留手动双面。
  document.querySelector('input[name="duplex"][value="simplex"]').checked = true;
  $("advanced-summary").textContent = "高级设置：Acrobat 默认";
  hideModal("modal-advanced");
  // 打印机探测失败不应阻断指标与预览刷新。
  try {
    await updatePrinterStatus($("printer-select").value);
  } catch (err) {
    window.__pushLog(now(), `无法刷新打印机状态：${errorText(err, "请检查打印机连接后重试。")}`, true);
  }
  await refreshMetrics();
  if (state.previewFile) {
    await renderPreview(Math.max(1, Math.min(state.previewTotal || 1,
      parseInt($("preview-page").value || "1", 10))));
  }
};
window.browseOutputDir = async function () {
  const res = await api().choose_output_directory();
  if (res.ok && res.path && advancedDraft) {
    advancedDraft.output_dir = res.path;
    $("adv-output_dir").value = res.path;
  } else if (!res.ok) {
    window.__pushLog(now(), res.message, true);
  }
};
window.focusPageSetup = function () {
  $("adv-paper-section").scrollIntoView({ behavior: "smooth", block: "start" });
};

/* ================= 日志抽屉（mousedown/mousemove/mouseup，document 级监听） ================= */
function initSash() {
  const sash = $("sash");
  let dragging = false;

  // 关键修复：pointer capture 在 WebView2 上偶发失效，改为 document 级
  // mousemove/mouseup，且 mousedown 即记录状态，不依赖事件冒泡顺序。
  sash.addEventListener("mousedown", (e) => {
    dragging = true;
    document.body.style.cursor = "sb_v_double_arrow";
    e.preventDefault();
  });
  document.addEventListener("mousemove", (e) => {
    if (!dragging) return;
    const h = window.innerHeight - e.clientY;
    const clamped = Math.max(96, Math.min(h, Math.floor(window.innerHeight / 2)));
    document.documentElement.style.setProperty("--log-h", clamped + "px");
  });
  document.addEventListener("mouseup", () => {
    if (!dragging) return;
    dragging = false;
    document.body.style.cursor = "";
  });
  // 触摸设备兜底（pointer 事件仅用于触摸，避免与 mouse 双触发）
  sash.addEventListener("touchstart", (e) => {
    dragging = true;
    e.preventDefault();
  }, { passive: false });
  document.addEventListener("touchmove", (e) => {
    if (!dragging) return;
    const t = e.touches[0];
    const h = window.innerHeight - t.clientY;
    const clamped = Math.max(96, Math.min(h, Math.floor(window.innerHeight / 2)));
    document.documentElement.style.setProperty("--log-h", clamped + "px");
  }, { passive: true });
  document.addEventListener("touchend", () => { dragging = false; });
}

async function exportLog() {
  if (!$("log-body").innerText.trim()) { window.__pushLog(now(), "暂无可导出的日志", true); return; }
  const res = await api().export_log($("log-body").innerText);
  if (res.ok && res.path) window.__pushLog(now(), `日志已导出：${res.path}`);
  else if (!res.ok) window.__pushLog(now(), res.message, true);
}

async function toggleLogDetail() {
  const detail = $("log-detail");
  detail.hidden = !detail.hidden;
  $("btn-toggle-detail").textContent = detail.hidden ? "页序/缩略图 ▾" : "页序/缩略图 ▴";
  if (detail.hidden || !state.previewFile) return;
  const strip = $("thumb-strip");
  strip.replaceChildren();
  const count = Math.min(3, state.previewTotal || 0);
  for (let page = 1; page <= count; page += 1) {
    const res = await api().preview_page(state.previewFile.iid, page, collectOptions());
    if (!res.ok) continue;
    const button = document.createElement("button");
    button.className = "thumb-button";
    button.title = `转到第 ${page} 面`;
    button.onclick = () => renderPreview(page);
    const img = document.createElement("img");
    img.src = res.dataurl;
    button.appendChild(img);
    strip.appendChild(button);
  }
}

/* ================= 初始化 ================= */
function onBridgeReady(fn) {
  if (window.pywebview && window.pywebview.api) { fn(); return; }
  const t = setInterval(() => {
    if (window.pywebview && window.pywebview.api) { clearInterval(t); fn(); }
  }, 100);
}

function requestNativeFileDrop(dataTransfer) {
  if (!window.chrome || !window.chrome.webview) return false;
  const files = dataTransfer && dataTransfer.files;
  if (!files || !window.chrome.webview.postMessageWithAdditionalObjects) return false;
  // 必须把 File 对象作为 additionalObjects 显式传给原生层；单独 postMessage
  // 只有一段文本，CoreWebView2File.Path 永远不会到达 Python。
  window.chrome.webview.postMessageWithAdditionalObjects("FilesDropped", files);
  return true;
}

onBridgeReady(async () => {
  await api().cleanup_stale();
  const st = await api().get_settings();
  if (st.ok) {
    state.schemes = st.schemes;
    applySettingsToUI(st.settings);
  }
  await loadPrinters();
  await refreshRecent();
  await loadPaperSources();
  initSash();
  renderFiles();

  $("btn-add").onclick = addFiles;
  $("btn-empty-add").onclick = addFiles;
  $("btn-add-dir").onclick = addFolder;
  $("btn-remove").onclick = async () => {
    const iids = [...state.selected];
    if (!iids.length) return;
    if (iids.length > 1) {
      const proceed = await confirmAction({
        title: "从列表移除文件",
        lines: [`将移除选中的 ${iids.length} 个文件。`, "原文件不会被删除，可随时重新添加。"],
        confirmLabel: "移除",
      });
      if (!proceed) return;
    }
    await api().remove_files(iids);
    state.files = state.files.filter((f) => !state.selected.has(f.iid));
    state.selected.clear();
    renderFiles();
    syncPreviewAfterFilesChanged();
    refreshMetrics();
  };
  $("btn-clear").onclick = async () => {
    if (!state.files.length) return;
    const proceed = await confirmAction({
      title: "清空待打印列表",
      lines: ["将移除列表中的全部文件，原文件不会被删除。", "清空后仍可在“最近”里重新添加。"],
      confirmLabel: "清空列表",
      danger: true,
    });
    if (!proceed) return;
    await api().clear_files();
    state.files = [];
    state.selected.clear();
    clearPreview();
    renderFiles();
    refreshMetrics();
    window.__pushLog(now(), "待打印列表已清空。");
  };
  $("btn-add-recent").onclick = async () => {
    const path = $("recent-select").value;
    if (path) await doAdd([path]);
  };
  $("btn-clear-recent").onclick = async () => { await api().clear_recent(); };

  const closeMore = () => { $("more-menu").open = false; };
  /* “更多”菜单统一作用于当前预览文件（回退已选首个）：默认全选时也只动这一个，
     两个方向不会再互相阻挡，资源管理器也不会永远定位到第一个选中文件。 */
  for (const [id, direction] of [["btn-move-up", -1], ["btn-move-down", 1]]) {
    $(id).onclick = async () => {
      const iid = currentFileIid();
      const rec = currentFileRecord();
      const name = basename((rec && (rec.source || rec.path)) || "");
      if (!iid) {
        window.__pushLog(now(), "请先选择要移动的文件。", true);
        closeMore();
        return;
      }
      const before = state.files.findIndex((f) => f.iid === iid);
      let res = null;
      try {
        res = await api().move_files([iid], direction);
      } catch (err) {
        window.__pushLog(now(), `无法移动《${name}》：${errorText(err, "请稍后重试。")}`, true);
        closeMore();
        return;
      }
      if (!res || !res.ok) {
        window.__pushLog(now(),
          `无法移动《${name}》：${errorText(res && res.message, "请稍后重试。")}`, true);
      } else {
        replaceFiles(res.files);
        const after = state.files.findIndex((f) => f.iid === iid);
        if (before === after) {
          window.__pushLog(now(), direction < 0
            ? `《${name}》已在最前，无法上移。` : `《${name}》已在最后，无法下移。`);
        } else {
          window.__pushLog(now(), `已${direction < 0 ? "上移" : "下移"}《${name}》。`);
        }
      }
      closeMore();
    };
  }
  $("btn-dedupe").onclick = async () => {
    let res = null;
    try {
      res = await api().dedupe_files();
    } catch (err) {
      window.__pushLog(now(), `去重失败：${errorText(err, "请稍后重试。")}`, true);
      closeMore();
      return;
    }
    if (!res || !res.ok) {
      window.__pushLog(now(), `去重失败：${errorText(res && res.message, "请稍后重试。")}`, true);
    } else {
      const removed = Array.isArray(res.removed) ? res.removed.length : 0;
      if (!removed) {
        window.__pushLog(now(), "未发现重复文件");
      } else {
        replaceFiles(res.files);
        window.__pushLog(now(), `已移除 ${removed} 个重复文件`);
      }
    }
    closeMore();
  };
  $("btn-show-explorer").onclick = async () => {
    const iid = currentFileIid();
    const rec = currentFileRecord();
    const name = basename((rec && (rec.source || rec.path)) || "");
    if (!iid) {
      window.__pushLog(now(), "请先选择要定位的文件。", true);
      closeMore();
      return;
    }
    let res = null;
    try {
      res = await api().show_in_explorer(iid);
    } catch (err) {
      window.__pushLog(now(),
        `无法在资源管理器中定位《${name}》：${errorText(err, "请稍后重试。")}`, true);
      closeMore();
      return;
    }
    if (!res || !res.ok) {
      window.__pushLog(now(),
        `无法在资源管理器中定位《${name}》：${errorText(res && res.message, "请稍后重试。")}`, true);
    } else {
      window.__pushLog(now(), `已在资源管理器中定位《${name}》。`);
    }
    closeMore();
  };
  // 选中操作条同样以当前预览文件为目标，默认作用对象固定为“当前文件”。
  $("btn-file-range").onclick = () => openFileOpt(currentFileIid(), "range", "current");
  $("btn-file-copies").onclick = () => openFileOpt(currentFileIid(), "copies", "current");
  $("btn-file-restore").onclick = async () => {
    const iid = currentFileIid();
    if (!iid) { window.__pushLog(now(), "请先选择要恢复的文件。", true); return; }
    let res = null;
    try {
      res = await api().set_files_options([iid], "", 1, true, false);
    } catch (err) {
      window.__pushLog(now(), `无法恢复全局范围：${errorText(err, "请稍后重试。")}`, true);
      return;
    }
    if (!res || !res.ok) {
      window.__pushLog(now(),
        `无法恢复全局范围：${errorText(res && res.message, "请稍后重试。")}`, true);
      return;
    }
    state.advanced.per_file = false;
    replaceFiles(res.files);
    refreshMetrics();
    persistSettings();
    window.__pushLog(now(), "已恢复该文件的全局范围与份数。");
  };
  $("btn-file-apply-all").onclick = () => openFileOpt(currentFileIid(), "range", "all");
  $("btn-fileopt-restore").onclick = () => saveFileOptions({ inherit: true });
  $("btn-fileopt-all").onclick = () => saveFileOptions({ applyAll: true });
  $("btn-last-range").onclick = () => window.pickFileRange("last");
  for (const radio of document.querySelectorAll('input[name="fileopt-scope"]')) {
    radio.onchange = () => window.setFileOptScope(radio.value);
  }

  $("btn-refresh-printers").onclick = async () => {
    await loadPrinters();
    await api().refresh_files();
    await loadPaperSources();
  };
  $("printer-select").onchange = async (e) => {
    state.settings.printer = e.target.value;
    updatePrinterStatus(e.target.value);
    await loadPaperSources();
    onSettingChange();
  };
  $("btn-print").onclick = startPrint;
  $("btn-pause").onclick = pauseToggle;
  $("btn-cancel").onclick = cancelPrint;
  $("btn-export-log").onclick = exportLog;
  $("btn-clear-log").onclick = () => { $("log-body").innerHTML = ""; $("log-summary").textContent = "暂无日志"; };
  $("btn-open-dir").onclick = () => api().open_log_dir();
  $("btn-toggle-detail").onclick = toggleLogDetail;
  $("btn-advanced").onclick = window.openAdvanced;
  $("btn-restore-defaults").onclick = async () => {
    const proceed = await confirmAction({
      title: "恢复默认设置",
      lines: ["将把全部打印设置恢复为应用默认值。", "最近文件记录会保留。"],
      confirmLabel: "恢复默认",
    });
    if (!proceed) return;
    const res = await api().restore_defaults(false);
    if (res.ok) {
      applySettingsToUI(res.settings);
      $("printer-select").value = res.settings.printer || $("printer-select").value;
      $("advanced-summary").textContent = "高级设置：全部为默认值";
      updatePrinterStatus($("printer-select").value);
      refreshMetrics();
      if (state.previewFile) renderPreview(1);
    }
  };
  $("btn-plan-detail").onclick = showPlanDetail;
  // 实时记录用户输入的指定范围原文；输入框在非 custom/recent 模式是 disabled，
  // 因此不会用“全部/各文件独立”之类的提示文案覆盖原文。
  $("range-text").addEventListener("input", (e) => { state.rangeText = e.target.value; });
  for (const id of ["copies", "range-mode", "range-text", "copy-exec",
                    "preview-scale-mode", "preview-scale-percent", "preview-nup", "preview-orientation"]) {
    $(id).onchange = onSettingChange;
  }
  document.querySelectorAll('input[name="duplex"],input[name="flow"],input[name="preview-view"]')
    .forEach((r) => { r.onchange = onSettingChange; });
  $("btn-refresh-preview").onclick = () => {
    if (state.previewFile) renderPreview(parseInt($("preview-page").value || "1", 10));
  };
  $("btn-preview-prev").onclick = () => changePreviewPage(-1);
  $("btn-preview-next").onclick = () => changePreviewPage(1);
  $("preview-page").onchange = (e) => {
    const target = Math.max(1, Math.min(state.previewTotal || 1,
      parseInt(e.target.value || "1", 10)));
    renderPreview(target);
  };

  // 拖放：用 postMessageWithAdditionalObjects 把 File 对象交给 WebView2，
  // Python 从 CoreWebView2File.Path 取得真实路径，再经原生通道加入列表。
  // 高亮只允许一层：空状态可见时高亮 #empty-state（实线蓝框+浅蓝底），
  // 已有文件时高亮 #file-list 内侧（inset 描边），不再额外画 outline 虚线。
  const dropZone = $("file-list");
  const dropEmpty = $("empty-state");
  const showDropHint = (on) => {
    const empty = state.files.length === 0;
    dropEmpty.classList.toggle("drag-over", on && empty);
    dropZone.classList.toggle("drag-over", on && !empty);
  };
  document.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  });
  document.addEventListener("dragenter", (e) => {
    if (e.dataTransfer.types.includes("Files")) showDropHint(true);
  });
  document.addEventListener("dragleave", (e) => {
    if (e.relatedTarget === null) showDropHint(false);
  });
  document.addEventListener("dragend", () => showDropHint(false));
  document.addEventListener("drop", (e) => {
    e.preventDefault();
    showDropHint(false);
    const notified = requestNativeFileDrop(e.dataTransfer);
    window.__pushLog(now(), notified ? "收到拖放文件，正在读取本地路径…"
                                     : "当前环境不支持原生拖放", !notified);
  });

  // 快捷键：Ctrl+O 添加、Delete 移除、F5 刷新、Ctrl+Enter/Ctrl+P 打印。
  document.addEventListener("keydown", (e) => {
    const editing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement?.tagName || "");
    if (e.ctrlKey && (e.key === "o" || e.key === "O")) { e.preventDefault(); addFiles(); }
    else if (e.ctrlKey && e.key === "Enter") { e.preventDefault(); startPrint(); }
    else if (e.ctrlKey && e.key === "p") { e.preventDefault(); startPrint(); }
    else if (e.key === "F5") {
      e.preventDefault();
      loadPrinters();
      api().refresh_files();
      loadPaperSources();
    } else if (e.key === "Escape") {
      if ($("modal-confirm").style.display !== "none") answerConfirm(false);
      else if ($("modal-flip").style.display !== "none") flipAnswer(false);
      else if ($("modal-phase").style.display !== "none") phaseAnswer(false);
      else {
        const openModals = [...document.querySelectorAll(".modal-mask")].filter((m) => m.style.display !== "none");
        if (openModals.length) openModals.forEach((m) => { m.style.display = "none"; });
        else { state.selected.clear(); renderFiles(); }
      }
    } else if (e.ctrlKey && e.key === "a" && !editing) {
      e.preventDefault();
      state.selected = new Set(state.files.map((f) => f.iid));
      renderFiles();
    } else if (e.key === "Delete" && state.selected.size && !editing) {
      $("btn-remove").click();
    }
  });

  window.__pushLog(now(), "就绪，可以添加文件。");
});
