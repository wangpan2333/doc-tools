"use strict";

/* ---------------- 工具配置（加功能 = 往这里加一项） ----------------
   cat      : 分类 id（侧栏单层导航）
   catName  : 分类显示名
   catColor : 分类代表色（侧栏小图标）
   items[]  : 该分类下的卡片
     id      工具 id（对应 converters.convert 的分支）
     name    卡片标题
     title   处理页大标题
     desc    描述
     engine  引擎文案（null 表示未实现占位）
     exts    支持的后缀（拖拽区提示 + 文件对话框类型）
     badge   彩色块中央的类型字符
     color   彩色块背景色
     enabled 是否已实现（false = 占位，进页后提示开发中、按钮禁用）
------------------------------------------------------------------- */
const TOOLS = [
  {
    cat: "pdf", catName: "PDF 转换", catColor: "#ef4444",
    items: [
      { id: "pdf2word", name: "PDF 转 Word", title: "PDF 转 Word",
        desc: "把 PDF 还原成可编辑的 Word 文档（扫描件自动 OCR）", engine: "exactdoc + LibreOffice 校验",
        exts: ".pdf", badge: "W", color: "#2563eb", enabled: true, needs: "exactdoc" },
      { id: "pdf2excel", name: "PDF 转 Excel", title: "PDF 转 Excel",
        desc: "把表格型 PDF 提取为 Excel 表格", engine: "—",
        exts: ".pdf", badge: "X", color: "#16a34a", enabled: false },
      { id: "pdf2ppt", name: "PDF 转 PPT", title: "PDF 转 PPT",
        desc: "把 PDF 每页转成一张幻灯片", exts: ".pdf",
        badge: "P", color: "#ea580c", enabled: false },
      { id: "pdf2img", name: "PDF 转图片", title: "PDF 转图片",
        desc: "把 PDF 每页渲染成高清图片", exts: ".pdf",
        badge: "IMG", color: "#0d9488", enabled: false },
      { id: "pdfmerge", name: "PDF 合并", title: "PDF 合并",
        desc: "把多个 PDF 合并成一个文件", exts: ".pdf",
        badge: "⊞", color: "#7c3aed", enabled: false },
      { id: "pdfcompress", name: "PDF 压缩", title: "PDF 压缩",
        desc: "减小 PDF 体积，便于分享发送", exts: ".pdf",
        badge: "⤓", color: "#dc2626", enabled: false },
      { id: "pdfsplit", name: "PDF 拆分", title: "PDF 拆分",
        desc: "按页码把 PDF 拆成多个文件", exts: ".pdf",
        badge: "✂", color: "#d97706", enabled: false },
    ],
  },
  {
    cat: "word", catName: "Word 转换", catColor: "#2563eb",
    items: [
      { id: "word2pdf", name: "Word 转 PDF", title: "Word 转 PDF",
        desc: "把 Word 文档批量转成 PDF", engine: "LibreOffice 无头",
        exts: ".docx / .doc", badge: "PDF", color: "#ef4444", enabled: true, needs: "libreoffice" },
      { id: "word2img", name: "Word 转图片", title: "Word 转图片",
        desc: "把 Word 每页渲染成图片", exts: ".docx / .doc",
        badge: "IMG", color: "#0d9488", enabled: false },
    ],
  },
  {
    cat: "image", catName: "图片转换", catColor: "#0d9488",
    items: [
      { id: "img2pdf", name: "图片转 PDF", title: "图片转 PDF",
        desc: "把多张图片合成为一个 PDF", exts: ".png / .jpg",
        badge: "PDF", color: "#ef4444", enabled: false },
      { id: "img2word", name: "图片转 Word", title: "图片转 Word",
        desc: "OCR 识别图片文字并生成 Word", exts: ".png / .jpg",
        badge: "W", color: "#2563eb", enabled: false },
    ],
  },
];

const STATUS_LABEL = { done: "完成", skipped: "已跳过", error: "失败", running: "转换中" };

/* ---------------- 状态 ---------------- */
let activeCat = "pdf";
let current = null;
let files = [];
let engine = { libreoffice: null, exactdoc: false, ocrmypdf: false, tesseract: false };
let running = false;

/* ---------------- DOM ---------------- */
const $ = (id) => document.getElementById(id);
const navEl = $("nav");
const cardsEl = $("cards");
const catTitle = $("catTitle");
const catSub = $("catSub");
const homeView = $("homeView");
const toolView = $("toolView");
const sidebarEl = $("sidebar");
const toolBadge = $("toolBadge");
const toolName = $("toolName");
const toolDesc = $("toolDesc");
const engineBadge = $("engineBadge");
const infoBanner = $("infoBanner");
const warnBanner = $("warnBanner");
const dropzone = $("dropzone");
const dzSub = $("dzSub");
const chipsEl = $("chips");
const outDirEl = $("outDir");
const conflictEl = $("conflict");
const convertBtn = $("convertBtn");
const tasksEl = $("tasks");
const tasksEmpty = $("tasksEmpty");
const tasksCount = $("tasksCount");
const toastEl = $("toast");

/* ---------------- 图标 ---------------- */
function badgeHTML(tool, large) {
  return (
    `<div class="tool-badge${large ? " lg" : ""}" style="background:${tool.color}">` +
    `<svg class="badge-doc"><use href="#i-doc"/></svg>` +
    `<span class="badge-tag">${tool.badge}</span>` +
    `</div>`
  );
}

/* ---------------- 渲染：侧栏分类 ---------------- */
function renderNav() {
  navEl.innerHTML = "";
  TOOLS.forEach((g) => {
    const item = document.createElement("div");
    item.className = "nav-item" + (g.cat === activeCat ? " active" : "");
    item.innerHTML =
      `<span class="nav-cat-dot" style="background:${g.catColor}">` +
      `<svg class="ic"><use href="#i-doc"/></svg></span>` +
      `<span>${g.catName}</span>`;
    item.addEventListener("click", () => setCategory(g.cat));
    navEl.appendChild(item);
  });
}

/* ---------------- 渲染：首页卡片 ---------------- */
function renderCards() {
  const group = TOOLS.find((g) => g.cat === activeCat);
  catTitle.textContent = group.catName;
  catSub.textContent = "选择下面的工具开始处理";
  cardsEl.innerHTML = "";
  group.items.forEach((it) => {
    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML =
      badgeHTML(it, false) +
      `<div class="card-body">` +
      `<div class="card-name">${it.name}</div>` +
      `<div class="card-desc">${it.desc}</div>` +
      `</div>`;
    card.addEventListener("click", () => openTool(it));
    cardsEl.appendChild(card);
  });
}

function setCategory(cat) {
  activeCat = cat;
  renderNav();
  renderCards();
}

function showHome() {
  toolView.hidden = true;
  homeView.hidden = false;
  sidebarEl.style.display = "";
  renderCards();
}

/* ---------------- 进入处理页 ---------------- */
function openTool(tool) {
  current = tool;
  flog("openTool", tool.id, tool.name);
  sidebarEl.style.display = "none"; // 处理页占满整个软件窗口
  homeView.hidden = true;
  toolView.hidden = false;

  toolBadge.innerHTML = badgeHTML(tool, true);
  toolName.textContent = tool.title;
  toolDesc.textContent = tool.desc;
  dzSub.textContent = "支持 " + (tool.exts || "—");

  // 重置文件与任务
  files = [];
  renderChips();
  tasksEl.innerHTML = "";
  tasksEmpty.style.display = "";
  tasksCount.textContent = "0 / 0 完成";
  convertBtn.disabled = false;

  if (!tool.enabled) {
    infoBanner.hidden = false;
    infoBanner.textContent = "该功能正在开发中，敬请期待。当前可使用的工具：PDF 转 Word、Word 转 PDF。";
    convertBtn.disabled = true;
    warnBanner.hidden = true;
    engineBadge.textContent = "";
  } else {
    infoBanner.hidden = true;
    engineBadge.textContent = tool.engine ? "引擎 · " + tool.engine : "";
    updateEngineWarning();
  }
}

/* ---------------- 引擎警告 ---------------- */
function updateEngineWarning() {
  const need = current && current.needs;
  if (!need) { warnBanner.hidden = true; return; }
  const missing =
    (need === "libreoffice" && !engine.libreoffice) ||
    (need === "exactdoc" && !engine.exactdoc);
  if (missing) {
    warnBanner.hidden = false;
    warnBanner.textContent =
      need === "libreoffice"
        ? "未检测到 LibreOffice。Word → PDF 需要它：请安装并把 soffice 加入 PATH（Windows 可用 winget install --id TheDocumentFoundation.LibreOffice）。"
        : "未检测到 exactdoc，PDF → Word 不可用，请先 pip install --pre exactdoc（详见 README）。";
  } else {
    warnBanner.hidden = true;
  }
}

/* ---------------- 文件选择 ---------------- */
function renderChips() {
  chipsEl.innerHTML = "";
  files.forEach((f, idx) => {
    const chip = document.createElement("div");
    chip.className = "chip";
    const name = f.split(/[\\/]/).pop();
    chip.innerHTML = `<span>${name}</span><span class="x" title="移除">×</span>`;
    chip.querySelector(".x").addEventListener("click", (e) => {
      e.stopPropagation();
      files.splice(idx, 1);
      renderChips();
    });
    chipsEl.appendChild(chip);
  });
}

async function pickFiles() {
  if (!current || !current.enabled) { toast("该功能开发中"); return; }
  if (!window.pywebview || !window.pywebview.api) { toast("请在桌面程序中运行"); return; }
  const picked = await window.pywebview.api.select_files(current.id);
  flog("pickFiles", current.id, "->", picked ? picked.length : 0, "个");
  if (picked && picked.length) {
    const set = new Set(files);
    picked.forEach((p) => set.add(p));
    files = [...set];
    renderChips();
  }
}

async function pickDir() {
  if (!window.pywebview || !window.pywebview.api) { toast("请在桌面程序中运行"); return; }
  const dir = await window.pywebview.api.select_dir();
  if (dir) outDirEl.value = dir;
}

/* ---------------- 转换 ---------------- */
function clearTasks() {
  tasksEl.innerHTML = "";
  tasksEmpty.style.display = "none";
}

function addTaskRow(name, i) {
  const row = document.createElement("div");
  row.className = "task";
  row.id = "task-" + i;
  row.innerHTML =
    `<div class="spinner"></div>` +
    `<div class="task-name">${escapeHtml(name)}</div>` +
    `<div class="task-out"></div>` +
    `<div class="task-status status-running">${STATUS_LABEL.running}</div>`;
  tasksEl.appendChild(row);
  return row;
}

// 由 Python 端 evaluate_js 调用
window.onTask = function (d) {
  let row = $("task-" + d.i);
  if (!row) row = addTaskRow(d.name, d.i);
  row.querySelector(".spinner")?.remove();
  const status = row.querySelector(".task-status");
  status.className = "task-status status-" + d.status;
  status.textContent = STATUS_LABEL[d.status] || d.status;
  const out = row.querySelector(".task-out");
  if (d.out) out.textContent = d.out.split(/[\\/]/).pop();
  updateCount();
};

window.onAllDone = function (d) {
  running = false;
  convertBtn.disabled = false;
  updateCount();
  toast(d.ok === d.total
    ? `转换完成：${d.ok} / ${d.total} 个文件成功`
    : `转换结束：${d.ok} / ${d.total} 个成功，其余失败`);
};

function updateCount() {
  const rows = tasksEl.querySelectorAll(".task");
  const done = tasksEl.querySelectorAll(".status-done, .status-skipped, .status-error").length;
  tasksCount.textContent = `${done} / ${rows.length} 完成`;
}

async function onConvert() {
  if (!current || !current.enabled) { toast("该功能开发中"); return; }
  if (!window.pywebview || !window.pywebview.api) { toast("请在桌面程序中运行"); return; }
  if (running) return;
  if (!files.length) { toast("请先选择文件"); return; }

  running = true;
  convertBtn.disabled = true;
  clearTasks();
  files.forEach((f, i) => addTaskRow(f.split(/[\\/]/).pop(), i + 1));
  updateCount();

  const payload = {
    files: files.slice(),
    out_dir: outDirEl.value || "./output",
    tool_id: current.id,
    conflict: conflictEl.value,
  };
  flog("onConvert", current.id, "files=", files.length, "out=", payload.out_dir, "conflict=", payload.conflict);
  try {
    await window.pywebview.api.convert(payload);
  } catch (e) {
    running = false;
    convertBtn.disabled = false;
    toast("转换出错：" + e);
  }
}

/* ---------------- 工具函数 ---------------- */
let toastTimer = null;
function toast(msg) {
  toastEl.textContent = msg;
  toastEl.hidden = false;
  requestAnimationFrame(() => toastEl.classList.add("show"));
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toastEl.classList.remove("show");
    setTimeout(() => (toastEl.hidden = true), 220);
  }, 2600);
}

function flog(...a) {
  try {
    if (window.pywebview && window.pywebview.api && window.pywebview.api.log) {
      window.pywebview.api.log(a.map(String).join(" "));
    }
  } catch (e) { /* 日志失败不影响主流程 */ }
}

window.addEventListener("error", (e) => {
  flog("[frontend error]", e.message, (e.filename || "") + ":" + e.lineno);
});

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}

/* ---------------- 事件绑定 ---------------- */
$("backBtn").addEventListener("click", showHome);
dropzone.addEventListener("click", pickFiles);
$("browseBtn").addEventListener("click", pickDir);
convertBtn.addEventListener("click", onConvert);

// 拖拽：webview 下拿不到路径，退化为点击选择
["dragover", "drop"].forEach((ev) =>
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    if (ev === "drop") pickFiles();
  })
);

/* ---------------- 启动 ---------------- */
function init() {
  renderNav();
  renderCards();
  if (window.pywebview && window.pywebview.api) {
    window.pywebview.api.check_engine().then((e) => {
      engine = e || engine;
    });
  }
}

if (window.pywebview && window.pywebview.ready) {
  window.pywebview.ready.then(init);
} else if (window.pywebview) {
  window.addEventListener("pywebviewready", init);
} else {
  // 纯浏览器预览（无 pywebview）也渲染布局，仅跳过引擎探测
  init();
}
