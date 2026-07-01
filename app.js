/**
 * SheetCast - Excel to CSV Converter
 * Main thread UI controller (communicates with Web Worker)
 */

// State management
let converterWorker = null;
let currentFileName = "";
let currentFileExt = "";
let activeSheetName = "";

// DOM Elements
const htmlElement = document.documentElement;
const themeToggleBtn = document.getElementById("theme-toggle");
const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("file-input");
const loadingSpinner = document.getElementById("loading-spinner");
const errorMessage = document.getElementById("error-message");
const errorText = document.getElementById("error-text");
const fileInfoCard = document.getElementById("file-info-card");
const infoFilename = document.getElementById("info-filename");
const infoFilesize = document.getElementById("info-filesize");
const infoSheetcount = document.getElementById("info-sheetcount");
const btnDownloadZip = document.getElementById("btn-download-zip");
const btnReset = document.getElementById("btn-reset");
const workspace = document.getElementById("workspace");
const sheetListContainer = document.getElementById("sheet-list-container");
const previewSheetName = document.getElementById("preview-sheet-name");
const previewTotalRows = document.getElementById("preview-total-rows");
const previewTotalCols = document.getElementById("preview-total-cols");
const btnDownloadSingle = document.getElementById("btn-download-single");
const previewTableHead = document.getElementById("preview-table-head");
const previewTableBody = document.getElementById("preview-table-body");

// Initialize application
document.addEventListener("DOMContentLoaded", () => {
  initTheme();
  setupEventListeners();
});

// ==========================================
// 1. Theme (Light/Dark Mode) Handling
// ==========================================
function initTheme() {
  const savedTheme = localStorage.getItem("theme");
  const systemPrefersDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
  
  if (savedTheme === "dark" || (!savedTheme && systemPrefersDark)) {
    htmlElement.classList.add("dark");
    htmlElement.classList.remove("light");
  } else {
    htmlElement.classList.add("light");
    htmlElement.classList.remove("dark");
  }
}

function toggleTheme() {
  if (htmlElement.classList.contains("dark")) {
    htmlElement.classList.remove("dark");
    htmlElement.classList.add("light");
    localStorage.setItem("theme", "light");
  } else {
    htmlElement.classList.remove("light");
    htmlElement.classList.add("dark");
    localStorage.setItem("theme", "dark");
  }
}

// ==========================================
// 2. Event Listeners Setup
// ==========================================
function setupEventListeners() {
  // Theme toggle
  themeToggleBtn.addEventListener("click", toggleTheme);

  // Drag and drop listeners
  dropzone.addEventListener("click", () => fileInput.click());
  
  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("drag-over");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("drag-over");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("drag-over");
    
    if (e.dataTransfer.files.length > 0) {
      handleFileSelection(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      handleFileSelection(e.target.files[0]);
    }
  });

  // Reset button
  btnReset.addEventListener("click", resetAppState);

  // Download all as ZIP
  btnDownloadZip.addEventListener("click", requestZipDownload);

  // Download active sheet as CSV
  btnDownloadSingle.addEventListener("click", requestSingleSheetDownload);
}

// ==========================================
// 3. Web Worker Communication & Lifecycle
// ==========================================
function initWorker() {
  // Terminate any existing worker to prevent memory leaks and clean state
  if (converterWorker) {
    converterWorker.terminate();
  }

  // Instantiate worker
  converterWorker = new Worker("worker.js");
  
  // Set up receiver for worker messages
  converterWorker.onmessage = function (e) {
    const { type, payload } = e.data;

    switch (type) {
      case "PARSE_COMPLETE":
        onParseComplete(payload.sheetsMeta, payload.firstSheetName);
        break;

      case "PREVIEW_DATA":
        onPreviewReceived(payload);
        break;

      case "EXPORT_SINGLE_COMPLETE":
        onExportSingleComplete(payload.fileName, payload.blob);
        break;

      case "EXPORT_ZIP_COMPLETE":
        onExportZipComplete(payload.fileName, payload.blob);
        break;

      case "ERROR":
        onWorkerError(payload.message);
        break;

      default:
        console.warn(`未知的 Worker 消息类型: ${type}`);
    }
  };

  converterWorker.onerror = function (err) {
    console.error("Worker Global Error:", err);
    showError("后台解析子线程加载失败，请刷新网页重试。");
    showLoading(false);
  };
}

// ==========================================
// 4. File Processing Trigger
// ==========================================
function handleFileSelection(file) {
  const fileName = file.name;
  const lastDot = fileName.lastIndexOf(".");
  const ext = fileName.substring(lastDot).toLowerCase();
  
  if (![".xlsx", ".xls", ".xlsm", ".xlsb"].includes(ext)) {
    showError("不支持的文件格式！请上传 Excel 文件 (.xlsx, .xls, .xlsm, .xlsb)");
    return;
  }

  hideError();
  showLoading(true);
  
  currentFileName = fileName.substring(0, lastDot);
  currentFileExt = ext;

  // Initialize fresh worker for file parsing
  initWorker();

  const reader = new FileReader();
  
  reader.onload = function (e) {
    const arrayBuffer = e.target.result;
    
    // Send arrayBuffer to the worker (transfer ownership to save memory)
    converterWorker.postMessage(
      {
        type: "PARSE_FILE",
        payload: { arrayBuffer }
      },
      [arrayBuffer] // Transferred arrayBuffer, main thread can no longer read it directly
    );
  };

  reader.onerror = function () {
    showError("读取文件错误，请重试。");
    showLoading(false);
  };

  // Keep file size formatted for display
  let sizeText = "";
  if (file.size < 1024) sizeText = file.size + " B";
  else if (file.size < 1024 * 1024) sizeText = (file.size / 1024).toFixed(2) + " KB";
  else sizeText = (file.size / (1024 * 1024)).toFixed(2) + " MB";

  infoFilename.textContent = file.name;
  infoFilesize.textContent = sizeText;

  reader.readAsArrayBuffer(file);
}

// ==========================================
// 5. Worker Message Receivers (Callbacks)
// ==========================================
function onParseComplete(sheetsMeta, firstSheetName) {
  infoSheetcount.textContent = sheetsMeta.length;

  showLoading(false);
  dropzone.classList.add("hidden");
  fileInfoCard.classList.remove("hidden");
  workspace.classList.remove("hidden");

  // Render Left Column (Sheet selection buttons)
  renderSheetList(sheetsMeta);

  // Automatically select and request preview for the first sheet
  if (firstSheetName) {
    selectSheet(firstSheetName);
  }
}

function renderSheetList(sheetsMeta) {
  sheetListContainer.innerHTML = "";

  sheetsMeta.forEach((meta) => {
    const button = document.createElement("button");
    button.className = `sheet-item-btn w-full flex items-center justify-between p-3.5 rounded-xl border text-left transition-all text-xs font-medium cursor-pointer border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300`;
    button.setAttribute("data-sheetname", meta.name);
    
    button.innerHTML = `
      <div class="flex items-center space-x-2.5 min-w-0 pr-2">
        <i class="fa-solid fa-sheet-plastic flex-shrink-0 text-sm opacity-80 text-emerald-500"></i>
        <span class="truncate block">${escapeHtml(meta.name)}</span>
      </div>
      <div class="flex-shrink-0 text-[10px] text-slate-400 dark:text-slate-500 bg-slate-100 dark:bg-slate-800 px-2 py-0.5 rounded-md">
        ${meta.rows}R × ${meta.cols}C
      </div>
    `;

    button.addEventListener("click", () => selectSheet(meta.name));
    sheetListContainer.appendChild(button);
  });
}

function selectSheet(sheetName) {
  activeSheetName = sheetName;
  
  // Highlight active button in the DOM list
  const buttons = sheetListContainer.querySelectorAll(".sheet-item-btn");
  buttons.forEach(btn => {
    const btnSheetName = btn.getAttribute("data-sheetname");
    if (btnSheetName === sheetName) {
      btn.className = "sheet-item-btn w-full flex items-center justify-between p-3.5 rounded-xl border text-left transition-all text-xs font-semibold border-emerald-500 bg-emerald-50/50 dark:bg-emerald-950/20 text-emerald-600 dark:text-emerald-400 cursor-pointer";
    } else {
      btn.className = "sheet-item-btn w-full flex items-center justify-between p-3.5 rounded-xl border text-left transition-all text-xs font-medium border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300 cursor-pointer";
    }
  });

  // Display a brief "loading preview" status in table
  previewTableHead.innerHTML = "";
  previewTableBody.innerHTML = `<tr><td class="p-8 text-center text-slate-400 font-light"><i class="fa-solid fa-spinner animate-spin mr-2"></i>正在加载预览数据...</td></tr>`;

  // Request preview from worker
  converterWorker.postMessage({
    type: "GET_PREVIEW",
    payload: { sheetName }
  });
}

function onPreviewReceived(payload) {
  const { sheetName, isEmpty, headers, rows, hasMoreCols, totalRows, totalCols } = payload;

  // Protect against race conditions (ignore message if user already switched sheets)
  if (sheetName !== activeSheetName) return;

  previewSheetName.querySelector("span").textContent = sheetName;
  previewTotalRows.textContent = totalRows || 0;
  previewTotalCols.textContent = totalCols || 0;

  if (isEmpty) {
    previewTableHead.innerHTML = `<tr><th class="p-4 text-center text-slate-400">表格内容为空</th></tr>`;
    previewTableBody.innerHTML = `<tr><td class="p-8 text-center text-slate-400 font-light">该 Sheet 内没有可显示的数据。</td></tr>`;
    return;
  }

  // Render Table Header Row (A, B, C...)
  let headHtml = "<tr>";
  headHtml += `<th class="p-2.5 border-b border-r border-slate-200 dark:border-slate-800 text-center w-12 bg-slate-100 dark:bg-slate-800 text-slate-500 select-none font-semibold">#</th>`;
  
  headers.forEach((colName) => {
    headHtml += `<th class="p-2.5 border-b border-r border-slate-200 dark:border-slate-800 text-center font-semibold bg-slate-100/80 dark:bg-slate-800/80 min-w-[120px] max-w-[200px] truncate text-slate-600 dark:text-slate-300">${colName}</th>`;
  });
  
  if (hasMoreCols) {
    headHtml += `<th class="p-2.5 border-b border-slate-200 dark:border-slate-800 text-center font-semibold bg-slate-100/50 dark:bg-slate-800/50 text-slate-400 italic">更多...</th>`;
  }
  headHtml += "</tr>";
  previewTableHead.innerHTML = headHtml;

  // Render Table Body Rows (with standard title tooltip to prevent layout jitter)
  let bodyHtml = "";
  rows.forEach((row) => {
    bodyHtml += `<tr class="hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors">`;
    bodyHtml += `<td class="p-2 border-b border-r border-slate-150 dark:border-slate-800 text-center bg-slate-50 dark:bg-slate-900/60 font-medium text-slate-500 select-none w-12">${row.index}</td>`;
    
    row.cells.forEach((cellVal) => {
      const escapedVal = escapeHtml(cellVal);
      bodyHtml += `<td class="p-2 border-b border-r border-slate-150 dark:border-slate-800 truncate text-slate-600 dark:text-slate-300 max-w-[200px]" title="${escapedVal}">${escapedVal}</td>`;
    });

    if (hasMoreCols) {
      bodyHtml += `<td class="p-2 border-b border-slate-150 dark:border-slate-800 text-slate-400/70 italic text-center select-none">...</td>`;
    }
    bodyHtml += "</tr>";
  });
  previewTableBody.innerHTML = bodyHtml;
}

// ==========================================
// 6. Action Triggers (Export commands to Worker)
// ==========================================
function requestSingleSheetDownload() {
  if (!converterWorker || !activeSheetName) return;

  btnDownloadSingle.disabled = true;
  const originalHtml = btnDownloadSingle.innerHTML;
  btnDownloadSingle.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> <span>正在导出...</span>`;

  // Store original HTML back to the DOM element after download complete callback
  btnDownloadSingle.setAttribute("data-orig-html", originalHtml);

  converterWorker.postMessage({
    type: "EXPORT_SINGLE",
    payload: {
      sheetName: activeSheetName,
      fileNamePrefix: currentFileName
    }
  });
}

function onExportSingleComplete(fileName, blob) {
  saveBlob(blob, fileName);

  // Restore button state
  btnDownloadSingle.disabled = false;
  const origHtml = btnDownloadSingle.getAttribute("data-orig-html");
  if (origHtml) {
    btnDownloadSingle.innerHTML = origHtml;
  }
}

function requestZipDownload() {
  if (!converterWorker) return;

  btnDownloadZip.disabled = true;
  const originalHtml = btnDownloadZip.innerHTML;
  btnDownloadZip.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> <span>正在生成 ZIP...</span>`;

  btnDownloadZip.setAttribute("data-orig-html", originalHtml);

  converterWorker.postMessage({
    type: "EXPORT_ZIP",
    payload: {
      fileNamePrefix: currentFileName
    }
  });
}

function onExportZipComplete(fileName, blob) {
  saveBlob(blob, fileName);

  btnDownloadZip.disabled = false;
  const origHtml = btnDownloadZip.getAttribute("data-orig-html");
  if (origHtml) {
    btnDownloadZip.innerHTML = origHtml;
  }
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  
  // Clean up ObjectURL reference
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// ==========================================
// 7. Error Handling & State Reset
// ==========================================
function onWorkerError(message) {
  showError(`转化数据失败: ${message}`);
  showLoading(false);

  // Reset button states in case they were loading
  btnDownloadSingle.disabled = false;
  const singleHtml = btnDownloadSingle.getAttribute("data-orig-html");
  if (singleHtml) btnDownloadSingle.innerHTML = singleHtml;

  btnDownloadZip.disabled = false;
  const zipHtml = btnDownloadZip.getAttribute("data-orig-html");
  if (zipHtml) btnDownloadZip.innerHTML = zipHtml;
}

function resetAppState() {
  if (converterWorker) {
    converterWorker.terminate();
    converterWorker = null;
  }

  currentFileName = "";
  currentFileExt = "";
  activeSheetName = "";
  fileInput.value = "";
  
  dropzone.classList.remove("hidden");
  fileInfoCard.classList.add("hidden");
  workspace.classList.add("hidden");
  hideError();
}

function showLoading(show) {
  if (show) {
    loadingSpinner.classList.remove("hidden");
    dropzone.classList.add("pointer-events-none");
    dropzone.classList.add("opacity-50");
  } else {
    loadingSpinner.classList.add("hidden");
    dropzone.classList.remove("pointer-events-none");
    dropzone.classList.remove("opacity-50");
  }
}

function showError(message) {
  errorText.textContent = message;
  errorMessage.classList.remove("hidden");
}

function hideError() {
  errorMessage.classList.add("hidden");
}

function escapeHtml(string) {
  if (string === null || string === undefined) return "";
  return String(string)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
