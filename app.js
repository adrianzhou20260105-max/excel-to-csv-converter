/**
 * SheetCast - Excel to CSV Converter
 * Core application logic
 */

// State management
let currentWorkbook = null;
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
  btnDownloadZip.addEventListener("click", downloadAllAsZip);

  // Download active sheet as CSV
  btnDownloadSingle.addEventListener("click", downloadActiveSheetAsCSV);
}

// ==========================================
// 3. File Processing & SheetJS Integration
// ==========================================
function handleFileSelection(file) {
  // Verify file type
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

  const reader = new FileReader();
  
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      
      // Read workbook using SheetJS with configurations to preserve cell formats
      const workbook = XLSX.read(data, {
        type: "array",
        cellNF: true,    // Retain number formats (dates, currencies, custom formatters)
        cellText: true,  // Retain the formatted text values (prevents scientific notation for big ints)
        cellStyles: true // Retain styles if present
      });

      processWorkbook(workbook, file.size);
    } catch (err) {
      console.error(err);
      showError("解析 Excel 文件失败，该文件可能已损坏，或者格式不受支持。");
      showLoading(false);
    }
  };

  reader.onerror = function() {
    showError("读取文件错误，请重试。");
    showLoading(false);
  };

  reader.readAsArrayBuffer(file);
}

function processWorkbook(workbook, fileSizeInBytes) {
  currentWorkbook = workbook;
  
  // Format file size
  let sizeText = "";
  if (fileSizeInBytes < 1024) sizeText = fileSizeInBytes + " B";
  else if (fileSizeInBytes < 1024 * 1024) sizeText = (fileSizeInBytes / 1024).toFixed(2) + " KB";
  else sizeText = (fileSizeInBytes / (1024 * 1024)).toFixed(2) + " MB";

  // Update file details card
  infoFilename.textContent = currentFileName + currentFileExt;
  infoFilesize.textContent = sizeText;
  infoSheetcount.textContent = workbook.SheetNames.length;

  // Toggle visible elements
  showLoading(false);
  dropzone.classList.add("hidden");
  fileInfoCard.classList.remove("hidden");
  workspace.classList.remove("hidden");

  // Populate sheet list
  renderSheetList();

  // Select first sheet by default
  if (workbook.SheetNames.length > 0) {
    selectSheet(workbook.SheetNames[0]);
  }
}

// ==========================================
// 4. Workspace & UI Rendering
// ==========================================
function renderSheetList() {
  sheetListContainer.innerHTML = "";

  currentWorkbook.SheetNames.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    
    // Count rows/cols from ref
    let rows = 0;
    let cols = 0;
    if (sheet["!ref"]) {
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      rows = range.e.r - range.s.r + 1;
      cols = range.e.c - range.s.c + 1;
    }

    // Create item button
    const button = document.createElement("button");
    button.className = `sheet-item-btn w-full flex items-center justify-between p-3.5 rounded-xl border text-left transition-all text-xs font-medium cursor-pointer ${
      sheetName === activeSheetName
        ? "border-emerald-500 bg-emerald-50/50 dark:bg-emerald-950/20 text-emerald-600 dark:text-emerald-400 font-semibold"
        : "border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300"
    }`;
    button.setAttribute("data-sheetname", sheetName);
    
    button.innerHTML = `
      <div class="flex items-center space-x-2.5 min-w-0 pr-2">
        <i class="fa-solid fa-sheet-plastic flex-shrink-0 text-sm opacity-80"></i>
        <span class="truncate block">${escapeHtml(sheetName)}</span>
      </div>
      <div class="flex-shrink-0 text-[10px] text-slate-400 dark:text-slate-500 bg-slate-100 dark:bg-slate-800 px-2 py-0.5 rounded-md">
        ${rows}R × ${cols}C
      </div>
    `;

    button.addEventListener("click", () => selectSheet(sheetName));
    sheetListContainer.appendChild(button);
  });
}

function selectSheet(sheetName) {
  activeSheetName = sheetName;
  
  // Highlight active button in list
  const buttons = sheetListContainer.querySelectorAll(".sheet-item-btn");
  buttons.forEach(btn => {
    const btnSheetName = btn.getAttribute("data-sheetname");
    if (btnSheetName === sheetName) {
      btn.className = "sheet-item-btn w-full flex items-center justify-between p-3.5 rounded-xl border text-left transition-all text-xs font-semibold border-emerald-500 bg-emerald-50/50 dark:bg-emerald-950/20 text-emerald-600 dark:text-emerald-400 cursor-pointer";
    } else {
      btn.className = "sheet-item-btn w-full flex items-center justify-between p-3.5 rounded-xl border text-left transition-all text-xs font-medium border-slate-200 dark:border-slate-800 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-600 dark:text-slate-300 cursor-pointer";
    }
  });

  // Render preview
  renderPreview();
}

function renderPreview() {
  const sheet = currentWorkbook.Sheets[activeSheetName];
  previewSheetName.querySelector("span").textContent = activeSheetName;

  if (!sheet || !sheet["!ref"]) {
    // Empty sheet handling
    previewTotalRows.textContent = "0";
    previewTotalCols.textContent = "0";
    previewTableHead.innerHTML = `<tr><th class="p-4 text-center text-slate-400">表格内容为空</th></tr>`;
    previewTableBody.innerHTML = `<tr><td class="p-8 text-center text-slate-400 font-light">该 Sheet 内没有可显示的数据。</td></tr>`;
    return;
  }

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const totalRows = range.e.r - range.s.r + 1;
  const totalCols = range.e.c - range.s.c + 1;
  
  previewTotalRows.textContent = totalRows;
  previewTotalCols.textContent = totalCols;

  // Maximum columns & rows to render in preview for performance reasons
  const maxPreviewRows = 15;
  const maxPreviewCols = 15;

  const startRow = range.s.r;
  const endRow = Math.min(startRow + maxPreviewRows - 1, range.e.r);
  
  const startCol = range.s.c;
  const endCol = Math.min(startCol + maxPreviewCols - 1, range.e.c);
  const hasMoreCols = range.e.c > endCol;

  // Generate Header Row
  let headHtml = "<tr>";
  // Top-left index header
  headHtml += `<th class="p-2.5 border-b border-r border-slate-200 dark:border-slate-800 text-center w-12 bg-slate-100 dark:bg-slate-800 text-slate-500 select-none font-semibold">#</th>`;
  
  for (let c = startCol; c <= endCol; c++) {
    const colName = XLSX.utils.encode_col(c);
    headHtml += `<th class="p-2.5 border-b border-r border-slate-200 dark:border-slate-800 text-center font-semibold bg-slate-100/80 dark:bg-slate-800/80 min-w-[120px] max-w-[200px] truncate text-slate-600 dark:text-slate-300">${colName}</th>`;
  }
  
  if (hasMoreCols) {
    headHtml += `<th class="p-2.5 border-b border-slate-200 dark:border-slate-800 text-center font-semibold bg-slate-100/50 dark:bg-slate-800/50 text-slate-400 italic">更多...</th>`;
  }
  headHtml += "</tr>";
  previewTableHead.innerHTML = headHtml;

  // Generate Body Rows
  let bodyHtml = "";
  for (let r = startRow; r <= endRow; r++) {
    bodyHtml += `<tr class="hover:bg-slate-50 dark:hover:bg-slate-800/30 transition-colors">`;
    
    // Row index label (e.g. 1, 2, 3...)
    bodyHtml += `<td class="p-2 border-b border-r border-slate-150 dark:border-slate-800 text-center bg-slate-50 dark:bg-slate-900/60 font-medium text-slate-500 select-none w-12">${r + 1}</td>`;
    
    for (let c = startCol; c <= endCol; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r: r, c: c });
      const cell = sheet[cellAddress];
      let val = "";
      
      if (cell) {
        // Prefer cell.w (formatted display string) over raw cell.v
        val = cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : "");
      }
      
      const escapedVal = escapeHtml(val);
      bodyHtml += `<td class="p-2 border-b border-r border-slate-150 dark:border-slate-800 truncate text-slate-600 dark:text-slate-300 max-w-[200px]" title="${escapedVal}">${escapedVal}</td>`;
    }
    
    if (hasMoreCols) {
      bodyHtml += `<td class="p-2 border-b border-slate-150 dark:border-slate-800 text-slate-400/70 italic text-center select-none">...</td>`;
    }
    
    bodyHtml += "</tr>";
  }
  previewTableBody.innerHTML = bodyHtml;
}

// ==========================================
// 5. CSV Conversion & Downloading
// ==========================================
function convertSheetToCSV(sheet) {
  if (!sheet) return "";
  
  // SheetJS conversion utility
  return XLSX.utils.sheet_to_csv(sheet, {
    blankrows: true, // Preserve spacing of rows
  });
}

function getCsvBlob(csvContent) {
  // Prepends UTF-8 BOM byte marker so Excel opens it correctly with Chinese characters
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
  return new Blob([bom, csvContent], { type: "text/csv;charset=utf-8;" });
}

function downloadActiveSheetAsCSV() {
  if (!currentWorkbook || !activeSheetName) return;

  const sheet = currentWorkbook.Sheets[activeSheetName];
  const csvContent = convertSheetToCSV(sheet);
  const blob = getCsvBlob(csvContent);
  
  // Sanitize sheet name for filenames
  const safeSheetName = activeSheetName.replace(/[\/\\?%*:|"<>\s]/g, "_");
  const fileName = `${currentFileName}_${safeSheetName}.csv`;
  
  saveBlob(blob, fileName);
}

function downloadAllAsZip() {
  if (!currentWorkbook) return;

  btnDownloadZip.disabled = true;
  const originalHtml = btnDownloadZip.innerHTML;
  btnDownloadZip.innerHTML = `<i class="fa-solid fa-spinner animate-spin"></i> <span>正在生成 ZIP...</span>`;

  const zip = new JSZip();

  currentWorkbook.SheetNames.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    const csvContent = convertSheetToCSV(sheet);
    const blob = getCsvBlob(csvContent);
    
    const safeSheetName = sheetName.replace(/[\/\\?%*:|"<>\s]/g, "_");
    const csvFileName = `${safeSheetName}.csv`;
    
    zip.file(csvFileName, blob);
  });

  zip.generateAsync({ type: "blob" })
    .then((zipBlob) => {
      const zipName = `${currentFileName}_CSVs.zip`;
      saveBlob(zipBlob, zipName);
      
      // Reset button state
      btnDownloadZip.disabled = false;
      btnDownloadZip.innerHTML = originalHtml;
    })
    .catch((err) => {
      console.error(err);
      alert("生成 ZIP 压缩包失败！");
      btnDownloadZip.disabled = false;
      btnDownloadZip.innerHTML = originalHtml;
    });
}

function saveBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  
  // Clean up
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

// ==========================================
// 6. Helpers & App UI State Control
// ==========================================
function resetAppState() {
  currentWorkbook = null;
  currentFileName = "";
  currentFileExt = "";
  activeSheetName = "";
  fileInput.value = "";
  
  // Toggle visibility back to default upload
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
