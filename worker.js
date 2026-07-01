/**
 * SheetCast Web Worker
 * Offloads Excel parsing, CSV generation, and ZIP compression from the main UI thread.
 */

// Import SheetJS and JSZip locally
self.importScripts("lib/xlsx.full.min.js");
self.importScripts("lib/jszip.min.js");

// Cached state inside the worker thread
let currentWorkbook = null;

// Listen to incoming commands from the main thread
self.onmessage = function (e) {
  const { type, payload } = e.data;

  try {
    switch (type) {
      case "PARSE_FILE":
        handleParseFile(payload.arrayBuffer);
        break;

      case "GET_PREVIEW":
        handleGetPreview(payload.sheetName);
        break;

      case "EXPORT_SINGLE":
        handleExportSingle(payload.sheetName, payload.fileNamePrefix);
        break;

      case "EXPORT_ZIP":
        handleExportZip(payload.fileNamePrefix);
        break;

      default:
        throw new Error(`未知命令类型: ${type}`);
    }
  } catch (error) {
    console.error("Worker Execution Error:", error);
    self.postMessage({
      type: "ERROR",
      payload: { message: error.message }
    });
  }
};

// ==========================================
// 1. Parsing Handler
// ==========================================
function handleParseFile(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);

  // Synchronous read (runs in background worker thread, does not block main UI)
  const workbook = XLSX.read(data, {
    type: "array",
    cellNF: true,    // Retain number format strings
    cellText: true,  // Retain cell display values
    cellStyles: false // Community version of SheetJS doesn't support cellStyles; set to false to save memory
  });

  currentWorkbook = workbook;

  // Extract sheet names and their sizes (R x C)
  const sheetsMeta = workbook.SheetNames.map((name) => {
    const sheet = workbook.Sheets[name];
    let rows = 0;
    let cols = 0;

    if (sheet && sheet["!ref"]) {
      const range = XLSX.utils.decode_range(sheet["!ref"]);
      rows = range.e.r - range.s.r + 1;
      cols = range.e.c - range.s.c + 1;
    }

    return { name, rows, cols };
  });

  self.postMessage({
    type: "PARSE_COMPLETE",
    payload: {
      sheetsMeta,
      firstSheetName: workbook.SheetNames[0]
    }
  });
}

// ==========================================
// 2. Preview Data Handler
// ==========================================
function handleGetPreview(sheetName) {
  if (!currentWorkbook) throw new Error("没有读取到当前工作簿。");

  const sheet = currentWorkbook.Sheets[sheetName];
  if (!sheet || !sheet["!ref"]) {
    self.postMessage({
      type: "PREVIEW_DATA",
      payload: {
        sheetName,
        isEmpty: true
      }
    });
    return;
  }

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const totalRows = range.e.r - range.s.r + 1;
  const totalCols = range.e.c - range.s.c + 1;

  // Max preview size in browser table
  const maxRows = 15;
  const maxCols = 15;

  const startRow = range.s.r;
  const endRow = Math.min(startRow + maxRows - 1, range.e.r);

  const startCol = range.s.c;
  const endCol = Math.min(startCol + maxCols - 1, range.e.c);
  const hasMoreCols = range.e.c > endCol;

  // Generate headers (A, B, C...)
  const headers = [];
  for (let c = startCol; c <= endCol; c++) {
    headers.push(XLSX.utils.encode_col(c));
  }

  // Generate rows data (read cell.w for display formatting)
  const rows = [];
  for (let r = startRow; r <= endRow; r++) {
    const rowCells = [];
    for (let c = startCol; c <= endCol; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c });
      const cell = sheet[cellAddress];
      let val = "";

      if (cell) {
        val = cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : "");
      }
      rowCells.push(val);
    }
    rows.push({
      index: r + 1,
      cells: rowCells
    });
  }

  self.postMessage({
    type: "PREVIEW_DATA",
    payload: {
      sheetName,
      isEmpty: false,
      headers,
      rows,
      hasMoreCols,
      totalRows,
      totalCols
    }
  });
}

// ==========================================
// 3. Export Single Sheet Handler
// ==========================================
function handleExportSingle(sheetName, fileNamePrefix) {
  if (!currentWorkbook) throw new Error("没有读取到当前工作簿。");

  const sheet = currentWorkbook.Sheets[sheetName];
  if (!sheet) throw new Error(`找不到工作表: ${sheetName}`);

  // Escape formula injection
  const safeSheet = getSanitizedSheet(sheet);
  
  // SheetJS conversion utility
  const csvContent = XLSX.utils.sheet_to_csv(safeSheet, { blankrows: true });
  
  // Generate file blob
  const csvBlob = getCsvBlob(csvContent);
  
  // Sanitize sheet name for filenames
  const safeSheetName = sheetName.replace(/[\/\\?%*:|"<>\s]/g, "_");
  const fileName = `${fileNamePrefix}_${safeSheetName}.csv`;

  self.postMessage({
    type: "EXPORT_SINGLE_COMPLETE",
    payload: {
      fileName,
      blob: csvBlob
    }
  });
}

// ==========================================
// 4. Export ZIP (All Sheets) Handler
// ==========================================
function handleExportZip(fileNamePrefix) {
  if (!currentWorkbook) throw new Error("没有读取到当前工作簿。");

  const zip = new JSZip();
  const usedNames = new Set();

  currentWorkbook.SheetNames.forEach((sheetName) => {
    const sheet = currentWorkbook.Sheets[sheetName];
    if (!sheet) return;

    // Escape formula injection
    const safeSheet = getSanitizedSheet(sheet);
    const csvContent = XLSX.utils.sheet_to_csv(safeSheet, { blankrows: true });
    const csvBlob = getCsvBlob(csvContent);

    // Resolve name collision (replace invalid chars and check Set)
    const baseSafeName = sheetName.replace(/[\/\\?%*:|"<>\s]/g, "_");
    let safeName = baseSafeName;
    let counter = 1;
    
    while (usedNames.has(safeName.toLowerCase())) {
      safeName = `${baseSafeName}_${counter}`;
      counter++;
    }
    usedNames.add(safeName.toLowerCase());

    const csvFileName = `${safeName}.csv`;
    zip.file(csvFileName, csvBlob);
  });

  // Generate ZIP blob asynchronously inside the worker thread
  zip.generateAsync({ type: "blob" })
    .then((zipBlob) => {
      const zipName = `${fileNamePrefix}_CSVs.zip`;
      
      self.postMessage({
        type: "EXPORT_ZIP_COMPLETE",
        payload: {
          fileName: zipName,
          blob: zipBlob
        }
      });
    })
    .catch((err) => {
      throw new Error(`打包压缩 ZIP 失败: ${err.message}`);
    });
}

// ==========================================
// 5. Helpers
// ==========================================
/**
 * Clones a sheet and escapes Excel formulas to protect against CSV Formula Injection.
 * Prepends a single quote to any cell starting with =, +, -, @.
 */
function getSanitizedSheet(sheet) {
  const newSheet = { ...sheet };

  for (const key in newSheet) {
    if (key[0] === "!") continue; // Skip metadata keys

    const cell = newSheet[key];
    if (cell) {
      // Clone the cell object to avoid mutating the original cached workbook sheet
      const newCell = { ...cell };
      
      if (newCell.w !== undefined) {
        const valStr = String(newCell.w);
        if (/^[=\+\-\@]/.test(valStr)) {
          newCell.w = "'" + valStr;
        }
      } else if (newCell.v !== undefined) {
        const valStr = String(newCell.v);
        if (/^[=\+\-\@]/.test(valStr)) {
          newCell.v = "'" + valStr;
        }
      }
      newSheet[key] = newCell;
    }
  }

  return newSheet;
}

/**
 * Wraps CSV string in a Blob and prepends the UTF-8 BOM byte marker
 * (0xEF, 0xBB, 0xBF) to prevent Excel from displaying scrambled Chinese text.
 */
function getCsvBlob(csvContent) {
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
  return new Blob([bom, csvContent], { type: "text/csv;charset=utf-8;" });
}
