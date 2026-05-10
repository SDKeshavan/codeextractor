
// ——— State ———
let lookupMap = null;   // Map<"matCode|shadeCode", newMatCode>
let resultRows = null;  // Array of row objects after merge
let resultCols = null;  // Column headers from the Excel file + New Material Code

// ——— DOM refs ———
const csvInput = document.getElementById("csv-input");
const excelInput = document.getElementById("excel-input");
const csvDropZone = document.getElementById("csv-drop-zone");
const excelDropZone = document.getElementById("excel-drop-zone");
const csvFileInfo = document.getElementById("csv-file-info");
const excelFileInfo = document.getElementById("excel-file-info");
const step1Card = document.getElementById("step1-card");
const step2Card = document.getElementById("step2-card");
const resultsSection = document.getElementById("results-section");
const statsRow = document.getElementById("stats-row");
const tableWrapper = document.getElementById("table-wrapper");
const btnDownload = document.getElementById("btn-download");
const btnCopy = document.getElementById("btn-copy");
const copyLabel = document.getElementById("copy-label");
const btnCopyCodes = document.getElementById("btn-copy-codes");
const copyCodesLabel = document.getElementById("copy-codes-label");
const btnReload1 = document.getElementById("btn-reload-1");
const btnReload2 = document.getElementById("btn-reload-2");

// Mode toggle & single lookup refs
const modeBulkBtn = document.getElementById("mode-bulk");
const modeSingleBtn = document.getElementById("mode-single");
const panelBulk = document.getElementById("panel-bulk");
const panelSingle = document.getElementById("panel-single");
const inputMatCode = document.getElementById("input-mat-code");
const inputShadeCode = document.getElementById("input-shade-code");
const btnLookup = document.getElementById("btn-lookup");
const singleResult = document.getElementById("single-result");
const singleResultValue = document.getElementById("single-result-value");
const btnCopySingle = document.getElementById("btn-copy-single");
const copySingleLabel = document.getElementById("copy-single-label");

// ——— Toast helper ———
function toast(msg, type = "info", durationMs = 3500) {
  const el = document.createElement("div");
  el.className = `toast ${type}`;
  el.textContent = msg;
  document.getElementById("toast-container").appendChild(el);
  setTimeout(() => {
    el.classList.add("out");
    el.addEventListener("animationend", () => el.remove());
  }, durationMs);
}

// ——— Drag-and-drop wiring ———
function wireDragDrop(zone, fileInput) {
  ["dragenter", "dragover"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add("drag-over"); })
  );
  ["dragleave", "drop"].forEach(evt =>
    zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove("drag-over"); })
  );
  zone.addEventListener("drop", e => {
    const file = e.dataTransfer.files[0];
    if (file) {
      // Manually set the file on the input & trigger change
      const dt = new DataTransfer();
      dt.items.add(file);
      fileInput.files = dt.files;
      fileInput.dispatchEvent(new Event("change"));
    }
  });
}

wireDragDrop(csvDropZone, csvInput);
wireDragDrop(excelDropZone, excelInput);

// ——— Step 1: Parse lookup file (CSV or Excel) ———
csvInput.addEventListener("change", () => {
  const file = csvInput.files[0];
  if (!file) return;

  csvFileInfo.textContent = "Processing…";
  csvFileInfo.classList.remove("error");
  csvDropZone.classList.add("loading");

  const ext = file.name.split(".").pop().toLowerCase();

  if (ext === "csv") {
    // Parse as CSV with Papa Parse
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete(results) {
        csvDropZone.classList.remove("loading");
        const rows = results.data;
        const cols = results.meta.fields.map(f => f.trim());
        buildLookup(file, rows, cols);
      },
      error(err) {
        csvDropZone.classList.remove("loading");
        csvFileInfo.textContent = `Parse error: ${err.message}`;
        csvFileInfo.classList.add("error");
        toast("Failed to parse file", "error");
      }
    });
  } else {
    // Parse as Excel with SheetJS
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
        const cols = rows.length ? Object.keys(rows[0]).map(c => c.trim()) : [];
        csvDropZone.classList.remove("loading");
        buildLookup(file, rows, cols);
      } catch (err) {
        csvDropZone.classList.remove("loading");
        csvFileInfo.textContent = `Parse error: ${err.message}`;
        csvFileInfo.classList.add("error");
        toast("Failed to parse file", "error");
      }
    };
    reader.readAsArrayBuffer(file);
  }
});

/**
 * Shared logic: validate columns & build the lookup map from parsed rows.
 */
function buildLookup(file, rows, cols) {
  const required = ["Old Material Code", "Old Shade Code", "New Material Code"];
  const missing = required.filter(c => !cols.includes(c));

  if (missing.length) {
    csvFileInfo.textContent = `Missing columns: ${missing.join(", ")}`;
    csvFileInfo.classList.add("error");
    toast("File is missing required columns", "error");
    return;
  }

  lookupMap = new Map();
  for (const row of rows) {
    const mat = String(row["Old Material Code"] ?? "").trim();
    const shade = String(row["Old Shade Code"] ?? "").trim();
    const newMat = String(row["New Material Code"] ?? "").trim();
    if (mat && shade) {
      lookupMap.set(`${mat}|${shade}`, newMat);
    }
  }

  csvFileInfo.textContent = `✓ ${file.name}  —  ${lookupMap.size} entries loaded`;
  step1Card.classList.add("done");
  btnReload1.classList.remove("hidden");
  toast(`Loaded ${lookupMap.size} lookup entries`, "success");

  // Enable step 2
  step2Card.classList.remove("disabled");
  excelInput.disabled = false;
  btnLookup.disabled = false;
}

// ——— Step 2: Parse Excel & merge ———
excelInput.addEventListener("change", () => {
  const file = excelInput.files[0];
  if (!file || !lookupMap) return;

  excelFileInfo.textContent = "Processing…";
  excelFileInfo.classList.remove("error");
  excelDropZone.classList.add("loading");

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

      if (rows.length === 0) {
        excelDropZone.classList.remove("loading");
        excelFileInfo.textContent = "Excel file is empty";
        excelFileInfo.classList.add("error");
        toast("Excel file contains no data", "error");
        return;
      }

      // Trim column headers
      const rawCols = Object.keys(rows[0]);
      const colMap = {};
      rawCols.forEach(c => { colMap[c.trim()] = c; });

      const required = ["Old Material Code", "Old Shade Code"];
      const missing = required.filter(c => !(c in colMap));
      if (missing.length) {
        excelDropZone.classList.remove("loading");
        excelFileInfo.textContent = `Missing columns: ${missing.join(", ")}`;
        excelFileInfo.classList.add("error");
        toast("Excel is missing required columns", "error");
        return;
      }

      // Merge
      let matched = 0;
      for (const row of rows) {
        const mat = String(row[colMap["Old Material Code"]] ?? "").trim();
        const shade = String(row[colMap["Old Shade Code"]] ?? "").trim();
        const key = `${mat}|${shade}`;
        const newMat = lookupMap.get(key) ?? "";
        row["New Material Code"] = newMat;
        if (newMat) matched++;
      }

      const total = rows.length;
      const unmatched = total - matched;

      excelDropZone.classList.remove("loading");
      excelFileInfo.textContent = `✓ ${file.name}  —  ${total} rows processed`;
      step2Card.classList.add("done");
      btnReload2.classList.remove("hidden");
      toast(`Matched ${matched} / ${total} rows`, matched === total ? "success" : "info");

      // Determine display columns — preserve original order, add New Material Code at end if not present
      resultCols = rawCols.map(c => c.trim());
      if (!resultCols.includes("New Material Code")) {
        resultCols.push("New Material Code");
      }
      resultRows = rows;

      renderResults(total, matched, unmatched);
    } catch (err) {
      excelDropZone.classList.remove("loading");
      excelFileInfo.textContent = `Error: ${err.message}`;
      excelFileInfo.classList.add("error");
      toast("Failed to read Excel file", "error");
    }
  };
  reader.readAsArrayBuffer(file);
});

// ——— Render results ———
function renderResults(total, matched, unmatched) {
  // Stats chips
  statsRow.innerHTML = `
    <span class="stat-chip total">Total: ${total}</span>
    <span class="stat-chip matched">Matched: ${matched}</span>
    ${unmatched > 0 ? `<span class="stat-chip unmatched">Unmatched: ${unmatched}</span>` : ""}
  `;

  // Table (show first 100 rows max for performance)
  const displayRows = resultRows.slice(0, 100);
  let html = "<table class='result-table'><thead><tr>";
  for (const col of resultCols) {
    html += `<th>${escHtml(col)}</th>`;
  }
  html += "</tr></thead><tbody>";
  for (const row of displayRows) {
    html += "<tr>";
    for (const col of resultCols) {
      const val = String(row[col] ?? "");
      if (col === "New Material Code") {
        if (val) {
          html += `<td class="cell-new-code">${escHtml(val)}</td>`;
        } else {
          html += `<td class="cell-missing">—</td>`;
        }
      } else {
        html += `<td>${escHtml(val)}</td>`;
      }
    }
    html += "</tr>";
  }
  html += "</tbody></table>";
  if (resultRows.length > 100) {
    html += `<p style="padding:0.6rem;font-size:0.78rem;color:var(--text-muted)">Showing first 100 of ${resultRows.length} rows</p>`;
  }
  tableWrapper.innerHTML = html;

  resultsSection.classList.remove("hidden");
  resultsSection.scrollIntoView({ behavior: "smooth", block: "start" });
}

function escHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

// ——— Download ———
btnDownload.addEventListener("click", () => {
  if (!resultRows || !resultCols) return;

  // Build sheet data in column order
  const aoa = [resultCols];
  for (const row of resultRows) {
    aoa.push(resultCols.map(col => row[col] ?? ""));
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Updated");
  XLSX.writeFile(wb, "updated_output.xlsx");
  toast("File downloaded!", "success");
});

// ——— Copy table to clipboard ———
btnCopy.addEventListener("click", async () => {
  if (!resultRows || !resultCols) return;

  // Build tab-separated string (all rows, not just displayed 100)
  const lines = [resultCols.join("\t")];
  for (const row of resultRows) {
    lines.push(resultCols.map(col => String(row[col] ?? "")).join("\t"));
  }
  const text = lines.join("\n");

  try {
    await navigator.clipboard.writeText(text);
    // Visual feedback
    copyLabel.textContent = "Copied!";
    btnCopy.classList.add("copied");
    toast("Table copied to clipboard", "success");
    setTimeout(() => {
      copyLabel.textContent = "Copy Table";
      btnCopy.classList.remove("copied");
    }, 2000);
  } catch (e) {
    toast("Copy failed — try again", "error");
  }
});

// ——— Copy only New Material Code column ———
btnCopyCodes.addEventListener("click", async () => {
  if (!resultRows) return;

  const codes = resultRows.map(row => String(row["New Material Code"] ?? ""));
  const text = codes.join("\n");

  try {
    await navigator.clipboard.writeText(text);
    copyCodesLabel.textContent = "Copied!";
    btnCopyCodes.classList.add("copied");
    toast("New Material Codes copied", "success");
    setTimeout(() => {
      copyCodesLabel.textContent = "Copy New Codes";
      btnCopyCodes.classList.remove("copied");
    }, 2000);
  } catch (e) {
    toast("Copy failed — try again", "error");
  }
});

// ——— Reload / re-upload handlers ———

/** Reset step 2 and results (keeps lookup intact) */
function resetStep2() {
  excelInput.value = "";
  excelFileInfo.textContent = "";
  excelFileInfo.classList.remove("error");
  step2Card.classList.remove("done");
  btnReload2.classList.add("hidden");
  resultsSection.classList.add("hidden");
  resultRows = null;
  resultCols = null;
  statsRow.innerHTML = "";
  tableWrapper.innerHTML = "";
  // Reset single mode
  inputMatCode.value = "";
  inputShadeCode.value = "";
  singleResult.classList.add("hidden");
  singleResult.classList.remove("not-found");
}

/** Reset step 1 (cascades to step 2) */
function resetStep1() {
  resetStep2();
  csvInput.value = "";
  csvFileInfo.textContent = "";
  csvFileInfo.classList.remove("error");
  step1Card.classList.remove("done");
  btnReload1.classList.add("hidden");
  step2Card.classList.add("disabled");
  excelInput.disabled = true;
  btnLookup.disabled = true;
  lookupMap = null;
}

btnReload1.addEventListener("click", (e) => {
  e.preventDefault();
  e.stopPropagation();
  resetStep1();
  // Trigger file picker immediately
  csvInput.click();
});

btnReload2.addEventListener("click", (e) => {
  e.preventDefault();
  e.stopPropagation();
  resetStep2();
  // Trigger file picker immediately
  if (!panelBulk.classList.contains("hidden")) {
    excelInput.click();
  } else {
    inputMatCode.focus();
  }
});

// ——— Mode toggle ———
function setMode(mode) {
  if (mode === "bulk") {
    modeBulkBtn.classList.add("active");
    modeSingleBtn.classList.remove("active");
    panelBulk.classList.remove("hidden");
    panelSingle.classList.add("hidden");
  } else {
    modeSingleBtn.classList.add("active");
    modeBulkBtn.classList.remove("active");
    panelSingle.classList.remove("hidden");
    panelBulk.classList.add("hidden");
  }
  // Hide bulk results when switching to single
  if (mode === "single") {
    resultsSection.classList.add("hidden");
  }
}

modeBulkBtn.addEventListener("click", () => setMode("bulk"));
modeSingleBtn.addEventListener("click", () => setMode("single"));

// ——— Single lookup ———
function doSingleLookup() {
  if (!lookupMap) return;

  const mat = inputMatCode.value.trim();
  const shade = inputShadeCode.value.trim();

  if (!mat || !shade) {
    toast("Enter both codes", "error");
    return;
  }

  const key = `${mat}|${shade}`;
  const result = lookupMap.get(key);

  singleResult.classList.remove("hidden", "not-found");

  if (result !== undefined && result !== "") {
    singleResultValue.textContent = result;
    btnCopySingle.style.display = "";
    toast("Match found!", "success");
  } else {
    singleResultValue.textContent = "No match found";
    singleResult.classList.add("not-found");
    btnCopySingle.style.display = "none";
    toast("No matching entry in lookup", "error");
  }
}

btnLookup.addEventListener("click", doSingleLookup);

// Allow Enter key to trigger lookup
[inputMatCode, inputShadeCode].forEach(input => {
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      doSingleLookup();
    }
  });
});

// Copy single result
btnCopySingle.addEventListener("click", async () => {
  const val = singleResultValue.textContent;
  if (!val) return;
  try {
    await navigator.clipboard.writeText(val);
    copySingleLabel.textContent = "Copied!";
    btnCopySingle.classList.add("copied");
    toast("Copied to clipboard", "success");
    setTimeout(() => {
      copySingleLabel.textContent = "Copy";
      btnCopySingle.classList.remove("copied");
    }, 2000);
  } catch (e) {
    toast("Copy failed", "error");
  }
});
