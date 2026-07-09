const DEFAULT_CONFIG = {
  soldCodeColumn: "รหัสสินค้า",
  vrichMatchColumn: "รหัสขาย",
  jstMatchColumn: "รหัสรูปแบบ",
  vrichQtyColumn: "จำนวน",
  jstQtyColumn: "จำนวน",
};

const state = {
  files: { vrich: null, jst: null, sold: null },
  tables: null,
  result: null,
  excludedCodes: new Set(),
  downloads: [],
};

const els = {
  vrichFile: document.getElementById("vrichFile"),
  jstFile: document.getElementById("jstFile"),
  soldFile: document.getElementById("soldFile"),
  vrichFileName: document.getElementById("vrichFileName"),
  jstFileName: document.getElementById("jstFileName"),
  soldFileName: document.getElementById("soldFileName"),
  soldCodeColumn: document.getElementById("soldCodeColumn"),
  vrichMatchColumn: document.getElementById("vrichMatchColumn"),
  jstMatchColumn: document.getElementById("jstMatchColumn"),
  vrichQtyColumn: document.getElementById("vrichQtyColumn"),
  jstQtyColumn: document.getElementById("jstQtyColumn"),
  inspectButton: document.getElementById("inspectButton"),
  processButton: document.getElementById("processButton"),
  clearButton: document.getElementById("clearButton"),
  applyExclusionsButton: document.getElementById("applyExclusionsButton"),
  statusPanel: document.getElementById("statusPanel"),
  preflightGrid: document.getElementById("preflightGrid"),
  summaryGrid: document.getElementById("summaryGrid"),
  missingPanel: document.getElementById("missingPanel"),
  missingTableWrap: document.getElementById("missingTableWrap"),
  downloadPanel: document.getElementById("downloadPanel"),
  downloadList: document.getElementById("downloadList"),
  runStatus: document.getElementById("runStatus"),
};

function getConfig() {
  return {
    soldCodeColumn: els.soldCodeColumn.value.trim() || DEFAULT_CONFIG.soldCodeColumn,
    vrichMatchColumn: els.vrichMatchColumn.value.trim() || DEFAULT_CONFIG.vrichMatchColumn,
    jstMatchColumn: els.jstMatchColumn.value.trim() || DEFAULT_CONFIG.jstMatchColumn,
    vrichQtyColumn: els.vrichQtyColumn.value.trim() || DEFAULT_CONFIG.vrichQtyColumn,
    jstQtyColumn: els.jstQtyColumn.value.trim() || DEFAULT_CONFIG.jstQtyColumn,
  };
}

function normalizeCode(value) {
  if (value === null || value === undefined) return "";
  let text = String(value);
  if (["nan", "none", "nat"].includes(text.toLowerCase())) return "";
  text = text
    .replace(/\u00a0/g, " ")
    .replace(/[\ufeff\u200b\u200c\u200d\u2060]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  if (/^.+\.0+$/.test(text)) text = text.replace(/\.0+$/, "");
  return text.toUpperCase();
}

function normalizeHeader(value) {
  return String(value ?? "")
    .replace(/\ufeff/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function setStatus(message, kind = "ok") {
  els.statusPanel.textContent = message;
  els.statusPanel.className = `status-panel visible ${kind}`;
}

function setRunStatus(status) {
  els.runStatus.textContent = status;
  els.runStatus.className = "status-pill";
  if (status === "PASS" || status === "PASS_WITH_EXCLUSION") {
    els.runStatus.classList.add("pass");
  } else if (status === "IDLE") {
    els.runStatus.classList.add("idle");
  } else if (status.includes("BLOCKED") || status === "FAIL") {
    els.runStatus.classList.add("fail");
  } else {
    els.runStatus.classList.add("warn");
  }
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function updateFileLabel(kind, file) {
  const label = {
    vrich: els.vrichFileName,
    jst: els.jstFileName,
    sold: els.soldFileName,
  }[kind];
  label.textContent = file ? file.name : "ยังไม่ได้เลือกไฟล์";
}

function readWorkbook(file) {
  return file.arrayBuffer().then((buffer) => {
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false, raw: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
    const headerRow = matrix.find((row) => row.some((cell) => normalizeHeader(cell)));
    if (!headerRow) throw new Error(`ไม่พบ header ในไฟล์ ${file.name}`);
    const headers = headerRow.map(normalizeHeader);
    while (headers.length && !headers[headers.length - 1]) headers.pop();
    const headerIndex = matrix.indexOf(headerRow);
    const rows = matrix
      .slice(headerIndex + 1)
      .map((row) => {
        const item = {};
        headers.forEach((header, index) => {
          item[header] = row[index] ?? "";
        });
        return item;
      })
      .filter((row) => headers.some((header) => String(row[header] ?? "").trim() !== ""));

    return {
      file,
      workbook,
      sheetName,
      headers,
      rows,
    };
  });
}

function requireColumns(table, columns, label) {
  const missing = columns.filter((column) => !table.headers.includes(column));
  if (missing.length) {
    throw new Error(`${label} ไม่มีคอลัมน์: ${missing.join(", ")}`);
  }
}

function buildIndex(rows, column) {
  const index = new Map();
  for (const row of rows) {
    const code = normalizeCode(row[column]);
    if (!code) continue;
    if (!index.has(code)) index.set(code, []);
    index.get(code).push(row);
  }
  return index;
}

function orderedSoldCodes(rows, column) {
  const seen = new Set();
  const result = [];
  for (const row of rows) {
    const code = normalizeCode(row[column]);
    if (!code || seen.has(code)) continue;
    seen.add(code);
    result.push({ code, raw: row[column] });
  }
  return result;
}

function duplicateCodesInSold(index, soldCodes) {
  const soldSet = new Set(soldCodes.map((item) => item.code));
  return [...index.entries()]
    .filter(([code, rows]) => soldSet.has(code) && rows.length > 1)
    .map(([code, rows]) => ({ code, count: rows.length }));
}

function inspectFiles() {
  const config = getConfig();
  if (!state.files.vrich || !state.files.jst || !state.files.sold) {
    throw new Error("กรุณาเลือกไฟล์ให้ครบ 3 ไฟล์");
  }

  return Promise.all([
    readWorkbook(state.files.vrich),
    readWorkbook(state.files.jst),
    readWorkbook(state.files.sold),
  ]).then(([vrich, jst, sold]) => {
    requireColumns(vrich, [config.vrichMatchColumn, config.vrichQtyColumn], "vRich");
    requireColumns(jst, [config.jstMatchColumn, config.jstQtyColumn], "JST");
    requireColumns(sold, [config.soldCodeColumn], "sold_today");
    state.tables = { vrich, jst, sold };
    renderPreflight();
    els.processButton.disabled = false;
    setStatus("ตรวจสอบไฟล์ผ่าน", "ok");
  });
}

function renderPreflight() {
  if (!state.tables) {
    els.preflightGrid.innerHTML = "";
    return;
  }
  const tableEntries = [
    ["vRich master", state.tables.vrich],
    ["JST stock", state.tables.jst],
    ["sold_today", state.tables.sold],
  ];
  els.preflightGrid.innerHTML = tableEntries
    .map(([label, table]) => {
      return `
        <article class="file-card">
          <h3>${escapeHtml(label)}</h3>
          <dl>
            <dt>ไฟล์</dt><dd>${escapeHtml(table.file.name)}</dd>
            <dt>sheet</dt><dd>${escapeHtml(table.sheetName)}</dd>
            <dt>แถว</dt><dd>${table.rows.length.toLocaleString()}</dd>
            <dt>คอลัมน์</dt><dd>${table.headers.length.toLocaleString()}</dd>
          </dl>
          <div class="columns">${escapeHtml(table.headers.join(", "))}</div>
        </article>
      `;
    })
    .join("");
}

function processData() {
  if (!state.tables) throw new Error("กรุณาตรวจสอบไฟล์ก่อนประมวลผล");
  const config = getConfig();
  const { vrich, jst, sold } = state.tables;

  requireColumns(vrich, [config.vrichMatchColumn, config.vrichQtyColumn], "vRich");
  requireColumns(jst, [config.jstMatchColumn, config.jstQtyColumn], "JST");
  requireColumns(sold, [config.soldCodeColumn], "sold_today");

  const soldCodes = orderedSoldCodes(sold.rows, config.soldCodeColumn);
  const vrichIndex = buildIndex(vrich.rows, config.vrichMatchColumn);
  const jstIndex = buildIndex(jst.rows, config.jstMatchColumn);
  const duplicateVrich = duplicateCodesInSold(vrichIndex, soldCodes);
  const duplicateJst = duplicateCodesInSold(jstIndex, soldCodes);

  const missingVrich = [];
  const missingJst = [];
  const updateRows = [];
  const blockedDuplicates = new Set([
    ...duplicateVrich.map((item) => item.code),
    ...duplicateJst.map((item) => item.code),
  ]);

  for (const soldCode of soldCodes) {
    const code = soldCode.code;
    const vrichRows = vrichIndex.get(code) || [];
    const jstRows = jstIndex.get(code) || [];
    if (!vrichRows.length) missingVrich.push(soldCode);
    if (!jstRows.length) missingJst.push(soldCode);
    if (vrichRows.length === 1 && jstRows.length === 1 && !blockedDuplicates.has(code) && !state.excludedCodes.has(code)) {
      const outputRow = {};
      for (const header of vrich.headers) outputRow[header] = vrichRows[0][header];
      outputRow[config.vrichQtyColumn] = jstRows[0][config.jstQtyColumn];
      updateRows.push(outputRow);
    }
  }

  const remainingMissingVrich = missingVrich.filter((item) => !state.excludedCodes.has(item.code));
  const remainingMissingJst = missingJst.filter((item) => !state.excludedCodes.has(item.code));
  const excluded = [...state.excludedCodes].filter((code) => soldCodes.some((item) => item.code === code));

  let status = "PASS";
  if (duplicateVrich.length || duplicateJst.length) status = "FAIL_DUPLICATE";
  if (remainingMissingVrich.length || remainingMissingJst.length) status = "FAIL";
  if (!remainingMissingVrich.length && !remainingMissingJst.length && excluded.length && !duplicateVrich.length && !duplicateJst.length) {
    status = "PASS_WITH_EXCLUSION";
  }

  state.result = {
    config,
    soldCodes,
    vrichIndex,
    jstIndex,
    missingVrich,
    missingJst,
    remainingMissingVrich,
    remainingMissingJst,
    duplicateVrich,
    duplicateJst,
    excluded,
    updateRows,
    status,
  };

  renderSummary();
  renderMissing();
  buildDownloads();
  renderDownloads();
  setRunStatus(status);
  setStatus(status === "PASS" || status === "PASS_WITH_EXCLUSION" ? "ประมวลผลสำเร็จ" : "ยังไม่ควร import", status === "FAIL" || status === "FAIL_DUPLICATE" ? "danger" : "ok");
}

function renderSummary() {
  const result = state.result;
  if (!result) {
    els.summaryGrid.innerHTML = "";
    return;
  }
  const metrics = [
    ["sold_today codes", result.soldCodes.length],
    ["match vRich", result.soldCodes.length - result.missingVrich.length],
    ["match JST", result.soldCodes.length - result.missingJst.length],
    ["update สำเร็จ", result.updateRows.length],
    ["missing vRich", result.remainingMissingVrich.length],
    ["missing JST", result.remainingMissingJst.length],
    ["duplicate vRich", result.duplicateVrich.length],
    ["duplicate JST", result.duplicateJst.length],
    ["excluded", result.excluded.length],
    ["status", result.status],
  ];
  els.summaryGrid.innerHTML = metrics
    .map(([label, value]) => `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`)
    .join("");
}

function renderMissing() {
  const result = state.result;
  if (!result) {
    els.missingPanel.classList.add("hidden");
    return;
  }
  const missingMap = new Map();
  for (const item of result.missingVrich) {
    if (!missingMap.has(item.code)) missingMap.set(item.code, { code: item.code, raw: item.raw, missingVrich: false, missingJst: false });
    missingMap.get(item.code).missingVrich = true;
  }
  for (const item of result.missingJst) {
    if (!missingMap.has(item.code)) missingMap.set(item.code, { code: item.code, raw: item.raw, missingVrich: false, missingJst: false });
    missingMap.get(item.code).missingJst = true;
  }
  const rows = [...missingMap.values()];
  if (!rows.length) {
    els.missingPanel.classList.add("hidden");
    els.missingTableWrap.innerHTML = "";
    return;
  }
  els.missingPanel.classList.remove("hidden");
  els.missingTableWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>exclude</th>
          <th>code</th>
          <th>missing vRich</th>
          <th>missing JST</th>
          <th>note</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map((row) => {
            const checked = state.excludedCodes.has(row.code) ? "checked" : "";
            const note = state.excludedCodes.has(row.code) ? "excluded โดยผู้ใช้ยืนยัน" : "";
            return `
              <tr>
                <td><input type="checkbox" data-exclude-code="${escapeHtml(row.code)}" ${checked}></td>
                <td>${escapeHtml(row.code)}</td>
                <td>${row.missingVrich ? "YES" : "NO"}</td>
                <td>${row.missingJst ? "YES" : "NO"}</td>
                <td>${escapeHtml(note)}</td>
              </tr>
            `;
          })
          .join("")}
      </tbody>
    </table>
  `;
}

function sheetFromRows(rows, headers) {
  const matrix = [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))];
  return XLSX.utils.aoa_to_sheet(matrix);
}

function downloadWorkbook(filename, sheets) {
  const workbook = XLSX.utils.book_new();
  for (const [sheetName, sheet] of sheets) {
    XLSX.utils.book_append_sheet(workbook, sheet, sheetName.slice(0, 31));
  }
  const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  return { filename, url };
}

function reportRowsForMissing(items, reason, config) {
  return items.map((item) => ({
    "รหัสที่ใช้จับคู่": item.code,
    "เหตุผล": reason,
    [config.soldCodeColumn]: item.raw,
  }));
}

function buildDownloads() {
  for (const item of state.downloads) URL.revokeObjectURL(item.url);
  state.downloads = [];
  const result = state.result;
  if (!result) return;
  const { config } = result;

  const summaryRows = [
    { รายการ: "status", ค่า: result.status },
    { รายการ: "workflow", ค่า: "Excel-only browser-only; no API; no backend; no upload; no auto import" },
    { รายการ: "SOLD_TODAY_CODE_COLUMN", ค่า: config.soldCodeColumn },
    { รายการ: "VRICH_MATCH_COLUMN", ค่า: config.vrichMatchColumn },
    { รายการ: "JST_MATCH_COLUMN", ค่า: config.jstMatchColumn },
    { รายการ: "VRICH_QTY_COLUMN", ค่า: config.vrichQtyColumn },
    { รายการ: "JST_QTY_COLUMN", ค่า: config.jstQtyColumn },
    { รายการ: "จำนวนรหัสใน sold_today", ค่า: result.soldCodes.length },
    { รายการ: "match vRich", ค่า: result.soldCodes.length - result.missingVrich.length },
    { รายการ: "match JST", ค่า: result.soldCodes.length - result.missingJst.length },
    { รายการ: "update สำเร็จ", ค่า: result.updateRows.length },
    { รายการ: "missing vRich หลัง exclude", ค่า: result.remainingMissingVrich.length },
    { รายการ: "missing JST หลัง exclude", ค่า: result.remainingMissingJst.length },
    { รายการ: "duplicate vRich ที่ชน sold_today", ค่า: result.duplicateVrich.length },
    { รายการ: "duplicate JST ที่ชน sold_today", ค่า: result.duplicateJst.length },
    { รายการ: "excluded codes", ค่า: result.excluded.join(", ") || "-" },
  ];

  const excludedRows = result.excluded.map((code) => ({
    code,
    note: "excluded โดยผู้ใช้ยืนยัน เพราะไม่พบใน JST",
  }));

  const duplicateVrichRows = result.duplicateVrich.map((item) => ({ code: item.code, count: item.count }));
  const duplicateJstRows = result.duplicateJst.map((item) => ({ code: item.code, count: item.count }));

  state.downloads.push(
    downloadWorkbook("vrich_import_update_qty.xlsx", [
      ["Sheet1", sheetFromRows(result.updateRows, state.tables.vrich.headers)],
    ])
  );
  state.downloads.push(
    downloadWorkbook("summary_report.xlsx", [
      ["summary", XLSX.utils.json_to_sheet(summaryRows)],
    ])
  );
  state.downloads.push(
    downloadWorkbook("report_missing_in_vrich.xlsx", [
      ["Sheet1", XLSX.utils.json_to_sheet(reportRowsForMissing(result.remainingMissingVrich, "อยู่ใน sold_today แต่ไม่พบใน vRich", config))],
    ])
  );
  state.downloads.push(
    downloadWorkbook("report_missing_in_jst.xlsx", [
      ["Sheet1", XLSX.utils.json_to_sheet(reportRowsForMissing(result.remainingMissingJst, "อยู่ใน sold_today แต่ไม่พบใน JST", config))],
    ])
  );
  state.downloads.push(
    downloadWorkbook("report_duplicate_vrich.xlsx", [
      ["Sheet1", XLSX.utils.json_to_sheet(duplicateVrichRows)],
    ])
  );
  state.downloads.push(
    downloadWorkbook("report_duplicate_jst.xlsx", [
      ["Sheet1", XLSX.utils.json_to_sheet(duplicateJstRows)],
    ])
  );
  if (excludedRows.length) {
    state.downloads.push(
      downloadWorkbook("excluded_codes_report.xlsx", [
        ["excluded_codes", XLSX.utils.json_to_sheet(excludedRows)],
      ])
    );
  }
}

function renderDownloads() {
  if (!state.downloads.length) {
    els.downloadPanel.classList.add("hidden");
    els.downloadList.innerHTML = "";
    return;
  }
  els.downloadPanel.classList.remove("hidden");
  els.downloadList.innerHTML = state.downloads
    .map((item) => `<a href="${item.url}" download="${escapeHtml(item.filename)}">${escapeHtml(item.filename)}</a>`)
    .join("");
}

function clearResults() {
  state.tables = null;
  state.result = null;
  state.excludedCodes.clear();
  for (const item of state.downloads) URL.revokeObjectURL(item.url);
  state.downloads = [];
  els.processButton.disabled = true;
  els.preflightGrid.innerHTML = "";
  els.summaryGrid.innerHTML = "";
  els.missingTableWrap.innerHTML = "";
  els.downloadList.innerHTML = "";
  els.missingPanel.classList.add("hidden");
  els.downloadPanel.classList.add("hidden");
  els.statusPanel.className = "status-panel";
  els.statusPanel.textContent = "";
  setRunStatus("IDLE");
}

function bindFileInput(input, kind) {
  input.addEventListener("change", () => {
    state.files[kind] = input.files[0] || null;
    updateFileLabel(kind, state.files[kind]);
    clearResults();
  });
}

bindFileInput(els.vrichFile, "vrich");
bindFileInput(els.jstFile, "jst");
bindFileInput(els.soldFile, "sold");

els.inspectButton.addEventListener("click", () => {
  inspectFiles().catch((error) => setStatus(error.message, "danger"));
});

els.processButton.addEventListener("click", () => {
  try {
    processData();
  } catch (error) {
    setStatus(error.message, "danger");
  }
});

els.applyExclusionsButton.addEventListener("click", () => {
  const checks = els.missingTableWrap.querySelectorAll("[data-exclude-code]");
  state.excludedCodes.clear();
  checks.forEach((check) => {
    if (check.checked) state.excludedCodes.add(check.dataset.excludeCode);
  });
  try {
    processData();
  } catch (error) {
    setStatus(error.message, "danger");
  }
});

els.clearButton.addEventListener("click", clearResults);

setRunStatus("IDLE");
