const DEFAULT_CONFIG = {
  soldCodeColumn: "รหัสสินค้า",
  vrichMatchColumn: "รหัสขาย",
  jstMatchColumn: "รหัสรูปแบบ",
  vrichQtyColumn: "จำนวน",
  jstQtyColumn: "จำนวน",
};

const STATUS_LABELS = {
  IDLE: "รอประมวลผล",
  PASS: "PASS: พร้อมตรวจขั้นสุดท้ายก่อนนำเข้า",
  PASS_WITH_EXCLUSION: "PASS_WITH_EXCLUSION: ผ่านแบบมีรหัสที่ผู้ใช้ยืนยันให้ข้าม",
  FAIL: "FAIL: ยังไม่ควรนำเข้า",
  FAIL_DUPLICATE: "FAIL_DUPLICATE: พบข้อมูลซ้ำที่กระทบรายการอัปเดต",
  BLOCKED: "BLOCKED: ต้องแก้ข้อมูลก่อนใช้งาน",
};

const FILE_SIZE_WARNING_MB = 30;
const FILE_SIZE_BLOCK_MB = 80;
const LARGE_FILE_ERROR_MESSAGE =
  "ไฟล์ Excel ใหญ่เกินกว่าที่ browser จะอ่านได้ในรอบเดียว กรุณาใช้ไฟล์ vRich แบบไม่มีรูปภาพ/ลดขนาดไฟล์ หรือใช้เวอร์ชัน Python สำหรับไฟล์ขนาดใหญ่";

const state = {
  files: { vrich: null, jst: null, sold: null },
  tables: null,
  result: null,
  excludedCodes: new Set(),
  downloads: [],
  isBusy: false,
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
  loadingOverlay: document.getElementById("loadingOverlay"),
  loadingTitle: document.getElementById("loadingTitle"),
  loadingMessage: document.getElementById("loadingMessage"),
  catTypingLottie: document.getElementById("catTypingLottie"),
};

const busyButtons = [els.inspectButton, els.processButton, els.clearButton, els.applyExclusionsButton];
busyButtons.forEach((button) => {
  button.dataset.defaultText = button.textContent;
});

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

function waitForPaint() {
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

function updateLoading(message, title = "กำลังทำงาน...") {
  els.loadingTitle.textContent = title;
  els.loadingMessage.textContent = message;
}

function setBusy(isBusy, options = {}) {
  state.isBusy = isBusy;
  els.loadingOverlay.classList.toggle("hidden", !isBusy);
  document.body.classList.toggle("is-loading", isBusy);
  if (isBusy) {
    updateLoading(options.message || "กรุณารอสักครู่", options.title || "กำลังทำงาน...");
  }
  for (const button of busyButtons) {
    button.disabled = isBusy || (button === els.processButton && !state.tables);
    button.textContent = button.dataset.defaultText;
  }
  if (isBusy && options.button) {
    options.button.textContent = options.busyText || options.button.dataset.defaultText;
  }
}

async function runLoadingTask(options, task) {
  setBusy(true, options);
  try {
    await waitForPaint();
    await task();
  } catch (error) {
    setStatus(normalizeWorkbookError(error).message, "danger");
  } finally {
    setBusy(false);
  }
}

function setRunStatus(status) {
  els.runStatus.textContent = STATUS_LABELS[status] || status;
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
  label.textContent = file ? `${file.name} (${formatFileSize(file.size)})` : "ยังไม่ได้เลือกไฟล์";
}

function fileSizeMb(file) {
  return file.size / (1024 * 1024);
}

function formatFileSize(bytes) {
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

function selectedFiles() {
  return [
    { label: "ไฟล์ vRich", file: state.files.vrich },
    { label: "ไฟล์ JST", file: state.files.jst },
    { label: "ไฟล์ sold_today", file: state.files.sold },
  ].filter((item) => item.file);
}

function validateFileSizes() {
  const files = selectedFiles();
  const blocked = files.filter((item) => fileSizeMb(item.file) > FILE_SIZE_BLOCK_MB);
  const warnings = files.filter((item) => fileSizeMb(item.file) > FILE_SIZE_WARNING_MB && fileSizeMb(item.file) <= FILE_SIZE_BLOCK_MB);
  if (blocked.length) {
    const names = blocked.map((item) => `${item.label}: ${item.file.name} (${formatFileSize(item.file.size)})`).join(", ");
    throw new Error(`${LARGE_FILE_ERROR_MESSAGE}\nไฟล์ที่ใหญ่เกิน 80 MB: ${names}`);
  }
  return warnings;
}

function isAllocationError(error) {
  const message = String(error && (error.message || error));
  return /array buffer allocation failed|out of memory|cannot allocate|allocation failed|invalid array length/i.test(message);
}

function normalizeWorkbookError(error) {
  if (isAllocationError(error)) return new Error(LARGE_FILE_ERROR_MESSAGE);
  return error;
}

async function readWorkbook(file, label) {
  try {
    if (label) {
      updateLoading(`กำลังอ่านไฟล์ ${label}...`, "กำลังตรวจสอบไฟล์");
      await waitForPaint();
    }
    const buffer = await file.arrayBuffer();
    const workbookLabel = label ? `จากไฟล์ ${label}` : "";
    updateLoading(`กำลังแปลงข้อมูล Excel ${workbookLabel}...`, "กำลังตรวจสอบไฟล์");
    await waitForPaint();
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false, raw: true });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
    const headerRow = matrix.find((row) => row.some((cell) => normalizeHeader(cell)));
    if (!headerRow) throw new Error(`ไม่พบแถวหัวคอลัมน์ในไฟล์ ${file.name}`);
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
  } catch (error) {
    throw normalizeWorkbookError(error);
  }
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

async function inspectFiles() {
  const config = getConfig();
  if (!state.files.vrich || !state.files.jst || !state.files.sold) {
    throw new Error("กรุณาเลือกไฟล์ให้ครบ 3 ไฟล์");
  }
  updateLoading("กำลังตรวจสอบขนาดไฟล์และเตรียมอ่านข้อมูล...", "กำลังตรวจสอบไฟล์");
  await waitForPaint();
  const sizeWarnings = validateFileSizes();

  const vrich = await readWorkbook(state.files.vrich, "vRich");
  const jst = await readWorkbook(state.files.jst, "JST");
  const sold = await readWorkbook(state.files.sold, "รหัสที่ต้องอัปเดต");
  updateLoading("กำลังตรวจคอลัมน์ที่จำเป็น...", "กำลังตรวจสอบไฟล์");
  await waitForPaint();
  requireColumns(vrich, [config.vrichMatchColumn, config.vrichQtyColumn], "vRich");
  requireColumns(jst, [config.jstMatchColumn, config.jstQtyColumn], "JST");
  requireColumns(sold, [config.soldCodeColumn], "sold_today");
  state.tables = { vrich, jst, sold };
  renderPreflight();
  els.processButton.disabled = false;
  if (sizeWarnings.length) {
    const names = sizeWarnings.map((item) => `${item.label}: ${item.file.name} (${formatFileSize(item.file.size)})`).join(", ");
    setStatus(`ตรวจสอบไฟล์ผ่าน แต่มีไฟล์ใหญ่กว่า ${FILE_SIZE_WARNING_MB} MB: ${names} หาก browser ช้าหรือค้าง ให้ใช้ไฟล์แบบไม่มีรูปภาพหรือใช้เวอร์ชัน Python`, "warn");
  } else {
    setStatus("ตรวจสอบไฟล์ผ่าน สามารถประมวลผลต่อได้", "ok");
  }
}

function renderPreflight() {
  if (!state.tables) {
    els.preflightGrid.innerHTML = "";
    return;
  }
  const tableEntries = [
    ["ไฟล์ vRich master", state.tables.vrich],
    ["ไฟล์ JST", state.tables.jst],
    ["ไฟล์ sold_today", state.tables.sold],
  ];
  els.preflightGrid.innerHTML = tableEntries
    .map(([label, table]) => {
      return `
        <article class="file-card">
          <h3>${escapeHtml(label)}</h3>
          <dl>
            <dt>ไฟล์</dt><dd>${escapeHtml(table.file.name)}</dd>
            <dt>ขนาด</dt><dd>${formatFileSize(table.file.size)}</dd>
            <dt>ชื่อชีต</dt><dd>${escapeHtml(table.sheetName)}</dd>
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

  updateLoading("กำลังสรุปผลและสร้างไฟล์ผลลัพธ์...", "กำลังสร้างไฟล์ผลลัพธ์");
  renderSummary();
  renderMissing();
  buildDownloads();
  renderDownloads();
  setRunStatus(status);
  setStatus(
    status === "PASS" || status === "PASS_WITH_EXCLUSION" ? STATUS_LABELS[status] : STATUS_LABELS[status] || "ยังไม่ควร import",
    status === "FAIL" || status === "FAIL_DUPLICATE" ? "danger" : "ok"
  );
}

function renderSummary() {
  const result = state.result;
  if (!result) {
    els.summaryGrid.innerHTML = "";
    return;
  }
  const metrics = [
    ["จำนวนรหัสใน sold_today", result.soldCodes.length],
    ["พบใน vRich", result.soldCodes.length - result.missingVrich.length],
    ["พบใน JST", result.soldCodes.length - result.missingJst.length],
    ["อัปเดตสำเร็จ", result.updateRows.length],
    ["ไม่พบใน vRich", result.remainingMissingVrich.length],
    ["ไม่พบใน JST", result.remainingMissingJst.length],
    ["รหัสซ้ำใน vRich", result.duplicateVrich.length],
    ["รหัสซ้ำใน JST", result.duplicateJst.length],
    ["รหัสที่ผู้ใช้ยืนยันให้ข้าม", result.excluded.length],
    ["สถานะ", STATUS_LABELS[result.status] || result.status],
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
          <th>ข้าม</th>
          <th>รหัส</th>
          <th>ไม่พบใน vRich</th>
          <th>ไม่พบใน JST</th>
          <th>หมายเหตุ</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map((row) => {
            const checked = state.excludedCodes.has(row.code) ? "checked" : "";
            const note = state.excludedCodes.has(row.code) ? "ผู้ใช้ยืนยันให้ข้ามรหัสนี้" : "";
            return `
              <tr>
                <td><input type="checkbox" data-exclude-code="${escapeHtml(row.code)}" ${checked}></td>
                <td>${escapeHtml(row.code)}</td>
                <td>${row.missingVrich ? "ใช่" : "ไม่ใช่"}</td>
                <td>${row.missingJst ? "ใช่" : "ไม่ใช่"}</td>
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
    { รายการ: "สถานะ", ค่า: result.status },
    { รายการ: "คำอธิบายสถานะ", ค่า: STATUS_LABELS[result.status] || result.status },
    { รายการ: "รูปแบบการทำงาน", ค่า: "ใช้ไฟล์ Excel เท่านั้น ประมวลผลในเบราว์เซอร์ ไม่ใช้ API ไม่มี backend ไม่อัปโหลดไฟล์ และไม่นำเข้า vRich เอง" },
    { รายการ: "SOLD_TODAY_CODE_COLUMN", ค่า: config.soldCodeColumn },
    { รายการ: "VRICH_MATCH_COLUMN", ค่า: config.vrichMatchColumn },
    { รายการ: "JST_MATCH_COLUMN", ค่า: config.jstMatchColumn },
    { รายการ: "VRICH_QTY_COLUMN", ค่า: config.vrichQtyColumn },
    { รายการ: "JST_QTY_COLUMN", ค่า: config.jstQtyColumn },
    { รายการ: "จำนวนรหัสใน sold_today", ค่า: result.soldCodes.length },
    { รายการ: "พบใน vRich", ค่า: result.soldCodes.length - result.missingVrich.length },
    { รายการ: "พบใน JST", ค่า: result.soldCodes.length - result.missingJst.length },
    { รายการ: "อัปเดตสำเร็จ", ค่า: result.updateRows.length },
    { รายการ: "ไม่พบใน vRich หลังข้ามรายการ", ค่า: result.remainingMissingVrich.length },
    { รายการ: "ไม่พบใน JST หลังข้ามรายการ", ค่า: result.remainingMissingJst.length },
    { รายการ: "รหัสซ้ำใน vRich ที่ชน sold_today", ค่า: result.duplicateVrich.length },
    { รายการ: "รหัสซ้ำใน JST ที่ชน sold_today", ค่า: result.duplicateJst.length },
    { รายการ: "รหัสที่ผู้ใช้ยืนยันให้ข้าม", ค่า: result.excluded.join(", ") || "-" },
  ];

  const excludedRows = result.excluded.map((code) => ({
    "รหัสที่ข้าม": code,
    "เหตุผล": "ผู้ใช้ยืนยันให้ข้าม เพราะไม่พบใน JST",
  }));

  const duplicateVrichRows = result.duplicateVrich.map((item) => ({ "รหัส": item.code, "จำนวนรายการซ้ำ": item.count }));
  const duplicateJstRows = result.duplicateJst.map((item) => ({ "รหัส": item.code, "จำนวนรายการซ้ำ": item.count }));

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
  const downloadLabels = {
    "vrich_import_update_qty.xlsx": "ไฟล์สำหรับนำเข้า vRich",
    "summary_report.xlsx": "รายงานสรุปผล",
    "report_missing_in_vrich.xlsx": "รายงานรหัสที่ไม่พบใน vRich",
    "report_missing_in_jst.xlsx": "รายงานรหัสที่ไม่พบใน JST",
    "report_duplicate_vrich.xlsx": "รายงานรหัสซ้ำใน vRich",
    "report_duplicate_jst.xlsx": "รายงานรหัสซ้ำใน JST",
    "excluded_codes_report.xlsx": "รายงานรหัสที่ผู้ใช้ยืนยันให้ข้าม",
  };
  els.downloadList.innerHTML = state.downloads
    .map((item) => {
      const label = downloadLabels[item.filename] || "ดาวน์โหลดไฟล์";
      return `<a href="${item.url}" download="${escapeHtml(item.filename)}"><span>${escapeHtml(label)}</span><small>${escapeHtml(item.filename)}</small></a>`;
    })
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
  runLoadingTask(
    {
      button: els.inspectButton,
      busyText: "กำลังตรวจสอบ...",
      title: "กำลังตรวจสอบไฟล์",
      message: "กำลังตรวจสอบไฟล์ กรุณารอสักครู่...",
    },
    inspectFiles
  );
});

els.processButton.addEventListener("click", () => {
  runLoadingTask(
    {
      button: els.processButton,
      busyText: "กำลังประมวลผล...",
      title: "กำลังประมวลผล",
      message: "กำลังประมวลผลข้อมูล...",
    },
    async () => {
      updateLoading("กำลังประมวลผล...", "กำลังประมวลผล");
      await waitForPaint();
      processData();
    }
  );
});

els.applyExclusionsButton.addEventListener("click", () => {
  runLoadingTask(
    {
      button: els.applyExclusionsButton,
      busyText: "กำลังยืนยัน...",
      title: "กำลังประมวลผล",
      message: "กำลังประมวลผลหลังยืนยันรหัสที่ข้าม...",
    },
    async () => {
      const checks = els.missingTableWrap.querySelectorAll("[data-exclude-code]");
      state.excludedCodes.clear();
      checks.forEach((check) => {
        if (check.checked) state.excludedCodes.add(check.dataset.excludeCode);
      });
      await waitForPaint();
      processData();
    }
  );
});

els.clearButton.addEventListener("click", clearResults);

els.downloadList.addEventListener("click", (event) => {
  const link = event.target.closest("a");
  if (!link) return;
  setBusy(true, {
    title: "กำลังดาวน์โหลดไฟล์",
    message: "กำลังเตรียมดาวน์โหลดไฟล์ผลลัพธ์...",
  });
  setTimeout(() => setBusy(false), 450);
});

setRunStatus("IDLE");
