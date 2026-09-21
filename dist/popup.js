// vendor/xlsx.mjs
var MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
var REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
var utf8 = new TextDecoder();
function u16(d, o) {
  return d.getUint16(o, true);
}
function u32(d, o) {
  return d.getUint32(o, true);
}
async function inflate(bytes) {
  if (typeof DecompressionStream === "undefined") throw new Error("This browser does not support local XLSX decompression.");
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function unzip(buf) {
  const data = new Uint8Array(buf), d = new DataView(data.buffer), files = /* @__PURE__ */ new Map();
  let eocd = -1;
  for (let i = data.length - 22; i >= Math.max(0, data.length - 65558); i--) if (u32(d, i) === 101010256) {
    eocd = i;
    break;
  }
  if (eocd < 0) throw new Error("Not an XLSX/ZIP workbook.");
  const count = u16(d, eocd + 10), cdSize = u32(d, eocd + 12), cdOffset = u32(d, eocd + 16);
  let p = cdOffset;
  for (let i = 0; i < count; i++) {
    if (u32(d, p) !== 33639248) break;
    const method = u16(d, p + 10), csize = u32(d, p + 20), usize = u32(d, p + 24), nlen = u16(d, p + 28), xlen = u16(d, p + 30), clen = u16(d, p + 32), off = u32(d, p + 42);
    const name = utf8.decode(data.slice(p + 46, p + 46 + nlen));
    const ld = off;
    const ln = u16(d, ld + 26), lx = u16(d, ld + 28);
    const raw = data.slice(ld + 30 + ln + lx, ld + 30 + ln + lx + csize);
    files.set(name, method === 0 ? raw : await inflate(raw));
    p += 46 + nlen + xlen + clen;
  }
  return files;
}
function xml(files, name) {
  const bytes = files.get(name);
  return bytes ? new DOMParser().parseFromString(utf8.decode(bytes), "application/xml") : null;
}
function colNum(ref) {
  let s = (ref.match(/^[A-Z]+/i)?.[0] || "").toUpperCase(), n = 0;
  for (const c of s) n = n * 26 + c.charCodeAt(0) - 64;
  return n;
}
async function read(data) {
  const files = await unzip(data);
  const wb = xml(files, "xl/workbook.xml");
  if (!wb) throw new Error("Workbook metadata not found.");
  const rels = xml(files, "xl/_rels/workbook.xml.rels");
  const relMap = /* @__PURE__ */ new Map();
  for (const r of [...rels?.getElementsByTagName("Relationship") || []]) relMap.set(r.getAttribute("Id"), r.getAttribute("Target"));
  const shared = xml(files, "xl/sharedStrings.xml");
  const strings = shared ? [...shared.getElementsByTagNameNS(MAIN, "si")].map((si) => [...si.getElementsByTagNameNS(MAIN, "t")].map((t) => t.textContent || "").join("")) : [];
  const sheetNames = [], Sheets = {};
  const sheets = wb.getElementsByTagNameNS(MAIN, "sheet");
  for (const sheet of sheets) {
    const name = sheet.getAttribute("name") || "Sheet";
    const rid = sheet.getAttributeNS(REL, "id");
    let target = relMap.get(rid) || "";
    if (target.startsWith("/")) target = target.slice(1);
    if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.\//, "");
    const doc = xml(files, target);
    const out = { "!ref": "A1:A1" };
    let maxR = 1, maxC = 1;
    for (const row of [...doc?.getElementsByTagNameNS(MAIN, "row") || []]) {
      for (const cell of [...row.getElementsByTagNameNS(MAIN, "c")]) {
        const ref = cell.getAttribute("r") || "";
        const t = cell.getAttribute("t") || "";
        const v = cell.getElementsByTagNameNS(MAIN, "v")[0]?.textContent || "";
        const is = cell.getElementsByTagNameNS(MAIN, "is")[0];
        let value = v;
        if (t === "s") value = strings[Number(v)] || "";
        else if (t === "inlineStr") value = is ? [...is.getElementsByTagNameNS(MAIN, "t")].map((x) => x.textContent || "").join("") : "";
        else if (t === "b") value = v === "1" ? "TRUE" : "FALSE";
        else if (t === "str") value = v;
        else if (v && cell.getAttribute("s")) value = v;
        const c = colNum(ref), r = Number(ref.match(/\d+/)?.[0] || 1);
        maxR = Math.max(maxR, r);
        maxC = Math.max(maxC, c);
        out[ref] = { v: value, w: String(value) };
      }
    }
    out["!ref"] = `A1:${String.fromCharCode(64 + Math.min(maxC, 26))}${maxR}`;
    sheetNames.push(name);
    Sheets[name] = out;
  }
  return { SheetNames: sheetNames, Sheets };
}
var XLSX = { read };
var xlsx_default = XLSX;

// src/excel/normalizer.ts
function normalizeRequirementId(input2) {
  const text2 = String(input2 ?? "").trim();
  if (!text2) return null;
  const sr = text2.match(/\bSR\s*[-_:]?\s*0*(\d+(?:\.\d+)*)\b/i);
  const req = text2.match(/\b(?:requirement|req)\s*#?\s*0*(\d+(?:\.\d+)*)\b/i);
  const bare = text2.match(/^\s*0*(\d+(?:\.\d+)*)\s*$/);
  const value = sr?.[1] ?? req?.[1] ?? bare?.[1];
  if (!value) return null;
  return value.split(".").map((part) => {
    const n = Number(part);
    return Number.isFinite(n) ? String(Number(n.toPrecision(12))) : part;
  }).join(".");
}

// src/excel/parser.ts
function text(v) {
  return v == null ? "" : String(v).replace(/\u00a0/g, " ").trim();
}
function colName(ref) {
  return (ref.match(/^[A-Z]+/i)?.[0] ?? "").toUpperCase();
}
function rowNumber(ref) {
  return Number(ref.match(/\d+/)?.[0] ?? 0);
}
function sheetToRows(sheet) {
  const cells = Object.keys(sheet).filter((k) => !k.startsWith("!"));
  const maxRow = Math.max(0, ...cells.map(rowNumber));
  const maxCol = Math.max(0, ...cells.map((k) => {
    let n = 0;
    for (const c of colName(k)) n = n * 26 + c.charCodeAt(0) - 64;
    return n;
  }));
  const rows = Array.from({ length: maxRow }, () => Array(maxCol).fill(""));
  const refs = Array.from({ length: maxRow }, () => Array(maxCol).fill(""));
  for (const ref of cells) {
    const c = sheet[ref];
    let n = 0;
    for (const ch of colName(ref)) n = n * 26 + ch.charCodeAt(0) - 64;
    rows[rowNumber(ref) - 1][n - 1] = text(c?.w ?? c?.v);
    refs[rowNumber(ref) - 1][n - 1] = ref;
  }
  return { rows, refs };
}
function parseWorkbook(workbook) {
  const output = [];
  for (const sheetName of workbook.SheetNames ?? []) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) continue;
    const { rows, refs } = sheetToRows(sheet);
    if (!rows.length) continue;
    const header = rows.slice(0, Math.min(12, rows.length)).find((r) => r.some((v) => /sr\s*(no|number)|requirement\s*(id|no|number)|vendor response|response|answer/i.test(v))) ?? rows[0];
    const reqCol = header.findIndex((v) => /sr\s*(no|number)|requirement\s*(id|no|number)/i.test(v));
    const responseCol = header.findIndex((v) => /vendor\s*response|response|answer|rationale/i.test(v));
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      let raw = "";
      let c = -1;
      if (reqCol >= 0) {
        raw = text(row[reqCol]);
        c = reqCol;
      }
      if (!raw) for (let i = 0; i < row.length; i++) {
        if (normalizeRequirementId(row[i])) {
          raw = text(row[i]);
          c = i;
          break;
        }
      }
      const id = normalizeRequirementId(raw);
      if (!id || r < 1) continue;
      let value = responseCol >= 0 ? text(row[responseCol]) : "";
      if (responseCol < 0 && !value) {
        const candidates = row.map((v, i) => ({ v: text(v), i })).filter((x) => x.i !== c && x.v);
        value = candidates.length ? candidates[candidates.length - 1].v : "";
      }
      output.push({ requirementId: id, rawRequirementId: raw, value, sheetName, row: r + 1, column: refs[r]?.[c] ? colName(refs[r][c]) : void 0, responseColumn: responseCol >= 0 ? colName(refs[r]?.[responseCol] ?? "") : void 0 });
    }
  }
  const dedup = /* @__PURE__ */ new Map();
  for (const item of output) {
    const prior = dedup.get(item.requirementId);
    if (!prior || !prior.value && item.value) dedup.set(item.requirementId, item);
  }
  return [...dedup.values()];
}
async function readWorkbookFile(file) {
  const name = file.name.toLowerCase();
  if (name.endsWith(".csv") || name.endsWith(".tsv")) {
    const raw = await file.text();
    const delim = name.endsWith(".tsv") ? "	" : ",";
    const rows = raw.split(/\r?\n/).map((line) => line.split(delim));
    return parseWorkbook({ SheetNames: ["CSV"], Sheets: { CSV: rows.reduce((s, row, r) => {
      row.forEach((v, c) => {
        s[`${String.fromCharCode(65 + c)}${r + 1}`] = { v, w: v };
      });
      return s;
    }, { "!ref": `A1:${String.fromCharCode(64 + (rows[0]?.length || 1))}${rows.length}` }) } });
  }
  const xlsx = globalThis.XLSX;
  if (!xlsx?.read) throw new Error("SheetJS runtime is missing. Run the build to bundle xlsx into vendor/xlsx.mjs.");
  const workbook = await xlsx.read(await file.arrayBuffer(), { type: "array", cellText: true, cellDates: true });
  return parseWorkbook(workbook);
}

// src/popup/popup.ts
var chromeApi = globalThis.chrome;
var requirements = [];
var activeTab = null;
var $ = (id) => document.getElementById(id);
var drop = $("dropZone");
var input = $("fileInput");
var dropPrompt = $("dropPrompt");
var selectedFile = $("selectedFile");
var selectedFileName = $("selectedFileName");
var clearFile = $("clearFile");
var workbookCard = $("workbookCard");
var fileName = $("fileName");
var excelCount = $("excelCount");
var status = $("status");
var fillBtn = $("fillBtn");
var replace = $("replaceExisting");
var overrideColour = $("overrideColour");
globalThis.XLSX = xlsx_default;
function setStatus(message, kind) {
  status.textContent = message;
  status.className = `status ${kind || ""}`;
}
function setWorkbookUi(name) {
  const hasFile = Boolean(name);
  drop.classList.toggle("has-file", hasFile);
  dropPrompt.classList.toggle("hidden", hasFile);
  selectedFile.classList.toggle("hidden", !hasFile);
  workbookCard.classList.toggle("hidden", !hasFile);
  if (hasFile) selectedFileName.textContent = name;
}
async function clearWorkbookForCurrentProcedure() {
  await tab();
  const procedureId = procedureIdForUrl(activeTab?.url);
  if (!procedureId) throw new Error("Open the CyberPass vendor questionnaire URL before clearing the workbook.");
  await chromeApi.runtime.sendMessage({ type: "CLEAR_WORKBOOK", procedureId });
  await send({ type: "INSTALL_HELPERS", requirements: [] }).catch(() => {
  });
  requirements = [];
  input.value = "";
  setWorkbookUi(null);
  fileName.textContent = "";
  excelCount.textContent = "\u2014";
  fillBtn.disabled = true;
  setStatus("Workbook cleared for this procedure.", "ok");
}
async function tab() {
  const tabs = await chromeApi.tabs.query({ active: true, currentWindow: true });
  activeTab = tabs[0];
  return activeTab;
}
function procedureIdForUrl(rawUrl) {
  try {
    const url = new URL(rawUrl || "");
    const match = url.pathname.match(/^\/procedures\/([^/]+)$/);
    return match ? decodeURIComponent(match[1]) : null;
  } catch {
    return null;
  }
}
function isAllowedTab(candidate) {
  try {
    const url = new URL(candidate?.url || "");
    return url.origin === "https://app.fido.cyber-pass.org" && procedureIdForUrl(candidate.url) !== null && url.searchParams.get("step") === "fido_user_authenticator_vendor_questionnaire";
  } catch {
    return false;
  }
}
async function send(message) {
  await tab();
  if (!isAllowedTab(activeTab)) throw new Error("Open the CyberPass vendor questionnaire URL before using the importer.");
  return chromeApi.tabs.sendMessage(activeTab.id, message);
}
async function sendWorker(message) {
  return chromeApi.runtime.sendMessage(message);
}
async function load(file) {
  let procedureId = null;
  try {
    await tab();
    if (!isAllowedTab(activeTab)) throw new Error("Open the CyberPass vendor questionnaire URL before selecting a workbook.");
    procedureId = procedureIdForUrl(activeTab.url);
    if (!procedureId) throw new Error("Could not determine the CyberPass procedure ID.");
    if (/\.xlsb?$/.test(file.name.toLowerCase())) setStatus("Legacy .xls/.xlsb files need a full SheetJS build; use .xlsx or .xlsm for this bundled reader.", "error");
    requirements = await readWorkbookFile(file);
    const response = await sendWorker({ type: "STORE_WORKBOOK", procedureId, fileName: file.name, requirements });
    if (response?.error) throw new Error(response.error);
    fileName.textContent = file.name;
    setWorkbookUi(file.name);
    excelCount.textContent = String(requirements.length);
    fillBtn.disabled = !requirements.length;
    await send({ type: "INSTALL_HELPERS", requirements }).catch(() => {
    });
    setStatus(requirements.length ? `Extracted ${requirements.length} requirement responses for procedure ${procedureId}.` : "No requirement rows detected.", "ok");
  } catch (e) {
    requirements = [];
    setWorkbookUi(null);
    if (procedureId) await sendWorker({ type: "CLEAR_WORKBOOK", procedureId }).catch(() => {
    });
    fillBtn.disabled = true;
    setStatus(String(e), "error");
  }
}
function confirmPageMovement() {
  return window.confirm("The CyberPass page will scroll while the extension loads dynamically rendered requirements. You will see the page moving. Continue?");
}
fillBtn.addEventListener("click", async () => {
  if (!confirmPageMovement()) return;
  try {
    const result = await send({ type: "FILL", requirements, options: { replaceExisting: replace.checked } });
    setStatus(`Filled ${result.filled}; skipped ${result.skipped}; ${result.mismatches} mismatch warnings; ${result.failed} failed. The form was not submitted.`, result.failed ? "error" : "ok");
  } catch (e) {
    setStatus(String(e), "error");
  }
});
drop.addEventListener("click", () => input.click());
clearFile.addEventListener("click", async (event) => {
  event.stopPropagation();
  try {
    await clearWorkbookForCurrentProcedure();
  } catch (error) {
    setStatus(String(error), "error");
  }
});
input.addEventListener("change", () => {
  const f = input.files?.[0];
  if (f) load(f);
});
for (const ev of ["dragenter", "dragover"]) drop.addEventListener(ev, (e) => {
  e.preventDefault();
  drop.classList.add("drag");
});
for (const ev of ["dragleave", "drop"]) drop.addEventListener(ev, (e) => {
  e.preventDefault();
  drop.classList.remove("drag");
});
drop.addEventListener("drop", (e) => {
  const f = e.dataTransfer.files?.[0];
  if (f) load(f);
});
overrideColour.addEventListener("change", async () => {
  const enabled = overrideColour.checked;
  await chromeApi.storage.local.set({ overrideRequirementColour: enabled });
  try {
    await send({ type: "SET_DESCRIPTION_COLOR", enabled });
    setStatus(enabled ? "Requirement descriptions are now black." : "Requirement colour override disabled.", "ok");
  } catch (e) {
    setStatus(String(e), "error");
  }
});
chromeApi.storage.local.get({ overrideRequirementColour: false }, (settings) => {
  overrideColour.checked = Boolean(settings.overrideRequirementColour);
});
async function restoreWorkbookForProcedure() {
  await tab();
  if (!isAllowedTab(activeTab)) return;
  const procedureId = procedureIdForUrl(activeTab.url);
  if (!procedureId) return;
  const response = await sendWorker({ type: "GET_WORKBOOK", procedureId });
  const workbook = response?.workbook;
  if (!workbook || !Array.isArray(workbook.requirements) || !workbook.requirements.length) return;
  requirements = workbook.requirements;
  fileName.textContent = workbook.fileName || "Restored workbook";
  setWorkbookUi(workbook.fileName || "Restored workbook");
  excelCount.textContent = String(requirements.length);
  fillBtn.disabled = false;
  await send({ type: "INSTALL_HELPERS", requirements }).catch(() => {
  });
  setStatus(`Restored ${requirements.length} requirement responses for procedure ${procedureId}.`, "ok");
}
chromeApi.storage.local.remove(["workbookRequirements", "workbookFileName"]);
(async () => {
  try {
    await restoreWorkbookForProcedure();
  } catch {
  }
})();
