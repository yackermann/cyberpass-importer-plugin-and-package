// src/excel/normalizer.ts
function normalizeRequirementId(input) {
  const text = String(input ?? "").trim();
  if (!text) return null;
  const sr = text.match(/\bSR\s*[-_:]?\s*0*(\d+(?:\.\d+)*)\b/i);
  const req = text.match(/\b(?:requirement|req)\s*#?\s*0*(\d+(?:\.\d+)*)\b/i);
  const bare = text.match(/^\s*0*(\d+(?:\.\d+)*)\s*$/);
  const value = sr?.[1] ?? req?.[1] ?? bare?.[1];
  if (!value) return null;
  return value.split(".").map((part) => {
    const n = Number(part);
    return Number.isFinite(n) ? String(Number(n.toPrecision(12))) : part;
  }).join(".");
}

// src/content/cyberpass.ts
var REQ_RE = /\b(?:Requirement|Req)\s*#?\s*([0-9]+(?:\.[0-9]+)*)\b/i;
function clean(s) {
  return s.replace(/\s+/g, " ").trim();
}
function controlInfo(el) {
  const x = el;
  return { tag: el.tagName.toLowerCase(), type: x.type, name: x.name, id: el.id, ariaLabel: el.getAttribute("aria-label") || void 0, value: readValue(el), selectorHint: el.id ? `#${CSS.escape(el.id)}` : void 0 };
}
function readValue(el) {
  const x = el;
  if (el instanceof HTMLSelectElement) return [...el.selectedOptions].map((o) => o.text).join(", ");
  if (el.isContentEditable) return el.textContent || "";
  return x.value || "";
}
function setNativeValue(el, value) {
  if (el instanceof HTMLSelectElement) {
    const option = [...el.options].find((o) => o.text.trim() === value.trim() || o.value === value);
    if (option) el.value = option.value;
    else {
      const fallback = [...el.options].find((o) => o.text.toLowerCase().includes(value.toLowerCase()));
      if (fallback) el.value = fallback.value;
    }
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  if (el.isContentEditable) {
    el.textContent = value;
    el.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: value }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
  const setter = Object.getOwnPropertyDescriptor(proto, "value")?.set;
  setter?.call(el, value);
  el.dispatchEvent(new Event("input", { bubbles: true }));
  el.dispatchEvent(new Event("change", { bubbles: true }));
  el.dispatchEvent(new Event("blur", { bubbles: true }));
}
function requirementContainer(node) {
  return node.closest(".input-node-view-builder-container") || node.parentElement || node;
}
function findFields() {
  const fields = [];
  const seen = /* @__PURE__ */ new Set();
  const candidates = [...document.querySelectorAll(".input-node-view-builder-header, [data-requirement], h1,h2,h3,h4,h5,h6")];
  for (const node of candidates) {
    const txt = clean(node.textContent || "");
    const m = txt.match(REQ_RE);
    if (!m) continue;
    const id = normalizeRequirementId(m[1]);
    if (!id || seen.has(id)) continue;
    const container = requirementContainer(node);
    const controls = [...container.querySelectorAll('textarea, input:not([type="hidden"]):not([type="file"]), select, [contenteditable="true"]')];
    if (!controls.length) continue;
    seen.add(id);
    const label = clean((node.nextElementSibling?.textContent || "").slice(0, 160));
    fields.push({ requirementId: id, label, containerText: clean(container.textContent || "").slice(0, 500), controls: controls.map(controlInfo) });
  }
  return { fields, debug: fields.map((f) => ({ requirementId: f.requirementId, text: f.containerText, controls: f.controls.map(({ tag, type, name, id, ariaLabel }) => ({ tag, type, name, id, ariaLabel })) })) };
}
function elementForField(field) {
  const ordered = [...field.controls].sort((a, b) => ({ textarea: 0, contenteditable: 1, select: 2, input: 3 }[a.tag] ?? 4) - ({ textarea: 0, contenteditable: 1, select: 2, input: 3 }[b.tag] ?? 4));
  const c = ordered[0];
  if (!c) return null;
  if (c.id) return document.getElementById(c.id);
  if (c.selectorHint) {
    try {
      return document.querySelector(c.selectorHint);
    } catch {
    }
  }
  return null;
}
function mark(el, kind) {
  const old = el.style.outline;
  el.style.outline = kind === "ok" ? "3px solid #2e9d62" : kind === "warn" ? "3px solid #d97706" : "3px solid #dc2626";
  el.style.outlineOffset = "2px";
  setTimeout(() => {
    el.style.outline = old;
    el.style.outlineOffset = "";
  }, 3500);
}
function scanPage() {
  return findFields();
}
function preview(requirements) {
  const scan = findFields();
  const byId = new Map(requirements.map((r) => [r.requirementId, r]));
  const rows = [];
  for (const f of scan.fields) {
    const r = byId.get(f.requirementId);
    const el = elementForField(f);
    const current = el ? readValue(el) : "";
    if (r) {
      const status = !r.value ? "empty" : current && current.trim() !== r.value.trim() ? "mismatch" : "matched";
      rows.push({ requirementId: f.requirementId, excel: r, cyberPass: f, status, currentValue: current, reason: status === "mismatch" ? "Existing value differs from workbook" : "" });
      if (status === "mismatch" && el) mark(el, "warn");
    } else rows.push({ requirementId: f.requirementId, cyberPass: f, status: "unmatched-page", currentValue: current });
  }
  for (const r of requirements) if (!scan.fields.some((f) => f.requirementId === r.requirementId)) rows.push({ requirementId: r.requirementId, excel: r, status: "unmatched-excel" });
  return { scan, rows };
}
function fill(requirements, options) {
  const scan = findFields();
  const byId = new Map(requirements.map((r) => [r.requirementId, r]));
  const out = { filled: 0, skipped: 0, failed: 0, mismatches: 0, errors: [] };
  for (const f of scan.fields) {
    const r = byId.get(f.requirementId);
    if (!r) continue;
    const el = elementForField(f);
    if (!el) {
      out.failed++;
      out.errors.push(`${f.requirementId}: editable control not found`);
      continue;
    }
    const current = readValue(el);
    if (!r.value) {
      out.skipped++;
      continue;
    }
    if (current.trim() && !options.replaceExisting) {
      if (current.trim() !== r.value.trim()) {
        out.mismatches++;
        mark(el, "warn");
      }
      out.skipped++;
      continue;
    }
    try {
      setNativeValue(el, r.value);
      mark(el, "ok");
      out.filled++;
    } catch (e) {
      out.failed++;
      mark(el, "fail");
      out.errors.push(`${f.requirementId}: ${String(e)}`);
    }
  }
  return out;
}
function installHelpers(requirements) {
  const byId = new Map(requirements.map((r) => [r.requirementId, r]));
  document.querySelectorAll("[data-cyberpass-helper]").forEach((e) => e.remove());
  for (const f of findFields().fields) {
    const r = byId.get(f.requirementId);
    if (!r) continue;
    const el = elementForField(f);
    if (!el) continue;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = "Fill from Excel";
    btn.dataset.cyberpassHelper = "1";
    btn.style.cssText = "margin:4px 0;padding:3px 7px;font:11px system-ui;cursor:pointer;background:#eef6ff;border:1px solid #6aa7df;border-radius:4px;";
    btn.addEventListener("click", () => {
      const current = readValue(el);
      if (current && current.trim() !== r.value.trim() && !confirm(`Requirement ${f.requirementId} already has content. Replace it?`)) return;
      setNativeValue(el, r.value);
      mark(el, "ok");
      btn.textContent = "Filled \u2713";
    });
    el.parentElement?.insertBefore(btn, el);
  }
}
function debugDom() {
  return findFields().debug;
}

// src/content/content.ts
chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  try {
    if (message.type === "SCAN_PAGE") sendResponse(scanPage());
    else if (message.type === "PREVIEW") sendResponse(preview(message.requirements));
    else if (message.type === "FILL") sendResponse(fill(message.requirements, message.options));
    else if (message.type === "DEBUG_DOM") sendResponse(debugDom());
    else if (message.type === "INSTALL_HELPERS") {
      installHelpers(message.requirements);
      sendResponse({ ok: true });
    }
  } catch (error) {
    sendResponse({ error: String(error) });
  }
  return true;
});
