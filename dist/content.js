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
function answerElementForField(field) {
  const control = field.controls.find((item) => item.id?.toLowerCase().endsWith(".answer"));
  return control?.id ? document.getElementById(control.id) : null;
}
function currentAnswerChoice(answer) {
  if (answer instanceof HTMLSelectElement) return normalizedAnswerLabel(answer.selectedOptions[0]?.text || answer.value);
  const select = answer.closest(".ant-select");
  const selected = select?.querySelector(".ant-select-content, .ant-select-selection-item");
  return normalizedAnswerLabel(selected?.textContent || readValue(answer));
}
function startsWithNotApplicable(value) {
  return /^\s*n\/a\b/i.test(value);
}
function desiredAnswer(value) {
  return startsWithNotApplicable(value) ? "N/A" : value.trim() ? "YES" : "NO";
}
function normalizedAnswerLabel(value) {
  const text = value.toLowerCase().replace(/[^a-z\/]+/g, " ").trim();
  if (/^n\s*\/\s*a(?:\b|$)/.test(text)) return "N/A";
  if (/^yes\b/.test(text)) return "YES";
  if (/^no\b/.test(text)) return "NO";
  return null;
}
async function setAnswerChoice(field, choice) {
  const answer = answerElementForField(field);
  if (!answer) return false;
  if (answer instanceof HTMLSelectElement) {
    const option = [...answer.options].find((item) => normalizedAnswerLabel(item.text) === choice);
    if (!option) return false;
    answer.value = option.value;
    answer.dispatchEvent(new Event("change", { bubbles: true }));
    mark(answer, "ok");
    return true;
  }
  if (answer instanceof HTMLInputElement) {
    const select = answer.closest(".ant-select") || answer.parentElement;
    (select || answer).dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, view: window }));
    (select || answer).click();
    answer.focus();
    await wait(120);
    const options = [...document.querySelectorAll(".ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option")];
    const fallback = [...document.querySelectorAll(".ant-select-item-option")];
    const option = [...options.length ? options : fallback].find((item) => normalizedAnswerLabel(item.querySelector(".ant-select-item-option-content")?.textContent || item.textContent || "") === choice);
    if (!option) return false;
    const optionContent = option.querySelector(".ant-select-item-option-content") || option;
    optionContent.click();
    await wait(80);
    const selected = currentAnswerChoice(answer);
    if (selected !== choice) return false;
    mark(answer, "ok");
    return true;
  }
  return false;
}
async function scanPage() {
  const fieldsById = /* @__PURE__ */ new Map();
  await walkLazyRequirements((fields2) => {
    for (const field of fields2) fieldsById.set(field.requirementId, field);
  });
  const fields = [...fieldsById.values()];
  return { fields, debug: fields.map((field) => ({ requirementId: field.requirementId, text: field.containerText, controls: field.controls.map(({ tag, type, name, id, ariaLabel }) => ({ tag, type, name, id, ariaLabel })) })) };
}
function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
function scrollTargets() {
  const targets = [];
  const root = document.scrollingElement;
  if (root) targets.push(root);
  for (const element of [...document.querySelectorAll("*")]) {
    if (element === root) continue;
    const style = getComputedStyle(element);
    if (element.scrollHeight - element.clientHeight > 80 && /(auto|scroll)/.test(style.overflowY)) targets.push(element);
  }
  return targets;
}
async function walkLazyRequirements(onFields) {
  const positions = /* @__PURE__ */ new Map();
  let stable = 0;
  for (let pass = 0; pass < 40 && stable < 2; pass++) {
    const before = findFields().fields;
    await onFields(before);
    const targets = scrollTargets();
    for (const target of targets) {
      if (!positions.has(target)) positions.set(target, target.scrollTop);
      target.scrollTop = target.scrollHeight;
    }
    window.scrollTo({ top: document.documentElement.scrollHeight, behavior: "auto" });
    await wait(180);
    const after = findFields().fields;
    const beforeIds = new Set(before.map((field) => field.requirementId));
    const added = after.some((field) => !beforeIds.has(field.requirementId));
    stable = added ? 0 : stable + 1;
  }
  await onFields(findFields().fields);
  for (const [target, top] of positions) target.scrollTop = top;
}
async function fillField(field, requirement, options, out) {
  const choice = desiredAnswer(requirement.value);
  const answer = answerElementForField(field);
  if (answer) {
    const currentAnswer = currentAnswerChoice(answer);
    if (currentAnswer && currentAnswer !== choice && !options.replaceExisting) {
      out.mismatches++;
      mark(answer, "warn");
    } else if (!currentAnswer || options.replaceExisting) {
      if (!await setAnswerChoice(field, choice)) {
        out.failed++;
        out.errors.push(`${field.requirementId}: could not set Response to ${choice}`);
      }
    }
  } else {
    out.errors.push(`${field.requirementId}: Response control not found`);
  }
  const el = elementForField(field);
  if (!el) {
    out.failed++;
    out.errors.push(`${field.requirementId}: editable comment control not found`);
    return;
  }
  if (!requirement.value) {
    out.skipped++;
    return;
  }
  const current = readValue(el);
  if (current.trim() && !options.replaceExisting) {
    if (current.trim() !== requirement.value.trim()) {
      out.mismatches++;
      mark(el, "warn");
    }
    out.skipped++;
    return;
  }
  try {
    setNativeValue(el, requirement.value);
    mark(el, "ok");
    out.filled++;
  } catch (error) {
    out.failed++;
    mark(el, "fail");
    out.errors.push(`${field.requirementId}: ${String(error)}`);
  }
}
async function preview(requirements) {
  const byId = new Map(requirements.map((r) => [r.requirementId, r]));
  const rowsById = /* @__PURE__ */ new Map();
  await walkLazyRequirements((fields2) => {
    for (const field of fields2) {
      const requirement = byId.get(field.requirementId);
      const el = elementForField(field);
      const current = el ? readValue(el) : "";
      if (requirement) {
        const status = !requirement.value ? "empty" : current && current.trim() !== requirement.value.trim() ? "mismatch" : "matched";
        rowsById.set(field.requirementId, { requirementId: field.requirementId, excel: requirement, cyberPass: field, status, currentValue: current, reason: status === "mismatch" ? "Existing value differs from workbook" : "" });
        if (status === "mismatch" && el) mark(el, "warn");
      } else rowsById.set(field.requirementId, { requirementId: field.requirementId, cyberPass: field, status: "unmatched-page", currentValue: current });
    }
  });
  for (const requirement of requirements) if (!rowsById.has(requirement.requirementId)) rowsById.set(requirement.requirementId, { requirementId: requirement.requirementId, excel: requirement, status: "unmatched-excel" });
  const fields = findFields();
  return { scan: fields, rows: [...rowsById.values()] };
}
async function fill(requirements, options) {
  const byId = new Map(requirements.map((r) => [r.requirementId, r]));
  const processed = /* @__PURE__ */ new Set();
  const out = { filled: 0, skipped: 0, failed: 0, mismatches: 0, errors: [] };
  await walkLazyRequirements(async (fields) => {
    for (const field of fields) {
      if (processed.has(field.requirementId)) continue;
      const requirement = byId.get(field.requirementId);
      if (!requirement) continue;
      processed.add(field.requirementId);
      await fillField(field, requirement, options, out);
    }
  });
  return out;
}
var helperRequirements = /* @__PURE__ */ new Map();
var dynamicObserver = null;
var refreshTimer;
var commentSuggestion = null;
var commentSuggestionAnchor = null;
function hideCommentSuggestion() {
  commentSuggestion?.remove();
  commentSuggestion = null;
  commentSuggestionAnchor = null;
}
function showCommentSuggestion(el, field, requirement) {
  hideCommentSuggestion();
  if (!requirement.value || !document.body) return;
  const card = document.createElement("div");
  card.setAttribute("role", "dialog");
  card.setAttribute("aria-label", "Excel autofill suggestion");
  card.style.cssText = "position:fixed;z-index:2147483646;width:300px;max-width:calc(100vw - 24px);padding:10px;background:#fff;border:1px solid #b8c9d8;border-radius:8px;box-shadow:0 5px 18px rgba(31,52,70,.18);font:12px system-ui;color:#243746;";
  const rect = el.getBoundingClientRect();
  card.style.left = `${Math.max(12, Math.min(rect.left, window.innerWidth - 312))}px`;
  card.style.top = `${Math.min(window.innerHeight - 100, rect.bottom + 6)}px`;
  const heading = document.createElement("div");
  heading.textContent = `Requirement ${field.requirementId} \xB7 Vendor Response`;
  heading.style.cssText = "font-weight:700;margin-bottom:5px;";
  const value = document.createElement("div");
  value.textContent = requirement.value;
  value.title = requirement.value;
  value.style.cssText = "color:#5d6b77;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;margin-bottom:8px;";
  const fillButton = document.createElement("button");
  fillButton.type = "button";
  fillButton.textContent = "Fill from Excel";
  fillButton.style.cssText = "display:block;margin-left:auto;border:0;border-radius:5px;padding:6px 9px;background:#246b9f;color:#fff;font:600 11px system-ui;cursor:pointer;";
  fillButton.addEventListener("click", async () => {
    await setAnswerChoice(field, desiredAnswer(requirement.value));
    setNativeValue(el, requirement.value);
    mark(el, "ok");
    hideCommentSuggestion();
  });
  card.append(heading, value, fillButton);
  document.body.appendChild(card);
  commentSuggestion = card;
  commentSuggestionAnchor = el;
}
function refreshHelpers() {
  for (const field of findFields().fields) {
    if (!helperRequirements.has(field.requirementId)) continue;
    const el = elementForField(field);
    if (!el) continue;
    if (el instanceof HTMLTextAreaElement && !el.dataset.cyberpassCommentAutofill) {
      el.dataset.cyberpassCommentAutofill = "1";
      el.addEventListener("click", async () => {
        const requirement = helperRequirements.get(field.requirementId);
        if (!requirement?.value) return;
        const current = readValue(el);
        if (!current.trim()) showCommentSuggestion(el, field, requirement);
        else if (current.trim() !== requirement.value.trim()) {
          hideCommentSuggestion();
          if (confirm(`Requirement ${field.requirementId} already has a comment. Replace it with the Vendor Response from Excel?`)) {
            await setAnswerChoice(field, desiredAnswer(requirement.value));
            setNativeValue(el, requirement.value);
            mark(el, "ok");
          }
        }
      });
    }
    const parent = el.parentElement;
    if (!parent || [...parent.children].some((child) => child instanceof HTMLElement && child.dataset.cyberpassHelperFor === field.requirementId)) continue;
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = "Fill from Excel";
    btn.dataset.cyberpassHelperFor = field.requirementId;
    btn.style.cssText = "display:block;margin:4px 0 4px auto;padding:3px 7px;font:11px system-ui;cursor:pointer;background:#eef6ff;border:1px solid #6aa7df;border-radius:4px;";
    btn.addEventListener("click", async () => {
      const requirement = helperRequirements.get(field.requirementId);
      if (!requirement) return;
      const current = readValue(el);
      if (requirement.value && current && current.trim() !== requirement.value.trim() && !confirm(`Requirement ${field.requirementId} already has content. Replace it?`)) return;
      await setAnswerChoice(field, desiredAnswer(requirement.value));
      if (requirement.value) {
        setNativeValue(el, requirement.value);
        mark(el, "ok");
      }
      btn.textContent = requirement.value ? "Filled \u2713" : "Marked NO \u2713";
    });
    parent.insertBefore(btn, el);
  }
}
function installHelpers(requirements) {
  helperRequirements = new Map(requirements.map((requirement) => [requirement.requirementId, requirement]));
  document.querySelectorAll("[data-cyberpass-helper-for]").forEach((element) => element.remove());
  refreshHelpers();
  if (!dynamicObserver) {
    const scheduleRefresh = () => {
      if (refreshTimer) window.clearTimeout(refreshTimer);
      refreshTimer = window.setTimeout(() => refreshHelpers(), 120);
    };
    dynamicObserver = new MutationObserver(scheduleRefresh);
    dynamicObserver.observe(document.body, { childList: true, subtree: true });
    document.addEventListener("click", (event) => {
      const target = event.target;
      if (commentSuggestion && target && target !== commentSuggestionAnchor && !commentSuggestion.contains(target)) hideCommentSuggestion();
    });
    window.addEventListener("scroll", () => hideCommentSuggestion(), { passive: true });
  }
}
function debugDom() {
  return findFields().debug;
}

// src/content/content.ts
function isCyberPassQuestionnairePage() {
  if (location.origin !== "https://app.fido.cyber-pass.org") return false;
  if (!/^\/procedures\/[^/]+$/.test(location.pathname)) return false;
  return new URLSearchParams(location.search).get("step") === "fido_user_authenticator_vendor_questionnaire";
}
function applyDescriptionColorOverride(enabled) {
  const styleId = "cyberpass-importer-description-color";
  document.getElementById(styleId)?.remove();
  if (!enabled) return;
  const style = document.createElement("style");
  style.id = styleId;
  style.textContent = ".input-node-view-builder-description, .input-node-view-builder-description * { color: #000 !important; }";
  document.documentElement.appendChild(style);
}
if (!isCyberPassQuestionnairePage()) {
} else {
  chrome.storage.local.get({ overrideRequirementColour: false }, (settings) => applyDescriptionColorOverride(Boolean(settings.overrideRequirementColour)));
  async function restoreSavedWorkbookHelpers() {
    const match = location.pathname.match(/^\/procedures\/([^/]+)$/);
    if (!match) return;
    const procedureId = decodeURIComponent(match[1]);
    const response = await chrome.runtime.sendMessage({ type: "GET_WORKBOOK", procedureId });
    const workbook = response?.workbook;
    if (Array.isArray(workbook?.requirements) && workbook.requirements.length) {
      installHelpers(workbook.requirements);
    }
  }
  restoreSavedWorkbookHelpers().catch((error) => console.error("CyberPass workbook restore failed", error));
  chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
    try {
      if (message.type === "SCAN_PAGE") scanPage().then(sendResponse).catch((error) => sendResponse({ error: String(error) }));
      else if (message.type === "PREVIEW") preview(message.requirements).then(sendResponse).catch((error) => sendResponse({ error: String(error) }));
      else if (message.type === "FILL") fill(message.requirements, message.options).then(sendResponse).catch((error) => sendResponse({ error: String(error) }));
      else if (message.type === "DEBUG_DOM") sendResponse(debugDom());
      else if (message.type === "INSTALL_HELPERS") {
        installHelpers(message.requirements);
        sendResponse({ ok: true });
      } else if (message.type === "SET_DESCRIPTION_COLOR") {
        applyDescriptionColorOverride(message.enabled);
        sendResponse({ ok: true });
      }
    } catch (error) {
      sendResponse({ error: String(error) });
    }
    return true;
  });
}
