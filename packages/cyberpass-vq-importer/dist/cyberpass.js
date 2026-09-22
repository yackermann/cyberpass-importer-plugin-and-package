import { normalizeRequirementId } from './normalizer.js';
import { responseChoice } from './workbook.js';
const REQUIREMENT_PATTERN = /\b(?:Requirement|Req)\s*#?\s*([0-9]+(?:\.[0-9]+)*)\b/i;
const clean = (value) => value.replace(/\s+/g, ' ').trim();
const wait = (milliseconds) => new Promise(resolve => setTimeout(resolve, milliseconds));
function readValue(element) {
    const input = element;
    if (element instanceof HTMLSelectElement)
        return [...element.selectedOptions].map(option => option.text).join(', ');
    if (element.isContentEditable)
        return element.textContent ?? '';
    return input.value ?? '';
}
function controlInfo(element) {
    const input = element;
    return { tag: element.tagName.toLowerCase(), type: input.type, name: input.name, id: element.id, ariaLabel: element.getAttribute('aria-label') ?? undefined, value: readValue(element) };
}
function containerFor(header) {
    return header.closest('.input-node-view-builder-container') ?? header.parentElement ?? header;
}
function findFields(document) {
    const fields = [];
    const seen = new Set();
    const candidates = [...document.querySelectorAll('.input-node-view-builder-header, [data-requirement], h1,h2,h3,h4,h5,h6')];
    for (const header of candidates) {
        const match = clean(header.textContent ?? '').match(REQUIREMENT_PATTERN);
        if (!match)
            continue;
        const requirementId = normalizeRequirementId(match[1]);
        if (!requirementId || seen.has(requirementId))
            continue;
        const container = containerFor(header);
        const controls = [...container.querySelectorAll('textarea, input:not([type="hidden"]):not([type="file"]), select, [contenteditable="true"]')];
        if (!controls.length)
            continue;
        seen.add(requirementId);
        fields.push({ requirementId, label: clean(header.nextElementSibling?.textContent ?? '').slice(0, 160), containerText: clean(container.textContent ?? '').slice(0, 500), controls: controls.map(controlInfo) });
    }
    return fields;
}
function elementForField(document, field) {
    const controls = [...field.controls].sort((a, b) => ({ textarea: 0, contenteditable: 1, select: 2, input: 3 }[a.tag] ?? 4) - ({ textarea: 0, contenteditable: 1, select: 2, input: 3 }[b.tag] ?? 4));
    const control = controls[0];
    return control?.id ? document.getElementById(control.id) : null;
}
function answerForField(document, field) {
    const control = field.controls.find(item => item.id?.toLowerCase().endsWith('.answer'));
    return control?.id ? document.getElementById(control.id) : null;
}
function normalizedAnswerLabel(value) {
    const text = value.toLowerCase().replace(/[^a-z/]+/g, ' ').trim();
    if (/^n\s*\/\s*a(?:\b|$)/.test(text))
        return 'N/A';
    if (/^yes\b/.test(text))
        return 'YES';
    if (/^no\b/.test(text))
        return 'NO';
    return null;
}
function currentAnswer(answer) {
    if (answer instanceof HTMLSelectElement)
        return normalizedAnswerLabel(answer.selectedOptions[0]?.text ?? answer.value);
    const select = answer.closest('.ant-select');
    return normalizedAnswerLabel(select?.querySelector('.ant-select-content, .ant-select-selection-item')?.textContent ?? readValue(answer));
}
function setValue(element, value) {
    if (element.isContentEditable) {
        element.textContent = value;
        element.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: value }));
        element.dispatchEvent(new Event('change', { bubbles: true }));
        return;
    }
    const prototype = element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(prototype, 'value')?.set?.call(element, value);
    element.dispatchEvent(new Event('input', { bubbles: true }));
    element.dispatchEvent(new Event('change', { bubbles: true }));
    element.dispatchEvent(new Event('blur', { bubbles: true }));
}
async function chooseAnswer(document, field, choice) {
    const answer = answerForField(document, field);
    if (!answer)
        return false;
    if (answer instanceof HTMLSelectElement) {
        const option = [...answer.options].find(item => normalizedAnswerLabel(item.text) === choice);
        if (!option)
            return false;
        answer.value = option.value;
        answer.dispatchEvent(new Event('change', { bubbles: true }));
        return true;
    }
    if (!(answer instanceof HTMLInputElement))
        return false;
    const select = answer.closest('.ant-select') ?? answer.parentElement;
    (select ?? answer).dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: document.defaultView }));
    (select ?? answer).click();
    answer.focus();
    await wait(120);
    const visible = [...document.querySelectorAll('.ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option')];
    const options = visible.length ? visible : [...document.querySelectorAll('.ant-select-item-option')];
    const option = options.find(item => normalizedAnswerLabel(item.querySelector('.ant-select-item-option-content')?.textContent ?? item.textContent ?? '') === choice);
    if (!option)
        return false;
    (option.querySelector('.ant-select-item-option-content') ?? option).click();
    await wait(80);
    return currentAnswer(answer) === choice;
}
async function walkLazy(document, callback) {
    const view = document.defaultView;
    if (!view)
        return callback(findFields(document));
    let stable = 0;
    for (let pass = 0; pass < 40 && stable < 2; pass++) {
        const before = findFields(document);
        await callback(before);
        const root = document.scrollingElement;
        const targets = [root, ...[...document.querySelectorAll('*')].filter(element => element !== root && element.scrollHeight - element.clientHeight > 80 && /(auto|scroll)/.test(view.getComputedStyle(element).overflowY))];
        for (const target of targets)
            if (target)
                target.scrollTop = target.scrollHeight;
        view.scrollTo({ top: document.documentElement.scrollHeight, behavior: 'auto' });
        await wait(180);
        const after = findFields(document);
        const ids = new Set(before.map(field => field.requirementId));
        stable = after.some(field => !ids.has(field.requirementId)) ? 0 : stable + 1;
    }
    await callback(findFields(document));
}
function mark(element, kind) {
    const old = element.style.outline;
    element.style.outline = kind === 'ok' ? '3px solid #2e9d62' : kind === 'warn' ? '3px solid #d97706' : '3px solid #dc2626';
    element.style.outlineOffset = '2px';
    setTimeout(() => { element.style.outline = old; element.style.outlineOffset = ''; }, 3500);
}
export class CyberPassVqImporter {
    document;
    constructor(document = globalThis.document) {
        this.document = document;
    }
    async scan(options = {}) {
        const fields = new Map();
        if (options.discoverLazyRequirements === false) {
            for (const field of findFields(this.document))
                fields.set(field.requirementId, field);
        }
        else {
            await walkLazy(this.document, current => { for (const field of current)
                fields.set(field.requirementId, field); });
        }
        return { fields: [...fields.values()] };
    }
    async fill(requirements, options = {}) {
        const byId = new Map(requirements.map(requirement => [requirement.requirementId, requirement]));
        const processed = new Set();
        const result = { filled: 0, skipped: 0, failed: 0, mismatches: 0, errors: [] };
        const process = async (fields) => {
            for (const field of fields) {
                if (processed.has(field.requirementId))
                    continue;
                const requirement = byId.get(field.requirementId);
                if (!requirement)
                    continue;
                processed.add(field.requirementId);
                const choice = responseChoice(requirement.value);
                const answer = answerForField(this.document, field);
                if (!answer)
                    result.errors.push(`${field.requirementId}: Response control not found`);
                else {
                    const existing = currentAnswer(answer);
                    if (existing && existing !== choice && !options.replaceExisting) {
                        result.mismatches++;
                        mark(answer, 'warn');
                    }
                    else if (!existing || options.replaceExisting) {
                        if (await chooseAnswer(this.document, field, choice))
                            mark(answer, 'ok');
                        else {
                            result.failed++;
                            result.errors.push(`${field.requirementId}: could not set Response to ${choice}`);
                        }
                    }
                }
                const element = elementForField(this.document, field);
                if (!element) {
                    result.failed++;
                    result.errors.push(`${field.requirementId}: editable comment control not found`);
                    continue;
                }
                if (!requirement.value) {
                    result.skipped++;
                    continue;
                }
                const existingValue = readValue(element);
                if (existingValue.trim() && !options.replaceExisting) {
                    if (existingValue.trim() !== requirement.value.trim()) {
                        result.mismatches++;
                        mark(element, 'warn');
                    }
                    result.skipped++;
                    continue;
                }
                try {
                    setValue(element, requirement.value);
                    mark(element, 'ok');
                    result.filled++;
                }
                catch (error) {
                    result.failed++;
                    mark(element, 'fail');
                    result.errors.push(`${field.requirementId}: ${String(error)}`);
                }
            }
        };
        if (options.discoverLazyRequirements === false)
            await process(findFields(this.document));
        else
            await walkLazy(this.document, process);
        return result;
    }
}
