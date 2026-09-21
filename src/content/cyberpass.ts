import type { CyberPassField, CyberPassControl, ScanResult, ExcelRequirement, FillOptions, FillResult } from '../shared/types';
import { normalizeRequirementId } from '../excel/normalizer';

const REQ_RE=/\b(?:Requirement|Req)\s*#?\s*([0-9]+(?:\.[0-9]+)*)\b/i;
function clean(s:string){return s.replace(/\s+/g,' ').trim();}
function controlInfo(el: HTMLElement): CyberPassControl { const x=el as HTMLInputElement; return {tag:el.tagName.toLowerCase(),type:x.type,name:x.name,id:el.id,ariaLabel:el.getAttribute('aria-label')||undefined,value:readValue(el),selectorHint:el.id?`#${CSS.escape(el.id)}`:undefined}; }
export function readValue(el: HTMLElement):string { const x=el as HTMLInputElement; if(el instanceof HTMLSelectElement)return [...el.selectedOptions].map(o=>o.text).join(', '); if(el.isContentEditable)return el.textContent||''; return x.value||''; }
function setNativeValue(el: HTMLElement,value:string){
  if(el instanceof HTMLSelectElement){const option=[...el.options].find(o=>o.text.trim()===value.trim()||o.value===value); if(option)el.value=option.value; else {const fallback=[...el.options].find(o=>o.text.toLowerCase().includes(value.toLowerCase())); if(fallback)el.value=fallback.value;} el.dispatchEvent(new Event('change',{bubbles:true})); return;}
  if(el.isContentEditable){el.textContent=value;el.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertText',data:value}));el.dispatchEvent(new Event('change',{bubbles:true}));return;}
  const proto=el instanceof HTMLTextAreaElement?HTMLTextAreaElement.prototype:HTMLInputElement.prototype; const setter=Object.getOwnPropertyDescriptor(proto,'value')?.set; setter?.call(el,value); el.dispatchEvent(new Event('input',{bubbles:true}));el.dispatchEvent(new Event('change',{bubbles:true}));el.dispatchEvent(new Event('blur',{bubbles:true}));
}
function requirementContainer(node:HTMLElement):HTMLElement { return node.closest('.input-node-view-builder-container') as HTMLElement || node.parentElement || node; }
function findFields(): ScanResult {
  const fields:CyberPassField[]=[]; const seen=new Set<string>();
  const candidates=[...document.querySelectorAll<HTMLElement>('.input-node-view-builder-header, [data-requirement], h1,h2,h3,h4,h5,h6')];
  for(const node of candidates){const txt=clean(node.textContent||''); const m=txt.match(REQ_RE); if(!m)continue; const id=normalizeRequirementId(m[1]); if(!id||seen.has(id))continue; const container=requirementContainer(node); const controls=[...container.querySelectorAll<HTMLElement>('textarea, input:not([type="hidden"]):not([type="file"]), select, [contenteditable="true"]')]; if(!controls.length)continue; seen.add(id); const label=clean((node.nextElementSibling?.textContent||'').slice(0,160)); fields.push({requirementId:id,label,containerText:clean(container.textContent||'').slice(0,500),controls:controls.map(controlInfo)}); }
  return {fields,debug:fields.map(f=>({requirementId:f.requirementId,text:f.containerText,controls:f.controls.map(({tag,type,name,id,ariaLabel})=>({tag,type,name,id,ariaLabel}))}))};
}
function elementForField(field:CyberPassField):HTMLElement|null { const ordered=[...field.controls].sort((a,b)=>({textarea:0,contenteditable:1,select:2,input:3}[a.tag]??4)-({textarea:0,contenteditable:1,select:2,input:3}[b.tag]??4)); const c=ordered[0]; if(!c)return null; if(c.id)return document.getElementById(c.id); if(c.selectorHint){try{return document.querySelector(c.selectorHint)}catch{}} return null; }
function mark(el:HTMLElement,kind:'ok'|'warn'|'fail'){const old=el.style.outline; el.style.outline=kind==='ok'?'3px solid #2e9d62':kind==='warn'?'3px solid #d97706':'3px solid #dc2626'; el.style.outlineOffset='2px'; setTimeout(()=>{el.style.outline=old;el.style.outlineOffset='';},3500);}
export async function scanPage(){
  const fieldsById=new Map<string,CyberPassField>();
  await walkLazyRequirements(fields=>{for(const field of fields) fieldsById.set(field.requirementId,field);});
  const fields=[...fieldsById.values()];
  return {fields,debug:fields.map(field=>({requirementId:field.requirementId,text:field.containerText,controls:field.controls.map(({tag,type,name,id,ariaLabel})=>({tag,type,name,id,ariaLabel}))}))};
}

function wait(ms:number):Promise<void>{return new Promise(resolve=>setTimeout(resolve,ms));}
function scrollTargets(): HTMLElement[] {
  const targets: HTMLElement[] = [];
  const root=document.scrollingElement as HTMLElement|null;
  if(root) targets.push(root);
  for(const element of [...document.querySelectorAll<HTMLElement>('*')]) {
    if(element===root) continue;
    const style=getComputedStyle(element);
    if(element.scrollHeight-element.clientHeight>80 && /(auto|scroll)/.test(style.overflowY)) targets.push(element);
  }
  return targets;
}

/** Walk lazy-loaded scroll regions until no new requirement containers appear. */
async function walkLazyRequirements(onFields:(fields:CyberPassField[])=>void):Promise<void> {
  const positions=new Map<HTMLElement,number>();
  let stable=0;
  for(let pass=0;pass<40 && stable<2;pass++) {
    const before=findFields().fields;
    onFields(before);
    const targets=scrollTargets();
    for(const target of targets) {
      if(!positions.has(target)) positions.set(target,target.scrollTop);
      target.scrollTop=target.scrollHeight;
    }
    window.scrollTo({top:document.documentElement.scrollHeight,behavior:'auto'});
    await wait(180);
    const after=findFields().fields;
    const beforeIds=new Set(before.map(field=>field.requirementId));
    const added=after.some(field=>!beforeIds.has(field.requirementId));
    stable=added?0:stable+1;
  }
  onFields(findFields().fields);
  for(const [target,top] of positions) target.scrollTop=top;
}

function fillField(field:CyberPassField, requirement:ExcelRequirement, options:FillOptions, out:FillResult):void {
  const el=elementForField(field);
  if(!el){out.failed++;out.errors.push(`${field.requirementId}: editable control not found`);return;}
  const current=readValue(el);
  if(!requirement.value){out.skipped++;return;}
  if(current.trim()&&!options.replaceExisting){
    if(current.trim()!==requirement.value.trim()){out.mismatches++;mark(el,'warn');}
    out.skipped++;return;
  }
  try{setNativeValue(el,requirement.value);mark(el,'ok');out.filled++;}
  catch(error){out.failed++;mark(el,'fail');out.errors.push(`${field.requirementId}: ${String(error)}`);}
}

export async function preview(requirements:ExcelRequirement[]){
  const byId=new Map(requirements.map(r=>[r.requirementId,r]));
  const rowsById=new Map<string,any>();
  await walkLazyRequirements(fields=>{
    for(const field of fields){
      const requirement=byId.get(field.requirementId); const el=elementForField(field); const current=el?readValue(el):'';
      if(requirement){const status=!requirement.value?'empty':current&&current.trim()!==requirement.value.trim()?'mismatch':'matched'; rowsById.set(field.requirementId,{requirementId:field.requirementId,excel:requirement,cyberPass:field,status,currentValue:current,reason:status==='mismatch'?'Existing value differs from workbook':''}); if(status==='mismatch'&&el)mark(el,'warn');}
      else rowsById.set(field.requirementId,{requirementId:field.requirementId,cyberPass:field,status:'unmatched-page',currentValue:current});
    }
  });
  for(const requirement of requirements) if(!rowsById.has(requirement.requirementId)) rowsById.set(requirement.requirementId,{requirementId:requirement.requirementId,excel:requirement,status:'unmatched-excel'});
  const fields=findFields();
  return {scan:fields,rows:[...rowsById.values()]};
}

export async function fill(requirements:ExcelRequirement[],options:FillOptions):Promise<FillResult> {
  const byId=new Map(requirements.map(r=>[r.requirementId,r]));
  const processed=new Set<string>();
  const out:FillResult={filled:0,skipped:0,failed:0,mismatches:0,errors:[]};
  await walkLazyRequirements(fields=>{
    for(const field of fields){if(processed.has(field.requirementId))continue;const requirement=byId.get(field.requirementId);if(!requirement)continue;processed.add(field.requirementId);fillField(field,requirement,options,out);}
  });
  return out;
}
let helperRequirements = new Map<string, ExcelRequirement>();
let dynamicObserver: MutationObserver | null = null;
let refreshTimer: number | undefined;

/** Safe to call after every lazy render: existing buttons are left in place. */
export function refreshHelpers(): void {
  for (const field of findFields().fields) {
    if (!helperRequirements.has(field.requirementId)) continue;
    const el = elementForField(field);
    const requirement = helperRequirements.get(field.requirementId);
    if (!el || !requirement) continue;
    if (el instanceof HTMLTextAreaElement && !el.dataset.cyberpassCommentAutofill) {
      el.dataset.cyberpassCommentAutofill = '1';
      el.addEventListener('click', () => {
        if (!requirement.value) return;
        const current = readValue(el);
        if (current.trim() && current.trim() !== requirement.value.trim() &&
            !confirm(`Requirement ${field.requirementId} already has a comment. Replace it with the Vendor Response from Excel?`)) return;
        if (!current.trim() || current.trim() !== requirement.value.trim()) {
          setNativeValue(el, requirement.value);
          mark(el, 'ok');
        }
      });
    }
    const parent = el.parentElement;
    if (!parent || [...parent.children].some(child =>
      child instanceof HTMLElement && child.dataset.cyberpassHelperFor === field.requirementId)) continue;
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.textContent = 'Fill from Excel';
    btn.dataset.cyberpassHelperFor = field.requirementId;
    btn.style.cssText = 'margin:4px 0;padding:3px 7px;font:11px system-ui;cursor:pointer;background:#eef6ff;border:1px solid #6aa7df;border-radius:4px;';
    btn.addEventListener('click', () => {
      const requirement = helperRequirements.get(field.requirementId);
      if (!requirement?.value) return;
      const current = readValue(el);
      if (current && current.trim() !== requirement.value.trim() &&
          !confirm(`Requirement ${field.requirementId} already has content. Replace it?`)) return;
      setNativeValue(el, requirement.value);
      mark(el, 'ok');
      btn.textContent = 'Filled ✓';
    });
    parent.insertBefore(btn, el);
  }
}

export function installHelpers(requirements: ExcelRequirement[]): void {
  helperRequirements = new Map(requirements.map(requirement => [requirement.requirementId, requirement]));
  document.querySelectorAll('[data-cyberpass-helper-for]').forEach(element => element.remove());
  refreshHelpers();
  if(!dynamicObserver){
    const scheduleRefresh=()=>{
      if(refreshTimer) window.clearTimeout(refreshTimer);
      refreshTimer=window.setTimeout(()=>refreshHelpers(),120);
    };
    dynamicObserver=new MutationObserver(scheduleRefresh);
    dynamicObserver.observe(document.body,{childList:true,subtree:true});
    window.addEventListener('scroll',scheduleRefresh,{passive:true});
  }
}
export function debugDom(){return findFields().debug;}
