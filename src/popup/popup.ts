import XLSX from '../../vendor/xlsx.mjs';
import { readWorkbookFile } from '../excel/parser';
import type { ExcelRequirement, MappingRow } from '../shared/types';
const chromeApi:any=(globalThis as any).chrome;
let requirements:ExcelRequirement[]=[]; let lastPreview:any=null; let activeTab:any=null;
const $=<T extends HTMLElement>(id:string)=>document.getElementById(id) as T;
const drop=$('dropZone'), input=$<HTMLInputElement>('fileInput'), fileName=$('fileName'), excelCount=$('excelCount'), pageCount=$('pageCount'), status=$('status'), scanBtn=$<HTMLButtonElement>('scanBtn'), previewBtn=$<HTMLButtonElement>('previewBtn'), fillBtn=$<HTMLButtonElement>('fillBtn'), replace=$<HTMLInputElement>('replaceExisting'), overrideColour=$<HTMLInputElement>('overrideColour'), preview=$('preview'), previewRows=$('previewRows'), summary=$('summary'), exportBtn=$<HTMLButtonElement>('exportBtn');
(globalThis as any).XLSX=XLSX;
function setStatus(message:string, kind?:'ok'|'error'){status.textContent=message;status.className=`status ${kind||''}`;}
async function tab(){const tabs=await chromeApi.tabs.query({active:true,currentWindow:true});activeTab=tabs[0];return activeTab;}
function isAllowedTab(candidate:any): boolean {
  try {
    const url=new URL(candidate?.url||'');
    return url.origin==='https://app.fido.cyber-pass.org' && /^\/procedures\/[^/]+$/.test(url.pathname) && url.searchParams.get('step')==='fido_user_authenticator_vendor_questionnaire';
  } catch { return false; }
}
async function send(message:any){await tab();if(!isAllowedTab(activeTab)) throw new Error('Open the CyberPass vendor questionnaire URL before using the importer.');return chromeApi.tabs.sendMessage(activeTab.id,message);}
async function load(file:File){try{if(/\.xlsb?$/.test(file.name.toLowerCase())) setStatus('Legacy .xls/.xlsb files need a full SheetJS build; use .xlsx or .xlsm for this bundled reader.','error'); requirements=await readWorkbookFile(file);await chromeApi.storage.local.set({workbookRequirements:requirements,workbookFileName:file.name});fileName.textContent=file.name;excelCount.textContent=String(requirements.length);previewBtn.disabled=!requirements.length;fillBtn.disabled=!requirements.length;setStatus(requirements.length?`Extracted ${requirements.length} requirement responses locally. Review the mapping before filling.`:'No requirement rows detected.','ok');}catch(e){requirements=[];await chromeApi.storage.local.remove(['workbookRequirements','workbookFileName']);previewBtn.disabled=true;fillBtn.disabled=true;setStatus(String(e),'error');}}
function renderRows(rows:MappingRow[]){previewRows.innerHTML='';for(const row of rows.slice(0,120)){const d=document.createElement('div');d.className='mapping-row';const value=row.excel?.value||row.currentValue||'—';const label=row.status==='matched'?'ready':row.status==='mismatch'?'warning':row.status.includes('unmatched')?'unmatched':row.status;const cls=row.status==='mismatch'?'warn':row.status.includes('unmatched')?'bad':'';d.innerHTML=`<span class="rid">${row.requirementId}</span><span class="val" title="${escapeHtml(value)}">${escapeHtml(value)}</span><span class="pill ${cls}">${label}</span>`;previewRows.appendChild(d);}if(rows.length>120){const more=document.createElement('div');more.style.cssText='font-size:10px;color:#778; padding:5px';more.textContent=`Showing 120 of ${rows.length} rows`;previewRows.appendChild(more);}}
function escapeHtml(value:string){return value.replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]||c));}
async function doPreview(){try{const result=await send({type:'PREVIEW',requirements});lastPreview=result;preview.classList.remove('hidden');renderRows(result.rows);const matched=result.rows.filter((r:any)=>r.status==='matched'||r.status==='mismatch').length;const unmatchedExcel=result.rows.filter((r:any)=>r.status==='unmatched-excel').length;const mismatch=result.rows.filter((r:any)=>r.status==='mismatch').length;pageCount.textContent=String(result.scan.fields.length);summary.classList.remove('hidden');summary.textContent=`${requirements.length} workbook rows · ${result.scan.fields.length} page requirements · ${matched} mapped · ${unmatchedExcel} workbook-only · ${mismatch} existing value warning`;setStatus('Preview ready. Warnings are highlighted on the CyberPass page; nothing has been written.','ok');await send({type:'INSTALL_HELPERS',requirements});}catch(e){setStatus(`Could not inspect the active tab: ${String(e)}`,'error');}}
scanBtn.addEventListener('click',async()=>{try{const result=await send({type:'SCAN_PAGE'});pageCount.textContent=String(result.fields.length);setStatus(`Found ${result.fields.length} CyberPass requirements. Preview mapping to compare values.`,'ok');}catch(e){setStatus('Open the CyberPass assessment tab before scanning.','error');}});
previewBtn.addEventListener('click',doPreview);
fillBtn.addEventListener('click',async()=>{try{const result=await send({type:'FILL',requirements,options:{replaceExisting:replace.checked}});setStatus(`Filled ${result.filled}; skipped ${result.skipped}; ${result.mismatches} mismatch warnings; ${result.failed} failed. The form was not submitted.`,result.failed?'error':'ok');if(lastPreview)await doPreview();}catch(e){setStatus(String(e),'error');}});
exportBtn.addEventListener('click',()=>{const payload={exportedAt:new Date().toISOString(),workbook:requirements.map(r=>({...r,value:r.value.slice(0,500)})),page:lastPreview?.scan?.debug||[]};const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}));a.download='cyberpass-debug.json';a.click();setStatus('Sanitized debug information exported locally.','ok');});
drop.addEventListener('click',()=>input.click());input.addEventListener('change',()=>{const f=input.files?.[0];if(f)load(f)});for(const ev of ['dragenter','dragover'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')});for(const ev of ['dragleave','drop'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')});drop.addEventListener('drop',(e:any)=>{const f=e.dataTransfer.files?.[0];if(f)load(f)});overrideColour.addEventListener('change',async()=>{const enabled=overrideColour.checked;await chromeApi.storage.local.set({overrideRequirementColour:enabled});try{await send({type:'SET_DESCRIPTION_COLOR',enabled});setStatus(enabled?'Requirement descriptions are now black.':'Requirement colour override disabled.','ok');}catch(e){setStatus(String(e),'error');}});
chromeApi.storage.local.get({overrideRequirementColour:false},(settings:any)=>{overrideColour.checked=Boolean(settings.overrideRequirementColour);});
chromeApi.storage.local.get({workbookRequirements:[],workbookFileName:''},(saved:any)=>{
  if(!Array.isArray(saved.workbookRequirements)||!saved.workbookRequirements.length)return;
  requirements=saved.workbookRequirements;
  fileName.textContent=saved.workbookFileName||'Restored workbook';
  excelCount.textContent=String(requirements.length);
  previewBtn.disabled=false;
  fillBtn.disabled=false;
  setStatus(`Restored ${requirements.length} requirement responses from local storage.`, 'ok');
});
tab().catch(()=>{});
