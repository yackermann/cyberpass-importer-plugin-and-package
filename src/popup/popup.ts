import XLSX from '../../vendor/xlsx.mjs';
import { readWorkbookFile } from '../excel/parser';
import type { ExcelRequirement } from '../shared/types';
const chromeApi:any=(globalThis as any).chrome;
let requirements:ExcelRequirement[]=[]; let activeTab:any=null;
const $=<T extends HTMLElement>(id:string)=>document.getElementById(id) as T;
const drop=$('dropZone'), input=$<HTMLInputElement>('fileInput'), dropPrompt=$('dropPrompt'), selectedFile=$('selectedFile'), selectedFileName=$('selectedFileName'), clearFile=$<HTMLButtonElement>('clearFile'), workbookCard=$('workbookCard'), fileName=$('fileName'), excelCount=$('excelCount'), status=$('status'), fillBtn=$<HTMLButtonElement>('fillBtn'), replace=$<HTMLInputElement>('replaceExisting'), overrideColour=$<HTMLInputElement>('overrideColour');
(globalThis as any).XLSX=XLSX;
function setStatus(message:string, kind?:'ok'|'error'){status.textContent=message;status.className=`status ${kind||''}`;}
function setWorkbookUi(name:string|null): void {
  const hasFile=Boolean(name);
  drop.classList.toggle('has-file',hasFile);
  dropPrompt.classList.toggle('hidden',hasFile);
  selectedFile.classList.toggle('hidden',!hasFile);
  workbookCard.classList.toggle('hidden',!hasFile);
  if(hasFile) selectedFileName.textContent=name!;
}
async function clearWorkbookForCurrentProcedure(): Promise<void> {
  await tab();
  const procedureId=procedureIdForUrl(activeTab?.url);
  if(!procedureId) throw new Error('Open the CyberPass vendor questionnaire URL before clearing the workbook.');
  await chromeApi.runtime.sendMessage({type:'CLEAR_WORKBOOK',procedureId});
  await send({type:'INSTALL_HELPERS',requirements:[]}).catch(()=>{});
  requirements=[];
  input.value='';
  setWorkbookUi(null);
  fileName.textContent='';
  excelCount.textContent='—';
  fillBtn.disabled=true;
  setStatus('Workbook cleared for this procedure.','ok');
}
async function tab(){const tabs=await chromeApi.tabs.query({active:true,currentWindow:true});activeTab=tabs[0];return activeTab;}
function procedureIdForUrl(rawUrl:any): string|null {
  try {
    const url=new URL(rawUrl||'');
    const match=url.pathname.match(/^\/procedures\/([^/]+)$/);
    return match ? decodeURIComponent(match[1]) : null;
  } catch { return null; }
}
function isAllowedTab(candidate:any): boolean {
  try {
    const url=new URL(candidate?.url||'');
    return url.origin==='https://app.fido.cyber-pass.org' && procedureIdForUrl(candidate.url)!==null && url.searchParams.get('step')==='fido_user_authenticator_vendor_questionnaire';
  } catch { return false; }
}
async function send(message:any){await tab();if(!isAllowedTab(activeTab)) throw new Error('Open the CyberPass vendor questionnaire URL before using the importer.');return chromeApi.tabs.sendMessage(activeTab.id,message);}
async function sendWorker(message:any){return chromeApi.runtime.sendMessage(message);}
async function load(file:File){let procedureId:string|null=null;try{await tab();if(!isAllowedTab(activeTab))throw new Error('Open the CyberPass vendor questionnaire URL before selecting a workbook.');procedureId=procedureIdForUrl(activeTab.url);if(!procedureId)throw new Error('Could not determine the CyberPass procedure ID.');if(/\.xlsb?$/.test(file.name.toLowerCase())) setStatus('Legacy .xls/.xlsb files need a full SheetJS build; use .xlsx or .xlsm for this bundled reader.','error');requirements=await readWorkbookFile(file);const response=await sendWorker({type:'STORE_WORKBOOK',procedureId,fileName:file.name,requirements});if(response?.error)throw new Error(response.error);fileName.textContent=file.name;setWorkbookUi(file.name);excelCount.textContent=String(requirements.length);fillBtn.disabled=!requirements.length;await send({type:'INSTALL_HELPERS',requirements}).catch(()=>{});setStatus(requirements.length?`Extracted ${requirements.length} requirement responses for procedure ${procedureId}.`:'No requirement rows detected.','ok');}catch(e){requirements=[];setWorkbookUi(null);if(procedureId)await sendWorker({type:'CLEAR_WORKBOOK',procedureId}).catch(()=>{});fillBtn.disabled=true;setStatus(String(e),'error');}}
function confirmPageMovement(): boolean {
  return window.confirm('The CyberPass page will scroll while the extension loads dynamically rendered requirements. You will see the page moving. Continue?');
}
fillBtn.addEventListener('click',async()=>{if(!confirmPageMovement())return;try{const result=await send({type:'FILL',requirements,options:{replaceExisting:replace.checked}});setStatus(`Filled ${result.filled}; skipped ${result.skipped}; ${result.mismatches} mismatch warnings; ${result.failed} failed. The form was not submitted.`,result.failed?'error':'ok');}catch(e){setStatus(String(e),'error');}});
drop.addEventListener('click',()=>input.click());
clearFile.addEventListener('click',async(event)=>{event.stopPropagation();try{await clearWorkbookForCurrentProcedure();}catch(error){setStatus(String(error),'error');}});input.addEventListener('change',()=>{const f=input.files?.[0];if(f)load(f)});for(const ev of ['dragenter','dragover'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')});for(const ev of ['dragleave','drop'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')});drop.addEventListener('drop',(e:any)=>{const f=e.dataTransfer.files?.[0];if(f)load(f)});overrideColour.addEventListener('change',async()=>{const enabled=overrideColour.checked;await chromeApi.storage.local.set({overrideRequirementColour:enabled});try{await send({type:'SET_DESCRIPTION_COLOR',enabled});setStatus(enabled?'Requirement descriptions are now black.':'Requirement colour override disabled.','ok');}catch(e){setStatus(String(e),'error');}});
chromeApi.storage.local.get({overrideRequirementColour:false},(settings:any)=>{overrideColour.checked=Boolean(settings.overrideRequirementColour);});
async function restoreWorkbookForProcedure(){
  await tab();
  if(!isAllowedTab(activeTab))return;
  const procedureId=procedureIdForUrl(activeTab.url);
  if(!procedureId)return;
  const response=await sendWorker({type:'GET_WORKBOOK',procedureId});
  const workbook=response?.workbook;
  if(!workbook||!Array.isArray(workbook.requirements)||!workbook.requirements.length)return;
  requirements=workbook.requirements;
  fileName.textContent=workbook.fileName||'Restored workbook';
  setWorkbookUi(workbook.fileName||'Restored workbook');
  excelCount.textContent=String(requirements.length);
  fillBtn.disabled=false;
  await send({type:'INSTALL_HELPERS',requirements}).catch(()=>{});
  setStatus(`Restored ${requirements.length} requirement responses for procedure ${procedureId}.`, 'ok');
}
chromeApi.storage.local.remove(['workbookRequirements','workbookFileName']);
(async()=>{try{await restoreWorkbookForProcedure();}catch{}})();
