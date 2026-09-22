import { readWorkbookFile } from '../excel/parser';
import type { ExcelRequirement } from '../shared/types';
const chromeApi:any=(globalThis as any).chrome;
let requirements:ExcelRequirement[]=[]; let activeTab:any=null;
const $=<T extends HTMLElement>(id:string)=>document.getElementById(id) as T;
const drop=$('dropZone'), input=$<HTMLInputElement>('fileInput'), dropPrompt=$('dropPrompt'), selectedFile=$('selectedFile'), selectedFileName=$('selectedFileName'), clearFile=$<HTMLButtonElement>('clearFile'), workbookCard=$('workbookCard'), fileName=$('fileName'), excelCount=$('excelCount'), status=$('status'), fillBtn=$<HTMLButtonElement>('fillBtn'), replace=$<HTMLInputElement>('replaceExisting'), overrideColour=$<HTMLInputElement>('overrideColour');
function errorText(error: unknown): string {
  if (error instanceof Error) return error.message;
  if (typeof error === 'string') return error;
  try { return JSON.stringify(error); } catch { return 'Unknown extension error'; }
}
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
  try { await send({type:'INSTALL_HELPERS',requirements:[]}); } catch(error) { setStatus(`Workbook cleared, but page helpers could not be removed: ${errorText(error)}`,'error'); }
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
async function send(message:any){await tab();if(!isAllowedTab(activeTab)) throw new Error('Open the CyberPass vendor questionnaire URL before using the importer.');const response=await chromeApi.tabs.sendMessage(activeTab.id,message);if(response?.error)throw new Error(response.error);return response;}
async function sendWorker(message:any){const response=await chromeApi.runtime.sendMessage(message);if(response?.error)throw new Error(response.error);return response;}
async function load(file:File){
  let procedureId:string|null=null;
  try {
    await tab();
    if(!isAllowedTab(activeTab)) throw new Error('Open the CyberPass vendor questionnaire URL before selecting a workbook.');
    procedureId=procedureIdForUrl(activeTab.url);
    if(!procedureId) throw new Error('Could not determine the CyberPass procedure ID.');
    const parsed=await readWorkbookFile(file);
    if(!parsed.length) throw new Error('No requirement rows detected. Check that the workbook contains a populated SR No. and Vendor Response column.');
    const response=await sendWorker({type:'STORE_WORKBOOK',procedureId,fileName:file.name,requirements:parsed});
    if(response?.error) throw new Error(response.error);
    requirements=parsed;
    fileName.textContent=file.name;
    setWorkbookUi(file.name);
    excelCount.textContent=String(requirements.length);
    fillBtn.disabled=false;
    try {
      await send({type:'INSTALL_HELPERS',requirements});
      setStatus(`Loaded ${requirements.length} requirement responses for procedure ${procedureId}.`,'ok');
    } catch(error) {
      setStatus(`Workbook saved, but page helpers could not be installed: ${errorText(error)}`,'error');
    }
  } catch(error) {
    requirements=[];
    setWorkbookUi(null);
    fillBtn.disabled=true;
    setStatus(errorText(error),'error');
  }
}
function confirmPageMovement(): boolean {
  return window.confirm('The CyberPass page will scroll while the extension loads dynamically rendered requirements. You will see the page moving. Continue?');
}
fillBtn.addEventListener('click',async()=>{if(!confirmPageMovement())return;try{const result=await send({type:'FILL',requirements,options:{replaceExisting:replace.checked}});if(!result||typeof result.filled!=='number')throw new Error('The page did not return a valid fill result.');const detail=result.errors?.length?` ${result.errors.slice(0,3).join(' | ')}`:'';setStatus(`Filled ${result.filled}; skipped ${result.skipped}; ${result.mismatches} mismatch warnings; ${result.failed} failed.${detail}`,result.failed?'error':'ok');}catch(e){setStatus(errorText(e),'error');}});
drop.addEventListener('click',()=>input.click());
clearFile.addEventListener('click',async(event)=>{event.stopPropagation();try{await clearWorkbookForCurrentProcedure();}catch(error){setStatus(errorText(error),'error');}});input.addEventListener('change',()=>{const f=input.files?.[0];if(f)load(f)});for(const ev of ['dragenter','dragover'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')});for(const ev of ['dragleave','drop'])drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')});drop.addEventListener('drop',(e:any)=>{const f=e.dataTransfer.files?.[0];if(f)load(f)});overrideColour.addEventListener('change',async()=>{const enabled=overrideColour.checked;await chromeApi.storage.local.set({overrideRequirementColour:enabled});try{await send({type:'SET_DESCRIPTION_COLOR',enabled});setStatus(enabled?'Requirement descriptions are now black.':'Requirement colour override disabled.','ok');}catch(e){setStatus(errorText(e),'error');}});
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
  try { await send({type:'INSTALL_HELPERS',requirements}); } catch(error) { setStatus(`Workbook restored, but page helpers could not be installed: ${errorText(error)}`,'error'); return; }
  setStatus(`Restored ${requirements.length} requirement responses for procedure ${procedureId}.`, 'ok');
}
chromeApi.storage.local.remove(['workbookRequirements','workbookFileName']);
(async()=>{try{await restoreWorkbookForProcedure();}catch(error){setStatus(`Could not restore this procedure's workbook: ${errorText(error)}`,'error');}})();
