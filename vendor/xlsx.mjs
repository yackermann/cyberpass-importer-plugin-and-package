/*
 * Local, dependency-free XLSX reader used by the extension build. It exposes the
 * small SheetJS-compatible surface this project needs (XLSX.read + SheetNames /
 * Sheets). When the optional SheetJS package is installed, scripts/build.mjs
 * replaces this file in the bundle with the full `xlsx` implementation.
 */
const MAIN='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const utf8=new TextDecoder();
function u16(d,o){return d.getUint16(o,true)} function u32(d,o){return d.getUint32(o,true)}
async function inflate(bytes){ if(typeof DecompressionStream==='undefined') throw new Error('This browser does not support local XLSX decompression.'); const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw')); return new Uint8Array(await new Response(stream).arrayBuffer()); }
async function unzip(buf){
  const data=new Uint8Array(buf), d=new DataView(data.buffer), files=new Map(); let eocd=-1;
  for(let i=data.length-22;i>=Math.max(0,data.length-65558);i--) if(u32(d,i)===0x06054b50){eocd=i;break;}
  if(eocd<0) throw new Error('Not an XLSX/ZIP workbook.'); const count=u16(d,eocd+10), cdSize=u32(d,eocd+12), cdOffset=u32(d,eocd+16); let p=cdOffset;
  for(let i=0;i<count;i++){ if(u32(d,p)!==0x02014b50) break; const method=u16(d,p+10), csize=u32(d,p+20), usize=u32(d,p+24), nlen=u16(d,p+28), xlen=u16(d,p+30), clen=u16(d,p+32), off=u32(d,p+42); const name=utf8.decode(data.slice(p+46,p+46+nlen)); const ld=off; const ln=u16(d,ld+26), lx=u16(d,ld+28); const raw=data.slice(ld+30+ln+lx,ld+30+ln+lx+csize); files.set(name, method===0?raw:await inflate(raw)); p+=46+nlen+xlen+clen; }
  return files;
}
function xml(files,name){const bytes=files.get(name); return bytes?new DOMParser().parseFromString(utf8.decode(bytes),'application/xml'):null;}
function children(el,tag){return [...(el?.getElementsByTagNameNS(MAIN,tag)||[])];}
function colNum(ref){let s=(ref.match(/^[A-Z]+/i)?.[0]||'').toUpperCase(),n=0;for(const c of s)n=n*26+c.charCodeAt(0)-64;return n;}
function serialDate(n){const date=new Date(Date.UTC(1899,11,30)+Number(n)*86400000);return date.toISOString().slice(0,10);}
async function read(data){
  const files=await unzip(data); const wb=xml(files,'xl/workbook.xml'); if(!wb) throw new Error('Workbook metadata not found.');
  const rels=xml(files,'xl/_rels/workbook.xml.rels'); const relMap=new Map(); for(const r of [...(rels?.getElementsByTagName('Relationship')||[])]) relMap.set(r.getAttribute('Id'),r.getAttribute('Target'));
  const shared=xml(files,'xl/sharedStrings.xml'); const strings=shared?[...shared.getElementsByTagNameNS(MAIN,'si')].map(si=>[...si.getElementsByTagNameNS(MAIN,'t')].map(t=>t.textContent||'').join('')):[];
  const sheetNames=[], Sheets={}; const sheets=wb.getElementsByTagNameNS(MAIN,'sheet');
  for(const sheet of sheets){const name=sheet.getAttribute('name')||'Sheet'; const rid=sheet.getAttributeNS(REL,'id'); let target=relMap.get(rid)||''; if(target.startsWith('/'))target=target.slice(1); if(!target.startsWith('xl/'))target='xl/'+target.replace(/^\.\//,''); const doc=xml(files,target); const out={'!ref':'A1:A1'}; let maxR=1,maxC=1;
    for(const row of [...(doc?.getElementsByTagNameNS(MAIN,'row')||[])]){for(const cell of [...row.getElementsByTagNameNS(MAIN,'c')]){const ref=cell.getAttribute('r')||''; const t=cell.getAttribute('t')||''; const v=cell.getElementsByTagNameNS(MAIN,'v')[0]?.textContent||''; const is=cell.getElementsByTagNameNS(MAIN,'is')[0]; let value=v; if(t==='s')value=strings[Number(v)]||''; else if(t==='inlineStr')value=is?[...is.getElementsByTagNameNS(MAIN,'t')].map(x=>x.textContent||'').join(''):''; else if(t==='b')value=v==='1'?'TRUE':'FALSE'; else if(t==='str')value=v; else if(v && cell.getAttribute('s')) value=v; const c=colNum(ref),r=Number(ref.match(/\d+/)?.[0]||1); maxR=Math.max(maxR,r);maxC=Math.max(maxC,c); out[ref]={v:value,w:String(value)}; }}
    out['!ref']=`A1:${String.fromCharCode(64+Math.min(maxC,26))}${maxR}`; sheetNames.push(name);Sheets[name]=out;
  }
  return {SheetNames:sheetNames,Sheets};
}
export const XLSX={read};
export default XLSX;
