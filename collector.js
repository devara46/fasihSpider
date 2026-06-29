'use strict';
// ────────────────────────────────────────────────────────────────────────────
// CRC32
// ────────────────────────────────────────────────────────────────────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=c&1?0xEDB88320^(c>>>1):c>>>1;t[i]=c;}
  return t;
})();
function crc32start() { return 0xFFFFFFFF; }
function crc32feed(state, buf) { let c=state; for(const b of buf) c=CRC_TABLE[(c^b)&0xFF]^(c>>>8); return c; }
function crc32end(state) { return (state^0xFFFFFFFF)>>>0; }
function concatU8(arrays){const n=arrays.reduce((s,a)=>s+a.length,0),out=new Uint8Array(n);let p=0;for(const a of arrays){out.set(a,p);p+=a.length;}return out;}

// ────────────────────────────────────────────────────────────────────────────
// ZIP writer with true DEFLATE streaming (no large in-memory accumulation).
// addCompressed() accepts already-compressed data + pre-computed CRC/sizes.
// ────────────────────────────────────────────────────────────────────────────
class ZipWriter {
  constructor(){this._entries=[];this._pos=0;}

  // Add a small file (compresses internally, fine for small files)
  async addFile(name,content){
    const enc=new TextEncoder(),nb=enc.encode(name);
    const raw=typeof content==='string'?enc.encode(content):content;
    const crc=crc32end(crc32feed(crc32start(),raw)),rawLen=raw.length;
    let db=raw,method=0;
    try{const c=await deflateRaw(raw);if(c.length<rawLen*0.95){db=c;method=8;}}catch(_){}
    this._push(nb,db,rawLen,crc,method);
  }

  // Add a file whose data was already compressed externally (for large worksheets)
  addCompressed(name,compData,rawLen,crc,method=8){
    const nb=new TextEncoder().encode(name);
    this._push(nb,compData,rawLen,crc,method);
  }

  _push(nb,db,rawLen,crc,method){
    const compLen=db.length,off=this._pos;
    const lfh=new Uint8Array(30+nb.length),lv=new DataView(lfh.buffer);
    lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);
    lv.setUint16(8,method,true);
    lv.setUint32(14,crc,true);lv.setUint32(18,compLen,true);lv.setUint32(22,rawLen,true);
    lv.setUint16(26,nb.length,true);lfh.set(nb,30);
    this._entries.push({lfh,db,nb,crc,rawLen,compLen,method,off});
    this._pos+=lfh.length+compLen;
  }

  finalize(){
    const parts=[],cdir=[];
    for(const e of this._entries){
      parts.push(e.lfh,e.db);
      const cde=new Uint8Array(46+e.nb.length),cv=new DataView(cde.buffer);
      cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);
      cv.setUint16(10,e.method,true);
      cv.setUint32(16,e.crc,true);cv.setUint32(20,e.compLen,true);cv.setUint32(24,e.rawLen,true);
      cv.setUint16(28,e.nb.length,true);cv.setUint32(42,e.off,true);cde.set(e.nb,46);
      cdir.push(cde);
    }
    const cd=concatU8(cdir),eo=new Uint8Array(22),ev=new DataView(eo.buffer);
    ev.setUint32(0,0x06054b50,true);ev.setUint16(8,this._entries.length,true);ev.setUint16(10,this._entries.length,true);
    ev.setUint32(12,cd.length,true);ev.setUint32(16,this._pos,true);
    return concatU8([...parts,cd,eo]);
  }
}

async function deflateRaw(data){
  const cs=new CompressionStream('deflate-raw'),w=cs.writable.getWriter(),r=cs.readable.getReader();
  w.write(data);w.close();
  const chunks=[];
  while(true){const{done,value}=await r.read();if(done)break;chunks.push(value);}
  return concatU8(chunks);
}

// Stream XML string parts through CompressionStream, computing CRC incrementally.
// Returns { compressed: Uint8Array, rawLen, crc } — no large in-memory XML string.
async function streamCompress(xmlPartGen) {
  const cs=new CompressionStream('deflate-raw');
  const w=cs.writable.getWriter(), r=cs.readable.getReader();
  const enc=new TextEncoder();
  const outChunks=[];
  let crcState=crc32start(), rawLen=0;

  // Read compressed output concurrently
  const reading=(async()=>{while(true){const{done,value}=await r.read();if(done)break;outChunks.push(value);}})();

  for await (const xml of xmlPartGen) {
    if(exportCancelled)break;
    const bytes=enc.encode(xml);
    crcState=crc32feed(crcState,bytes);
    rawLen+=bytes.length;
    await w.write(bytes);
    await sleep(0); // yield to browser between parts
  }
  await w.close();
  await reading;
  return{compressed:concatU8(outChunks),rawLen,crc:crc32end(crcState)};
}

// ────────────────────────────────────────────────────────────────────────────
// XLSX static content (small files, compressed via addFile)
// ────────────────────────────────────────────────────────────────────────────
const xmlEsc=s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
function colLetter(i){let s='';i++;while(i>0){const r=(i-1)%26;s=String.fromCharCode(65+r)+s;i=Math.floor((i-1)/26);}return s;}
const XLSX_MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

async function addXLSXStatics(zip, sheetName) {
  await zip.addFile('[Content_Types].xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>`);
  await zip.addFile('_rels/.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  await zip.addFile('xl/workbook.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${xmlEsc(sheetName)}" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  await zip.addFile('xl/_rels/workbook.xml.rels',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  await zip.addFile('xl/styles.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><sz val="11"/><name val="Calibri"/><b/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>`);
}

// Build one XLSX file from a row array (used for T1 which stays in-memory)
async function buildXLSXFromArray(sheetName, headers, rows, onProgress) {
  const zip=new ZipWriter();
  await addXLSXStatics(zip, sheetName);
  const {compressed,rawLen,crc}=await streamCompress((async function*(){
    let hdr=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">`;
    headers.forEach((h,ci)=>hdr+=`<c r="${colLetter(ci)}1" t="inlineStr" s="1"><is><t>${xmlEsc(h)}</t></is></c>`);
    yield hdr+'</row>';
    const CHUNK=500;
    for(let i=0;i<rows.length;i+=CHUNK){
      let xml='';
      const end=Math.min(i+CHUNK,rows.length);
      for(let ri=i;ri<end;ri++){
        const rn=ri+2;
        xml+=`<row r="${rn}">`;
        headers.forEach((h,ci)=>xml+=`<c r="${colLetter(ci)}${rn}" t="inlineStr"><is><t>${xmlEsc(rows[ri][h]??'')}</t></is></c>`);
        xml+='</row>';
      }
      if(onProgress)onProgress(end,rows.length);
      yield xml;
    }
    yield'</sheetData></worksheet>';
  })());
  zip.addCompressed('xl/worksheets/sheet1.xml',compressed,rawLen,crc);
  return zip.finalize();
}

// ────────────────────────────────────────────────────────────────────────────
// IndexedDB — streaming store (write during collection, read during export)
// ────────────────────────────────────────────────────────────────────────────
const DB = {
  _db: null,
  STORES: ['t2_results','t3_results','t2_errors','t3_errors','t6_results'],

  async open() {
    if(this._db)return;
    await new Promise((res,rej)=>{
      // Bumped to 2 so onupgradeneeded fires again on existing installs and
      // creates the new t6_results store alongside the originals.
      const req=indexedDB.open('fasih_collector',2);
      req.onupgradeneeded=e=>{
        const db=e.target.result;
        for(const s of this.STORES)if(!db.objectStoreNames.contains(s))db.createObjectStore(s,{autoIncrement:true,keyPath:'_iid'});
      };
      req.onsuccess=e=>{this._db=e.target.result;res();};
      req.onerror=e=>rej(e.target.error);
    });
  },

  async add(store,obj){
    return new Promise((res,rej)=>{
      const tx=this._db.transaction(store,'readwrite'),r=tx.objectStore(store).add(obj);
      r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);
    });
  },

  async count(store){
    return new Promise((res,rej)=>{
      const tx=this._db.transaction(store,'readonly'),r=tx.objectStore(store).count();
      r.onsuccess=()=>res(r.result);r.onerror=()=>rej(r.error);
    });
  },

  async clear(store){
    return new Promise((res,rej)=>{
      const tx=this._db.transaction(store,'readwrite'),r=tx.objectStore(store).clear();
      r.onsuccess=()=>res();r.onerror=()=>rej(r.error);
    });
  },

  // Stream records through callback(chunk[]). Returns total records read.
  async stream(store, chunkSize, onChunk) {
    return new Promise((res,rej)=>{
      const tx=this._db.transaction(store,'readonly');
      const cursor=tx.objectStore(store).openCursor();
      let chunk=[],total=0,pending=Promise.resolve();
      cursor.onsuccess=e=>{
        const c=e.target.result;
        if(c){
          chunk.push(c.value);total++;
          if(chunk.length>=chunkSize){const b=chunk;chunk=[];pending=pending.then(()=>onChunk(b));}
          c.continue();
        } else {
          if(chunk.length)pending=pending.then(()=>onChunk(chunk));
          pending.then(()=>res(total)).catch(rej);
        }
      };
      cursor.onerror=()=>rej(cursor.error);
    });
  },

  // Build an XLSX from the store, auto-splitting every splitAt rows.
  // Downloads each split immediately so memory is freed between files.
  async exportXLSX(store, sheetName, headers, rowMapper, filenameBase, splitAt=50000, onProgress) {
    const total=await this.count(store);
    const splits=Math.ceil(total/splitAt)||1;
    let fileNum=0,batchRows=[],rowsDone=0;

    const flushFile=async()=>{
      if(!batchRows.length)return;
      fileNum++;
      const label=splits>1?`_part${fileNum}of${splits}`:'';
      if(onProgress)onProgress(rowsDone,total,fileNum-1,splits,`Building file ${fileNum}/${splits}...`);
      const zip=new ZipWriter();
      await addXLSXStatics(zip,`${sheetName}${splits>1?` (${fileNum}/${splits})`:''}`)
      const rows=batchRows.splice(0);
      const{compressed,rawLen,crc}=await streamCompress((async function*(){
        let hdr=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">`;
        headers.forEach((h,ci)=>hdr+=`<c r="${colLetter(ci)}1" t="inlineStr" s="1"><is><t>${xmlEsc(h)}</t></is></c>`);
        yield hdr+'</row>';
        const CHUNK=500;
        for(let i=0;i<rows.length;i+=CHUNK){
          let xml='';
          const end=Math.min(i+CHUNK,rows.length);
          for(let ri=i;ri<end;ri++){
            const rn=ri+2;xml+=`<row r="${rn}">`;
            headers.forEach((h,ci)=>xml+=`<c r="${colLetter(ci)}${rn}" t="inlineStr"><is><t>${xmlEsc(rows[ri][h]??'')}</t></is></c>`);
            xml+='</row>';
          }
          yield xml;
        }
        yield'</sheetData></worksheet>';
      })());
      zip.addCompressed('xl/worksheets/sheet1.xml',compressed,rawLen,crc);
      if(!exportCancelled){
        const fn=`${filenameBase}${label}_${nowTs()}.xlsx`;
        downloadBlob(new Blob([zip.finalize()],{type:XLSX_MIME}),fn);
        await sleep(200);// brief gap between downloads
      }
    };

    await this.stream(store,500,async chunk=>{
      for(const rec of chunk){if(exportCancelled)return;batchRows.push(rowMapper(rec));}
      rowsDone+=chunk.length;
      if(onProgress)onProgress(rowsDone,total,fileNum,splits);
      if(batchRows.length>=splitAt)await flushFile();
    });
    if(!exportCancelled)await flushFile();
  },

  // Stream CSV directly to disk via File System Access API.
  async exportCSV(store, headers, rowMapper, filename) {
    let handle;
    try{handle=await window.showSaveFilePicker({suggestedName:filename,types:[{description:'CSV',accept:{'text/csv':['.csv']}}]});}
    catch(e){if(e.name==='AbortError')return 0;throw e;}
    const writable=await handle.createWritable();
    await writable.write('\uFEFF'+headers.map(h=>`"${h}"`).join(',')+'\r\n');
    let done=0,total=await this.count(store);
    await this.stream(store,1000,async chunk=>{
      if(exportCancelled){return;}
      let csv='';
      for(const rec of chunk){
        const row=rowMapper(rec);
        csv+=headers.map(h=>`"${String(row[h]??'').replace(/"/g,'""')}"`).join(',')+'\r\n';
        done++;
      }
      await writable.write(csv);
      if(updateExportProgress)updateExportProgress(done,total,0,1);
      await sleep(0);
    });
    await writable.close();
    return done;
  },

  // Full rows (not just the key field) — used to drive a retry-in-place run.
  async getAll(store){
    const rows=[];
    await this.stream(store,1000,async chunk=>{rows.push(...chunk);});
    return rows;
  },

  // Export all IDs/codes from a store to a simple CSV for retry
  async exportErrorCSV(store, keyField, filename) {
    const rows=[];
    await this.stream(store,1000,async chunk=>{rows.push(...chunk.map(r=>r[keyField]).filter(Boolean));});
    if(!rows.length)return;
    const csv='\uFEFF'+keyField+'\r\n'+rows.map(v=>`"${v}"`).join('\r\n');
    downloadBlob(new Blob([csv],{type:'text/csv;charset=utf-8'}),filename);
    return rows.length;
  },
};

// ────────────────────────────────────────────────────────────────────────────
// Utilities
// ────────────────────────────────────────────────────────────────────────────
const sleep=ms=>new Promise(r=>setTimeout(r,ms));
function nowTs(){return new Date().toISOString().replace(/[:.]/g,'-').slice(0,19);}
function escHtml(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function downloadBlob(blob,filename){const url=URL.createObjectURL(blob),a=Object.assign(document.createElement('a'),{href:url,download:filename});document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(url);a.remove();},1500);}
function buildCSV(headers,rows){const l=[headers.join(',')];for(const r of rows)l.push(headers.map(h=>`"${String(r[h]??'').replace(/"/g,'""')}"`).join(','));return'\uFEFF'+l.join('\r\n');}

// ────────────────────────────────────────────────────────────────────────────
// Export progress overlay
// ────────────────────────────────────────────────────────────────────────────
let exportCancelled=false;
function showExportProgress(label='Exporting...'){exportCancelled=false;document.getElementById('exp-overlay').style.display='flex';document.getElementById('exp-label').textContent=label;document.getElementById('exp-detail').textContent='Starting...';document.getElementById('exp-bar').style.width='0%';}
function hideExportProgress(){document.getElementById('exp-overlay').style.display='none';}
function updateExportProgress(done,total,si=0,totalFiles=1,extraLabel=''){
  if(exportCancelled)return;
  const pct=total>0?Math.min(100,Math.round(done/total*100)):0;
  document.getElementById('exp-detail').textContent=(totalFiles>1?`File ${si+1}/${totalFiles} — `:'')+(extraLabel||`${done.toLocaleString()} / ${total.toLocaleString()} rows (${pct}%)`);
  document.getElementById('exp-bar').style.width=pct+'%';
}

// ────────────────────────────────────────────────────────────────────────────
// WorkQueue
// ────────────────────────────────────────────────────────────────────────────
class WorkQueue{constructor(s){this.slots=s;this.active=0;this.pending=[];}run(fn){return new Promise((res,rej)=>{this.pending.push({fn,res,rej});this._drain();});}_drain(){while(this.active<this.slots&&this.pending.length){const{fn,res,rej}=this.pending.shift();this.active++;fn().then(v=>{this.active--;res(v);this._drain();}).catch(e=>{this.active--;rej(e);this._drain();});}}}

// ────────────────────────────────────────────────────────────────────────────
// Answer extractor + file parser
// ────────────────────────────────────────────────────────────────────────────
function extractAnswer(answer){
  if(answer===null||answer===undefined)return'';
  if(typeof answer==='boolean')return answer?'true':'false';
  if(typeof answer==='number')return String(answer);
  if(typeof answer==='string')return answer;
  if(Array.isArray(answer)){
    if(!answer.length)return'';
    const f=answer[0];
    if(f===null||typeof f!=='object')return answer.map(a=>String(a??'')).join('; ');
    if('signature'in f||(('type'in f)&&String(f.type).startsWith('image/')))return'[signature]';
    if('filename'in f&&('url'in f||'uri'in f))return answer.map(a=>a.url||a.filename||'').join('; ');
    if('label'in f)return answer.map(a=>String(a.label??a.value??'')).join('; ');
    return JSON.stringify(answer);
  }
  if(typeof answer==='object')return JSON.stringify(answer);
  return String(answer);
}

async function parseFileToIds(file){
  const text=await file.text(),ext=file.name.split('.').pop().toLowerCase();
  try{
    if(ext==='json'){
      let data=JSON.parse(text);
      if(data&&!Array.isArray(data)){for(const k of['data','results','rows','items']){if(Array.isArray(data[k])){data=data[k];break;}}}
      if(!Array.isArray(data))data=[data];if(!data.length)return[];
      if(typeof data[0]==='string')return data.filter(Boolean);
      const s=data[0],key=Object.keys(s).find(k=>/assignment.?id/i.test(k))||Object.keys(s).find(k=>/full.?code/i.test(k))||Object.keys(s)[0];
      return data.map(r=>String(r[key]??'')).filter(Boolean);
    }
    if(ext==='csv'){
      const lines=text.replace(/\r/g,'').split('\n').filter(l=>l.trim());if(!lines.length)return[];
      const hdr=lines[0].split(',').map(h=>h.trim().replace(/^"|"$/g,''));
      let col=hdr.findIndex(h=>/assignment.?id/i.test(h));if(col<0)col=hdr.findIndex(h=>/full.?code/i.test(h));if(col<0)col=hdr.findIndex(h=>/code/i.test(h));if(col<0)col=0;
      return lines.slice(1).map(l=>{const c=l.split(',').map(c=>c.trim().replace(/^"|"$/g,''));return c[col];}).filter(Boolean);
    }
    return text.replace(/\r/g,'').split('\n').map(l=>l.trim()).filter(Boolean);
  }catch(e){throw new Error(`File parse error: ${e.message}`);}
}

// ────────────────────────────────────────────────────────────────────────────
// TAB 1: Region Hierarchy (in-memory, manageable size)
// ────────────────────────────────────────────────────────────────────────────
const T1={
  GROUP_ID:'a45adac1-e711-4c15-b3f9-1f30fc151565',
  BASE:'https://fasih-sm.bps.go.id/app/api/region/api/v1/region',
  PARAM:{3:'level2FullCode',4:'level3FullCode',5:'level4FullCode',6:'level5FullCode'},
  NAMES:{3:'Kecamatan',4:'Desa/Kel',5:'SLS',6:'Sub-SLS'},
  collecting:false,stopped:false,rows:[],errorItems:[],wq:null,timer:null,startMs:0,
  stats:{req:0,rows:0,err:0,active:0},lcount:{3:0,4:0,5:0,6:0},

  async fetchLevel(level,code,retries=3){
    const url=`${this.BASE}/level${level}?groupId=${this.GROUP_ID}&${this.PARAM[level]}=${code}`;
    for(let a=1;a<=retries;a++){if(this.stopped)return[];
      try{const r=await fetch(url,{credentials:'include'});if(!r.ok)throw new Error(`HTTP ${r.status}`);const j=await r.json();if(j.success&&Array.isArray(j.data))return j.data;return[];}
      catch(e){if(a===retries){this.stats.err++;this.errorItems.push({level,parentCode:code});this.log(`ERR lv${level}[${code}]: ${e.message}`,'r');return[];}await sleep(300*a);}
    }return[];
  },
  async collectNode(chain,code,level,maxLv,delay){
    if(this.stopped)return;
    const data=await this.wq.run(async()=>{this.stats.active++;updateStats('t1',this.stats);const d=await this.fetchLevel(level,code);this.stats.req++;this.stats.active--;if(delay>0)await sleep(delay);return d;});
    updateStats('t1',this.stats);
    if(!data.length){if(Object.keys(chain).length){this.rows.push({...chain});this.stats.rows++;updateStats('t1',this.stats);}return;}
    this.lcount[level]=(this.lcount[level]||0)+data.length;document.getElementById(`lc${level}`).textContent=this.lcount[level];
    this.log(`L${level}[${code}] → ${data.length} ${this.NAMES[level]}`,'i');
    await Promise.all(data.map(item=>{if(this.stopped)return Promise.resolve();const nc={...chain,[`level${level}`]:item};if(level>=maxLv){this.rows.push(nc);this.stats.rows++;updateStats('t1',this.stats);return Promise.resolve();}return this.collectNode(nc,item.fullCode,level+1,maxLv,delay);}));
  },
  async start(){
    const l2=document.getElementById('t1-l2').value.trim(),maxLv=parseInt(document.getElementById('t1-maxlv').value),conc=parseInt(document.getElementById('t1-conc').value)||3,delay=parseInt(document.getElementById('t1-delay').value)||0;
    if(!l2)return alert('Enter a Level 2 Full Code.');
    this.collecting=true;this.stopped=false;this.rows=[];this.errorItems=[];this.stats={req:0,rows:0,err:0,active:0};
    Object.keys(this.lcount).forEach(k=>this.lcount[k]=0);[3,4,5,6].forEach(l=>document.getElementById(`lc${l}`).textContent=0);
    this.wq=new WorkQueue(conc);clearLog('t1');hideDL('t1');setStatus('t1','running');this.startMs=Date.now();
    this.timer=setInterval(()=>document.getElementById('t1-elapsed').textContent=((Date.now()-this.startMs)/1000|0)+'s',1000);
    this.log(`Start: l2=${l2} maxLv=${maxLv} conc=${conc} delay=${delay}ms`,'s');
    try{await this.collectNode({},l2,3,maxLv,delay);}catch(e){this.log(`Fatal: ${e.message}`,'r');}
    clearInterval(this.timer);this.collecting=false;
    const sec=((Date.now()-this.startMs)/1000).toFixed(1);setStatus('t1',this.stopped?'stopped':'done');
    this.log(`Done in ${sec}s — ${this.rows.length} rows, ${this.stats.err} errors.`,this.stopped?'w':'s');
    T2.refreshInfo();T6.refreshInfo();saveStorage('t1data',{level2FullCode:l2,collectedAt:new Date().toISOString(),rows:this.rows});showDL('t1',this.rows.length);
  },
  stop(){this.stopped=true;this.log('Stop...','w');document.getElementById('t1-stop').disabled=true;},
  log(msg,cls='i'){addLog('t1',msg,cls);},
  // Our level numbers already match the Report Progress payload's
  // region1Id..region10Id numbering (level3=Kecamatan=region3, etc.), so the
  // fetched item's own "id" is surfaced as regionNId — ready to paste
  // straight into that payload's region filter.
  buildHeaders(){const lvls=new Set();for(const r of this.rows)Object.keys(r).forEach(k=>{const n=parseInt(k.replace('level',''));if(n>=3&&n<=7)lvls.add(n);});const sl=[...lvls].sort((a,b)=>a-b),h=[];sl.forEach(l=>h.push(`level${l}_fullCode`,`level${l}_code`,`level${l}_name`,`region${l}Id`));return{headers:h,levels:sl};},
  toFlat(){const{headers,levels}=this.buildHeaders();return{headers,rows:this.rows.map(row=>{const f={};levels.forEach(l=>{const it=row[`level${l}`]||{};f[`level${l}_fullCode`]=it.fullCode||'';f[`level${l}_code`]=it.code||'';f[`level${l}_name`]=it.name||'';f[`region${l}Id`]=it.id??'';});return f;})};},
  getFullCodes(pref='deepest'){const c=new Set();for(const r of this.rows){if(pref==='deepest'){for(let l=7;l>=3;l--){if(r[`level${l}`]){c.add(r[`level${l}`].fullCode);break;}}}else{const lv=parseInt(pref);if(r[`level${lv}`])c.add(r[`level${lv}`].fullCode);}}return[...c];},
  // One entry per unique region-id chain, using each row's deepest
  // available level (rows can stop short of maxLv when a branch has no
  // children). Deduped since multiple rows can share the same leaf.
  getDeepestRegionChains(){
    const seen=new Map();
    for(const r of this.rows){
      const chain={};let deepest=null;
      for(let l=3;l<=6;l++){const it=r[`level${l}`];if(it){chain[l]=it.id??null;deepest=l;}}
      if(deepest===null)continue;
      const key=JSON.stringify(chain);
      if(!seen.has(key))seen.set(key,{chain,deepest});
    }
    return[...seen.values()];
  },
  exportJSON(){if(!this.rows.length)return alert('No data.');downloadBlob(new Blob([JSON.stringify({collectedAt:new Date().toISOString(),totalRows:this.rows.length,data:this.rows},null,2)],{type:'application/json'}),`region_${nowTs()}.json`);this.log('Exported JSON','s');},
  async exportCSV(){if(!this.rows.length)return alert('No data.');const{headers,rows}=this.toFlat();showExportProgress('Writing CSV...');try{const csv=buildCSV(headers,rows);downloadBlob(new Blob([csv],{type:'text/csv;charset=utf-8'}),`region_${nowTs()}.csv`);}finally{hideExportProgress();}this.log('Exported CSV','s');},
  async exportXLSX(){if(!this.rows.length)return alert('No data.');const{headers,rows}=this.toFlat();showExportProgress('Building XLSX...');try{const xlsx=await buildXLSXFromArray('Region Hierarchy',headers,rows,(d,t)=>updateExportProgress(d,t));if(!exportCancelled)downloadBlob(new Blob([xlsx],{type:XLSX_MIME}),`region_${nowTs()}.xlsx`);}finally{hideExportProgress();}this.log('Exported XLSX','s');},
  exportErrors(){if(!this.errorItems.length)return;downloadBlob(new Blob([buildCSV(['level','parentCode'],this.errorItems.map(e=>({level:String(e.level),parentCode:e.parentCode})))],{type:'text/csv;charset=utf-8'}),`errors_t1_${nowTs()}.csv`);this.log(`Exported ${this.errorItems.length} error entries`,'s');},
};

// ────────────────────────────────────────────────────────────────────────────
// TAB 2: Assignment List — streams to IndexedDB
// ────────────────────────────────────────────────────────────────────────────
const T2={
  BASE:'https://fasih-sm.bps.go.id/app/api/assignment-general/api/assignments/get-principal-values-by-smallest-code/fd68e454-ba45-4b85-8205-f3bf777ded24',
  collecting:false,stopped:false,sourceMode:'tab1',fileCodes:[],
  // In-memory tracking only (no full rows in memory):
  _assignmentIds:new Set(),// for T3 source
  stats:{total:0,done:0,found:0,err:0,active:0},wq:null,timer:null,startMs:0,

  async fetch(fullCode,retries=3){
    for(let a=1;a<=retries;a++){if(this.stopped)return null;
      try{const r=await fetch(`${this.BASE}/${fullCode}`,{credentials:'include'});if(!r.ok)throw new Error(`HTTP ${r.status}`);return await r.json();}
      catch(e){if(a===retries){this.stats.err++;await DB.add('t2_errors',{fullCode});this.log(`ERR[${fullCode}]: ${e.message}`,'r');return null;}await sleep(300*a);}
    }return null;
  },
  getSourceCodes(){return this.sourceMode==='tab1'?T1.getFullCodes(document.getElementById('t2-level-sel').value):this.fileCodes;},
  getAssignmentIds(){return[...this._assignmentIds];},

  async start(){
    const codes=this.getSourceCodes();if(!codes.length)return alert('No fullCodes available.');
    const conc=parseInt(document.getElementById('t2-conc').value)||3,delay=parseInt(document.getElementById('t2-delay').value)||0;
    this.collecting=true;this.stopped=false;this._assignmentIds=new Set();
    await DB.clear('t2_results');await DB.clear('t2_errors');
    this.stats={total:codes.length,done:0,found:0,err:0,active:0};
    this.wq=new WorkQueue(conc);clearLog('t2');hideDL('t2');setStatus('t2','running');this.startMs=Date.now();
    this.timer=setInterval(()=>document.getElementById('t2-elapsed').textContent=((Date.now()-this.startMs)/1000|0)+'s',1000);
    this.log(`Start: ${codes.length} fullCodes → IndexedDB`,'s');
    await Promise.all(codes.map(code=>async()=>{
      if(this.stopped)return;
      const j=await this.wq.run(async()=>{this.stats.active++;updateStats('t2',this.stats);const d=await this.fetch(code);this.stats.active--;if(delay>0)await sleep(delay);return d;});
      this.stats.done++;
      const items=j?.success&&Array.isArray(j.data)?j.data:(j?.success&&j?.data?[j.data]:[]);
      for(const item of items){await DB.add('t2_results',{fullCode:code,...item});if(item.assignmentId)this._assignmentIds.add(item.assignmentId);this.stats.found++;}
      if(items.length)this.log(`OK[${code}] → ${items.length}`,'i');
      updateStats('t2',this.stats);
    }).map(t=>t()));
    clearInterval(this.timer);this.collecting=false;
    const total=await DB.count('t2_results');
    setStatus('t2',this.stopped?'stopped':'done');
    this.log(`Done — ${total} rows in DB, ${this.stats.err} errors.`,this.stopped?'w':'s');
    T3.refreshInfo();showDL('t2',total);
  },
  // Re-fetches only the fullCodes currently sitting in t2_errors, merging
  // successes into the existing t2_results rather than wiping the DB —
  // unlike start(), which always clears both stores first.
  async retryErrors(){
    if(this.collecting)return;
    const codes=(await DB.getAll('t2_errors')).map(r=>r.fullCode).filter(Boolean);
    if(!codes.length)return alert('No errors to retry.');
    const conc=parseInt(document.getElementById('t2-conc').value)||3,delay=parseInt(document.getElementById('t2-delay').value)||0;
    this.collecting=true;this.stopped=false;
    await DB.clear('t2_errors');// about to re-fetch exactly these; failures get re-added by fetch()
    this.stats={total:codes.length,done:0,found:0,err:0,active:0};
    this.wq=new WorkQueue(conc);setStatus('t2','running');document.getElementById('t2-retry-err').disabled=true;this.startMs=Date.now();
    this.timer=setInterval(()=>document.getElementById('t2-elapsed').textContent=((Date.now()-this.startMs)/1000|0)+'s',1000);
    this.log(`Retrying ${codes.length} failed fullCode(s) — merging into existing results`,'s');
    await Promise.all(codes.map(code=>async()=>{
      if(this.stopped)return;
      const j=await this.wq.run(async()=>{this.stats.active++;updateStats('t2',this.stats);const d=await this.fetch(code);this.stats.active--;if(delay>0)await sleep(delay);return d;});
      this.stats.done++;
      const items=j?.success&&Array.isArray(j.data)?j.data:(j?.success&&j?.data?[j.data]:[]);
      for(const item of items){await DB.add('t2_results',{fullCode:code,...item});if(item.assignmentId)this._assignmentIds.add(item.assignmentId);this.stats.found++;}
      if(items.length)this.log(`OK[${code}] → ${items.length}`,'i');
      updateStats('t2',this.stats);
    }).map(t=>t()));
    clearInterval(this.timer);this.collecting=false;
    const total=await DB.count('t2_results'),remainingErr=await DB.count('t2_errors');
    setStatus('t2',this.stopped?'stopped':'done');
    this.log(`Retry done — ${total} rows in DB total, ${remainingErr} still failing.`,remainingErr?'w':'s');
    T3.refreshInfo();showDL('t2',total);
  },
  stop(){this.stopped=true;this.log('Stop...','w');document.getElementById('t2-stop').disabled=true;},
  log(msg,cls='i'){addLog('t2',msg,cls);},
  refreshInfo(){const n=T1.rows.length,el=document.getElementById('t2-tab1-info');el.innerHTML=n>0?`<strong>${n}</strong> rows from Region Hierarchy Tab. Deepest codes: <strong>${T1.getFullCodes('deepest').length}</strong>`:'No Region Hierarchy Tab results yet.';},
  async loadFile(file){try{this.fileCodes=await parseFileToIds(file);const el=document.getElementById('t2-file-info');el.style.display='block';el.innerHTML=`<strong>${this.fileCodes.length}</strong> codes from <em>${escHtml(file.name)}</em>`;this.log(`Loaded ${this.fileCodes.length} codes`,'s');}catch(e){alert(e.message);}},

  T2_HEADERS:['fullCode','assignmentId','data1','data2','data3','data4','data5','data6','data7','data8','data9','data10'],
  rowMapper(rec){const f={};this.T2_HEADERS.forEach(h=>f[h]=String(rec[h]??''));return f;},

  exportJSON(){this.log('JSON export not supported for large DB datasets — use CSV or XLSX.','w');alert('Use CSV or XLSX export for DB-backed data.');},
  async exportCSV(){showExportProgress('Writing CSV to disk (File Save dialog will open)...');try{const n=await DB.exportCSV('t2_results',this.T2_HEADERS,r=>this.rowMapper(r),`assignments_list_${nowTs()}.csv`);this.log(n?`Exported ${n} rows via File Save`:'Export cancelled','s');}catch(e){this.log(`CSV error: ${e.message}`,'r');}finally{hideExportProgress();}},
  async exportXLSX(){showExportProgress('Building XLSX...');try{await DB.exportXLSX('t2_results','Assignment List',this.T2_HEADERS,r=>this.rowMapper(r),`assignments_list`,50000,updateExportProgress);this.log('XLSX export complete','s');}catch(e){this.log(`XLSX error: ${e.message}`,'r');}finally{hideExportProgress();}},
  async exportErrors(){const n=await DB.exportErrorCSV('t2_errors','fullCode',`errors_t2_${nowTs()}.csv`);this.log(n?`Exported ${n} failed fullCodes`:'No errors','s');},
  async clearDB(){if(!confirm('Clear all Assignment List Tab data from the local database?'))return;await DB.clear('t2_results');await DB.clear('t2_errors');this._assignmentIds=new Set();hideDL('t2');this.log('DB cleared','w');},
};

// ────────────────────────────────────────────────────────────────────────────
// TAB 3: Assignment Detail — streams to IndexedDB, pre-parses answers
// ────────────────────────────────────────────────────────────────────────────
const T3={
  BASE:'https://fasih-sm.bps.go.id/app/api/assignment-general/api/assignment/get-by-assignment-id',
  META:['_id','code_identity','assignment_status_alias','assignment_status_id',
    'data1','data2','data3','data4','data5','data6','data7','data8','data9','data10',
    'longitude','latitude','date_created','date_modified',
    'current_user_fullname','current_user_username','current_user_survey_role_name',
    'source_from','sum_error','sum_remark','sum_clean','done'],
  collecting:false,stopped:false,sourceMode:'tab2',fileIds:[],exportMode:'data',
  dataKeys:new Set(),predefKeys:new Set(),// in-memory key registries (just strings, small)
  stats:{total:0,done:0,err:0,active:0},wq:null,timer:null,startMs:0,

  async fetch(id,retries=3){
    const url=`${this.BASE}?assignmentId=${id}`;
    for(let a=1;a<=retries;a++){if(this.stopped)return null;
      try{const r=await fetch(url,{credentials:'include'});if(!r.ok)throw new Error(`HTTP ${r.status}`);return await r.json();}
      catch(e){if(a===retries){this.stats.err++;await DB.add('t3_errors',{assignmentId:id});this.log(`ERR[${id}]: ${e.message}`,'r');return null;}await sleep(400*a);}
    }return null;
  },

  parseRaw(id,raw){
    const meta={};this.META.forEach(k=>meta[k]=raw[k]??null);
    let dataMap={},predefMap={};
    try{const obj=JSON.parse(raw.data||'{}');(obj.answers||[]).forEach(({dataKey,answer})=>{if(dataKey){dataMap[dataKey]=extractAnswer(answer);this.dataKeys.add(dataKey);}});}catch(e){this.log(`Parse ERR data[${id}]: ${e.message}`,'r');}
    try{const obj=JSON.parse(raw.pre_defined_data||'{}');(obj.predata||[]).forEach(({dataKey,answer})=>{if(dataKey){predefMap[dataKey]=extractAnswer(answer);this.predefKeys.add(dataKey);}});}catch(e){this.log(`Parse ERR predef[${id}]: ${e.message}`,'r');}
    return{assignmentId:id,meta,dataMap,predefMap};
  },

  getSourceIds(){return this.sourceMode==='tab2'?T2.getAssignmentIds():this.fileIds;},

  async start(){
    const ids=this.getSourceIds();
    if(!ids.length){return alert(this.sourceMode==='tab2'?'No assignmentIds from Assignment List Tab.\n\nRun Assignment List Tab first, or switch source to "Upload File".':'No IDs in uploaded file.');}
    const conc=parseInt(document.getElementById('t3-conc').value)||2,delay=parseInt(document.getElementById('t3-delay').value)||0;
    this.collecting=true;this.stopped=false;this.dataKeys=new Set();this.predefKeys=new Set();
    await DB.clear('t3_results');await DB.clear('t3_errors');
    this.stats={total:ids.length,done:0,err:0,active:0};
    this.wq=new WorkQueue(conc);clearLog('t3');hideDL('t3');setStatus('t3','running');this.startMs=Date.now();
    this.timer=setInterval(()=>document.getElementById('t3-elapsed').textContent=((Date.now()-this.startMs)/1000|0)+'s',1000);
    this.log(`Start: ${ids.length} IDs → IndexedDB, mode=${this.exportMode}`,'s');
    await Promise.all(ids.map(id=>async()=>{
      if(this.stopped)return;
      const j=await this.wq.run(async()=>{this.stats.active++;updateStats('t3',this.stats);const d=await this.fetch(id);this.stats.active--;if(delay>0)await sleep(delay);return d;});
      this.stats.done++;
      if(j?.success&&j?.data){await DB.add('t3_results',this.parseRaw(id,j.data));this.log(`OK[${id}]`,'i');}
      updateStats('t3',this.stats);
    }).map(t=>t()));
    clearInterval(this.timer);this.collecting=false;
    const total=await DB.count('t3_results');
    setStatus('t3',this.stopped?'stopped':'done');
    this.log(`Done — ${total} assignments in DB, ${this.stats.err} errors.`,this.stopped?'w':'s');
    showDL('t3',total);
  },
  stop(){this.stopped=true;this.log('Stop...','w');document.getElementById('t3-stop').disabled=true;},
  log(msg,cls='i'){addLog('t3',msg,cls);},
  refreshInfo(){const n=T2.getAssignmentIds().length,el=document.getElementById('t3-tab2-info');el.innerHTML=n>0?`<strong>${n}</strong> unique assignmentIds from Assignment List Tab (${n} total).`:'No Assignment List Tab results yet.';},
  async loadFile(file){try{this.fileIds=await parseFileToIds(file);const el=document.getElementById('t3-file-info');el.style.display='block';el.innerHTML=`<strong>${this.fileIds.length}</strong> IDs from <em>${escHtml(file.name)}</em>`;this.log(`Loaded ${this.fileIds.length} IDs`,'s');}catch(e){alert(e.message);}},

  getHeaders(){
    const mode=this.exportMode;
    const ansKeys=[...(mode==='data'?this.dataKeys:this.predefKeys)];
    return[...this.META,...ansKeys];
  },
  rowMapper(rec){
    const mode=this.exportMode,headers=this.getHeaders();
    const flat={};
    this.META.forEach(k=>flat[k]=String(rec.meta?.[k]??''));
    const amap=mode==='data'?rec.dataMap:rec.predefMap;
    headers.slice(this.META.length).forEach(k=>flat[k]=amap?.[k]??'');
    return flat;
  },

  exportJSON(){this.log('Use CSV or XLSX for DB-backed data. JSON of 250K assignments would be too large.','w');alert('JSON export is not recommended for large datasets.\n\nUse CSV (streaming to disk) or XLSX (split files).');},
  async exportCSV(){
    const headers=this.getHeaders();
    showExportProgress('Writing CSV to disk (File Save dialog will open)...');
    try{const n=await DB.exportCSV('t3_results',headers,r=>this.rowMapper(r),`assignments_detail_${this.exportMode}_${nowTs()}.csv`);this.log(n?`Exported ${n} rows via File Save`:'Export cancelled','s');}
    catch(e){this.log(`CSV error: ${e.message}`,'r');}finally{hideExportProgress();}
  },
  async exportXLSX(){
    const headers=this.getHeaders();
    const total=await DB.count('t3_results');
    const splits=Math.ceil(total/50000);
    showExportProgress(`Building XLSX${splits>1?` (${splits} files, 50K rows each)`:''}...`);
    try{await DB.exportXLSX('t3_results',`Assignment Detail (${this.exportMode})`,headers,r=>this.rowMapper(r),`assignments_detail_${this.exportMode}`,50000,updateExportProgress);this.log('XLSX export complete','s');}
    catch(e){this.log(`XLSX error: ${e.message}`,'r');}finally{hideExportProgress();}
  },
  async exportErrors(){const n=await DB.exportErrorCSV('t3_errors','assignmentId',`errors_t3_${nowTs()}.csv`);this.log(n?`Exported ${n} failed assignmentIds`:'No errors','s');},
  async clearDB(){if(!confirm('Clear all Tab 3 data from the local database?'))return;await DB.clear('t3_results');await DB.clear('t3_errors');this.dataKeys=new Set();this.predefKeys=new Set();hideDL('t3');this.log('DB cleared','w');},
};

// ────────────────────────────────────────────────────────────────────────────
// TAB 6: Assignment Status — DataTables-style POST, driven by the unique
// region-id chains from the Region Hierarchy tab. For each chain, every
// page is collected (start/length) before moving to the next chain — runs
// strictly sequentially across chains too, by request, rather than
// concurrently like Tabs 2/3, since the requirement is "finish one region
// before starting the next."
// ────────────────────────────────────────────────────────────────────────────
const T6={
  ENDPOINT:'https://fasih-sm.bps.go.id/app/api/analytic/api/v2/assignment/datatable-all-user-survey-periode',
  MAX_RETRIES:4,RETRY_BASE_DELAY_MS:1500,RETRYABLE_STATUSES:[429,500,502,503,504],
  running:false,stopped:false,errorItems:[],wq:null,sourceMode:'tab1',fileChains:[],
  stats:{chainsTotal:0,chainsDone:0,records:0,err:0,active:0},timer:null,startMs:0,

  log(msg,cls='i'){addLog('t6',msg,cls);},

  refreshInfo(){
    const n=T1.rows.length,el=document.getElementById('t6-tab1-info');
    if(n>0){const chains=T1.getDeepestRegionChains().length;el.innerHTML=`<strong>${n}</strong> rows from Region Hierarchy Tab &mdash; <strong>${chains}</strong> unique region chain${chains===1?'':'s'} to collect.`;}
    else el.innerHTML='No Region Hierarchy results yet.';
  },
  getSourceChains(){return this.sourceMode==='tab1'?T1.getDeepestRegionChains():this.fileChains;},
  async loadFile(file){
    try{
      const chains=await parseRegionIdChainsFile(file);
      this.fileChains=chains;
      const info=document.getElementById('t6-file-info');
      info.style.display='block';
      info.innerHTML=`<strong>${chains.length}</strong> region chain${chains.length===1?'':'s'} loaded from <em>${escHtml(file.name)}</em>`;
      this.log(`Loaded ${chains.length} region chains from ${file.name}`,'s');
    }catch(e){
      const info=document.getElementById('t6-file-info');
      info.style.display='block';
      info.innerHTML=`<strong style="color:var(--red);">Failed to load:</strong> ${escHtml(e.message)}`;
      this.log(`File load error: ${e.message}`,'r');
    }
  },

  buildPayload(chain,start,length,surveyPeriodId){
    const region={};
    for(let l=1;l<=6;l++)region[`region${l}Id`]=chain[l]??null;
    const dataCols=Array.from({length:10},(_,i)=>({data:`data${i+1}`,orderable:true}));
    return{
      start,length,
      columns:[{data:'id',orderable:true},{data:'codeIdentity',orderable:true},...dataCols],
      order:[],
      search:{value:'',regex:false},
      assignmentExtraParam:{...region,surveyPeriodId,assignmentErrorStatusType:-1,filterTargetType:'TARGET_ONLY'}
    };
  },

  // This endpoint's actual shape is {searchData:[...], ...} — not the plain
  // DataTables {data:[...]} convention, and not the {success,data:{...}}
  // wrap most other endpoints in this app use either. Check all of them.
  extractRows(res){
    const directKeys=['searchData','data','content','items','rows','list'];
    for(const key of directKeys)if(Array.isArray(res?.[key]))return res[key];
    for(const key of directKeys)if(Array.isArray(res?.data?.[key]))return res.data[key];
    return[];
  },
  extractTotal(res){
    const keys=['recordsFiltered','recordsTotal','totalRecords','totalElements','total','count'];
    for(const key of keys){
      if(typeof res?.[key]==='number')return res[key];
      if(typeof res?.data?.[key]==='number')return res.data[key];
    }
    return null;
  },

  async fetchPage(payload){
    let attempt=0;
    while(true){
      let res;
      try{
        const r=await fetch(this.ENDPOINT,{method:'POST',headers:{'Content-Type':'application/json','Accept':'application/json',...await T4.csrfHeaders()},credentials:'include',body:JSON.stringify(payload)});
        if(!r.ok){let bodyText='';try{bodyText=(await r.text()).slice(0,300);}catch(_){}throw Object.assign(new Error(`HTTP ${r.status} ${r.statusText}`),{status:r.status,bodyText});}
        res=await r.json();
      }catch(e){
        const status=e.status??null;
        const retryable=status===null||this.RETRYABLE_STATUSES.includes(status);
        attempt++;
        if(!retryable||attempt>this.MAX_RETRIES)throw e;
        const delay=this.RETRY_BASE_DELAY_MS*2**(attempt-1);
        this.log(`${e.message}, retrying (${attempt}/${this.MAX_RETRIES}) in ${Math.round(delay/1000)}s...`,'r');
        await sleep(delay);
        continue;
      }
      return res;
    }
  },

  flattenChain(chain){const out={};for(const[lvl,id]of Object.entries(chain))out[`region${lvl}Id`]=id;return out;},

  // Pages through one region chain to exhaustion before returning.
  async collectChain(chainInfo,length,delay,surveyPeriodId){
    let start=0,total=null,collected=0;
    while(true){
      if(this.stopped)return;
      const payload=this.buildPayload(chainInfo.chain,start,length,surveyPeriodId);
      let res;
      try{res=await this.fetchPage(payload);}
      catch(e){
        this.stats.err++;updateErrBtn('t6',this.errorItems.length+1);
        this.errorItems.push({...this.flattenChain(chainInfo.chain),start,message:e.message});
        this.log(`ERR chain=${JSON.stringify(chainInfo.chain)} start=${start}: ${e.message}`,'r');
        return;
      }
      const rows=this.extractRows(res);
      for(const row of rows){await DB.add('t6_results',{...row,...this.flattenChain(chainInfo.chain)});}
      collected+=rows.length;this.stats.records+=rows.length;
      updateStatsT6();
      if(total===null)total=this.extractTotal(res);
      const reachedTotal=total!==null&&collected>=total;
      const emptyPage=rows.length===0;
      const shortPage=rows.length>0&&rows.length<length&&total===null;
      if(reachedTotal||emptyPage||shortPage)return;
      start+=length;
      await sleep(delay);
    }
  },

  async start(){
    if(this.running)return;
    const chains=this.getSourceChains();
    if(!chains.length){this.log(this.sourceMode==='tab1'?'No region chains found — run the Region Hierarchy tab first.':'No region chains loaded — upload a file first.','r');return;}
    const conc=parseInt(document.getElementById('t6-conc').value)||1;
    const length=parseInt(document.getElementById('t6-length').value)||10;
    const delay=parseInt(document.getElementById('t6-delay').value)||0;
    const surveyPeriodId=document.getElementById('t6-survey').value.trim()||T4.DEFAULT_PAYLOAD.surveyPeriodId;

    this.running=true;this.stopped=false;this.errorItems=[];
    this.stats={chainsTotal:chains.length,chainsDone:0,records:0,err:0,active:0};
    await DB.clear('t6_results');
    this.wq=new WorkQueue(conc);
    clearLog('t6');hideDL('t6');setStatusT6('running');updateErrBtn('t6',0);this.startMs=Date.now();
    this.timer=setInterval(()=>document.getElementById('t6-elapsed').textContent=((Date.now()-this.startMs)/1000|0)+'s',1000);
    document.getElementById('t6-start').disabled=true;document.getElementById('t6-stop').disabled=false;
    this.log(`Start: ${chains.length} unique region chain(s), concurrency=${conc}, page length=${length}`,'s');
    updateStatsT6();

    // Each worker still pages one chain to exhaustion before picking up the
    // next — concurrency is across chains, never within a single chain.
    await Promise.all(chains.map(chainInfo=>this.wq.run(async()=>{
      if(this.stopped)return;
      this.stats.active++;updateStatsT6();
      this.log(`Region chain (deepest level ${chainInfo.deepest}): ${JSON.stringify(chainInfo.chain)}`,'i');
      await this.collectChain(chainInfo,length,delay,surveyPeriodId);
      this.stats.active--;this.stats.chainsDone++;
      updateStatsT6();
    })));
    if(this.stopped)this.log('Stopped by user.','w');

    clearInterval(this.timer);this.running=false;
    document.getElementById('t6-start').disabled=false;document.getElementById('t6-stop').disabled=true;
    const total=await DB.count('t6_results');
    setStatusT6(this.stopped?'stopped':'done');
    this.log(`Done — ${total} rows in DB across ${this.stats.chainsDone}/${chains.length} chains, ${this.stats.err} chain error${this.stats.err===1?'':'s'}.`,this.stopped?'w':'s');
    showDL('t6',total);
  },
  stop(){this.stopped=true;this.log('Stop requested...','w');document.getElementById('t6-stop').disabled=true;},

  T6_HEADERS:['region1Id','region2Id','region3Id','region4Id','region5Id','region6Id','id','codeIdentity','data1','data2','data3','data4','data5','data6','data7','data8','data9','data10','assignmentStatusId','assignmentStatusAlias','strata','currentUserFullname','currentUserUsername','currentUserSurveyRoleName','dateCreated','dateModified'],
  rowMapper(rec){const f={};this.T6_HEADERS.forEach(h=>f[h]=String(rec[h]??''));return f;},

  exportJSON(){this.log('JSON export not supported for large DB datasets — use CSV or XLSX.','w');alert('Use CSV or XLSX export for DB-backed data.');},
  async exportCSV(){showExportProgress('Writing CSV to disk (File Save dialog will open)...');try{const n=await DB.exportCSV('t6_results',this.T6_HEADERS,r=>this.rowMapper(r),`assignment_status_${nowTs()}.csv`);this.log(n?`Exported ${n} rows via File Save`:'Export cancelled','s');}catch(e){this.log(`CSV error: ${e.message}`,'r');}finally{hideExportProgress();}},
  async exportXLSX(){showExportProgress('Building XLSX...');try{await DB.exportXLSX('t6_results','Assignment Status',this.T6_HEADERS,r=>this.rowMapper(r),`assignment_status`,50000,updateExportProgress);this.log('XLSX export complete','s');}catch(e){this.log(`XLSX error: ${e.message}`,'r');}finally{hideExportProgress();}},
  // region1Id..region6Id columns mean this file can be re-uploaded directly
  // as the "Upload File" source below to retry just the chains that failed.
  exportErrors(){
    if(!this.errorItems.length)return;
    const headers=['region1Id','region2Id','region3Id','region4Id','region5Id','region6Id','start','message'];
    downloadBlob(new Blob([buildCSV(headers,this.errorItems)],{type:'text/csv;charset=utf-8'}),`errors_t6_${nowTs()}.csv`);
    this.log(`Exported ${this.errorItems.length} chain error entries`,'s');
  },
  async clearDB(){if(!confirm('Clear all Assignment Status data from the local database?'))return;await DB.clear('t6_results');this.errorItems=[];hideDL('t6');updateErrBtn('t6',0);this.log('DB cleared','w');},
};
function setStatusT6(state){
  const b=document.getElementById('t6-badge');
  b.textContent={idle:'Idle',running:'Running',done:'Done',stopped:'Stopped'}[state]||state;
  b.className=`badge ${state}`;
  document.getElementById('t6-start').disabled=state==='running';
  document.getElementById('t6-stop').disabled=state!=='running';
}
function updateStatsT6(){
  const s=T6.stats;
  document.getElementById('t6-chains-done').textContent=s.chainsDone;
  document.getElementById('t6-chains-total').textContent=s.chainsTotal;
  document.getElementById('t6-records').textContent=s.records;
  document.getElementById('t6-err').textContent=s.err;
  document.getElementById('t6-act').textContent=s.active;
  const p=document.getElementById('t6-prog');
  if(p)p.style.width=(s.chainsTotal>0?s.chainsDone/s.chainsTotal*100:0)+'%';
}

// ────────────────────────────────────────────────────────────────────────────
// TAB 4: Report Progress by Responsibility — paginated collector with
// resumable progress (chrome.storage.local), explode-by-region and
// aggregate-by-iddesa exports. Ported from the standalone BPS Report
// Collector extension; uses fetch() directly (same-extension page, host
// permission already covers fasih-sm.bps.go.id, so no executeScript hop
// into a separate tab is needed here).
// ────────────────────────────────────────────────────────────────────────────
const T4={
  ENDPOINT:'https://fasih-sm.bps.go.id/app/api/analytic/api/v2/assignment/report-progress-by-responsibility',
  STORAGE_KEY:'t4Progress',
  MAX_PAGES:2000,DELAY_MS:300,MAX_RETRIES:4,RETRY_BASE_DELAY_MS:1500,
  RETRYABLE_STATUSES:[429,500,502,503,504],
  DEFAULT_PAYLOAD:{
    surveyPeriodId:'fd68e454-ba45-4b85-8205-f3bf777ded24',
    surveyRoleId:'6d7d919a-45e5-4779-bb87-2905b49fd31a',
    size:10,page:0,search:'',target:'TARGET_ONLY', // this endpoint's "page" is 0-indexed
    region:{region1Id:null,region2Id:null,region3Id:null,region4Id:null,region5Id:null,region6Id:null,region7Id:null,region8Id:null,region9Id:null,region10Id:null},
    regionSummaryLevel:6
  },
  collecting:false,stopped:false,collected:[],resumeState:null,pendingResume:null,
  stats:{page:0,records:0,err:0},timer:null,startMs:0,

  log(msg,cls='i'){addLog('t4',msg,cls);},

  async checkResumable(){
    const saved=await this._loadProgress();
    if(saved&&Array.isArray(saved.collected)&&saved.collected.length>0){
      this.pendingResume=saved;
      document.getElementById('t4-resume-text').textContent=`${saved.collected.length} record${saved.collected.length===1?'':'s'} collected up to page ${saved.lastPage} from a previous run that didn't finish.`;
      document.getElementById('t4-resume').style.display='block';
    }
  },
  hideResumeBanner(){document.getElementById('t4-resume').style.display='none';this.pendingResume=null;},
  _saveProgress(state){return new Promise(res=>chrome.storage.local.set({[this.STORAGE_KEY]:state},res));},
  _loadProgress(){return new Promise(res=>chrome.storage.local.get([this.STORAGE_KEY],r=>res(r?.[this.STORAGE_KEY]||null)));},
  _clearProgress(){return new Promise(res=>chrome.storage.local.remove([this.STORAGE_KEY],res));},

  // Spring Security's double-submit-cookie convention: the CSRF cookie has
  // to be echoed back as a request header, or the server returns
  // "Invalid CSRF Token". The extension page's own document.cookie can't
  // see fasih-sm.bps.go.id's cookies (different origin), so this reads them
  // via chrome.cookies instead.
  async csrfHeaders(){
    const names={'XSRF-TOKEN':'X-XSRF-TOKEN','CSRF-TOKEN':'X-CSRF-TOKEN','csrftoken':'X-CSRFToken','_csrf':'X-CSRF-Token'};
    const headers={};
    try{
      const cookies=await new Promise(res=>chrome.cookies.getAll({url:'https://fasih-sm.bps.go.id/'},res));
      for(const c of cookies)if(names[c.name])headers[names[c.name]]=decodeURIComponent(c.value);
    }catch(_){}
    return headers;
  },

  async fetchPage(payload,page,retries=undefined){
    let attempt=0;
    while(true){
      let res;
      try{
        const r=await fetch(this.ENDPOINT,{method:'POST',headers:{'Content-Type':'application/json','Accept':'application/json',...await this.csrfHeaders()},credentials:'include',body:JSON.stringify(payload)});
        if(!r.ok){let bodyText='';try{bodyText=(await r.text()).slice(0,300);}catch(_){}throw Object.assign(new Error(`HTTP ${r.status} ${r.statusText}`),{status:r.status,bodyText});}
        res=await r.json();
      }catch(e){
        const status=e.status??null;
        const retryable=status===null||this.RETRYABLE_STATUSES.includes(status);
        attempt++;
        if(!retryable||attempt>this.MAX_RETRIES)throw e;
        const delay=this.RETRY_BASE_DELAY_MS*2**(attempt-1);
        this.log(`${e.message}, retrying (${attempt}/${this.MAX_RETRIES}) in ${Math.round(delay/1000)}s...`,'r');
        await sleep(delay);
        continue;
      }
      return res;
    }
  },

  extractRows(response){
    if(Array.isArray(response))return response;
    const directKeys=['content','data','result','items','rows','list'];
    for(const key of directKeys)if(Array.isArray(response?.[key]))return response[key];
    for(const key of['data','result'])if(Array.isArray(response?.[key]?.content))return response[key].content;
    return[];
  },
  readNumber(response,keys){
    for(const key of keys){
      if(typeof response?.[key]==='number')return response[key];
      for(const wrap of['data','result'])if(typeof response?.[wrap]?.[key]==='number')return response[wrap][key];
    }
    return null;
  },

  async start(){
    if(this.collecting)return;
    let basePayload,startPage,seed=[];
    if(this.resumeState){
      basePayload=this.resumeState.basePayload;seed=this.resumeState.collected;startPage=this.resumeState.lastPage+1;
      document.getElementById('t4-payload').value=JSON.stringify(basePayload,null,2);
    }else{
      try{basePayload=JSON.parse(document.getElementById('t4-payload').value);}
      catch(e){this.log('Payload is not valid JSON. Fix it and try again.','r');return;}
      // Don't use `|| 0` here — this endpoint's page is 0-indexed, and
      // `Number(0) || 0` is fine, but `||` would also clobber an explicit 0
      // with a fallback if the fallback weren't 0 too. Being explicit about
      // the NaN check keeps an intentional page:0 from ever being coerced.
      const parsedPage=Number(basePayload.page);
      startPage=Number.isFinite(parsedPage)?parsedPage:0;
    }
    this.resumeState=null;this.collecting=true;this.stopped=false;this.collected=seed;
    this.stats={page:0,records:seed.length,err:0};
    clearLog('t4');hideDL('t4');setStatus('t4','running');this.startMs=Date.now();
    document.getElementById('t4-start').disabled=true;document.getElementById('t4-stop').disabled=false;
    this.timer=setInterval(()=>document.getElementById('t4-elapsed').textContent=((Date.now()-this.startMs)/1000|0)+'s',1000);
    if(seed.length)this.log(`Resuming with ${seed.length} records already collected.`,'s');
    this.updateStats();

    const size=Number(basePayload.size)||10;
    const conc=parseInt(document.getElementById('t4-conc').value)||1;
    let page=startPage,totalPages=null,totalElements=null,done=false;

    // Pages are fetched conc-at-a-time (instead of strictly one-by-one) to
    // shrink the wall-clock window over which this live, offset-paginated
    // dataset can shift underneath us — e.g. TARGET_ONLY rows dropping out
    // as enumerators submit mid-collection, which otherwise causes an
    // undercount well below the true total.
    while(!done){
      if(this.stopped){
        this.log('Stopped by user.','w');
        await this._saveProgress({collected:this.collected,lastPage:page-1,basePayload});
        break;
      }
      const batchPages=Array.from({length:conc},(_,i)=>page+i);
      const results=await Promise.all(batchPages.map(p=>
        this.fetchPage({...basePayload,page:p},p).then(response=>({p,response})).catch(e=>({p,error:e}))
      ));

      for(const{p,response,error}of results){
        if(this.stopped)break;
        if(error){
          this.stats.err++;this.log(`Request failed after retries (page ${p}): ${error.message}${error.bodyText?` Server said: ${error.bodyText}`:''}`,'r');
          await this._saveProgress({collected:this.collected,lastPage:p-1,basePayload});
          setStatus('t4','stopped');
          done=true;break;
        }
        const rows=this.extractRows(response);
        this.collected.push(...rows);
        this.stats.page=p;this.stats.records=this.collected.length;
        this.log(`Page ${p}: ${rows.length} record${rows.length===1?'':'s'}`,'i');
        this.updateStats();
        await this._saveProgress({collected:this.collected,lastPage:p,basePayload});

        if(totalPages===null)totalPages=this.readNumber(response,['totalPages','totalPage']);
        if(totalElements===null)totalElements=this.readNumber(response,['totalElements','totalItems','total']);

        const reachedTotalPages=totalPages!==null&&p>=totalPages;
        const reachedTotalElements=totalElements!==null&&this.collected.length>=totalElements;
        const emptyPage=rows.length===0;
        const shortPage=rows.length>0&&rows.length<size&&totalPages===null&&totalElements===null;

        if(reachedTotalPages||reachedTotalElements||emptyPage||shortPage){
          this.log(`Done, ${this.collected.length} record${this.collected.length===1?'':'s'} total.`,'s');
          await this._clearProgress();
          done=true;break;
        }
        if(p>=this.MAX_PAGES){
          this.log(`Stopped at the ${this.MAX_PAGES}-page safety limit.`,'r');
          await this._saveProgress({collected:this.collected,lastPage:p,basePayload});
          done=true;break;
        }
      }
      if(!done){page+=conc;await sleep(this.DELAY_MS);}
    }

    clearInterval(this.timer);this.collecting=false;
    document.getElementById('t4-start').disabled=false;document.getElementById('t4-stop').disabled=true;
    setStatus('t4',this.stopped?'stopped':(this.stats.err&&this.collected.length===0?'stopped':'done'));

    if(this.collected.length>0){
      document.getElementById('t4-result-info').innerHTML=`<strong>${this.collected.length}</strong> records collected.`;
      const aggRows=this.aggregateByIddesa();
      document.getElementById('t4-agg-info').innerHTML=`<strong>${aggRows.length}</strong> villages (iddesa).`;
      const aggKecRows=this.aggregateByIdkec();
      document.getElementById('t4-agg-kec-info').innerHTML=`<strong>${aggKecRows.length}</strong> kecamatan (idkec).`;
      const aggUserRows=this.aggregateByUsername();
      document.getElementById('t4-agg-username-info').innerHTML=`<strong>${aggUserRows.length}</strong> username${aggUserRows.length===1?'':'s'}.`;
      showDL('t4',this.collected.length);
      this.updateDlSummary();
    }
  },
  stop(){this.stopped=true;this.log('Stop requested...','w');document.getElementById('t4-stop').disabled=true;},
  updateStats(){
    document.getElementById('t4-page').textContent=this.stats.page;
    document.getElementById('t4-records').textContent=this.stats.records;
    document.getElementById('t4-err').textContent=this.stats.err;
  },

  // Exploding regionSummary (and statusBreakdown within it) into flat rows,
  // then aggregating by iddesa (first 10 digits of regionCode), mirrors the
  // standalone BPS Report Collector's export logic.
  flattenObject(obj,prefix=''){
    const out={};
    for(const[key,value]of Object.entries(obj||{})){
      const fullKey=prefix?`${prefix}.${key}`:key;
      if(value!==null&&typeof value==='object'&&!Array.isArray(value))Object.assign(out,this.flattenObject(value,fullKey));
      else if(Array.isArray(value))out[fullKey]=JSON.stringify(value);
      else out[fullKey]=value;
    }
    return out;
  },
  explodeRegionSummary(){
    const rows=[];
    for(const record of this.collected){
      const{regionSummary,total,...rest}=record;
      const parentFlat=this.flattenObject(rest);parentFlat.total=total;
      if(!Array.isArray(regionSummary)||regionSummary.length===0){rows.push({...parentFlat,regionCode:null,regionTotal:null});continue;}
      for(const region of regionSummary){
        const{statusBreakdown,total:regionTotal,regionCode,...regionRest}=region||{};
        const row={...parentFlat,...this.flattenObject(regionRest,'region'),regionCode,regionTotal};
        if(Array.isArray(statusBreakdown))for(const entry of statusBreakdown)if(entry&&entry.status!==undefined)row[`status.${entry.status}`]=entry.count;
        rows.push(row);
      }
    }
    return rows;
  },
  // total_assignment = every status in the list; submitted = everything
  // except OPEN and DRAFT; progress = everything except OPEN (DRAFT still
  // counts as "in progress"). Applied uniformly at every aggregation level
  // (raw, iddesa, idkec) since they all share the same status.* column
  // shape. Named "total_assignment" (not "total") since "total" is already
  // the API's own per-record field.
  TOTAL_STATUS_KEYS:['status.OPEN','status.SUBMITTED BY Pencacah','status.DRAFT','status.APPROVED BY Pengawas','status.REJECTED BY Pengawas','status.EDITED BY Pengawas','status.REVOKED BY Pengawas','status.SUBMITTED RESPONDENT','status.REJECTED BY Admin Kabupaten'],
  SUBMITTED_STATUS_KEYS:['status.SUBMITTED BY Pencacah','status.APPROVED BY Pengawas','status.REJECTED BY Pengawas','status.EDITED BY Pengawas','status.REVOKED BY Pengawas','status.SUBMITTED RESPONDENT','status.REJECTED BY Admin Kabupaten'],
  PROGRESS_STATUS_KEYS:['status.SUBMITTED BY Pencacah','status.DRAFT','status.APPROVED BY Pengawas','status.REJECTED BY Pengawas','status.EDITED BY Pengawas','status.REVOKED BY Pengawas','status.SUBMITTED RESPONDENT','status.REJECTED BY Admin Kabupaten'],
  COMPUTED_COLS:['total_assignment','submitted','submitted_percentage','progress','progress_percentage'],
  statusSum(row,keys){return keys.reduce((s,k)=>s+(Number(row[k])||0),0);},
  addStatusAggregates(row){
    const total=this.statusSum(row,this.TOTAL_STATUS_KEYS);
    const submitted=this.statusSum(row,this.SUBMITTED_STATUS_KEYS);
    const progress=this.statusSum(row,this.PROGRESS_STATUS_KEYS);
    row.total_assignment=total;
    row.submitted=submitted;
    row.submitted_percentage=total>0?Number((submitted/total*100).toFixed(2)):0;
    row.progress=progress;
    row.progress_percentage=total>0?Number((progress/total*100).toFixed(2)):0;
    return row;
  },
  buildExportRows(){return this.explodeRegionSummary().map(r=>this.addStatusAggregates(this.flattenObject(r)));},
  // Generic prefix-of-regionCode rollup — iddesa is the first 10 digits
  // (province+regency+district+village), idkec is the first 7
  // (province+regency+district). regionTotal and every status.* column sum
  // across whichever regionCode entries share that prefix.
  // Generic rollup: keyFn picks the group key for each exploded row,
  // regionTotal and every status.* column sum within each group, then the
  // same total_assignment/submitted/progress columns get computed on top.
  aggregateBy(keyName,keyFn){
    const rows=this.buildExportRows();
    const groups=new Map(),order=[];
    for(const row of rows){
      const keyVal=keyFn(row);
      if(!groups.has(keyVal)){groups.set(keyVal,{[keyName]:keyVal,regionCodeCount:0,regionTotal:0});order.push(keyVal);}
      const agg=groups.get(keyVal);
      if(row.regionCode)agg.regionCodeCount+=1;
      agg.regionTotal+=Number(row.regionTotal)||0;
      for(const[key,value]of Object.entries(row))if(key.startsWith('status.'))agg[key]=(agg[key]||0)+(Number(value)||0);
    }
    return order.map(key=>this.addStatusAggregates(groups.get(key)));
  },
  aggregateByPrefix(prefixLen,keyName){
    return this.aggregateBy(keyName,row=>{const rc=row.regionCode;return rc?String(rc).slice(0,prefixLen):'(unassigned)';});
  },
  aggregateByIddesa(){return this.aggregateByPrefix(10,'iddesa');},
  aggregateByIdkec(){return this.aggregateByPrefix(7,'idkec');},
  // The endpoint's exact field name for the username isn't fixed across
  // payload variants, so this looks for any flattened key containing
  // "username" (preferring an exact match) rather than hardcoding one.
  findUsernameField(rows){
    const keys=new Set();
    rows.forEach(r=>Object.keys(r).forEach(k=>keys.add(k)));
    const candidates=[...keys].filter(k=>/username/i.test(k));
    return candidates.includes('username')?'username':(candidates[0]||null);
  },
  aggregateByUsername(){
    const field=this.findUsernameField(this.buildExportRows());
    return this.aggregateBy('username',row=>field?String(row[field]??'(unknown)'):'(unknown)');
  },
  // Computed columns are pulled out and appended last regardless of where
  // they fall in each row's own key order, so they consistently land at
  // the end of the exported header row.
  headerUnion(rows){
    const s=new Set();
    rows.forEach(r=>Object.keys(r).forEach(k=>{if(!this.COMPUTED_COLS.includes(k))s.add(k);}));
    const present=this.COMPUTED_COLS.filter(c=>rows.some(r=>c in r));
    return[...s,...present];
  },

  // Single dropdown (#t4-level-sel) picks which level JSON/CSV/XLSX export.
  LEVEL_META:{
    raw:{noun:'records',sheet:'Report Progress',base:'report_progress'},
    iddesa:{noun:'villages',sheet:'By iddesa',base:'report_progress_by_iddesa'},
    idkec:{noun:'kecamatan',sheet:'By idkec',base:'report_progress_by_idkec'},
    username:{noun:'usernames',sheet:'By username',base:'report_progress_by_username'}
  },
  selectedLevel(){return document.getElementById('t4-level-sel').value;},
  levelRows(level){
    if(level==='iddesa')return this.aggregateByIddesa();
    if(level==='idkec')return this.aggregateByIdkec();
    if(level==='username')return this.aggregateByUsername();
    return this.buildExportRows();
  },
  updateDlSummary(){
    if(!this.collected.length)return;
    const level=this.selectedLevel(),meta=this.LEVEL_META[level];
    document.getElementById('t4-dl-count').textContent=this.levelRows(level).length;
    document.getElementById('t4-dl-noun').textContent=meta.noun;
  },

  exportJSON(){
    if(!this.collected.length)return alert('No data.');
    const level=this.selectedLevel(),meta=this.LEVEL_META[level];
    const data=level==='raw'?this.collected:this.levelRows(level);
    downloadBlob(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),`${meta.base}_${nowTs()}.json`);
    this.log(`Exported ${meta.noun} JSON`,'s');
  },
  exportCSV(){
    if(!this.collected.length)return alert('No data.');
    const level=this.selectedLevel(),meta=this.LEVEL_META[level];
    const rows=this.levelRows(level),headers=this.headerUnion(rows);
    downloadBlob(new Blob([buildCSV(headers,rows)],{type:'text/csv;charset=utf-8'}),`${meta.base}_${nowTs()}.csv`);
    this.log(`Exported ${meta.noun} CSV`,'s');
  },
  async exportXLSX(){
    if(!this.collected.length)return alert('No data.');
    const level=this.selectedLevel(),meta=this.LEVEL_META[level];
    const rows=this.levelRows(level),headers=this.headerUnion(rows);
    showExportProgress('Building XLSX...');
    try{const xlsx=await buildXLSXFromArray(meta.sheet,headers,rows,(d,t)=>updateExportProgress(d,t));if(!exportCancelled)downloadBlob(new Blob([xlsx],{type:XLSX_MIME}),`${meta.base}_${nowTs()}.xlsx`);}
    finally{hideExportProgress();}
    this.log(`Exported ${meta.noun} XLSX`,'s');
  },
};

// ────────────────────────────────────────────────────────────────────────────
// TAB 7: Anomaly — Dashboard SE2026 single-fetch collector
// ────────────────────────────────────────────────────────────────────────────
const T7={
  CONFIGS:{
    usaha:{
      base:'https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab?kode_kabupaten={kab}&indikator=128,129,130,131,132,133,134,135&sudah_indikator=40,41,42,43,44,45,46,46&type=usaha&anomali_no={no}',
      count:8,
    },
    keluarga:{
      base:'https://dashboard-se2026.apps.bps.go.id/api/mikro/anomali-case-kab?kode_kabupaten={kab}&indikator=136,137,139,140,141,142&sudah_indikator=47,48,50,51,52,53&type=keluarga&anomali_no={no}',
      count:7,
    },
  },
  fetching:false,stopped:false,
  store:{usaha:{data:[],headers:[]},keluarga:{data:[],headers:[]}},
  enrichSource:'api',enrichFileMap:null,

  log(msg,cls='i'){addLog('t7',msg,cls);},

  buildUrl(type,no){
    const kab=document.getElementById('t7-kab').value.trim()||'5206';
    return this.CONFIGS[type].base.replace('{kab}',encodeURIComponent(kab)).replace('{no}',no);
  },

  normaliseRows(json){
    let rows=json;
    if(!Array.isArray(rows)){
      for(const k of['data','result','items','rows','list','content']){
        if(Array.isArray(json[k])){rows=json[k];break;}
      }
    }
    return Array.isArray(rows)?rows:[json];
  },

  MAX_RETRIES:4,RETRY_BASE_MS:1500,RETRYABLE:[429,500,502,503,504],

  async fetchWithRetry(url,label){
    let attempt=0;
    while(true){
      try{
        const r=await window.fetch(url,{credentials:'include'});
        if(!r.ok){
          const err=Object.assign(new Error(`HTTP ${r.status} ${r.statusText}`),{status:r.status});
          throw err;
        }
        return await r.json();
      }catch(e){
        const retryable=e.status==null||this.RETRYABLE.includes(e.status);
        attempt++;
        if(!retryable||attempt>this.MAX_RETRIES){
          this.log(`${label} failed after ${attempt} attempt${attempt===1?'':'s'}: ${e.message}`,'r');
          return null;
        }
        const delay=this.RETRY_BASE_MS*2**(attempt-1);
        this.log(`${label} ${e.message}, retry ${attempt}/${this.MAX_RETRIES} in ${Math.round(delay/1000)}s...`,'w');
        await sleep(delay);
      }
    }
  },

  async fetchLoop(type){
    const cfg=this.CONFIGS[type];
    const allRows=[];
    const hSet=new Set();
    let errors=0;
    for(let no=1;no<=cfg.count;no++){
      if(this.stopped)break;
      const url=this.buildUrl(type,no);
      this.log(`${type} anomali_no=${no}: ${url}`,'i');
      const json=await this.fetchWithRetry(url,`${type} no=${no}`);
      if(json===null){errors++;continue;}
      const rows=this.normaliseRows(json);
      rows.forEach(row=>Object.keys(row).forEach(k=>hSet.add(k)));
      allRows.push(...rows);
      this.log(`${type} no=${no} → ${rows.length} record${rows.length===1?'':'s'} (total ${allRows.length})`,'i');
      document.getElementById('t7-records').textContent=
        this.store.usaha.data.length+this.store.keluarga.data.length+allRows.length;
    }
    if(errors)this.log(`${type}: ${errors} request${errors===1?'':'s'} failed and skipped`,'w');
    return{data:allRows,headers:[...hSet]};
  },

  ASSIGNMENT_BASE:'https://fasih-sm.bps.go.id/app/api/assignment-general/api/assignment/get-by-assignment-id',

  async fetchAssignment(id,retries=3){
    const url=`${this.ASSIGNMENT_BASE}?assignmentId=${id}`;
    for(let a=1;a<=retries;a++){if(this.stopped)return null;
      try{const r=await fetch(url,{credentials:'include'});if(!r.ok)throw new Error(`HTTP ${r.status}`);return await r.json();}
      catch(e){if(a===retries){this.log(`ERR[${id}]: ${e.message}`,'r');return null;}await sleep(400*a);}
    }return null;
  },

  // Builds an assignmentId→{data1,data2} map from whichever source is
  // already at hand, instead of always re-hitting the live API per id.
  async buildLookupFromTab2(){
    const rows=await DB.getAll('t2_results');
    const map=new Map();
    for(const r of rows)if(r.assignmentId)map.set(r.assignmentId,{data1:r.data1??'',data2:r.data2??''});
    return map;
  },
  refreshTab2Info(){
    DB.count('t2_results').then(n=>{
      const el=document.getElementById('t7-enrich-tab2-info');
      if(el)el.innerHTML=n>0?`<strong>${n}</strong> rows in the Assignment List Tab's local database.`:'No Assignment List Tab results yet.';
    });
  },
  async loadEnrichFile(file){
    try{
      const map=await parseAssignmentLookupFile(file);
      this.enrichFileMap=map;
      const info=document.getElementById('t7-enrich-file-info');
      info.style.display='block';
      info.innerHTML=`<strong>${map.size}</strong> assignment(s) loaded from <em>${escHtml(file.name)}</em>`;
      this.log(`Loaded ${map.size} assignments from ${file.name}`,'s');
    }catch(e){
      this.enrichFileMap=null;
      const info=document.getElementById('t7-enrich-file-info');
      info.style.display='block';
      info.innerHTML=`<strong style="color:var(--red);">Failed to load:</strong> ${escHtml(e.message)}`;
      this.log(`File load error: ${e.message}`,'r');
    }
  },

  // Merges a lookup map into every collected row, adding data1/data2 to
  // each store's export headers. Shared by all three enrichment sources.
  applyLookup(cache){
    const types=['usaha','keluarga'];
    let matched=0,total=0;
    types.forEach(t=>{
      this.store[t].data.forEach(r=>{
        total++;
        const hit=r.assignment_id?cache.get(r.assignment_id):null;
        if(hit)matched++;
        r.data1=hit?hit.data1:'';
        r.data2=hit?hit.data2:'';
      });
      if(!this.store[t].headers.includes('data1'))this.store[t].headers.push('data1');
      if(!this.store[t].headers.includes('data2'))this.store[t].headers.push('data2');
    });
    return{matched,total};
  },

  async enrichAssignments(){
    if(this.enrichSource==='api')return this.enrichViaApi();
    let cache;
    if(this.enrichSource==='tab2'){
      cache=await this.buildLookupFromTab2();
      if(!cache.size){this.log('No Assignment List Tab results found in the local database.','r');return;}
      this.log(`Using ${cache.size} assignment(s) from Assignment List Tab DB`,'i');
    }else{
      if(!this.enrichFileMap){this.log('No lookup file loaded — upload an Assignment List export first.','r');return;}
      cache=this.enrichFileMap;
      this.log(`Using ${cache.size} assignment(s) from uploaded file`,'i');
    }
    const{matched,total}=this.applyLookup(cache);
    this.log(`Done — ${matched}/${total} record(s) matched.`,matched<total?'w':'s');
  },

  async enrichViaApi(){
    const types=['usaha','keluarga'];
    const ids=new Set();
    types.forEach(t=>this.store[t].data.forEach(r=>{if(r.assignment_id)ids.add(r.assignment_id);}));
    const idList=[...ids];
    if(!idList.length)return;
    const conc=parseInt(document.getElementById('t7-enrich-conc').value)||2,delay=parseInt(document.getElementById('t7-enrich-delay').value)||0;
    const stats={total:idList.length,done:0,err:0,active:0};
    const wq=new WorkQueue(conc);
    this.log(`Start: ${idList.length} assignment_id(s) → enrich data1/data2`,'s');
    const cache=new Map();
    await Promise.all(idList.map(id=>async()=>{
      if(this.stopped)return;
      const j=await wq.run(async()=>{stats.active++;const d=await this.fetchAssignment(id);stats.active--;if(delay>0)await sleep(delay);return d;});
      stats.done++;
      if(j?.success&&j?.data){cache.set(id,{data1:j.data.data1??'',data2:j.data.data2??''});this.log(`OK[${id}]`,'i');}
      else stats.err++;
    }).map(t=>t()));
    const{matched,total}=this.applyLookup(cache);
    this.log(`Done — ${stats.total-stats.err} fetched ok, ${stats.err} fetch errors, ${matched}/${total} record(s) matched.`,stats.err?'w':'s');
  },

  async fetch(){
    if(this.fetching)return;
    this.fetching=true;this.stopped=false;
    this.store={usaha:{data:[],headers:[]},keluarga:{data:[],headers:[]}};
    const startMs=Date.now();
    clearLog('t7');hideDL('t7');
    document.getElementById('t7-badge').textContent='Running';document.getElementById('t7-badge').className='badge running';
    document.getElementById('t7-records').textContent='0';
    document.getElementById('t7-fetch').disabled=true;
    document.getElementById('t7-stop').disabled=false;
    const timer=setInterval(()=>document.getElementById('t7-elapsed').textContent=((Date.now()-startMs)/1000|0)+'s',1000);
    try{
      this.store.usaha=await this.fetchLoop('usaha');
      if(!this.stopped)this.store.keluarga=await this.fetchLoop('keluarga');
      if(!this.stopped&&document.getElementById('t7-enrich').checked)await this.enrichAssignments();
      const total=this.store.usaha.data.length+this.store.keluarga.data.length;
      document.getElementById('t7-records').textContent=total;
      const state=this.stopped?'stopped':'done';
      this.log(`${this.stopped?'Stopped':'Done'} — usaha: ${this.store.usaha.data.length}, keluarga: ${this.store.keluarga.data.length}`,this.stopped?'w':'s');
      document.getElementById('t7-badge').textContent=this.stopped?'Stopped':'Done';
      document.getElementById('t7-badge').className=`badge ${state}`;
      if(total>0){showDL('t7',0);this.updateDlCount();}
    }catch(e){
      this.log(`Error: ${e.message}`,'r');
      document.getElementById('t7-badge').textContent='Stopped';document.getElementById('t7-badge').className='badge stopped';
    }finally{
      clearInterval(timer);this.fetching=false;
      document.getElementById('t7-fetch').disabled=false;
      document.getElementById('t7-stop').disabled=true;
    }
  },

  stop(){this.stopped=true;this.log('Stop requested...','w');document.getElementById('t7-stop').disabled=true;},

  selectedType(){return document.getElementById('t7-type-dl').value;},
  current(){return this.store[this.selectedType()];},
  updateDlCount(){
    const c=this.current();
    document.getElementById('t7-dl-count').textContent=c.data.length;
  },

  rowFlat(r,headers){const f={};headers.forEach(h=>f[h]=String(r[h]??''));return f;},

  exportJSON(){
    const type=this.selectedType();
    const{data,headers}=this.store[type]||{data:[],headers:[]};
    this.log(`Exporting JSON — type: ${type}, records: ${data.length}`,'i');
    if(!data.length)return alert('No data for '+type+'.');
    downloadBlob(new Blob([JSON.stringify(data,null,2)],{type:'application/json'}),`anomaly_${type}_${nowTs()}.json`);
    this.log(`Exported ${type} JSON (${data.length} records)`,'s');
  },
  exportCSV(){
    const type=this.selectedType();
    const{data,headers}=this.store[type]||{data:[],headers:[]};
    this.log(`Exporting CSV — type: ${type}, records: ${data.length}`,'i');
    if(!data.length)return alert('No data for '+type+'.');
    const rows=data.map(r=>this.rowFlat(r,headers));
    downloadBlob(new Blob([buildCSV(headers,rows)],{type:'text/csv;charset=utf-8'}),`anomaly_${type}_${nowTs()}.csv`);
    this.log(`Exported ${type} CSV (${data.length} records)`,'s');
  },
  async exportXLSX(){
    const type=this.selectedType();
    const{data,headers}=this.store[type]||{data:[],headers:[]};
    this.log(`Exporting XLSX — type: ${type}, records: ${data.length}`,'i');
    if(!data.length)return alert('No data for '+type+'.');
    const rows=data.map(r=>this.rowFlat(r,headers));
    showExportProgress('Building XLSX...');
    try{
      const xlsx=await buildXLSXFromArray('Anomaly '+type,headers,rows,(d,t)=>updateExportProgress(d,t));
      if(!exportCancelled)downloadBlob(new Blob([xlsx],{type:XLSX_MIME}),`anomaly_${type}_${nowTs()}.xlsx`);
    }finally{hideExportProgress();}
    this.log(`Exported ${type} XLSX (${data.length} records)`,'s');
  },
};

// ────────────────────────────────────────────────────────────────────────────
// TAB 5: SQL Lab — close every open SQL Lab tab on the Superset dashboard.
// Runs via chrome.scripting.executeScript inside the dashboard tab itself
// (not via fetch), since this is a DOM-click task, not an API call.
// ────────────────────────────────────────────────────────────────────────────
const T5={
  URL_PATTERN:'https://fasih-dashboard.bps.go.id/superset/sqllab*',
  SELECTOR:'button.ant-tabs-tab-remove[aria-label="remove"]',
  running:false,stopRequested:false,errorRows:[],

  log(msg,cls='i'){addLog('t5',msg,cls);},

  async start(){
    if(this.running)return;
    const delay=parseInt(document.getElementById('t5-delay').value)||0;
    // If multiple tabs match (e.g. a stale background tab left open
    // elsewhere), prefer the one that's actually active/focused right now —
    // a background tab can be sitting on an internal error frame that still
    // reports a normal title/url to chrome.tabs.query but fails
    // scripting.executeScript with "Frame ... is showing error page."
    const allMatches=await new Promise(res=>chrome.tabs.query({url:this.URL_PATTERN},res));
    let tab=allMatches&&allMatches.length?(allMatches.find(t=>t.active)||allMatches.find(t=>t.status==='complete')||allMatches[0]):null;
    if(!tab){this.log('No open tab found on fasih-dashboard.bps.go.id/superset/sqllab. Open it there first.','r');return;}

    this.running=true;this.stopRequested=false;this.errorRows=[];
    clearLog('t5');setStatus5('running');updateErrBtn('t5',0);
    document.getElementById('t5-start').disabled=true;document.getElementById('t5-stop').disabled=false;
    if(allMatches.length>1)this.log(`Found ${allMatches.length} matching tabs, using: ${tab.title||tab.url} (tabId ${tab.id})`,'i');
    this.log(`Using tab: ${tab.title||tab.url}`,'s');

    let totalClosed=0,linksOpened=0;
    try{
      totalClosed+=await this.closeAllTabs(tab.id,delay);

      if(!this.stopRequested&&this.linkRows.length){
        this.log(`Loaded reference has ${this.linkRows.length} link${this.linkRows.length===1?'':'s'} — opening each, then closing the tabs it produces.`,'i');
        for(const row of this.linkRows){
          if(this.stopRequested){this.log('Stopped by user.','w');break;}
          this.log(`Opening: ${row.kode||row.link}${row.keterangan?` — ${row.keterangan}`:''}`,'i');
          const ok=await this.openLinkAndWait(tab.id,row.link);
          if(!ok){this.log(`Could not open/load: ${row.link}`,'r');continue;}
          linksOpened++;
          if(this.stopRequested){this.log('Stopped by user.','w');break;}

          const outcome=await this.runQueryWithRetries(tab.id,3);
          if(this.stopRequested||outcome.status==='stopped'){this.log('Stopped by user.','w');break;}
          if(outcome.status==='no-run-button')this.log('Could not find the Run button — skipping results for this link.','r');
          else if(outcome.status==='timeout')this.log('Timed out waiting for the query to finish.','r');
          else if(outcome.status==='error'){
            this.log(`DB engine error persisted after 3 attempts — recording and moving on: "${outcome.text}"`,'r');
            this.errorRows.push({kode:row.kode,keterangan:row.keterangan,link:row.link,error:outcome.text});
            updateErrBtn('t5',this.errorRows.length);
          }
          else if(outcome.status==='warning')this.log(`Query finished with no download available: "${outcome.text}"`,'w');
          else{
            this.log('Results loaded.','s');
            const dl=await this.triggerDownloadCsv(tab.id);
            this.log(dl?'Triggered Download to CSV.':'Could not find the Download to CSV link.',dl?'s':'r');
            if(dl)await sleep(1500); // let the download kick off before closing the tab
          }

          if(this.stopRequested){this.log('Stopped by user.','w');break;}
          totalClosed+=await this.closeAllTabs(tab.id,delay);
        }
      }
    }catch(e){this.log(`Error: ${e.message}`,'r');}

    this.running=false;
    document.getElementById('t5-start').disabled=false;document.getElementById('t5-stop').disabled=true;
    setStatus5(this.stopRequested?'stopped':'done');
    this.log(`Done — ${linksOpened} link${linksOpened===1?'':'s'} opened, ${totalClosed} tab close button${totalClosed===1?'':'s'} clicked in total.`,'s');
  },
  stop(){this.stopRequested=true;this.log('Stop requested...','w');document.getElementById('t5-stop').disabled=true;},

  // Repeatedly clicks the right-most close button until none remain (or
  // stop is requested / the tab is lost). Returns how many were clicked.
  async closeAllTabs(tabId,delay){
    let clicked=0;
    while(true){
      if(this.stopRequested)break;
      const remaining=await this.clickLastTab(tabId);
      if(remaining===null){this.log('Lost the target tab (closed or navigated away).','r');break;}
      if(remaining===0)break;
      clicked++;
      this.log(`Closed 1 tab (${remaining-1} remaining before this click took effect).`,'i');
      await sleep(delay);
    }
    return clicked;
  },

  // Navigates the target tab to a link and waits for it to finish loading.
  // Returns false if the link is missing/invalid or the tab never reaches
  // "complete" within the timeout.
  async openLinkAndWait(tabId,url,timeoutMs=30000){
    if(!/^https?:\/\//i.test(url||''))return false;
    await new Promise((resolve,reject)=>chrome.tabs.update(tabId,{url},(t)=>{
      if(chrome.runtime.lastError){reject(new Error(chrome.runtime.lastError.message));return;}
      resolve(t);
    }));
    await new Promise((resolve)=>{
      let done=false;
      const timer=setTimeout(()=>{if(!done){chrome.tabs.onUpdated.removeListener(listener);done=true;resolve();}},timeoutMs);
      function listener(id,info){
        if(id===tabId&&info.status==='complete'&&!done){
          done=true;clearTimeout(timer);chrome.tabs.onUpdated.removeListener(listener);resolve();
        }
      }
      chrome.tabs.onUpdated.addListener(listener);
    });
    // Give the SPA a moment to finish rendering the new SQL Lab tab after
    // the underlying page load event fires.
    await sleep(1200);
    return true;
  },

  // Runs func inside the target tab and returns its return value.
  async execInTab(tabId,func,args=[]){
    const results=await new Promise((resolve,reject)=>{
      chrome.scripting.executeScript({target:{tabId},func,args},(res)=>{
        if(chrome.runtime.lastError){reject(new Error(chrome.runtime.lastError.message));return;}
        resolve(res);
      });
    });
    if(!results||!results.length)return undefined;
    return results[0].result;
  },

  // Superset's Run button is matched by text, not a stable class —
  // "ant-btn superset-button cta" is shared with other action buttons.
  runButtonPresent(tabId){
    return this.execInTab(tabId,()=>!!Array.from(document.querySelectorAll('button.ant-btn.superset-button.cta')).find(b=>b.textContent.trim()==='Run'&&!b.disabled));
  },
  clickRunButton(tabId){
    return this.execInTab(tabId,()=>{
      const btn=Array.from(document.querySelectorAll('button.ant-btn.superset-button.cta')).find(b=>b.textContent.trim()==='Run');
      if(!btn)return false;
      btn.click();
      return true;
    });
  },
  // Polls for the Run button instead of guessing a fixed delay — the SPA
  // can take a variable amount of time to finish mounting the editor after
  // the underlying page navigation reports "complete."
  async waitForRunButton(tabId,timeoutMs=15000,pollMs=400){
    const deadline=Date.now()+timeoutMs;
    while(Date.now()<deadline){
      if(this.stopRequested)return false;
      if(await this.runButtonPresent(tabId))return true;
      await sleep(pollMs);
    }
    return false;
  },

  // The results panel container (id ends in "-panel-Results", random
  // query-id prefix) can exist in the DOM *while still loading* — it's just
  // a sign the Results tab rendered, not that the query finished. The real
  // "done" signal is either the Download to CSV link (success, has rows)
  // or a *warning/error* ant-alert (no data / a query error). A plain
  // ant-alert-info ("Running statement 1 out of 1") is still in progress
  // and must NOT be treated as terminal, or it bails out mid-query.
  // 'warning' (e.g. "The query returned no data") just means nothing to
  // download — move on. 'error' (e.g. a DB engine error) gets recorded
  // against this link's kode/keterangan before moving on, per request.
  queryOutcome(tabId){
    return this.execInTab(tabId,()=>{
      if(document.querySelector('a[href*="/api/v1/sqllab/export/"]'))return{status:'ok'};
      // The DB engine error box isn't an AntD Alert at all — it's a plain
      // div[role="alert"] with a <strong>DB engine Error</strong> heading
      // and the actual message inside ".error-body".
      const errBox=Array.from(document.querySelectorAll('div[role="alert"]')).find(el=>/db engine error/i.test(el.querySelector('strong')?.textContent||''));
      if(errBox){
        const body=errBox.querySelector('.error-body');
        return{status:'error',text:(body?body.textContent:errBox.textContent).trim()};
      }
      const warn=document.querySelector('.ant-alert-warning .ant-alert-message');
      if(warn)return{status:'warning',text:warn.textContent.trim()};
      return{status:'pending'};
    });
  },

  async waitForResultsPanel(tabId,timeoutMs=60000,pollMs=600){
    const deadline=Date.now()+timeoutMs;
    while(Date.now()<deadline){
      if(this.stopRequested)return{status:'stopped'};
      const outcome=await this.queryOutcome(tabId);
      if(outcome.status!=='pending')return outcome;
      await sleep(pollMs);
    }
    return{status:'timeout'};
  },

  // Clicks Run and waits for an outcome; if it's a DB engine error, re-runs
  // the same query up to maxAttempts times before giving up on it. Only a
  // DB engine error triggers a retry — 'ok', 'warning' (no data), and
  // 'timeout' are all treated as final on the first try.
  async runQueryWithRetries(tabId,maxAttempts=3){
    let last={status:'no-run-button'};
    for(let attempt=1;attempt<=maxAttempts;attempt++){
      if(this.stopRequested)return{status:'stopped'};
      const buttonReady=await this.waitForRunButton(tabId);
      if(this.stopRequested)return{status:'stopped'};
      if(!buttonReady)return{status:'no-run-button'};
      await this.clickRunButton(tabId);
      this.log(attempt===1?'Clicked Run, waiting for results...':`DB engine error — retry ${attempt-1}/${maxAttempts-1}, waiting for results...`,'i');
      last=await this.waitForResultsPanel(tabId);
      if(this.stopRequested||last.status==='stopped')return{status:'stopped'};
      if(last.status!=='error')return last;
      if(attempt<maxAttempts)await sleep(800);
    }
    return last;
  },

  // The export href has a random query id too (".../export/D1jn7Jh7GR/"),
  // so match on the stable path prefix instead.
  //
  // Triggers the download via chrome.downloads instead of clicking the
  // page's <a> — clicking it repeatedly in a row is a *script-triggered*
  // download from the page's own JS, and Chrome's anti-abuse heuristic
  // silently blocks automatic downloads after a few in quick succession.
  // chrome.downloads.download() goes through the extension's own download
  // permission instead, so it isn't subject to that per-page throttle, and
  // it still carries the browser's normal session cookies for the request.
  async triggerDownloadCsv(tabId){
    const href=await this.execInTab(tabId,()=>{
      const a=document.querySelector('a[href*="/api/v1/sqllab/export/"]');
      return a?new URL(a.getAttribute('href'),location.href).href:null;
    });
    if(!href)return false;
    await new Promise((resolve,reject)=>{
      chrome.downloads.download({url:href,saveAs:false},(downloadId)=>{
        if(chrome.runtime.lastError){reject(new Error(chrome.runtime.lastError.message));return;}
        resolve(downloadId);
      });
    });
    return true;
  },

  // Clicks the last matching close button found in the tab right now.
  // Returns the count of matching buttons *before* the click, or null if
  // the tab is gone. Re-queries fresh each call since clicking removes
  // the element and may shift what's "last".
  async clickLastTab(tabId){
    const results=await new Promise((resolve,reject)=>{
      chrome.scripting.executeScript({
        target:{tabId},
        func:(selector)=>{
          const btns=Array.from(document.querySelectorAll(selector));
          if(!btns.length)return 0;
          btns[btns.length-1].click();
          return btns.length;
        },
        args:[this.SELECTOR]
      },(res)=>{
        if(chrome.runtime.lastError){reject(new Error(chrome.runtime.lastError.message));return;}
        resolve(res);
      });
    });
    if(!results||!results.length)return null;
    return results[0].result;
  },

  // ── Link reference loader (kode / keterangan / link columns) ──
  linkRows:[],
  async loadFile(file){
    try{
      const rows=await parseKodeKeteranganLinkFile(file);
      this.linkRows=rows;
      renderLinkTable(rows);
      const info=document.getElementById('t5-file-info');
      info.style.display='block';
      info.innerHTML=`<strong>${rows.length}</strong> row${rows.length===1?'':'s'} loaded from <em>${escHtml(file.name)}</em>`;
      this.log(`Loaded ${rows.length} rows from ${file.name}`,'s');
    }catch(e){
      const info=document.getElementById('t5-file-info');
      info.style.display='block';
      info.innerHTML=`<strong style="color:var(--red);">Failed to load:</strong> ${escHtml(e.message)}`;
      this.log(`File load error: ${e.message}`,'r');
    }
  },

  exportErrors(){
    if(!this.errorRows.length)return;
    const headers=['kode','keterangan','link','error'];
    downloadBlob(new Blob([buildCSV(headers,this.errorRows)],{type:'text/csv;charset=utf-8'}),`sqllab_errors_${nowTs()}.csv`);
    this.log(`Exported ${this.errorRows.length} DB engine error${this.errorRows.length===1?'':'s'}`,'s');
  },
};

function renderLinkTable(rows){
  const wrap=document.getElementById('t5-link-wrap'),table=document.getElementById('t5-link-table');
  if(!rows.length){wrap.style.display='none';table.innerHTML='';return;}
  const esc=escHtml;
  const safeHref=link=>/^https?:\/\//i.test(link)?esc(link):'';
  const head='<thead><tr><th>kode</th><th>keterangan</th><th>link</th></tr></thead>';
  const body='<tbody>'+rows.map(r=>{
    const href=safeHref(r.link);
    const linkCell=href?`<a href="${href}" target="_blank" rel="noopener noreferrer">${esc(r.link)}</a>`:esc(r.link);
    return `<tr><td>${esc(r.kode)}</td><td>${esc(r.keterangan)}</td><td>${linkCell}</td></tr>`;
  }).join('')+'</tbody>';
  table.innerHTML=head+body;
  wrap.style.display='block';
}

// ────────────────────────────────────────────────────────────────────────────
// Minimal file parsers — .csv/.txt directly; for .xlsx it unzips and parses
// the worksheet XML by hand (no SheetJS bundled), since this project only
// ever needed to *write* xlsx until now (see buildXLSXFromArray above).
// ────────────────────────────────────────────────────────────────────────────
async function readTableFromFile(file){
  const ext=file.name.split('.').pop().toLowerCase();
  const table=(ext==='xlsx'||ext==='xls')?await parseXlsxToTable(await file.arrayBuffer()):parseDelimitedText(await file.text());
  if(!table.length)throw new Error('File is empty.');
  return table;
}

function requireColumns(headerRow,names){
  const colIdx={};
  for(const name of names){
    const idx=headerRow.indexOf(name);
    if(idx===-1)throw new Error(`Missing required column "${name}". Found: ${headerRow.filter(Boolean).join(', ')||'(no header row detected)'}`);
    colIdx[name]=idx;
  }
  return colIdx;
}

// The "kode / keterangan / link" reference sheet (SQL Lab tab).
async function parseKodeKeteranganLinkFile(file){
  const table=await readTableFromFile(file);
  const headerRow=table[0].map(h=>String(h??'').trim().toLowerCase());
  const colIdx=requireColumns(headerRow,['kode','keterangan','link']);
  const out=[];
  for(let i=1;i<table.length;i++){
    const r=table[i];
    if(!r||r.every(v=>v===undefined||v===null||v===''))continue;
    out.push({
      kode:String(r[colIdx.kode]??'').trim(),
      keterangan:String(r[colIdx.keterangan]??'').trim(),
      link:String(r[colIdx.link]??'').trim()
    });
  }
  return out;
}

// A region-id chain sheet (Assignment Status tab) — at least one of
// region1Id..region6Id, same shape T1.getDeepestRegionChains() produces, so
// this can come straight from a previously exported error CSV to retry just
// the failed chains.
async function parseRegionIdChainsFile(file){
  const table=await readTableFromFile(file);
  const headerRow=table[0].map(h=>String(h??'').trim().toLowerCase());
  const colIdx={};
  for(let l=1;l<=6;l++){const idx=headerRow.indexOf(`region${l}id`);if(idx!==-1)colIdx[l]=idx;}
  if(!Object.keys(colIdx).length)throw new Error(`No region1Id..region6Id columns found. Found: ${headerRow.filter(Boolean).join(', ')||'(no header row detected)'}`);

  const out=[];
  for(let i=1;i<table.length;i++){
    const r=table[i];
    if(!r||r.every(v=>v===undefined||v===null||v===''))continue;
    const chain={};let deepest=null;
    for(const[l,idx]of Object.entries(colIdx)){
      const val=String(r[idx]??'').trim();
      if(val){chain[Number(l)]=val;deepest=Math.max(deepest??0,Number(l));}
    }
    if(deepest===null)continue;
    out.push({chain,deepest});
  }
  return out;
}

// Assignment List export (assignmentId, data1, data2, ...) used as an
// offline lookup table for the Anomaly tab's enrichment step — same shape
// T2.exportCSV/exportXLSX produce, so a previous Assignment List run's
// export can be fed straight back in instead of re-hitting the live API.
async function parseAssignmentLookupFile(file){
  const ext=file.name.split('.').pop().toLowerCase();
  const map=new Map();
  if(ext==='json'){
    let data=JSON.parse(await file.text());
    if(data&&!Array.isArray(data)){for(const k of['data','results','rows','items']){if(Array.isArray(data[k])){data=data[k];break;}}}
    if(!Array.isArray(data))data=[data];
    for(const r of data){
      if(!r||typeof r!=='object')continue;
      const idKey=Object.keys(r).find(k=>/^assignment.?id$/i.test(k));
      if(!idKey)continue;
      const id=String(r[idKey]??'').trim();if(!id)continue;
      const d1Key=Object.keys(r).find(k=>/^data1$/i.test(k));
      const d2Key=Object.keys(r).find(k=>/^data2$/i.test(k));
      map.set(id,{data1:String(r[d1Key]??''),data2:String(r[d2Key]??'')});
    }
    return map;
  }
  const table=await readTableFromFile(file);
  const headerRow=table[0].map(h=>String(h??'').trim().toLowerCase());
  const colIdx=requireColumns(headerRow,['assignmentid','data1','data2']);
  for(let i=1;i<table.length;i++){
    const r=table[i];
    if(!r||r.every(v=>v===undefined||v===null||v===''))continue;
    const id=String(r[colIdx.assignmentid]??'').trim();
    if(!id)continue;
    map.set(id,{data1:String(r[colIdx.data1]??''),data2:String(r[colIdx.data2]??'')});
  }
  return map;
}

function parseDelimitedText(text){
  const lines=text.replace(/\r/g,'').split('\n').filter(l=>l.length>0);
  const delim=lines[0]&&lines[0].includes('\t')&&!lines[0].includes(',')?'\t':',';
  const parseLine=line=>{
    const out=[];let cur='',inQ=false;
    for(let i=0;i<line.length;i++){
      const ch=line[i];
      if(inQ){
        if(ch==='"'){if(line[i+1]==='"'){cur+='"';i++;}else inQ=false;}
        else cur+=ch;
      }else{
        if(ch==='"')inQ=true;
        else if(ch===delim){out.push(cur);cur='';}
        else cur+=ch;
      }
    }
    out.push(cur);
    return out;
  };
  return lines.map(parseLine);
}

// ── Hand-rolled xlsx reader: unzip via local file headers (no central
// directory walk needed for well-formed, non-streamed zips like Excel's
// output) + DecompressionStream('deflate-raw') for inflate, then a small
// XML read of sharedStrings.xml and the first worksheet. ──
async function parseXlsxToTable(arrayBuffer){
  const entries=await unzipLocalEntries(arrayBuffer);
  const dec=new TextDecoder();
  const sheetBytes=entries['xl/worksheets/sheet1.xml'];
  if(!sheetBytes)throw new Error('Could not find a worksheet inside this xlsx file.');
  const sharedBytes=entries['xl/sharedStrings.xml'];
  const shared=sharedBytes?parseSharedStringsXml(dec.decode(sharedBytes)):[];
  return parseWorksheetXml(dec.decode(sheetBytes),shared);
}

async function unzipLocalEntries(arrayBuffer){
  const view=new DataView(arrayBuffer),bytes=new Uint8Array(arrayBuffer);
  const entries={};
  let pos=0;
  while(pos+30<=bytes.length&&view.getUint32(pos,true)===0x04034b50){
    const method=view.getUint16(pos+8,true);
    const compSize=view.getUint32(pos+18,true);
    const nameLen=view.getUint16(pos+26,true);
    const extraLen=view.getUint16(pos+28,true);
    const nameStart=pos+30;
    const name=new TextDecoder().decode(bytes.subarray(nameStart,nameStart+nameLen));
    const dataStart=nameStart+nameLen+extraLen;
    const dataEnd=dataStart+compSize;
    const raw=bytes.subarray(dataStart,Math.min(dataEnd,bytes.length));
    if(method===0)entries[name]=raw;
    else if(method===8)entries[name]=await inflateRawRead(raw);
    pos=dataEnd;
  }
  return entries;
}
async function inflateRawRead(bytes){
  if(!bytes.length)return new Uint8Array(0);
  const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
function colIndexFromCellRef(ref){
  const m=/^([A-Z]+)/.exec(ref||'');if(!m)return 0;
  let col=0;for(const ch of m[1])col=col*26+(ch.charCodeAt(0)-64);
  return col-1;
}
function parseSharedStringsXml(xml){
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  return Array.from(doc.getElementsByTagName('si')).map(si=>Array.from(si.getElementsByTagName('t')).map(t=>t.textContent).join(''));
}
function parseWorksheetXml(xml,shared){
  const doc=new DOMParser().parseFromString(xml,'application/xml');
  const rows=[];
  for(const rowEl of Array.from(doc.getElementsByTagName('row'))){
    const row=[];
    for(const c of Array.from(rowEl.children)){
      const idx=colIndexFromCellRef(c.getAttribute('r'));
      const type=c.getAttribute('t');
      const vEl=c.getElementsByTagName('v')[0];
      const isEl=c.getElementsByTagName('is')[0];
      let value='';
      if(type==='s'&&vEl)value=shared[parseInt(vEl.textContent,10)]??'';
      else if(type==='inlineStr'&&isEl)value=Array.from(isEl.getElementsByTagName('t')).map(t=>t.textContent).join('');
      else if(vEl)value=vEl.textContent;
      row[idx]=value;
    }
    rows.push(row);
  }
  return rows;
}
function setStatus5(state){
  const b=document.getElementById('t5-badge');
  b.textContent={idle:'Idle',running:'Running',done:'Done',stopped:'Stopped'}[state]||state;
  b.className=`badge ${state}`;
}

// ────────────────────────────────────────────────────────────────────────────
// Shared UI helpers
// ────────────────────────────────────────────────────────────────────────────
const LOG_BUF={t1:[],t2:[],t3:[],t4:[],t5:[],t6:[],t7:[]};
function addLog(tab,msg,cls='i'){const ts=new Date().toLocaleTimeString('en-GB');LOG_BUF[tab].unshift({ts,msg,cls});if(LOG_BUF[tab].length>300)LOG_BUF[tab].length=300;renderLog(tab);}
function clearLog(tab){LOG_BUF[tab]=[];renderLog(tab);}
function renderLog(tab){const el=document.getElementById(`${tab}-log-area`);if(!el)return;el.innerHTML=LOG_BUF[tab].map(e=>`<div class="le l${e.cls}"><span class="lts">${e.ts}</span>${escHtml(e.msg)}</div>`).join('');}
function setStatus(tab,state){const b=document.getElementById(`${tab}-badge`);b.textContent={idle:'Idle',running:'Running',done:'Done',stopped:'Stopped'}[state]||state;b.className=`badge ${state}`;document.getElementById(`${tab}-start`).disabled=state==='running';document.getElementById(`${tab}-stop`).disabled=state!=='running';}
function updateErrBtn(tab,count){
  const btn=document.getElementById(`${tab}-err-exp`);if(btn){btn.disabled=count===0;btn.textContent=count>0?`\u26a0 ${count} failed`:'\u26a0 No errors';}
  const retryBtn=document.getElementById(`${tab}-retry-err`);if(retryBtn)retryBtn.disabled=count===0;
}
function updateStats(tab,s){
  const set=(id,v)=>{const el=document.getElementById(id);if(el)el.textContent=v;};
  if(tab==='t1'){set('t1-req',s.req);set('t1-rows',s.rows);set('t1-err',s.err);set('t1-act',s.active);}
  if(tab==='t2'){set('t2-total',s.total);set('t2-done',s.done);set('t2-found',s.found);set('t2-err',s.err);set('t2-act',s.active);const p=document.getElementById('t2-prog');if(p)p.style.width=(s.total>0?s.done/s.total*100:0)+'%';}
  if(tab==='t3'){set('t3-total',s.total);set('t3-done',s.done);set('t3-err',s.err);set('t3-act',s.active);const p=document.getElementById('t3-prog');if(p)p.style.width=(s.total>0?s.done/s.total*100:0)+'%';}
  updateErrBtn(tab,s.err);
}
function showDL(tab,count){document.getElementById(`${tab}-dl`).style.display='block';document.getElementById(`${tab}-dl-count`).textContent=count;}
function hideDL(tab){document.getElementById(`${tab}-dl`).style.display='none';}
function saveStorage(key,data){try{chrome.storage.local.set({[key]:data});}catch(_){}}
function loadStorage(){
  try{chrome.storage.local.get(['t1data'],res=>{
    if(!res?.t1data?.rows?.length)return;
    const{level2FullCode,collectedAt,rows}=res.t1data;T1.rows=rows;T1.stats.rows=rows.length;
    const prev=document.getElementById('t1-prev');prev.style.display='block';
    document.getElementById('t1-prev-info').textContent=`Previous: ${level2FullCode} — ${rows.length} rows (${new Date(collectedAt).toLocaleString()})`;
    document.getElementById('t1-load-prev').onclick=()=>{showDL('t1',rows.length);T2.refreshInfo();T6.refreshInfo();addLog('t1',`Loaded ${rows.length} rows from storage`,'s');};
    T2.refreshInfo();T6.refreshInfo();
  });}catch(_){}
}

// ────────────────────────────────────────────────────────────────────────────
// INIT
// ────────────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', async ()=>{
  await DB.open();// initialise IndexedDB first

  // Restore DB counts on load
  for(const tab of['t2','t3','t6']){
    const store=`${tab}_results`;
    const n=await DB.count(store);
    if(n>0){showDL(tab,n);addLog(tab,`Found ${n} rows in local database from previous session.`,'s');}
  }
  for(const tab of['t2','t3']){
    const errCount=await DB.count(`${tab}_errors`);
    if(errCount>0)updateErrBtn(tab,errCount);
  }

  // Nav switching (vertical sidebar)
  document.querySelectorAll('.nav-btn').forEach(btn=>btn.addEventListener('click',()=>{
    document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c=>c.classList.remove('active'));
    btn.classList.add('active');document.getElementById(btn.dataset.tab).classList.add('active');
    if(btn.dataset.tab==='t2')T2.refreshInfo();if(btn.dataset.tab==='t3')T3.refreshInfo();if(btn.dataset.tab==='t6')T6.refreshInfo();if(btn.dataset.tab==='t7'){renderLog('t7');T7.refreshTab2Info();}
  }));

  document.getElementById('exp-cancel').addEventListener('click',()=>{exportCancelled=true;hideExportProgress();});

  // Tab 1
  document.getElementById('t1-start').addEventListener('click',()=>{if(!T1.collecting)T1.start();});
  document.getElementById('t1-stop').addEventListener('click',()=>T1.stop());
  document.getElementById('t1-json').addEventListener('click',()=>T1.exportJSON());
  document.getElementById('t1-csv').addEventListener('click',()=>T1.exportCSV());
  document.getElementById('t1-xlsx').addEventListener('click',()=>T1.exportXLSX());
  document.getElementById('t1-err-exp').addEventListener('click',()=>T1.exportErrors());

  // Tab 2
  document.getElementById('t2-start').addEventListener('click',()=>{if(!T2.collecting)T2.start();});
  document.getElementById('t2-stop').addEventListener('click',()=>T2.stop());
  document.getElementById('t2-json').addEventListener('click',()=>T2.exportJSON());
  document.getElementById('t2-csv').addEventListener('click',()=>T2.exportCSV());
  document.getElementById('t2-xlsx').addEventListener('click',()=>T2.exportXLSX());
  document.getElementById('t2-err-exp').addEventListener('click',()=>T2.exportErrors());
  document.getElementById('t2-retry-err').addEventListener('click',()=>{if(!T2.collecting)T2.retryErrors();});
  document.getElementById('t2-clear-db').addEventListener('click',()=>T2.clearDB());

  // Tab 3
  document.getElementById('t3-start').addEventListener('click',()=>{if(!T3.collecting)T3.start();});
  document.getElementById('t3-stop').addEventListener('click',()=>T3.stop());
  document.getElementById('t3-json').addEventListener('click',()=>T3.exportJSON());
  document.getElementById('t3-csv').addEventListener('click',()=>T3.exportCSV());
  document.getElementById('t3-xlsx').addEventListener('click',()=>T3.exportXLSX());
  document.getElementById('t3-err-exp').addEventListener('click',()=>T3.exportErrors());
  document.getElementById('t3-clear-db').addEventListener('click',()=>T3.clearDB());

  // Tab 6
  document.getElementById('t6-start').addEventListener('click',()=>{if(!T6.running)T6.start();});
  document.getElementById('t6-stop').addEventListener('click',()=>T6.stop());
  document.getElementById('t6-json').addEventListener('click',()=>T6.exportJSON());
  document.getElementById('t6-csv').addEventListener('click',()=>T6.exportCSV());
  document.getElementById('t6-xlsx').addEventListener('click',()=>T6.exportXLSX());
  document.getElementById('t6-err-exp').addEventListener('click',()=>T6.exportErrors());
  document.getElementById('t6-clear-db').addEventListener('click',()=>T6.clearDB());

  function makeSourceToggle(srcBtnId,fileBtnId,fromSrcId,fromFileId,obj,primaryMode,dropId,inputId){
    document.getElementById(srcBtnId).addEventListener('click',()=>{obj.sourceMode=primaryMode;document.getElementById(srcBtnId).classList.add('active');document.getElementById(fileBtnId).classList.remove('active');document.getElementById(fromSrcId).style.display='';document.getElementById(fromFileId).style.display='none';});
    document.getElementById(fileBtnId).addEventListener('click',()=>{obj.sourceMode='file';document.getElementById(fileBtnId).classList.add('active');document.getElementById(srcBtnId).classList.remove('active');document.getElementById(fromFileId).style.display='';document.getElementById(fromSrcId).style.display='none';});
    const drop=document.getElementById(dropId),input=document.getElementById(inputId);
    drop.addEventListener('click',()=>input.click());
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('drag');});
    drop.addEventListener('dragleave',()=>drop.classList.remove('drag'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('drag');const f=e.dataTransfer.files[0];if(f){drop.textContent=f.name;obj.loadFile(f);}});
    input.addEventListener('change',()=>{const f=input.files[0];if(f){drop.textContent=f.name;obj.loadFile(f);}});
  }
  makeSourceToggle('t2-src-tab1','t2-src-file','t2-from-tab1','t2-from-file',T2,'tab1','t2-drop','t2-file-input');
  makeSourceToggle('t3-src-tab2','t3-src-file','t3-from-tab2','t3-from-file',T3,'tab2','t3-drop','t3-file-input');
  makeSourceToggle('t6-src-tab1','t6-src-file','t6-from-tab1','t6-from-file',T6,'tab1','t6-drop','t6-file-input');

  document.getElementById('t3-mode-data').addEventListener('click',()=>{T3.exportMode='data';document.getElementById('t3-mode-data').classList.add('active');document.getElementById('t3-mode-predef').classList.remove('active');});
  document.getElementById('t3-mode-predef').addEventListener('click',()=>{T3.exportMode='predef';document.getElementById('t3-mode-predef').classList.add('active');document.getElementById('t3-mode-data').classList.remove('active');});

  // Tab 4
  document.getElementById('t4-payload').value=JSON.stringify(T4.DEFAULT_PAYLOAD,null,2);
  document.getElementById('t4-start').addEventListener('click',()=>T4.start());
  document.getElementById('t4-stop').addEventListener('click',()=>T4.stop());
  document.getElementById('t4-resume-btn').addEventListener('click',()=>{
    if(!T4.pendingResume)return;
    T4.resumeState=T4.pendingResume;T4.hideResumeBanner();T4.start();
  });
  document.getElementById('t4-startover-btn').addEventListener('click',async()=>{
    if(!confirm('Discard the records collected so far and start over from page 1?'))return;
    await T4._clearProgress();T4.hideResumeBanner();
  });
  document.getElementById('t4-json').addEventListener('click',()=>T4.exportJSON());
  document.getElementById('t4-csv').addEventListener('click',()=>T4.exportCSV());
  document.getElementById('t4-xlsx').addEventListener('click',()=>T4.exportXLSX());
  document.getElementById('t4-level-sel').addEventListener('change',()=>T4.updateDlSummary());
  T4.checkResumable();

  // Tab 5
  document.getElementById('t5-start').addEventListener('click',()=>T5.start());
  document.getElementById('t5-stop').addEventListener('click',()=>T5.stop());
  document.getElementById('t5-err-exp').addEventListener('click',()=>T5.exportErrors());
  {
    const drop=document.getElementById('t5-drop'),input=document.getElementById('t5-file-input');
    drop.addEventListener('click',()=>input.click());
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('drag');});
    drop.addEventListener('dragleave',()=>drop.classList.remove('drag'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('drag');const f=e.dataTransfer.files[0];if(f)T5.loadFile(f);});
    input.addEventListener('change',()=>{const f=input.files[0];if(f)T5.loadFile(f);});
  }

  // Tab 7
  document.getElementById('t7-fetch').addEventListener('click',()=>T7.fetch());
  document.getElementById('t7-stop').addEventListener('click',()=>T7.stop());
  document.getElementById('t7-type-dl').addEventListener('change',()=>T7.updateDlCount());
  document.getElementById('t7-json').addEventListener('click',()=>T7.exportJSON());
  document.getElementById('t7-csv').addEventListener('click',()=>T7.exportCSV());
  document.getElementById('t7-xlsx').addEventListener('click',()=>T7.exportXLSX());

  {
    const srcBtns={api:'t7-enrich-src-api',tab2:'t7-enrich-src-tab2',file:'t7-enrich-src-file'};
    const panels={api:'t7-enrich-from-api',tab2:'t7-enrich-from-tab2',file:'t7-enrich-from-file'};
    Object.entries(srcBtns).forEach(([mode,btnId])=>{
      document.getElementById(btnId).addEventListener('click',()=>{
        T7.enrichSource=mode;
        Object.entries(srcBtns).forEach(([m,id])=>document.getElementById(id).classList.toggle('active',m===mode));
        Object.entries(panels).forEach(([m,id])=>document.getElementById(id).style.display=m===mode?'':'none');
        if(mode==='tab2')T7.refreshTab2Info();
      });
    });
    const drop=document.getElementById('t7-enrich-drop'),input=document.getElementById('t7-enrich-file-input');
    drop.addEventListener('click',()=>input.click());
    drop.addEventListener('dragover',e=>{e.preventDefault();drop.classList.add('drag');});
    drop.addEventListener('dragleave',()=>drop.classList.remove('drag'));
    drop.addEventListener('drop',e=>{e.preventDefault();drop.classList.remove('drag');const f=e.dataTransfer.files[0];if(f)T7.loadEnrichFile(f);});
    input.addEventListener('change',()=>{const f=input.files[0];if(f)T7.loadEnrichFile(f);});
  }

  ['t1','t2','t3','t4'].forEach(t=>setStatus(t,'idle'));
  setStatus5('idle');
  setStatusT6('idle');
  T6.refreshInfo();
  loadStorage();
});
