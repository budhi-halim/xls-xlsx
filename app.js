'use strict';
(()=>{
const q=id=>document.getElementById(id),ui={input:q('file-input'),zone:q('dropzone'),list:q('file-list'),count:q('queue-count'),clear:q('clear-btn'),convert:q('convert-btn'),label:q('convert-label'),zip:q('zip-btn'),progress:q('overall-progress'),fill:q('progress-fill'),track:q('progress-track'),percent:q('progress-percent'),progressLabel:q('progress-label'),message:q('message')};
const jobs=[],formats={biff:'Binary Excel XLS',xml:'Excel 2003 XML',html:'HTML spreadsheet',xlsx:'XLSX'};let sequence=0,busy=false,worker=null,readyPromise=null;
function el(tag,cls,txt){const e=document.createElement(tag);if(cls)e.className=cls;if(txt!=null)e.textContent=txt;return e}
function msg(s,error){ui.message.textContent=s||'';ui.message.hidden=!s;ui.message.classList.toggle('error',!!error)}
function pretty(n){return n<1024?n+' B':n<1048576?(n/1024).toFixed(1)+' KB':(n/1048576).toFixed(1)+' MB'}
function detect(b){
 if(b.length>=8&&[208,207,17,224,161,177,26,225].every((x,i)=>b[i]===x))return 'biff';
 if(b.length>=4&&b[0]===80&&b[1]===75&&[3,5,7].includes(b[2]))return 'xlsx';
 if(b.length>=4&&b[0]===9&&[0,2,4,8].includes(b[1]))return 'biff';
 let t;try{t=(b[0]===255&&b[1]===254?new TextDecoder('utf-16le'):b[0]===254&&b[1]===255?new TextDecoder('utf-16be'):new TextDecoder('utf-8')).decode(b).replace(/^\uFEFF/,'').trimStart()}catch{return null}
 if(/<(?:[\w-]+:)?Workbook\b/i.test(t.slice(0,8000))&&/schemas-microsoft-com:office:spreadsheet/i.test(t))return 'xml';
 if(/^(?:<!doctype\s+html[^>]*>\s*)?<html\b/i.test(t)||/^<table\b/i.test(t))return 'html';return null;
}
function unique(name,used){let result=name,i=2;while(used.has(result.toLowerCase()))result=name.slice(0,-5)+' ('+(i++)+').xlsx';used.add(result.toLowerCase());return result}
function showProgress(n,total,description){ui.progress.hidden=false;const p=Math.round(100*n/Math.max(1,total));ui.fill.style.width=p+'%';ui.track.setAttribute('aria-valuenow',String(p));ui.percent.textContent=p+'%';ui.progressLabel.textContent=description}
function draw(){
 ui.list.replaceChildren();ui.count.textContent=jobs.length;ui.clear.hidden=!jobs.length||busy;let pending=0,done=0,checking=0;
 jobs.forEach(j=>{
  if(j.status==='pending'||j.status==='error')pending++;if(j.status==='done')done++;if(j.status==='detecting')checking++;
  const row=el('div','file-row'),symbol=el('div','file-symbol','XLS'),center=el('div','file-center'),title=el('div','file-name',j.file.name),meta=el('div','file-meta'),actions=el('div','file-actions');
  title.title=j.file.name;meta.append(el('span','',pretty(j.file.size)),el('span','','·'),el('span','',j.format?formats[j.format]:'Checking format'));
  if(j.status==='done')meta.append(el('span','','·'),el('span','',j.sheets+' sheets'));
  if(j.status==='error')meta.append(el('span','status error',j.error));
  center.append(title,meta);
  if(j.status==='done'){const button=el('button','small-button','Download XLSX');button.type='button';button.onclick=()=>download(j);actions.append(button);if(j.warnings.length)actions.append(el('span','status warning','Review warnings'))}
  else actions.append(el('span',j.status==='error'?'status error':'status',j.status==='error'?'Failed':j.status==='working'?(j.stage||'Converting'):j.status==='pending'?'Ready':'Checking'));
  if(!busy){const remove=el('button','remove-button','×');remove.type='button';remove.title='Remove file';remove.setAttribute('aria-label','Remove '+j.file.name);remove.onclick=()=>{jobs.splice(jobs.indexOf(j),1);draw()};actions.append(remove)}
  row.append(symbol,center,actions);ui.list.append(row)
 });
 if(!jobs.length)ui.list.append(el('p','empty','No files selected.'));
 ui.convert.disabled=busy||!pending||!!checking;ui.label.textContent=busy?'Converting…':pending?'Convert '+pending+' file'+(pending===1?'':'s'):done?'Conversion complete':'Convert files';ui.zip.hidden=busy||done<2;
}
async function addFiles(files){
 if(busy)return msg('Wait until conversion finishes before adding files.',true);msg('');ui.progress.hidden=true;
 const names=new Set(jobs.map(j=>j.out.toLowerCase())),selected=Array.from(files||[]).slice(0,Math.max(0,60-jobs.length)),added=[];
 for(const file of selected){
  if(!/\.(xls|xml|xlsx)$/i.test(file.name)||jobs.some(j=>j.file.name===file.name&&j.file.size===file.size&&j.file.lastModified===file.lastModified))continue;
  const stem=file.name.replace(/\.(xls|xml|xlsx)$/i,'').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,120);
  const j={id:++sequence,file,out:unique((stem||'workbook')+'.xlsx',names),status:'detecting',format:null,error:'',warnings:[],bytes:null};jobs.push(j);added.push(j);
 }draw();
 for(const j of added){try{if(j.file.size>75*1048576)throw Error('File exceeds 75 MB.');j.format=detect(new Uint8Array(await j.file.slice(0,16384).arrayBuffer()));if(!j.format)throw Error('Unsupported file content.');j.status='pending'}catch(e){j.status='error';j.error=String(e.message||e)}draw()}
 if(!added.length)msg('No new supported files were added.',true)
}
function initWorker(){
 if(readyPromise)return readyPromise;
 readyPromise=new Promise((resolve,reject)=>{
  try{worker=new Worker('./worker.js')}catch(e){reject(e);return}
  const timeout=setTimeout(()=>{clear();reject(Error('Spreadsheet library load timed out.'))},35000);
  function clear(){clearTimeout(timeout);worker.removeEventListener('message',listener)}
  function listener(e){if(e.data?.type==='ready'){clear();resolve(worker)}else if(e.data?.type==='engine-error'){clear();reject(Error(e.data.message))}}
  worker.addEventListener('message',listener);worker.addEventListener('error',()=>{clear();reject(Error('Conversion worker failed to start.'))},{once:true});
 }).catch(e=>{worker?.terminate();worker=null;readyPromise=null;throw e});return readyPromise
}
function convertOne(w,j,buffer){
 return new Promise((resolve,reject)=>{
  const timeout=setTimeout(()=>{clear();reject(Error('Conversion timed out.'))},120000);
  function clear(){clearTimeout(timeout);w.removeEventListener('message',listener);w.removeEventListener('error',crash)}
  function crash(){clear();reject(Error('Conversion worker crashed.'))}
  function listener(e){const d=e.data;if(d?.id!==j.id)return;if(d.type==='stage'){j.stage=d.stage;draw()}else if(d.type==='result'){clear();resolve(d)}else if(d.type==='failure'){clear();reject(Error(d.error))}}
  w.addEventListener('message',listener);w.addEventListener('error',crash);w.postMessage({id:j.id,format:j.format,buffer},[buffer])
 })
}
async function convert(){
 if(busy)return;const targets=jobs.filter(j=>j.format&&(j.status==='pending'||j.status==='error'));if(!targets.length)return;
 busy=true;draw();msg('');let finished=0,failed=0,warning=0;
 try{let w=await initWorker();for(const j of targets){
  j.status='working';j.stage='Reading';draw();showProgress(finished,targets.length,'Processing '+j.file.name);
  try{const result=await convertOne(w,j,await j.file.arrayBuffer());j.bytes=new Uint8Array(result.buffer);j.sheets=result.sheets||0;j.warnings=result.warnings||[];j.status='done';if(j.warnings.length)warning++}
  catch(e){failed++;j.status='error';j.error=String(e.message||e).slice(0,200);if(/crashed|timed out/i.test(j.error)){w.terminate();readyPromise=null;worker=null;if(finished+1<targets.length)w=await initWorker()}}
  finished++;showProgress(finished,targets.length,finished+' of '+targets.length+' processed');draw()
 }msg((finished-failed)+' converted'+(failed?', '+failed+' failed':'')+'.'+(warning?' '+warning+' file(s) require review.':''))}
 catch(e){msg(String(e.message||e),true)}finally{busy=false;draw()}
}
function save(blob,name){const url=URL.createObjectURL(blob),a=el('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000)}
function download(j){if(!j.bytes)return;save(new Blob([j.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),j.out);if(j.warnings.length)msg(j.out+': '+j.warnings.join(' ').slice(0,280))}
async function downloadZip(){
 const complete=jobs.filter(j=>j.status==='done'&&j.bytes);if(complete.length<2)return;ui.zip.disabled=true;ui.zip.textContent='Preparing ZIP…';
 try{
  if(typeof JSZip==='undefined')await new Promise((resolve,reject)=>{const s=el('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';s.onload=resolve;s.onerror=()=>reject(Error('ZIP library unavailable; download files individually.'));document.head.append(s)});
  const zip=new JSZip();complete.forEach(j=>zip.file(j.out,j.bytes));save(await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:4}}),'converted-xlsx.zip')
 }catch(e){msg(String(e.message||e),true)}finally{ui.zip.disabled=false;ui.zip.textContent='Download all (.zip)'}
}
ui.input.onchange=async e=>{await addFiles(e.target.files);ui.input.value=''};ui.clear.onclick=()=>{if(busy)return;jobs.length=0;msg('');ui.progress.hidden=true;draw()};ui.convert.onclick=convert;ui.zip.onclick=downloadZip;
['dragover','dragenter'].forEach(t=>ui.zone.addEventListener(t,e=>{e.preventDefault();ui.zone.classList.add('dragging')}));
['dragleave','dragend'].forEach(t=>ui.zone.addEventListener(t,e=>{e.preventDefault();ui.zone.classList.remove('dragging')}));
ui.zone.addEventListener('drop',e=>{e.preventDefault();ui.zone.classList.remove('dragging');addFiles(e.dataTransfer.files)});draw()
})();