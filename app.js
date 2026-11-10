'use strict';
(()=>{
 const $=id=>document.getElementById(id),ui={input:$('file-input'),drop:$('dropzone'),list:$('file-list'),count:$('queue-count'),add:$('add-btn'),clear:$('clear-btn'),convert:$('convert-btn'),zip:$('zip-btn'),progress:$('overall-progress'),fill:$('progress-fill'),track:$('progress-track'),percent:$('progress-percent'),label:$('progress-label'),message:$('message')};
 const jobs=[];let nextId=0,busy=false,worker=null,initPromise=null;
 const names={xml:'Excel 2003 XML',html:'HTML spreadsheet',biff:'Binary XLS',xlsx:'XLSX'};
 function el(tag,cls,text){const x=document.createElement(tag);if(cls)x.className=cls;if(text!==undefined)x.textContent=text;return x;}
 function size(n){return n<1048576?`${(n/1024).toFixed(1)} KB`:`${(n/1048576).toFixed(1)} MB`;}
 function message(text='',error=false){ui.message.textContent=text;ui.message.hidden=!text;ui.message.classList.toggle('error',error);}
 function progress(done,total,label){ui.progress.hidden=false;const n=Math.round(done/Math.max(1,total)*100);ui.track.setAttribute('aria-valuenow',String(n));ui.fill.style.width=n+'%';ui.percent.textContent=n+'%';ui.label.textContent=label;}
 function downloadBlob(data,name){const blob=new Blob([data],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),url=URL.createObjectURL(blob),a=el('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);}
 function render(){
   ui.list.replaceChildren();ui.count.textContent=jobs.length;ui.clear.hidden=busy||!jobs.length;ui.add.hidden=busy||!jobs.length;
   const waiting=jobs.filter(x=>x.status==='ready'||x.status==='error').length,loading=jobs.some(x=>x.status==='checking'),completed=jobs.filter(x=>x.status==='done').length;
   ui.convert.disabled=busy||!waiting||loading;ui.convert.textContent=busy?'Converting…':waiting?'Convert '+waiting+' file'+(waiting===1?'':'s'):'Convert to XLSX';ui.zip.hidden=busy||completed<2;
   if(!jobs.length)ui.list.append(el('p','empty','No files selected'));
   for(const job of jobs){
    const row=el('div','file-row'),icon=el('div','file-icon','XLS'),info=el('div','file-info'),name=el('div','file-name',job.file.name),meta=el('div','file-meta',`${size(job.file.size)} · ${job.type?names[job.type]:'Checking format'}`),actions=el('div','file-controls');name.title=job.file.name;
    if(job.status==='error'){meta.textContent=job.error;meta.classList.add('error');actions.append(el('span','status error','Failed'));}
    if(job.status==='checking')actions.append(el('span','status','Checking'));
    if(job.status==='ready')actions.append(el('span','status','Ready'));
    if(job.status==='working')actions.append(el('span','status',job.stage||'Converting'));
    if(job.status==='done'){const btn=el('button','download','Download');btn.type='button';btn.addEventListener('click',()=>downloadBlob(job.bytes,job.output));actions.append(btn);}
    if(!busy){const remove=el('button','remove','×');remove.type='button';remove.setAttribute('aria-label','Remove '+job.file.name);remove.addEventListener('click',()=>{jobs.splice(jobs.indexOf(job),1);render();});actions.append(remove);}
    info.append(name,meta);row.append(icon,info,actions);ui.list.append(row);
   }
 }
 function outputName(filename){return filename.replace(/\.(xls|xml|xlsx)$/i,'').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,120)+'.xlsx';}
 async function add(files){
  if(busy)return;message('');ui.progress.hidden=true;
  const added=[],existing=new Set(jobs.map(j=>j.output.toLowerCase()));
  for(const f of Array.from(files||[]).slice(0,Math.max(0,60-jobs.length))){
   if(!/\.(xls|xml|xlsx)$/i.test(f.name))continue;
   if(jobs.some(x=>x.file.name===f.name&&x.file.size===f.size&&x.file.lastModified===f.lastModified))continue;
   let candidate=outputName(f.name),k=2;while(existing.has(candidate.toLowerCase()))candidate=outputName(f.name).slice(0,-5)+` (${k++}).xlsx`;existing.add(candidate.toLowerCase());
   const j={id:++nextId,file:f,output:candidate,type:null,status:'checking',bytes:null,error:''};jobs.push(j);added.push(j);
  }
  render();for(const j of added){try{if(j.file.size>75*1048576)throw Error('File exceeds 75 MB');const bytes=new Uint8Array(await j.file.slice(0,20000).arrayBuffer());j.type=LegacyConverter.detect(bytes);if(!j.type)throw Error('Unsupported file format');j.status='ready';}catch(e){j.status='error';j.error=String(e.message||e);}render();}
 }
 function workerReady(){
  if(initPromise)return initPromise;
  initPromise=new Promise((resolve,reject)=>{
    try{worker=new Worker('./worker.js?v=3');}catch(e){reject(e);return;}
    const timer=setTimeout(()=>{cleanup();reject(Error('Spreadsheet engine unavailable'));},30000);
    function cleanup(){clearTimeout(timer);worker.removeEventListener('message',onmessage);}
    function onmessage(e){if(e.data?.type==='ready'){cleanup();resolve(worker);}if(e.data?.type==='engine-error'){cleanup();reject(Error(e.data.message));}}
    worker.addEventListener('message',onmessage);worker.addEventListener('error',()=>{cleanup();reject(Error('Conversion worker failed'));},{once:true});
  }).catch(e=>{worker?.terminate();worker=null;initPromise=null;throw e;});return initPromise;
 }
 function runWorker(w,j,buffer){return new Promise((resolve,reject)=>{
   const timer=setTimeout(()=>{cleanup();reject(Error('Conversion timed out'));},120000);
   function cleanup(){clearTimeout(timer);w.removeEventListener('message',response);w.removeEventListener('error',failed);}
   function failed(){cleanup();reject(Error('Conversion worker failed'));}
   function response(e){const d=e.data;if(d?.id!==j.id)return;if(d.type==='result'){cleanup();resolve(new Uint8Array(d.buffer));}else if(d.type==='failure'){cleanup();reject(Error(d.error));}else if(d.type==='stage'){j.stage=d.stage;render();}}
   w.addEventListener('message',response);w.addEventListener('error',failed);w.postMessage({id:j.id,buffer,format:j.type},[buffer]);
 });}
 async function convert(){
  if(busy)return;const pending=jobs.filter(j=>j.status==='ready'||j.status==='error');if(!pending.length)return;
  busy=true;message('');render();let completed=0,failed=0;
  for(const j of pending){j.status='working';j.stage='Converting';render();progress(completed,pending.length,`${completed+1} of ${pending.length}`);
   try{
    const buffer=await j.file.arrayBuffer();let bytes;
    if(j.type==='xml'||j.type==='html'){
       await new Promise(r=>setTimeout(r,0)); // allow the progress indicator to paint
       bytes=LegacyConverter.convert(buffer,j.type).bytes;
    }else if(j.type==='xlsx'){bytes=new Uint8Array(buffer);}
    else {const w=await workerReady();bytes=await runWorker(w,j,buffer);}
    j.bytes=bytes;j.status='done';j.error='';
   }catch(e){j.status='error';j.error=String(e.message||e);failed++;}
   completed++;progress(completed,pending.length,`${completed} of ${pending.length}`);render();
  }
  busy=false;render();message(failed?`${pending.length-failed} converted, ${failed} failed`: `${completed} file${completed===1?'':'s'} converted`,false);
 }
 async function zip(){
   const done=jobs.filter(x=>x.status==='done');if(done.length<2)return;ui.zip.disabled=true;ui.zip.textContent='Preparing ZIP…';
   try{if(!window.JSZip)await new Promise((resolve,reject)=>{const s=el('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';s.onload=resolve;s.onerror=()=>reject(Error('ZIP library unavailable'));document.head.append(s);});
    const z=new JSZip();for(const j of done)z.file(j.output,j.bytes);const archive=await z.generateAsync({type:'blob',compression:'DEFLATE'});const url=URL.createObjectURL(archive),a=el('a');a.href=url;a.download='converted-xlsx.zip';document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
   }catch(e){message(e.message||String(e),true);}finally{ui.zip.disabled=false;ui.zip.textContent='Download ZIP';}
 }
 ui.input.onchange=async e=>{await add(e.target.files);e.target.value='';};ui.add.onclick=()=>ui.input.click();ui.clear.onclick=()=>{if(!busy){jobs.length=0;ui.progress.hidden=true;message('');render();}};ui.convert.onclick=convert;ui.zip.onclick=zip;
 ['dragenter','dragover'].forEach(t=>ui.drop.addEventListener(t,e=>{e.preventDefault();ui.drop.classList.add('dragging');}));
 ['dragleave','dragend'].forEach(t=>ui.drop.addEventListener(t,e=>{e.preventDefault();ui.drop.classList.remove('dragging');}));
 ui.drop.addEventListener('drop',e=>{e.preventDefault();ui.drop.classList.remove('dragging');add(e.dataTransfer.files);});render();
})();
