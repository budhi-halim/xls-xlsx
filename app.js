'use strict';
(() => {
  const get = id => document.getElementById(id);
  const ui = {
    input:get('file-input'), drop:get('dropzone'), list:get('file-list'),
    count:get('queue-count'), add:get('add-btn'), clear:get('clear-btn'),
    convert:get('convert-btn'), zip:get('zip-btn'), progress:get('overall-progress'),
    fill:get('progress-fill'), track:get('progress-track'),
    percent:get('progress-percent'), label:get('progress-label'),
    message:get('message')
  };
  const files = [];
  const typeNames = {xml:'Excel 2003 XML',html:'HTML report',biff:'Binary XLS',xlsx:'XLSX'};
  let serial = 0, busy = false, pythonReady = typeof window.convertWithPython === 'function';
  window.addEventListener('python-converter-ready', () => {
    pythonReady = true;
    message('');
    render();
  });

  function element(tag,cls,text) {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text !== undefined) node.textContent = text;
    return node;
  }
  const size = n => n < 1048576 ? (n/1024).toFixed(1)+' KB' : (n/1048576).toFixed(1)+' MB';
  const outputName = s => (s.replace(/\.(xls|xml|xlsx)$/i,'').replace(/[\\/:*?"<>|\x00-\x1f]/g,'_').slice(0,120)||'workbook')+'.xlsx';
  function message(text='',error=false) {
    ui.message.textContent=text;
    ui.message.classList.toggle('error',!!error);
    ui.message.hidden=!text;
  }
  function stage(done,total) {
    const value = Math.round(done/Math.max(1,total)*100);
    ui.progress.hidden=false;
    ui.fill.style.width=value+'%';
    ui.track.setAttribute('aria-valuenow',String(value));
    ui.percent.textContent=value+'%';
    ui.label.textContent=done+' / '+total;
  }
  function detect(bytes) {
    const magic=[208,207,17,224,161,177,26,225];
    if (magic.every((b,i)=>bytes[i]===b)) return 'biff';
    if (bytes[0]===80 && bytes[1]===75) return 'xlsx';
    const decoder = bytes[0]===255 && bytes[1]===254 ? new TextDecoder('utf-16le'):
      bytes[0]===254 && bytes[1]===255 ? new TextDecoder('utf-16be'):new TextDecoder('utf-8');
    const text=decoder.decode(bytes).replace(/^\uFEFF/,'').trimStart().toLowerCase();
    if (text.includes('schemas-microsoft-com:office:spreadsheet') && /<(?:[\w-]+:)?workbook\b/.test(text)) return 'xml';
    if (text.includes('<html') || text.includes('<table')) return 'html';
    if (bytes[0]===9 && [0,2,4,8].includes(bytes[1])) return 'biff';
    return null;
  }
  function saveBytes(bytes,name) {
    const blob=new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    saveBlob(blob,name);
  }
  function saveBlob(blob,name) {
    const href=URL.createObjectURL(blob);
    const a=element('a');
    a.href=href;a.download=name;
    document.body.append(a);a.click();a.remove();
    setTimeout(()=>URL.revokeObjectURL(href),60000);
  }
  function fromBase64(base64) {
    const binary=atob(base64), out=new Uint8Array(binary.length);
    for(let i=0;i<binary.length;i++) out[i]=binary.charCodeAt(i);
    return out;
  }
  function render() {
    ui.list.replaceChildren();
    ui.count.textContent=String(files.length);
    ui.clear.hidden=busy||!files.length;
    ui.add.hidden=busy||!files.length;
    const pending=files.filter(f=>f.status==='ready'||f.status==='error').length;
    const done=files.filter(f=>f.status==='done').length;
    ui.convert.disabled=busy||!pythonReady||!pending||files.some(f=>f.status==='checking');
    ui.convert.textContent=busy?'Converting…':!pythonReady?'Loading converter…':pending?'Convert '+pending+' file'+(pending===1?'':'s'):'Convert to XLSX';
    ui.zip.hidden=busy||done<2;
    if(!files.length)ui.list.append(element('p','empty','No files'));
    for(const file of files) {
      const row=element('div','file-row'),icon=element('div','file-icon','XLS'),details=element('div','file-info');
      const title=element('div','file-name',file.file.name);
      title.title=file.file.name;
      const hint=element('div','file-meta',file.status==='error'?file.error:
        size(file.file.size)+' · '+(typeNames[file.type]||'Checking'));
      if(file.status==='error')hint.classList.add('error');
      details.append(title,hint);
      const controls=element('div','file-controls');
      if(file.status==='done'){
        const btn=element('button','download','Download');
        btn.type='button';btn.onclick=()=>saveBytes(file.bytes,file.output);
        controls.append(btn);
      } else controls.append(element('span',file.status==='error'?'status error':'status',
        file.status==='checking'?'Checking':file.status==='ready'?'Ready':file.status==='working'?'Converting':'Failed'));
      if(!busy){
        const remove=element('button','remove','×');
        remove.type='button';
        remove.setAttribute('aria-label','Remove '+file.file.name);
        remove.onclick=()=>{files.splice(files.indexOf(file),1);render();};
        controls.append(remove);
      }
      row.append(icon,details,controls);ui.list.append(row);
    }
  }
  async function add(selected) {
    if(busy)return;
    message('');ui.progress.hidden=true;
    const used=new Set(files.map(x=>x.output.toLowerCase())),batch=[];
    for(const f of Array.from(selected||[]).slice(0,Math.max(0,60-files.length))){
      if(!/\.(xls|xlsx|xml)$/i.test(f.name))continue;
      if(files.some(j=>j.file.name===f.name&&j.file.size===f.size&&j.file.lastModified===f.lastModified))continue;
      let name=outputName(f.name),seq=2;
      while(used.has(name.toLowerCase()))name=outputName(f.name).slice(0,-5)+' ('+(seq++)+').xlsx';
      used.add(name.toLowerCase());
      const job={id:++serial,file:f,output:name,status:'checking',type:null,error:'',bytes:null};
      files.push(job);batch.push(job);
    }
    render();
    for(const job of batch){
      try{
        if(job.file.size>75*1048576)throw Error('File exceeds 75 MB');
        job.type=detect(new Uint8Array(await job.file.slice(0,35000).arrayBuffer()));
        if(!job.type)throw Error('Unsupported file format');
        job.status='ready';
      }catch(e){job.status='error';job.error=String(e.message||e);}
      render();
    }
  }
  async function convert() {
    if(busy || !pythonReady)return;
    const queued=files.filter(j=>(j.status==='ready'||j.status==='error')&&j.type);
    if(!queued.length)return;
    busy=true;message('');render();stage(0,queued.length);
    let errors=0;
    for(let n=0;n<queued.length;n++){
      const item=queued[n];
      item.status='working';render();
      try{
        const result=JSON.parse(await window.convertWithPython(item.file));
        if(!result.base64)throw Error('No XLSX data returned');
        const bytes=fromBase64(result.base64);
        if(bytes[0]!==80||bytes[1]!==75)throw Error('Invalid XLSX result');
        item.bytes=bytes;
        item.status='done';
        item.error='';
      }catch(e){
        item.status='error';
        item.error=(e?.message||String(e)).slice(0,250);
        errors++;
      }
      stage(n+1,queued.length);render();
      await new Promise(r=>setTimeout(r,0));
    }
    busy=false;render();
    if(errors)message((queued.length-errors)+' converted, '+errors+' failed',true);
  }
  async function downloadZip() {
    const complete=files.filter(j=>j.status==='done');
    if(complete.length<2)return;
    ui.zip.disabled=true;
    try{
      if(!window.JSZip)await new Promise((resolve,reject)=>{
        const s=element('script');s.src='https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
        s.onload=resolve;s.onerror=()=>reject(Error('Unable to load ZIP library'));
        document.head.append(s);
      });
      const zip=new window.JSZip();
      for(const file of complete)zip.file(file.output,file.bytes);
      saveBlob(await zip.generateAsync({type:'blob',compression:'DEFLATE'}),'converted-xlsx.zip');
    }catch(e){message(e.message||String(e),true);}
    finally{ui.zip.disabled=false;}
  }
  ui.input.addEventListener('change',async e=>{await add(e.target.files);e.target.value='';});
  ui.add.onclick=()=>ui.input.click();
  ui.clear.onclick=()=>{if(busy)return;files.length=0;ui.progress.hidden=true;message('');render();};
  ui.convert.onclick=convert;
  ui.zip.onclick=downloadZip;
  for(const event of ['dragenter','dragover'])ui.drop.addEventListener(event,e=>{e.preventDefault();ui.drop.classList.add('dragging');});
  for(const event of ['dragleave','dragend'])ui.drop.addEventListener(event,e=>{e.preventDefault();ui.drop.classList.remove('dragging');});
  ui.drop.addEventListener('drop',e=>{e.preventDefault();ui.drop.classList.remove('dragging');add(e.dataTransfer.files);});
  render();
})();
