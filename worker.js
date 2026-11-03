'use strict';
try{
 importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
 if(typeof XLSX==='undefined')throw Error('Spreadsheet engine unavailable.');
 self.postMessage({type:'ready'});
}catch(e){self.postMessage({type:'engine-error',message:String(e.message||e)})}
function signature(ws){
 return Object.keys(ws).filter(a=>a[0]!=='!'&&ws[a]&&(ws[a].v!==undefined||ws[a].f)).sort().map(a=>[a,String(ws[a].v??''),String(ws[a].f||'')]);
}
function verify(original,converted){
 const warns=[];
 if(original.SheetNames.length!==converted.SheetNames.length)warns.push('Worksheet count differs.');
 for(const name of original.SheetNames){
  const a=original.Sheets[name],b=converted.Sheets[name];
  if(!b){warns.push('Missing worksheet: '+name);continue}
  const aa=signature(a),bb=signature(b);
  if(aa.length!==bb.length)warns.push(name+': cell count differs.');
  if(aa.length<=300000&&bb.length<=300000){
   let diff=0;for(let i=0;i<Math.max(aa.length,bb.length);i++){if(JSON.stringify(aa[i])!==JSON.stringify(bb[i])){diff++;if(diff===3)break}}
   if(diff)warns.push(name+': differences in cell values or formulas.');
  }else warns.push(name+': complete cell comparison skipped due to size.');
  if((a['!merges']||[]).length!==(b['!merges']||[]).length)warns.push(name+': merge count differs.');
 }
 return warns
}
self.onmessage=e=>{
 const {id,buffer,format}=e.data||{};if(!id||!buffer)return;
 try{
  self.postMessage({id,type:'stage',stage:'Reading workbook'});
  const original=XLSX.read(new Uint8Array(buffer),{type:'array',cellFormula:true,cellStyles:true,cellDates:false,cellNF:true,sheetStubs:false});
  if(!original.SheetNames.length)throw Error('Workbook has no worksheets.');
  if(format==='xlsx'){self.postMessage({id,type:'result',buffer:buffer.slice(0),sheets:original.SheetNames.length,warnings:[]});return}
  self.postMessage({id,type:'stage',stage:'Writing XLSX'});
  const output=XLSX.write(original,{bookType:'xlsx',type:'array',compression:true,cellStyles:true});
  self.postMessage({id,type:'stage',stage:'Verifying'});
  const test=XLSX.read(output,{type:'array',cellFormula:true,cellStyles:true,cellDates:false,sheetStubs:false});
  const warnings=verify(original,test);
  // SheetJS writes a Uint8Array. postMessage transfer lists accept ArrayBuffer,
  // not Uint8Array, so transfer the exact underlying byte range.
  const exactBytes = output instanceof Uint8Array ? output : new Uint8Array(output);
  const resultBuffer = (exactBytes.byteOffset === 0 && exactBytes.byteLength === exactBytes.buffer.byteLength)
    ? exactBytes.buffer : exactBytes.slice().buffer;
  self.postMessage({id,type:'result',buffer:resultBuffer,sheets:test.SheetNames.length,warnings},[resultBuffer]);
 }catch(err){self.postMessage({id,type:'failure',error:String(err?.message||err).slice(0,450)})}
};