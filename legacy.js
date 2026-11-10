/* Excel 2003 XML and HTML-table XLS converter. No network or third-party dependencies. */
'use strict';
(() => {
  const NS = 'urn:schemas-microsoft-com:office:spreadsheet';
  const enc = new TextEncoder();
  const esc = value => String(value ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  const attr = (node,name) => node.getAttributeNS?.(NS,name) ?? node.getAttribute?.('ss:'+name) ?? node.getAttribute?.(name);
  const children = (node,tag) => Array.from(node?.children||[]).filter(x=>x.localName===tag);
  const first = (node,tag) => children(node,tag)[0]||null;
  const rgb = s => { if(!s)return null; const m=String(s).trim().match(/^#([0-9A-Fa-f]{6})$/);return m?m[1].toUpperCase():null; };
  const xlCol = i => {let s='';for(let k=i+1;k;k=Math.floor((k-1)/26))s=String.fromCharCode(65+(k-1)%26)+s;return s;};
  const loc = (r,c)=>xlCol(c)+(r+1);
  const num = s => {
    let v=String(s||'').trim().replace(/\u00a0/g,' ');
    const neg=/^\s*\(/.test(v)||/^-/.test(v);v=v.replace(/[^0-9.,]/g,'');
    if(v.includes(',')&&v.includes('.')){
      v=v.lastIndexOf(',')>v.lastIndexOf('.')?v.replace(/\./g,'').replace(',','.'):v.replace(/,/g,'');
    } else if(v.includes(',')) {
      const parts=v.split(',');v=parts.length===2&&parts[1].length<=2?parts.join('.'):parts.join('');
    } else if((v.match(/\./g)||[]).length>1) v=v.replace(/\./g,'');
    else if(/\.\d{3}$/.test(v))v=v.replace('.','');
    const n=Number(v);return Number.isFinite(n)?(neg?-n:n):null;
  };
  const stable = v => JSON.stringify(v||{});
  class Styles {
    constructor(){this.fonts=[{name:'Calibri',size:11}];this.fills=[null,{pattern:'gray125'}];this.borders=[{}];this.fmts=[];this.xfs=[{fontId:0,fillId:0,borderId:0,numFmtId:0}];this.keys=new Map();this.fontKeys=new Map([[stable(this.fonts[0]),0]]);this.fillKeys=new Map();this.borderKeys=new Map([[stable({}),0]]);this.fmtKeys=new Map();}
    addUnique(arr,map,x){const key=stable(x);if(map.has(key))return map.get(key);const i=arr.length;arr.push(x);map.set(key,i);return i;}
    style(s={}){
      const key=stable(s);if(this.keys.has(key))return this.keys.get(key);
      const f=s.font||this.fonts[0],fo=this.addUnique(this.fonts,this.fontKeys,f);
      const fill=s.fill?this.addUnique(this.fills,this.fillKeys,s.fill):0;
      const b=s.border?this.addUnique(this.borders,this.borderKeys,s.border):0;
      let nf=0;if(s.numFmt&&s.numFmt!=='General') {if(!this.fmtKeys.has(s.numFmt)){this.fmtKeys.set(s.numFmt,164+this.fmts.length);this.fmts.push(s.numFmt);}nf=this.fmtKeys.get(s.numFmt);}
      const i=this.xfs.length;this.xfs.push({fontId:fo,fillId:fill,borderId:b,numFmtId:nf,alignment:s.alignment||null});this.keys.set(key,i);return i;
    }
    xml(){
      const font=f=>`<font><sz val="${f.size||11}"/><name val="${esc(f.name||'Calibri')}"/>${f.bold?'<b/>':''}${f.italic?'<i/>':''}${f.underline?'<u/>':''}${f.color?`<color rgb="FF${f.color}"/>`:''}</font>`;
      const fill=f=>f?.pattern==='gray125'?'<fill><patternFill patternType="gray125"/></fill>':!f?'<fill><patternFill patternType="none"/></fill>':`<fill><patternFill patternType="solid"><fgColor rgb="FF${f.color}"/><bgColor indexed="64"/></patternFill></fill>`;
      const border=b=>`<border>${['left','right','top','bottom'].map(e=>b[e]?`<${e} style="thin"><color rgb="FF${b[e]}"/></${e}>`:`<${e}/>`).join('')}<diagonal/></border>`;
      const xf=x=>`<xf numFmtId="${x.numFmtId}" fontId="${x.fontId}" fillId="${x.fillId}" borderId="${x.borderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1" applyAlignment="1">${x.alignment?`<alignment${x.alignment.horizontal?` horizontal="${x.alignment.horizontal}"`:''}${x.alignment.vertical?` vertical="${x.alignment.vertical}"`:''}${x.alignment.wrapText?' wrapText="1"':''}/>`:''}</xf>`;
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${this.fmts.length}">${this.fmts.map((s,i)=>`<numFmt numFmtId="${i+164}" formatCode="${esc(s)}"/>`).join('')}</numFmts><fonts count="${this.fonts.length}">${this.fonts.map(font).join('')}</fonts><fills count="${this.fills.length}">${this.fills.map(fill).join('')}</fills><borders count="${this.borders.length}">${this.borders.map(border).join('')}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${this.xfs.length}">${this.xfs.map(xf).join('')}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
    }
  }
  function parseStyleXML(el){
    const s={};const ft=first(el,'Font'),interior=first(el,'Interior'),al=first(el,'Alignment'),nf=first(el,'NumberFormat'),br=first(el,'Borders');
    if(ft)s.font={name:attr(ft,'FontName')||'Calibri',size:Number(attr(ft,'Size'))||10,bold:attr(ft,'Bold')==='1',italic:attr(ft,'Italic')==='1',color:rgb(attr(ft,'Color'))||'000000'};
    if(interior&&rgb(attr(interior,'Color')))s.fill={color:rgb(attr(interior,'Color'))};
    if(al){const h=(attr(al,'Horizontal')||'').toLowerCase(),v=(attr(al,'Vertical')||'').toLowerCase();s.alignment={horizontal:({center:'center',right:'right',left:'left',justify:'justify'})[h],vertical:v==='center'?'center':v||undefined,wrapText:attr(al,'WrapText')==='1'};}
    if(nf&&attr(nf,'Format'))s.numFmt=attr(nf,'Format');
    if(br){const b={};for(const x of children(br,'Border')){const p=(attr(x,'Position')||'').toLowerCase(),color=rgb(attr(x,'Color'))||'D9D9D9';if(['top','left','bottom','right'].includes(p))b[p]=color;}if(Object.keys(b).length)s.border=b;}
    return s;
  }
  function readXml(text){
    const doc=new DOMParser().parseFromString(text,'application/xml');if(doc.querySelector('parsererror'))throw Error('Invalid Excel 2003 XML.');
    if(doc.documentElement.localName!=='Workbook')throw Error('Not an Excel 2003 XML workbook.');
    const definitions={};for(const style of Array.from(doc.getElementsByTagNameNS(NS,'Style')))definitions[attr(style,'ID')]=parseStyleXML(style);
    const sheets=[];
    for(const src of Array.from(doc.getElementsByTagNameNS(NS,'Worksheet'))){
      const table=first(src,'Table');if(!table)continue;
      const w={name:attr(src,'Name')||'Sheet',rows:[],merges:[],widths:{}};
      let ci=0;for(const c of children(table,'Column')){const index=attr(c,'Index')?Number(attr(c,'Index'))-1:ci;const width=Number(attr(c,'Width'));if(width>0)w.widths[index]=Math.min(65,Math.max(5,width*1.333/7));ci=index+1;}
      let ri=0;for(const row of children(table,'Row')){
        const current=attr(row,'Index')?Number(attr(row,'Index'))-1:ri;const record={index:current,height:Number(attr(row,'Height'))||null,cells:[]};let col=0;
        for(const c of children(row,'Cell')){
          const currentCol=attr(c,'Index')?Number(attr(c,'Index'))-1:col;const data=first(c,'Data'),raw=data?data.textContent||'':'',t=attr(data,'Type')||'String',style=definitions[attr(c,'StyleID')]||{},numeric=t==='Number';
          const value=numeric?Number(raw):raw;
          record.cells.push({c:currentCol,value:Number.isFinite(value)||!numeric?value:raw,numeric,style,formula:attr(c,'Formula')||null});
          const across=+(attr(c,'MergeAcross')||0),down=+(attr(c,'MergeDown')||0);
          if(across||down)w.merges.push({r:current,c:currentCol,rs:down+1,cs:across+1});col=currentCol+across+1;
        }
        w.rows.push(record);ri=current+1;
      }
      sheets.push(w);
    }
    if(!sheets.length)throw Error('No worksheets found.');return sheets;
  }
  function cssDeclarations(text){const out={};for(const chunk of String(text||'').split(';')){const i=chunk.indexOf(':');if(i>0)out[chunk.slice(0,i).trim().toLowerCase()]=chunk.slice(i+1).trim().replace(/!important/i,'');}return out;}
  function htmlStyleMap(doc){const m={};for(const tag of doc.querySelectorAll('style')){
    const data=(tag.textContent||'').replace(/\/\*[\s\S]*?\*\//g,'');const re=/([^{}]+)\{([^{}]*)\}/g;let result;
    while((result=re.exec(data))){const css=cssDeclarations(result[2]);for(const key of result[1].split(',')){const k=key.trim().toLowerCase();if(k)Object.assign(m[k]||(m[k]={}),css);}}
  }return m;}
  function readHtml(text){
    const doc=new DOMParser().parseFromString(text,'text/html'),table=doc.querySelector('table');if(!table)throw Error('No HTML table found.');
    const rules=htmlStyleMap(doc),sheets=[{name:'Monitoring',rows:[],merges:[],widths:{}}],sheet=sheets[0];
    const occupied=new Set();let nr=0;
    for(const tr of Array.from(table.rows)){
      const cells=Array.from(tr.cells);if(!cells.length)continue;
      const row={index:nr,height:null,cells:[]};let col=0;
      for(const cell of cells){while(occupied.has(nr+':'+col))col++;
        const colspan=Math.max(1,Math.min(40,parseInt(cell.getAttribute('colspan'),10)||1)),rowspan=Math.max(1,Math.min(1000,parseInt(cell.getAttribute('rowspan'),10)||1));
        const css={...(rules.body||{}),...(rules[cell.tagName.toLowerCase()]||{})};
        const classes=(cell.getAttribute('class')||'').split(/\s+/).filter(Boolean);
        for(const cls of classes)Object.assign(css,rules['.'+cls.toLowerCase()]||{});
        Object.assign(css,cssDeclarations(cell.getAttribute('style')));
        const style={font:{name:'Calibri',size:10,bold:cell.tagName==='TH'},alignment:{vertical:'center'}};
        const align=(css['text-align']||'').trim().toLowerCase();if(['center','left','right'].includes(align))style.alignment.horizontal=align;
        const color=rgb(css.color),bg=rgb(css['background-color']);if(color)style.font.color=color;if(bg)style.fill={color:bg};
        if(css['font-weight']==='bold'||+css['font-weight']>=600)style.font.bold=true;
        const fontsize=/([\d.]+)(pt|px)/.exec(css['font-size']||'');if(fontsize)style.font.size=+fontsize[1]*(fontsize[2]==='px'?.75:1);
        const border=(css.border||''),edge=rgb((border.match(/#[\da-fA-F]{6}/)||[])[0]);
        if(edge)style.border={top:edge,left:edge,right:edge,bottom:edge};
        if(css['vertical-align']==='top')style.alignment.vertical='top';
        const raw=(cell.textContent||'').replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
        const isNum=classes.includes('num')||!!cell.getAttribute('x:num');
        const parsed=isNum?num(raw):null;const numeric=isNum&&parsed!==null&&raw!=='';
        if(numeric)style.numFmt='"Rp. "#,##0.00';
        else if(classes.includes('str')||classes.includes('text-cell'))style.numFmt='@';
        row.cells.push({c:col,value:numeric?parsed:raw,numeric,style});
        if(colspan>1||rowspan>1){sheet.merges.push({r:nr,c:col,rs:rowspan,cs:colspan});for(let rr=nr;rr<nr+rowspan;rr++)for(let cc=col;cc<col+colspan;cc++)if(rr!==nr||cc!==col)occupied.add(rr+':'+cc);}
        col+=colspan;
      }
      sheet.rows.push(row);nr++;
    }
    const widths=[6,14,29,12,25,12,19,19,13,20,20,22];widths.forEach((v,i)=>sheet.widths[i]=v);
    return sheets;
  }
  function renderSheet(sheet,styles){
    const count=sheet.rows.length;const maxCol=Math.max(0,...sheet.rows.flatMap(r=>r.cells.map(c=>c.c)),...sheet.merges.map(m=>m.c+m.cs-1));
    const cols=Object.keys(sheet.widths).map(k=>`<col min="${+k+1}" max="${+k+1}" width="${sheet.widths[k]}" customWidth="1"/>`).join('');
    const rows=sheet.rows.map(r=>{
      const cells=r.cells.map(c=>{
        const a=`r="${loc(r.index,c.c)}" s="${styles.style(c.style)}"`;
        if(c.numeric && typeof c.value==='number')return `<c ${a}><v>${Number.isFinite(c.value)?c.value:0}</v></c>`;
        if(c.formula&&/^=/.test(c.formula)&&!c.formula.includes('R['))return `<c ${a}><f>${esc(c.formula.slice(1))}</f>${c.numeric?`<v>${c.value}</v>`:''}</c>`;
        return `<c ${a} t="inlineStr"><is><t xml:space="preserve">${esc(c.value)}</t></is></c>`;
      }).join('');return `<row r="${r.index+1}"${r.height?` ht="${r.height}" customHeight="1"`:''}>${cells}</row>`;
    }).join('');
    const merges=sheet.merges.length?`<mergeCells count="${sheet.merges.length}">${sheet.merges.map(m=>`<mergeCell ref="${loc(m.r,m.c)}:${loc(m.r+m.rs-1,m.c+m.cs-1)}"/>`).join('')}</mergeCells>`:'';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${loc(Math.max(0,count-1),maxCol)}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"/></sheetViews><sheetFormatPr defaultRowHeight="16"/><cols>${cols}</cols><sheetData>${rows}</sheetData>${merges}</worksheet>`;
  }
  function crc32(bytes){let c=0xffffffff;for(const b of bytes){c^=b;for(let i=0;i<8;i++)c=c&1?(c>>>1)^0xedb88320:c>>>1;}return (c^0xffffffff)>>>0;}
  function zipFiles(entries){
    const locparts=[],central=[];let pos=0;const u16=(d,o,v)=>d.setUint16(o,v,true),u32=(d,o,v)=>d.setUint32(o,v>>>0,true);
    for(const [name,data] of entries){const nb=enc.encode(name),bytes=enc.encode(data),crc=crc32(bytes),head=new Uint8Array(30+nb.length),dv=new DataView(head.buffer);
      u32(dv,0,0x04034b50);u16(dv,4,20);u16(dv,6,0x0800);u16(dv,8,0);u32(dv,14,crc);u32(dv,18,bytes.length);u32(dv,22,bytes.length);u16(dv,26,nb.length);head.set(nb,30);locparts.push(head,bytes);
      const cent=new Uint8Array(46+nb.length),cv=new DataView(cent.buffer);u32(cv,0,0x02014b50);u16(cv,4,20);u16(cv,6,20);u16(cv,8,0x0800);u16(cv,10,0);u32(cv,16,crc);u32(cv,20,bytes.length);u32(cv,24,bytes.length);u16(cv,28,nb.length);u32(cv,42,pos);cent.set(nb,46);central.push(cent);pos+=head.length+bytes.length;
    }
    const centralSize=central.reduce((n,x)=>n+x.length,0),end=new Uint8Array(22),dv=new DataView(end.buffer);u32(dv,0,0x06054b50);u16(dv,8,entries.length);u16(dv,10,entries.length);u32(dv,12,centralSize);u32(dv,16,pos);
    const result=new Uint8Array(pos+centralSize+end.length);let off=0;for(const part of [...locparts,...central,end]){result.set(part,off);off+=part.length;}return result;
  }
  function writeXlsx(sheets){
    const styles=new Styles();const xmls=sheets.map(s=>renderSheet(s,styles));
    const types=`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${xmls.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`;
    const relations=`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`;
    const book=`<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${sheets.map((s,i)=>`<sheet name="${esc(s.name.replace(/[\\\/\[\]\*\?:]/g,' ').slice(0,31)||'Sheet')}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`;
    const rels=`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${sheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
    return zipFiles([['[Content_Types].xml',types],['_rels/.rels',relations],['xl/workbook.xml',book],['xl/_rels/workbook.xml.rels',rels],['xl/styles.xml',styles.xml()],...xmls.map((x,i)=>[`xl/worksheets/sheet${i+1}.xml`,x])]);
  }
  function decode(bytes){if(bytes[0]===0xff&&bytes[1]===0xfe)return new TextDecoder('utf-16le').decode(bytes);if(bytes[0]===0xfe&&bytes[1]===0xff)return new TextDecoder('utf-16be').decode(bytes);return new TextDecoder('utf-8').decode(bytes).replace(/^\uFEFF/,'');}
  function detect(bytes){
    if(bytes.length>=8&&[208,207,17,224,161,177,26,225].every((x,i)=>bytes[i]===x))return 'biff';
    if(bytes.length>=4&&bytes[0]===80&&bytes[1]===75)return 'xlsx';
    const t=decode(bytes.slice(0,20000)).trimStart();
    if(/<(?:[\w-]+:)?Workbook\b/i.test(t.slice(0,10000))&&/schemas-microsoft-com:office:spreadsheet/i.test(t))return 'xml';
    if(/<html\b/i.test(t.slice(0,4000))||/<table\b/i.test(t.slice(0,4000)))return 'html';
    return null;
  }
  function convert(buffer,format){const b=new Uint8Array(buffer),kind=format||detect(b);if(kind!=='xml'&&kind!=='html')throw Error('Not a supported legacy text workbook.');const source=decode(b);const sheets=kind==='xml'?readXml(source):readHtml(source);return {bytes:writeXlsx(sheets),sheets:sheets.length,format:kind};}
  globalThis.LegacyConverter={detect,convert};
})();
