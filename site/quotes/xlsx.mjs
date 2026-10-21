import {validate} from './core.mjs';
import {validateIO,ioLayout,ioContact} from './workflow.mjs?v=3';
import {ioTemplate} from './io-template.mjs?v=3';

// Edit the uploaded OOXML package in place so logos, signatures and cell styles survive.
// No private workbook is shipped with the site. Only the browser holds its bytes.
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const all=(root,tag)=>Array.from(root.getElementsByTagNameNS('*',tag));
const one=(root,tag)=>all(root,tag)[0];
const xml=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(one(d,'parsererror'))throw new Error('엑셀 XML을 읽을 수 없어요.');return d;};
const serialize=d=>new XMLSerializer().serializeToString(d);
const text=(root,tag)=>all(root,tag).map(n=>n.textContent).join('');
export async function readWorkbook(bytes,kind){
  if(!globalThis.JSZip)throw new Error('엑셀 도구를 불러오지 못했어요. 새로고침해 주세요.');
  if(bytes.byteLength>10*1024*1024)throw new Error('10MB 이하의 xlsx 파일을 선택해 주세요.');
  let zip;try{zip=await JSZip.loadAsync(bytes);}catch{throw new Error('암호가 없는 .xlsx 파일을 선택해 주세요.');}
  const wbPath='xl/workbook.xml', wb=xml(await zip.file(wbPath)?.async('string')||'');
  const sheets=all(wb,'sheet');
  if(sheets.length!==1)throw new Error('첨부한 원본처럼 시트가 1개인 견적서 또는 장비리스트를 선택해 주세요.');
  const rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  const rid=sheets[0].getAttribute('r:id');
  const rel=all(rels,'Relationship').find(r=>r.getAttribute('Id')===rid);
  const path=(rel?.getAttribute('Target')||'').replace(/^\//,'').replace(/^(?!xl\/)/,'xl/');
  const entry=zip.file(path);if(!entry)throw new Error('시트 파일을 찾을 수 없어요.');
  const doc=xml(await entry.async('string'));
  const shared=zip.file('xl/sharedStrings.xml')?all(xml(await zip.file('xl/sharedStrings.xml').async('string')),'si'):[];
  const strings=shared.map(si=>text(si,'t'));
  const cells=new Map(all(doc,'c').map(c=>[c.getAttribute('r'),c]));
  const get=ref=>{const c=cells.get(ref);if(!c)return '';const v=one(c,'v')?.textContent??'';if(c.getAttribute('t')==='s')return strings[+v]??'';if(c.getAttribute('t')==='inlineStr')return text(c,'t');return v===''?'':Number.isFinite(+v)?+v:v;};
  const isQuote=String(get('C1')).includes('견적서')&&String(get('F9')).includes('단가');
  const isGear=String(get('C1')).includes('장비 리스트')&&String(get('E9')).includes('수량');
  if(kind==='quote'&&!isQuote||kind==='equipment'&&!isGear)throw new Error(kind==='quote'?'견적서 양식 파일을 선택해 주세요.':'장비리스트 파일을 선택해 주세요.');
  // These coordinates are the provided 2026 template contract; reject unrelated layouts.
  if(isQuote&&(!String(get('A34')).includes('소 계')||!String(get('A43')).includes('SUB TOTAL')))throw new Error('장비 10~33행·총괄내역 36~45행의 원본 견적서 양식이 필요해요.');
  const items=[];let category='';
  for(let r=10;r<=(isQuote?33:34);r++){
    const model=String(get(`${isQuote?'B':'C'}${r}`)||'').trim();
    if(get(`A${r}`))category=String(get(`A${r}`)).trim();
    if(model)items.push({name:category,model,maker:String(get(`${isQuote?'C':'G'}${r}`)||'').trim(),quantity:Number(get(`${isQuote?'D':'E'}${r}`))||1,unit:String(get(`${isQuote?'E':'F'}${r}`)||'EA'),rate:isQuote?get(`F${r}`):'',note:isQuote?String(get(`H${r}`)||''):''});
  }
  const fees=isQuote?[38,39,40,41].map(r=>({name:String(get(`A${r}`)),rate:get(`B${r}`),quantity:0,days:1,quantityUnit:String(get(`F${r}`)||'인'),daysUnit:String(get(`D${r}`)||'일')})):[];
  return {zip,wb,rels,doc,path,get,items,fees,shared,sheetName:sheets[0].getAttribute('name')};
}
function element(doc,tag,attrs={}){const n=doc.createElementNS(NS,tag);Object.entries(attrs).forEach(([k,v])=>n.setAttribute(k,String(v)));return n;}
function cell(doc,ref){
  let c=all(doc,'c').find(c=>c.getAttribute('r')===ref);if(c)return c;
  const r=+ref.match(/\d+/)[0],data=one(doc,'sheetData');let row=all(data,'row').find(x=>+x.getAttribute('r')===r);
  if(!row){row=element(doc,'row',{r});const next=all(data,'row').find(x=>+x.getAttribute('r')>r);data.insertBefore(row,next||null);}
  c=element(doc,'c',{r:ref});const next=all(row,'c').find(x=>x.getAttribute('r').replace(/\d+/,'')>ref.replace(/\d+/,''));row.insertBefore(c,next||null);return c;
}
function write(doc,ref,value,formula){
  const c=cell(doc,ref);while(c.firstChild)c.removeChild(c.firstChild);c.removeAttribute('t');
  if(formula){const f=element(doc,'f');f.textContent=formula;c.append(f);}
  if(value===''||value===null||value===undefined)return;
  if(typeof value==='number'){const v=element(doc,'v');v.textContent=String(value);c.append(v);}
  else {c.setAttribute('t','inlineStr');const is=element(doc,'is'),t=element(doc,'t');t.setAttribute('xml:space','preserve');t.textContent=String(value);is.append(t);c.append(is);}
}
function shiftRef(ref,at,count){return ref.replace(/(\$?[A-Z]+\$?)(\d+)/g,(_,a,b)=>a+(+b>=at?+b+count:+b));}
function insertRows(doc,at,count,sourceRow){
  if(!count)return;
  const data=one(doc,'sheetData');const source=all(data,'row').find(r=>+r.getAttribute('r')===sourceRow).cloneNode(true);
  for(const row of all(data,'row')){const r=+row.getAttribute('r');if(r>=at){row.setAttribute('r',r+count);for(const c of all(row,'c'))c.setAttribute('r',shiftRef(c.getAttribute('r'),at,count));}}
  for(const tag of ['mergeCell','dimension','autoFilter','selection'])for(const n of all(doc,tag))for(const a of ['ref','sqref','activeCell'])if(n.hasAttribute(a))n.setAttribute(a,shiftRef(n.getAttribute(a),at,count));
  for(const f of all(doc,'f'))f.textContent=shiftRef(f.textContent,at,count);
  const next=all(data,'row').find(r=>+r.getAttribute('r')===at+count);
  for(let i=0;i<count;i++){const row=source.cloneNode(true);row.setAttribute('r',at+i);for(const c of all(row,'c')){c.setAttribute('r',c.getAttribute('r').replace(/\d+$/,String(at+i)));while(c.firstChild)c.removeChild(c.firstChild);c.removeAttribute('t');}data.insertBefore(row,next);}
}
function rowHeight(doc,r,h){const row=all(doc,'row').find(n=>+n.getAttribute('r')===r);if(row){row.setAttribute('ht',h);row.setAttribute('customHeight','1');}}
async function prepare(book){
  // Convert original shared-string cells before removing the string table, so old client
  // names and notes cannot remain recoverable in an exported workbook's unused strings.
  for(const c of all(book.doc,'c'))if(c.getAttribute('t')==='s'){
    const si=book.shared[+one(c,'v').textContent],is=element(book.doc,'is');
    for(const child of Array.from(si.childNodes))is.append(book.doc.importNode(child,true));
    while(c.firstChild)c.removeChild(c.firstChild);c.setAttribute('t','inlineStr');c.append(is);
  }
  for(const rel of all(book.rels,'Relationship'))if(/\/(sharedStrings|calcChain)$/.test(rel.getAttribute('Type')||''))rel.remove();
  book.zip.remove('xl/sharedStrings.xml');book.zip.remove('xl/calcChain.xml');
  const types=xml(await book.zip.file('[Content_Types].xml').async('string'));
  for(const n of all(types,'Override'))if(/\/(sharedStrings|calcChain)\.xml$/.test(n.getAttribute('PartName')||''))n.remove();
  book.zip.file('[Content_Types].xml',serialize(types));
  if(book.zip.file('docProps/core.xml'))book.zip.file('docProps/core.xml','<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>사운드 오늘</dc:creator><cp:lastModifiedBy>사운드 오늘</cp:lastModifiedBy></cp:coreProperties>');
  book.styles=xml(await book.zip.file('xl/styles.xml').async('string'));
  // Retain original image extents even when wrapped headings require taller rows.
  for(const name of Object.keys(book.zip.files).filter(n=>/^xl\/drawings\/drawing\d+\.xml$/.test(n))){
    const drawing=xml(await book.zip.file(name).async('string'));
    for(const anchor of all(drawing,'twoCellAnchor')){
      const ext=one(one(anchor,'xfrm'),'ext');if(!ext)continue;
      const replacement=drawing.createElementNS(anchor.namespaceURI,'xdr:oneCellAnchor');
      replacement.append(one(anchor,'from').cloneNode(true));const size=drawing.createElementNS(anchor.namespaceURI,'xdr:ext');size.setAttribute('cx',ext.getAttribute('cx'));size.setAttribute('cy',ext.getAttribute('cy'));replacement.append(size);
      for(const tag of ['pic','clientData']){const child=one(anchor,tag);if(child)replacement.append(child.cloneNode(true));}
      anchor.parentNode.replaceChild(replacement,anchor);
    }
    book.zip.file(name,serialize(drawing));
  }
}
function wrap(book,refs){
  const xfs=one(book.styles,'cellXfs');
  for(const ref of refs){const c=cell(book.doc,ref),base=all(xfs,'xf')[+(c.getAttribute('s')||0)];const xf=base.cloneNode(true);let a=one(xf,'alignment');if(!a){a=element(book.styles,'alignment');xf.append(a);}a.setAttribute('wrapText','1');a.setAttribute('vertical','center');xf.setAttribute('applyAlignment','1');c.setAttribute('s',all(xfs,'xf').length);xfs.append(xf);}
  xfs.setAttribute('count',all(xfs,'xf').length);
}
const textLines=(value,width)=>String(value??'').split('\n').reduce((n,line)=>n+Math.max(1,Math.ceil([...line].reduce((s,c)=>s+(c.charCodeAt(0)>255?2:1),0)/width)),0);
function fitTextRow(book,r,values,min=14){
  const refs=Object.keys(values).map(col=>`${col}${r}`);wrap(book,refs);
  const height=Math.max(min,...Object.entries(values).map(([col,{value,width}])=>textLines(value,width)*11+3));
  rowHeight(book.doc,r,height);if(height>min)book.extraHeight=(book.extraHeight||0)+height-min;
}
function header(book,q,s){
  const d=book.doc;
  write(d,'A2',`DATE : ${q.date} / ${q.number}`);
  write(d,'A3',`CLIENT : ${q.client}`);
  write(d,'A4',`PROJECT : ${q.start} ~ ${q.end}`);
  write(d,'A5',`${q.project}${q.venue?' / '+q.venue:''}`);
  write(d,'G4',`${s.name}\n사업자번호: ${s.number}\n대표: ${s.representative} (인)`);
  for(const r of [2,3,4,5])fitTextRow(book,r,{A:{value:text(cell(d,`A${r}`),'t'),width:38}},28);
}
async function finish(book,lastRow){
  const {zip,doc,wb,rels,path,sheetName}=book;
  let names=one(wb,'definedNames');if(!names){names=element(wb,'definedNames');wb.documentElement.insertBefore(names,one(wb,'calcPr')||null);}
  for(const n of all(names,'definedName'))if(n.getAttribute('name')==='_xlnm.Print_Area')n.remove();
  const area=element(wb,'definedName',{name:'_xlnm.Print_Area',localSheetId:0});area.textContent=`'${sheetName.replaceAll("'","''")}'!$A$1:$H$${lastRow}`;names.append(area);
  let calc=one(wb,'calcPr');if(!calc){calc=element(wb,'calcPr');wb.documentElement.append(calc);}calc.setAttribute('calcMode','auto');calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('forceFullCalc','1');
  const dim=one(doc,'dimension');if(dim)dim.setAttribute('ref',`A1:H${lastRow}`);
  for(const tag of ['rowBreaks','colBreaks'])for(const n of all(doc,tag))n.remove();
  let setup=one(doc,'pageSetup');if(setup){setup.setAttribute('paperSize','9');setup.setAttribute('orientation','portrait');setup.setAttribute('fitToWidth','1');setup.setAttribute('fitToHeight',lastRow<=52&&(book.extraHeight||0)<80?'1':'0');}
  for(const v of all(doc,'sheetView'))v.setAttribute('showGridLines','0');
  zip.file(path,serialize(doc));zip.file('xl/workbook.xml',serialize(wb));zip.file('xl/_rels/workbook.xml.rels',serialize(rels));zip.file('xl/styles.xml',serialize(book.styles));
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
export async function exportQuote(bytes,q,s){
  const t=validate(q),b=await readWorkbook(bytes,'quote');await prepare(b);
  const d=b.doc,extra=Math.max(0,q.items.length-24);insertRows(d,34,extra,33);
  const extraFee=Math.max(0,q.fees.length-4);insertRows(d,42+extra,extraFee,41+extra);
  const R=n=>n+extra+(n>=42?extraFee:0),count=Math.max(24,q.items.length);
  header(b,q,s);
  write(d,'A7','▼ 하기와 같이 견적합니다. (유효기간: 견적일로부터 1개월)');
  for(let i=0;i<count;i++){
    const r=10+i,it=q.items[i];for(const col of 'ABCDEFGH')write(d,`${col}${r}`,'');
    if(it){['name','model','maker','quantity','unit','rate',null,'note'].forEach((k,j)=>{if(k)write(d,`${'ABCDEFGH'[j]}${r}`,['quantity','rate'].includes(k)?Number(it[k]):it[k]);});write(d,`G${r}`,t.lines[i],`ROUND(D${r}*F${r},0)`);fitTextRow(b,r,{A:{value:it.name,width:24},B:{value:it.model,width:16},C:{value:it.maker,width:16},H:{value:it.note,width:16}});}
  }
  write(d,`G${R(34)}`,t.daily,`SUM(G10:G${9+count})`);
  const rental=R(36),setup=R(37);
  write(d,`B${rental}`,t.daily,`G${R(34)}`);write(d,`C${rental}`,+q.rentalDays);write(d,`E${rental}`,1);write(d,`G${rental}`,t.rental,`B${rental}*C${rental}*E${rental}`);
  write(d,`A${setup}`,'설치일 장비 임대료');write(d,`H${setup}`,`정상가의 ${q.setupPercent}% 청구`);cell(d,`H${setup}`).setAttribute('s',cell(d,`A${setup}`).getAttribute('s'));
  write(d,`B${setup}`,t.setupRate,`ROUNDDOWN(B${rental}*${Number(q.setupPercent)}/100,0)`);write(d,`C${setup}`,+q.setupDays);write(d,`E${setup}`,1);write(d,`G${setup}`,t.setup,`B${setup}*C${setup}*E${setup}`);
  for(let i=0;i<Math.max(4,q.fees.length);i++){
    const r=38+extra+i,f=q.fees[i];for(const col of 'ABCDEFGH')write(d,`${col}${r}`,'');
    if(f){write(d,`A${r}`,f.name);write(d,`B${r}`,+f.rate);write(d,`C${r}`,+f.days);write(d,`D${r}`,f.daysUnit);write(d,`E${r}`,+f.quantity);write(d,`F${r}`,f.quantityUnit);write(d,`G${r}`,t.feeLines[i],`ROUND(B${r}*C${r}*E${r},0)`);fitTextRow(b,r,{A:{value:f.name,width:24}});}
  }
  write(d,`G${R(42)}`,-t.discount);write(d,`G${R(43)}`,t.supply,`SUM(G${rental}:G${R(42)})`);
  write(d,`G${R(44)}`,t.vat,q.vatOn?`ROUNDDOWN(G${R(43)}*10%,0)`:'0');
  write(d,`A${R(44)}`,q.vatOn?'부가세 / VAT (10%)':'부가세 / VAT 미적용');
  write(d,`G${R(45)}`,t.total,`SUM(G${R(43)}:G${R(44)})`);
  write(d,'G8',t.total,`G${R(45)}`);write(d,'C8',t.total,'G8');
  // The source C8 used an accounting format that could display a dollar sign.
  cell(d,'C8').setAttribute('s',cell(d,'G8').getAttribute('s'));
  const remark=`작성자: ${q.issuer||s.representative} · ${s.phone}${q.contact?'\n거래처 담당: '+q.contact:''}${q.notes?'\n'+q.notes:''}`;
  write(d,`B${R(46)}`,remark);const height=Math.max(25,Math.ceil(remark.length/65)*10+8);rowHeight(d,R(46),height);rowHeight(d,R(47),height);
  wrap(b,[`B${R(46)}`,`H${setup}`]);rowHeight(d,setup,27);
  write(d,`A${R(48)}`,`${s.address}  |  ${s.email}`);
  return finish(b,R(48));
}
export async function exportEquipment(bytes,q,s,packing=null){
  validate(q);const items=packing||q.items;if(!items.length)throw new Error('장비리스트에 넣을 장비를 추가해 주세요.');
  const b=await readWorkbook(bytes,'equipment');await prepare(b);const d=b.doc;
  const extra=Math.max(0,items.length-25);insertRows(d,35,extra,34);header(b,q,s);
  const merges=one(d,'mergeCells');
  // The original wireless category spans four rows, overlapping individual merges.
  // Rebuild only the item-area merges so arbitrary selections don't inherit it.
  for(const m of all(merges,'mergeCell')){const rows=m.getAttribute('ref').match(/\d+/g).map(Number);if(rows[0]>=10&&rows[0]<35+extra||packing&&m.getAttribute('ref')==='G9:H9')m.remove();}
  if(packing){write(d,'C1','출고 장비 리스트');write(d,'G9','출고');write(d,'H9','반납');cell(d,'H9').setAttribute('s',cell(d,'G9').getAttribute('s'));write(d,'C9','모델 / 연결 장비·비고');}
  for(let i=0;i<Math.max(25,items.length);i++){
    const r=10+i,it=items[i];for(const col of 'ABCDEFGH')write(d,`${col}${r}`,'');
    for(const ref of [`A${r}:B${r}`,`C${r}:D${r}`,...(packing?[]:[`G${r}:H${r}`])])merges.append(element(d,'mergeCell',{ref}));
    if(it){const name=it.name+(packing?'\n'+it.kind:'');const model=packing?[it.model,it.note,it.source&&it.source!=='견적 포함'?'연결: '+it.source:''].filter(Boolean).join('\n'):it.model;
      write(d,`A${r}`,name);write(d,`C${r}`,model);write(d,`E${r}`,+it.quantity);write(d,`F${r}`,it.unit);const note=[it.maker,it.note].filter(Boolean).join(' / ');
      write(d,`G${r}`,packing?(it.out?'완료':'□'):note);if(packing)write(d,`H${r}`,it.returned?'완료':'□');
      fitTextRow(b,r,{A:{value:name,width:38},C:{value:model,width:28},G:{value:packing?'□':note,width:packing?12:28}},packing?28:14);}
  }
  merges.setAttribute('count',all(merges,'mergeCell').length);
  write(d,`E${35+extra}`,items.reduce((s,i)=>s+Number(i.quantity),0),`SUM(E10:E${34+extra})`);
  write(d,`A${37+extra}`,packing?'REMARK: 견적 장비 + 추가 준비물. 출고·반납 시 실물을 확인해 주세요.':'REMARK: 장비는 현장 상황에 따라 동급 모델로 대체될 수 있습니다.');
  write(d,`A${38+extra}`,`${s.address}  |  ${s.email}`);
  return finish(b,38+extra);
}

export async function exportIO(q,s){
  validateIO(q);const layout=ioLayout(q);
  const zip=await JSZip.loadAsync(Uint8Array.from(atob(ioTemplate),c=>c.charCodeAt(0)));
  const wb=xml(await zip.file('xl/workbook.xml').async('string')),styles=xml(await zip.file('xl/styles.xml').async('string'));
  const rels=xml(await zip.file('xl/_rels/workbook.xml.rels').async('string'));
  const sheet=all(wb,'sheet')[0],target=all(rels,'Relationship').find(r=>r.getAttribute('Id')===sheet.getAttribute('r:id')).getAttribute('Target');
  const p=target.replace(/^\//,'').replace(/^(?!xl\/)/,'xl/'),doc=xml(await zip.file(p).async('string')),b={doc,styles};
  const data=one(doc,'sheetData'),merges=one(doc,'mergeCells');
  const samples=Object.fromEntries([6,9,33,34,40].map(r=>[r,all(data,'row').find(n=>+n.getAttribute('r')===r).cloneNode(true)]));
  const style=(sample,col)=>all(samples[sample],'c').find(c=>c.getAttribute('r').replace(/\d+$/,'')===col)?.getAttribute('s')||'0';
  for(const row of all(data,'row'))if(+row.getAttribute('r')>=9)row.remove();
  for(const merge of all(merges,'mergeCell'))if(+merge.getAttribute('ref').match(/\d+/)[0]>=9)merge.remove();
  const merge=ref=>merges.append(element(doc,'mergeCell',{ref}));
  const appendRow=(r,sample=9)=>{const row=samples[sample].cloneNode(true);row.setAttribute('r',r);for(const c of all(row,'c')){c.setAttribute('r',c.getAttribute('r').replace(/\d+$/,String(r)));while(c.firstChild)c.removeChild(c.firstChild);c.removeAttribute('t');}data.append(row);return row;};
  const title=q.project?q.project+' I/O Sheet':'I/O Sheet';write(doc,'C1',title);wrap(b,['C1']);
  const titleHeight=Math.max(23,Math.ceil(textLines(title,50)*24/3));for(const r of [1,2,3])rowHeight(doc,r,titleHeight);
  const meta='일정: '+(q.start||'')+(q.end&&q.end!==q.start?' ~ '+q.end:'')+'    장소: '+(q.venue||'');
  write(doc,'A4',meta);write(doc,'A5',ioContact(q,s));
  fitTextRow(b,4,{A:{value:meta,width:140}},22);fitTextRow(b,5,{A:{value:ioContact(q,s),width:140}},24);
  for(let i=0;i<layout.count;i++){
    const r=9+i,input=layout.inputs[i],output=layout.outputs[i];appendRow(r);
    const values={A:input.channel,B:input.name,C:input.std,D:input.note};
    if(output){values.E=output.channel;if(output.span){values.F=output.name;values.G=output.note;if(output.span>1){merge(`F${r}:F${r+output.span-1}`);merge(`G${r}:G${r+output.span-1}`);}}}
    else {const header=i===layout.outputCount;values.E=header?'Check List':layout.checklist[i-layout.outputCount-1]||'';merge(`E${r}:G${r}`);for(const col of 'EFG')cell(doc,col+r).setAttribute('s',style(header?33:34,col));}
    if(i===layout.count-1)for(const col of 'ABCDEFG')cell(doc,col+r).setAttribute('s',style(40,col));
    for(const [col,value] of Object.entries(values))write(doc,col+r,value);
    const widths={A:7,B:28,C:9,D:34,E:output?7:70,F:28,G:34};
    fitTextRow(b,r,Object.fromEntries(Object.entries(values).map(([col,value])=>[col,{value,width:widths[col]}])),23);
  }
  let last=8+layout.count;
  if(q.io.notes){appendRow(++last,6);merge(`A${last}:G${last}`);write(doc,'A'+last,'현장 기술 메모');appendRow(++last);merge(`A${last}:G${last}`);write(doc,'A'+last,q.io.notes);fitTextRow(b,last,{A:{value:q.io.notes,width:140}},28);}
  merges.setAttribute('count',all(merges,'mergeCell').length);
  const dim=one(doc,'dimension');if(dim)dim.setAttribute('ref',`A1:G${last}`);
  let setup=one(doc,'pageSetup');if(!setup){setup=element(doc,'pageSetup');doc.documentElement.insertBefore(setup,one(doc,'headerFooter')||one(doc,'drawing')||null);}
  Object.entries({paperSize:9,orientation:'portrait',fitToWidth:1,fitToHeight:layout.count<=32&&(b.extraHeight||0)<100?'1':'0'}).forEach(([k,v])=>setup.setAttribute(k,String(v)));
  let margins=one(doc,'pageMargins');if(!margins){margins=element(doc,'pageMargins');doc.documentElement.insertBefore(margins,setup);}Object.entries({left:.2,right:.2,top:.25,bottom:.25,header:.1,footer:.1}).forEach(([k,v])=>margins.setAttribute(k,String(v)));
  let prop=one(doc,'sheetPr');if(!prop){prop=element(doc,'sheetPr');doc.documentElement.insertBefore(prop,doc.documentElement.firstChild);}let fit=one(prop,'pageSetUpPr');if(!fit){fit=element(doc,'pageSetUpPr');prop.append(fit);}fit.setAttribute('fitToPage','1');
  let names=one(wb,'definedNames');if(!names){names=element(wb,'definedNames');wb.documentElement.insertBefore(names,one(wb,'calcPr')||null);}
  for(const n of all(names,'definedName'))if(/^_xlnm.Print_/.test(n.getAttribute('name')||''))n.remove();
  const sheetName=sheet.getAttribute('name').replaceAll("'","''");for(const [name,ref] of [['_xlnm.Print_Area',`$A$1:$G$${last}`],['_xlnm.Print_Titles','$1:$8']]){const n=element(wb,'definedName',{name,localSheetId:0});n.textContent=`'${sheetName}'!${ref}`;names.append(n);}
  for(const v of all(doc,'sheetView'))v.setAttribute('showGridLines','0');
  zip.file(p,serialize(doc));zip.file('xl/styles.xml',serialize(styles));zip.file('xl/workbook.xml',serialize(wb));
  return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
