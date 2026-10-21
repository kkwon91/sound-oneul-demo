import {VERSION,DEFAULT_SUPPLIER,today,uid,clone,freshQuote,calculate,validate,toJob,mergeCatalog,money} from './core.mjs?v=2';
import {readWorkbook,exportQuote,exportEquipment,exportIO} from './xlsx.mjs?v=3';
import {normalizeQuote,searchCatalog,itemFromCatalog,reconcileCatalog,packingRows,validatePacking,validateRules,IO_FIELDS,blankIO,validateIO,ioLayout,ioContact} from './workflow.mjs?v=3';
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let db,profile={supplier:{...DEFAULT_SUPPLIER},catalog:[],fees:[]},quotes=[],q,saveTimer,toastTimer,lastURL,busy=false;
let catalogTarget=null,modelDraft,quickResults=[],quickIndex=-1;
const status=s=>$('#save-state').textContent=s;
const toast=s=>{$('#toast').textContent=s;$('#toast').hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('#toast').hidden=true,5000);};
const openDB=()=>new Promise((resolve,reject)=>{const r=indexedDB.open('sound-oneul-quotes',1);r.onupgradeneeded=()=>{r.result.createObjectStore('state');r.result.createObjectStore('quotes',{keyPath:'id'});};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
const read=(store,key)=>new Promise((resolve,reject)=>{const r=db.transaction(store).objectStore(store)[key===undefined?'getAll':'get'](key);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
const put=(store,data,key)=>new Promise((resolve,reject)=>{const tx=db.transaction(store,'readwrite');if(key===undefined)tx.objectStore(store).put(data);else tx.objectStore(store).put(data,key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);});
const saveProfile=()=>put('state',profile,'profile');
function meaningful(){return q&&!!(q.project||q.client||q.items.length||q.fees.some(f=>+f.quantity>0)||q.packingExtras?.length||q.io?.inputs.length||q.io?.outputs.length||q.io?.notes||q.io?.checklist||q.io?.contactName||q.io?.contactPhone||q.io?.contactEmail);}
async function save(){
  clearTimeout(saveTimer);if(!db||!q)return;
  const snapshot=clone(q);snapshot.updatedAt=new Date().toISOString();
  await put('quotes',snapshot);await put('state',snapshot.id,'current');
  if(q.id===snapshot.id){q.updatedAt=snapshot.updatedAt;status('자동 저장 완료');}
  const i=quotes.findIndex(x=>x.id===snapshot.id);if(i<0)quotes.push(snapshot);else quotes[i]=snapshot;
  renderHistory();
}
function changed(){status('저장 중…');clearTimeout(saveTimer);saveTimer=setTimeout(()=>save().catch(()=>{status('저장 실패');toast('저장 공간을 확인하고 전체 백업을 받아 주세요.');}),450);}
async function safe(action){if(busy)return;busy=true;try{await action();}catch(e){toast(e.message||'작업을 완료하지 못했어요.');console.error(e);}finally{busy=false;}}
function renderSetup(){
  $('#quote-name').textContent=profile.quote?.name||'2026년 사운드오늘 견적서.xlsx';
  $('#gear-name').textContent=profile.gear?.name||'장비리스트 사운드오늘.xlsx';
  $('#setup-status').textContent=profile.quote&&profile.gear?`원본 양식 2개 · 기본 장비 ${profile.catalog.length}개 등록됨`:profile.quote?'견적서 등록됨 · 장비리스트를 등록해 주세요.':'처음 한 번, 원본 엑셀 2개를 등록해 주세요.';
  $('#supplier').innerHTML=Object.entries({name:'상호',representative:'대표자',number:'사업자등록번호',phone:'영업 연락처',email:'이메일',address:'사업장 주소'}).map(([k,label])=>`<label>${label}<input data-supplier="${k}" maxlength="150" value="${esc(profile.supplier[k])}"></label>`).join('');
}
function renderHistory(){
  const search=$('#history-search').value.trim().toLowerCase();
  const sorted=quotes.slice().sort((a,b)=>(b.updatedAt||'').localeCompare(a.updatedAt||''));
  const html=x=>{let total='금액 입력 중';try{total=money(calculate(x).total)+'원';}catch{}return `<button class="history-item ${x.id===q?.id?'active':''}" data-open="${esc(x.id)}"><b>${esc(x.project||'제목 없는 견적')}</b><small>${esc(x.client||'거래처 미입력')} · ${esc(x.date)}</small><strong>${total}</strong></button>`;};
  const active=sorted.filter(x=>!x.archived);
  $('#saved-count').textContent=active.length;
  $('#history-list').innerHTML=active.filter(x=>`${x.project} ${x.client}`.toLowerCase().includes(search)).map(html).join('')||'<p class="hint empty">저장된 견적이 없습니다.</p>';
  $('#archive-list').innerHTML=sorted.filter(x=>x.archived).map(x=>html(x)+`<button data-restore="${esc(x.id)}">보관 해제</button>`).join('')||'<p class="hint">보관한 견적이 없습니다.</p>';
}
function rowInput(type,index,key,value,label,{number=false,min=0,max=1000000000,step=1}={}){
  return `<input aria-label="${esc(label)}" data-${type}="${index}" data-key="${key}" ${number?`type="number" min="${min}" max="${max}" step="${step}"`:'maxlength="100"'} value="${esc(value)}">`;
}
function renderRows(){
  $('#items-empty').hidden=!!q.items.length;$('#item-count').textContent=q.items.length+'개 품목';
  $('#items').innerHTML=q.items.map((it,i)=>`<tr><td>${rowInput('item',i,'name',it.name,`${i+1}번 장비명`)}${rowInput('item',i,'model',it.model,`${i+1}번 모델명`)}${rowInput('item',i,'maker',it.maker,`${i+1}번 제조사`)}</td><td>${rowInput('item',i,'quantity',it.quantity,`${i+1}번 수량`,{number:true,min:1,max:100000})}</td><td>${rowInput('item',i,'unit',it.unit,`${i+1}번 단위`)}</td><td>${rowInput('item',i,'rate',it.rate,`${i+1}번 단가`,{number:true})}</td><td class="amount" data-line="${i}"></td><td>${rowInput('item',i,'note',it.note||'',`${i+1}번 비고`)}</td><td><button type="button" class="remove" data-remove-item="${i}" aria-label="${i+1}번 장비 삭제">×</button></td></tr>`).join('');
  const unit=(i,key,value,options)=>`<select aria-label="${i+1}번 비용 ${key==='daysUnit'?'일수':'수량'} 단위" data-fee="${i}" data-key="${key}">${[...new Set([value,...options])].map(v=>`<option${v===value?' selected':''}>${esc(v)}</option>`).join('')}</select>`;
  $('#fees').innerHTML=q.fees.map((f,i)=>`<tr><td>${rowInput('fee',i,'name',f.name,`${i+1}번 비용명`)}</td><td>${rowInput('fee',i,'rate',f.rate,`${i+1}번 비용 단가`,{number:true})}</td><td>${rowInput('fee',i,'days',f.days,`${i+1}번 비용 일수·대수`,{number:true,max:3660,step:0.5})}${unit(i,'daysUnit',f.daysUnit,['일','대','회'])}</td><td>${rowInput('fee',i,'quantity',f.quantity,`${i+1}번 비용 인원·횟수`,{number:true,max:100000})}${unit(i,'quantityUnit',f.quantityUnit,['인','회','번','식'])}</td><td class="amount" data-fee-line="${i}"></td><td><button type="button" class="remove" data-remove-fee="${i}" aria-label="${i+1}번 비용 삭제">×</button></td></tr>`).join('');
  $$('#items tr').forEach((row,i)=>{const button=document.createElement('button');button.type='button';button.className='model-pick';button.dataset.pickModel=String(i);button.textContent='모델 선택';row.cells[0].append(button);});
  renderTotals();renderPacking();renderDeviceOptions();
}
function renderTotals(){
  try{
    const c=calculate(q);$('#calc-error').hidden=true;
    for(const [id,key] of Object.entries({daily:'daily',rental:'rental','setup-total':'setup','fees-total':'fees','discount-total':'discount',supply:'supply',vat:'vat'}))$('#'+id).textContent=(id==='discount-total'&&c[key]?'-':'')+money(c[key])+'원';
    $('#total').innerHTML=money(c.total)+'<span>원</span>';$('#vat-label').textContent=q.vatOn?'부가세 포함':'부가세 미적용';
    $$('[data-line]').forEach(n=>n.textContent=money(c.lines[+n.dataset.line]));$$('[data-fee-line]').forEach(n=>n.textContent=money(c.feeLines[+n.dataset.feeLine]));
  }catch(e){
    $('#calc-error').textContent=e.message;$('#calc-error').hidden=false;
    for(const id of ['daily','rental','setup-total','fees-total','discount-total','supply','vat','total'])$('#'+id).textContent='—';
    $$('[data-line],[data-fee-line]').forEach(n=>n.textContent='—');
  }
}
function render(){
  normalizeQuote(q);
  $$('[data-q]').forEach(n=>n.value=String(q[n.dataset.q]??''));renderRows();renderHistory();status(q.updatedAt?'저장된 견적':'새 견적');
  $('#archive').textContent=q.archived?'보관 해제':'보관';
  renderIO();hideQuick();
}
function renderCatalog(){
  const items=searchCatalog(profile.catalog,$('#catalog-search').value,{includeInactive:$('#catalog-inactive').checked});
  $('#catalog-list').innerHTML=items.map(i=>`<div class="catalog-row"><div><strong>${esc(i.model)}${i.inactive?' (사용 중지)':''}</strong><small>${esc(i.name)} · ${esc(i.maker)} · ${esc(i.unit)}</small><button class="text-button" data-edit-model="${esc(i.id)}">모델·추가 준비물 설정</button></div><label>1일 단가 (원)<input aria-label="${esc(i.model)} 기본 단가" type="number" min="0" step="1" data-price="${esc(i.id)}" value="${esc(i.rate)}"></label><button data-add="${esc(i.id)}" ${i.inactive?'disabled':''}>${catalogTarget===null?'추가':'선택'}</button></div>`).join('')||'<p class="empty">검색 결과가 없습니다.<br><small>검색어를 바꾸거나 보유 모델을 등록해 주세요.</small></p>';
}
async function importTemplate(file,kind){
  if(!file)return;const bytes=await file.arrayBuffer();const book=await readWorkbook(bytes,kind);
  const next=clone(profile);next[kind==='quote'?'quote':'gear']={name:file.name,bytes};
  const other=kind==='quote'?next.gear:next.quote;
  const otherBook=other?await readWorkbook(other.bytes,kind==='quote'?'equipment':'quote'):null;
  const qb=kind==='quote'?book:otherBook,gb=kind==='equipment'?book:otherBook;
  next.catalog=reconcileCatalog(mergeCatalog(qb?.items||[],gb?.items||[]),profile.catalog);if(qb)next.fees=qb.fees;
  await put('state',next,'profile');profile=next;
  if(!q.fees.length&&profile.fees.length){q.fees=clone(profile.fees).map(f=>({...f,quantity:0,days:1}));if(meaningful())await save();}
  renderSetup();renderRows();$('#setup').open=!(profile.quote&&profile.gear);toast(`${kind==='quote'?'견적서 양식':'장비리스트'} 등록 완료 · 장비 ${profile.catalog.length}개`);
}
function addItem(item,quantity=1){
  if(!item)throw new Error('장비를 다시 선택해 주세요.');
  if(catalogTarget!==null){const previous=q.items[catalogTarget];q.items[catalogTarget]={...itemFromCatalog(item,previous.quantity),lineId:previous.lineId,note:previous.note||''};catalogTarget=null;$('#catalog-dialog').close();changed();renderRows();return;}
  if(q.items.length>=120)throw new Error('장비는 120개 품목까지 추가할 수 있어요.');
  const existing=q.items.find(i=>i.catalogId===item.id&&i.rate===item.rate&&JSON.stringify(i.accessories)===JSON.stringify(item.accessories||[]));
  if(existing)existing.quantity=Number(existing.quantity)+quantity;
  else q.items.push(itemFromCatalog(item,quantity));
  changed();renderRows();
}
function download(bytes,name,mime){
  if(lastURL)URL.revokeObjectURL(lastURL);
  lastURL=URL.createObjectURL(new Blob([bytes],{type:mime}));const a=$('#last-download');a.href=lastURL;a.download=name;a.hidden=false;a.textContent=name+' 다시 받기';a.click();
}
const xlsxMime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const filename=(label)=>`${q.date}_${q.project}_${label}.xlsx`.replace(/[\\/:*?"<>|\x00-\x1f]/g,'_');
async function exportFile(kind){
  if(kind==='io')validateIO(q);else validate(q);const source=kind==='quote'?profile.quote:profile.gear;
  if(kind!=='io'&&!source){$('#setup').open=true;throw new Error(`${kind==='quote'?'견적서':'장비리스트'} 원본 엑셀을 먼저 등록해 주세요.`);}
  if(!profile.supplier.name.trim()||!profile.supplier.representative.trim()||!profile.supplier.number.trim())throw new Error('공급자 상호·대표자·사업자등록번호를 확인해 주세요.');
  if(meaningful())await save();const out=kind==='io'?await exportIO(q,profile.supplier):kind==='quote'?await exportQuote(source.bytes,q,profile.supplier):await exportEquipment(source.bytes,q,profile.supplier,validatePacking(q));
  download(out,filename(kind==='quote'?'견적서':kind==='io'?'인풋아웃풋리스트':'출고장비리스트'),xlsxMime);toast('엑셀 파일을 만들었습니다.');
}
function preview(){
  const c=validate(q),s=profile.supplier;
  const lines=q.items.map((it,i)=>`<tr><td>${esc(it.name)}<br><small>${esc(it.model)} · ${esc(it.maker)}</small></td><td>${it.quantity} ${esc(it.unit)}</td><td class="amount">${money(it.rate)}</td><td class="amount">${money(c.lines[i])}</td></tr>`).join('');
  const fees=q.fees.filter(f=>+f.quantity>0).map(f=>`<p>${esc(f.name)}: ${money(f.rate)}원 × ${f.days}${esc(f.daysUnit)} × ${f.quantity}${esc(f.quantityUnit)}</p>`).join('');
  $('#preview').innerHTML=`<article class="preview-paper"><div class="preview-title"><div><h3>견적서</h3><small>PROJECT ESTIMATE</small></div><img src="../assets/oneul-logo-original.png" alt="사운드 오늘"></div><div class="preview-meta">${[['거래처',q.client],['견적일 / 번호',q.date+' / '+q.number],['행사',q.project],['일정',q.start+' ~ '+q.end],['장소',q.venue||'미정'],['공급자',s.name+' / '+s.representative]].map(([k,v])=>`<p><strong>${k}</strong>${esc(v)}</p>`).join('')}</div><table><thead><tr><th>장비</th><th>수량</th><th>1일 단가</th><th>1일 금액</th></tr></thead><tbody>${lines}</tbody></table><p>장비 1일 소계 ${money(c.daily)}원 × 임대 ${q.rentalDays}일</p><p>설치 ${q.setupDays}일 · 정상가의 ${q.setupPercent}% 청구</p>${fees}<dl><div><dt>공급가액 (네고 ${money(c.discount)}원 반영)</dt><dd>${money(c.supply)}원</dd></div><div><dt>부가세 ${q.vatOn?'10%':'미적용'}</dt><dd>${money(c.vat)}원</dd></div></dl><p class="preview-total">총 ${money(c.total)}원</p><pre>${esc(q.notes)}</pre><p class="hint">작성자 ${esc(q.issuer)} · ${esc(s.phone)}<br>${esc(s.address)}<br>${esc(s.email)}</p></article>`;
  $('#preview-dialog').showModal();
}
async function backup(){
  if(meaningful())await save();const z=new JSZip(),meta=clone(profile);delete meta.quote;delete meta.gear;
  z.file('oneul-quotes.json',JSON.stringify({version:VERSION,exportedAt:new Date().toISOString(),profile:meta,quotes},null,2));
  if(profile.quote)z.file('quote-template.xlsx',profile.quote.bytes);if(profile.gear)z.file('equipment-template.xlsx',profile.gear.bytes);
  download(await z.generateAsync({type:'uint8array',compression:'DEFLATE'}),`사운드오늘_견적백업_${today()}.zip`,'application/zip');toast('양식·장비·견적 전체 백업을 만들었습니다.');
}
function checkDraft(value){
  if(!value||typeof value.id!=='string'||!Array.isArray(value.items)||!Array.isArray(value.fees)||value.items.length>120||value.fees.length>40)throw new Error('올바른 견적 백업이 아니에요.');
  // Drafts may legitimately contain incomplete numeric fields. Preserve them as drafts.
  const base=normalizeQuote(freshQuote());for(const k of Object.keys(base))if(k in value)base[k]=value[k];base.archived=Boolean(value.archived);return normalizeQuote(base);
}
async function restore(file){
  if(!file)return;if(file.size>30*1024*1024)throw new Error('30MB 이하의 백업 파일을 선택해 주세요.');
  const z=await JSZip.loadAsync(await file.arrayBuffer()),entry=z.file('oneul-quotes.json');if(!entry)throw new Error('사운드 오늘 견적 백업 파일이 아니에요.');
  const data=JSON.parse(await entry.async('string'));if(![1,VERSION].includes(data.version)||!Array.isArray(data.quotes)||!Array.isArray(data.profile?.catalog)||!Array.isArray(data.profile?.fees))throw new Error('지원하지 않는 백업 형식입니다.');
  const incoming=data.quotes.map(checkDraft);if(meaningful())await save();
  const next={...profile,...data.profile,supplier:{...DEFAULT_SUPPLIER,...data.profile.supplier}};
  for(const [file,key,kind] of [['quote-template.xlsx','quote','quote'],['equipment-template.xlsx','gear','equipment']])if(z.file(file)){const bytes=await z.file(file).async('arraybuffer');await readWorkbook(bytes,kind);next[key]={name:file,bytes};}
  for(const quote of incoming){const existing=quotes.find(x=>x.id===quote.id);if(existing&&JSON.stringify(existing)!==JSON.stringify(quote)){quote.id=uid();quote.number+='-복원';}await put('quotes',quote);}
  await put('state',next,'profile');profile=next;quotes=await read('quotes');renderSetup();renderHistory();toast(`${incoming.length}개 견적 복원 완료. 기존 견적도 유지됩니다.`);
}
$('#quote-form').addEventListener('submit',e=>e.preventDefault());
document.addEventListener('input',e=>{
  const n=e.target;
  if(n.dataset.q){const k=n.dataset.q;q[k]=k==='vatOn'?n.value==='true':n.value;if(k==='start'&&!q.end){q.end=q.start;$('[data-q="end"]').value=q.end;}changed();renderTotals();}
  if(n.dataset.item!==undefined||n.dataset.fee!==undefined){const row=n.dataset.item!==undefined?q.items[+n.dataset.item]:q.fees[+n.dataset.fee];row[n.dataset.key]=n.value;if(n.dataset.item!==undefined&&['name','model','maker'].includes(n.dataset.key)){delete row.catalogId;row.accessories=[];}changed();renderTotals();renderPacking();renderDeviceOptions();}
  if(n.dataset.supplier){profile.supplier[n.dataset.supplier]=n.value;saveProfile().catch(()=>toast('공급자 정보 저장 실패'));}
  if(n.dataset.price){const it=profile.catalog.find(i=>i.id===n.dataset.price);it.rate=n.value;saveProfile().catch(()=>toast('장비 단가 저장 실패'));}
  if(n.id==='history-search')renderHistory();if(n.id==='catalog-search'||n.id==='catalog-inactive')renderCatalog();
  if(n.id==='equipment-search')renderQuick();
  if(n.dataset.extra!==undefined){const extra=q.packingExtras.find(x=>x.id===n.dataset.extra);extra[n.dataset.key]=n.value;changed();updatePackingStatus();}
  if(n.dataset.io){q.io[n.dataset.io][+n.dataset.row][n.dataset.key]=n.value;changed();}
  if(n.dataset.ioMeta){q.io[n.dataset.ioMeta]=n.value;changed();}
  if(n.dataset.rule!==undefined)modelDraft.accessories[+n.dataset.rule][n.dataset.key]=n.value;
});
document.addEventListener('change',e=>{
  const n=e.target;if(n.id==='quote-template')safe(()=>importTemplate(n.files[0],'quote'));if(n.id==='gear-template')safe(()=>importTemplate(n.files[0],'equipment'));if(n.id==='restore')safe(()=>restore(n.files[0]));
  if(n.dataset.packKey){const row=packingRows(q).find(r=>r.key===n.dataset.packKey);if(n.dataset.check==='included'){q.packingExcluded=q.packingExcluded.filter(k=>k!==row.key);if(!n.checked)q.packingExcluded.push(row.key);}else{q.packingChecks[row.key]={signature:row.signature,out:row.out,returned:row.returned,[n.dataset.check]:n.checked};}changed();renderPacking();}
});
document.addEventListener('click',e=>{
  const n=e.target.closest('button,[data-close]');if(!n)return;
  if(n.dataset.close){$('#'+n.dataset.close).close();return;}
  safe(async()=>{
    if(n.dataset.open){if(meaningful())await save();q=clone(quotes.find(x=>x.id===n.dataset.open));await put('state',q.id,'current');render();}
    if(n.dataset.restore){const x=quotes.find(x=>x.id===n.dataset.restore);x.archived=false;await put('quotes',x);if(q.id===x.id)q.archived=false;render();}
    if(n.dataset.add){addItem(profile.catalog.find(i=>i.id===n.dataset.add));toast('장비를 적용했습니다.');}
    if(n.dataset.quick){catalogTarget=null;addItem(profile.catalog.find(i=>i.id===n.dataset.quick));$('#equipment-search').value='';hideQuick();$('#equipment-search').focus();}
    if(n.dataset.pickModel!==undefined){catalogTarget=+n.dataset.pickModel;$('#catalog-search').value=q.items[catalogTarget].name;renderCatalog();$('#catalog-dialog').showModal();}
    if(n.dataset.editModel)openModel(n.dataset.editModel);
    if(n.dataset.removeRule!==undefined){modelDraft.accessories.splice(+n.dataset.removeRule,1);renderRules();}
    if(n.dataset.removeExtra){q.packingExtras=q.packingExtras.filter(x=>x.id!==n.dataset.removeExtra);changed();renderPacking();}
    if(n.dataset.addIo){addIORows(n.dataset.addIo,+(n.dataset.count||1));}
    if(n.dataset.removeIo){q.io[n.dataset.removeIo].splice(+n.dataset.row,1);changed();renderIO();}
    if(n.dataset.removeItem!==undefined){q.items.splice(+n.dataset.removeItem,1);changed();renderRows();}
    if(n.dataset.removeFee!==undefined){q.fees.splice(+n.dataset.removeFee,1);changed();renderRows();}
    switch(n.id){
      case 'new':if(meaningful())await save();q=freshQuote(profile.fees);render();await put('state',null,'current');break;
      case 'save':await save();toast('견적을 저장했습니다.');break;
      case 'duplicate':await save();q={...clone(q),id:uid(),number:q.number+'-복사',date:today(),archived:false,updatedAt:'',packingChecks:{}};await save();render();toast('복사한 견적을 열었습니다. 출고·반납 체크는 초기화했습니다.');break;
      case 'archive':q.archived=!q.archived;await save();render();toast(q.archived?'보관함으로 옮겼습니다. 언제든 복원할 수 있어요.':'보관을 해제했습니다.');break;
      case 'catalog-open':catalogTarget=null;renderCatalog();$('#catalog-dialog').showModal();break;
      case 'catalog-new':openModel();break;
      case 'rule-add':modelDraft.accessories.push({id:uid(),name:'',model:'',quantity:1,mode:'unit',unit:'EA'});renderRules();break;
      case 'source-pack':catalogTarget=null;if(!profile.catalog.length)throw new Error('원본 엑셀을 먼저 등록해 주세요.');for(const it of profile.catalog.filter(i=>!i.inactive&&i.sourceQuantity))addItem(it,Number(it.sourceQuantity)||1);toast('원본 장비 구성을 추가했습니다.');break;
      case 'custom-item':catalogTarget=null;addItem({id:uid(),name:'',model:'',maker:'',unit:'EA',rate:'',note:''});break;
      case 'packing-extra':q.packingExtras.push({id:uid(),name:'',model:'',quantity:1,unit:'EA',note:''});changed();renderPacking();break;
      case 'packing-refresh':for(const i of q.items){const model=profile.catalog.find(m=>m.id===i.catalogId);if(model)i.accessories=clone(model.accessories||[]);}changed();renderPacking();toast('현재 모델별 준비물 규칙을 적용했습니다.');break;
      case 'add-fee':if(q.fees.length>=40)throw new Error('비용은 40개 항목까지 추가할 수 있어요.');q.fees.push({name:'기타 비용',rate:'',days:1,quantity:1,daysUnit:'일',quantityUnit:'식'});changed();renderRows();break;
      case 'download-quote':await exportFile('quote');break;
      case 'download-gear':await exportFile('equipment');break;
      case 'download-io':await exportFile('io');break;
      case 'download-all':await exportBundle();break;
      case 'io-preview':previewIO();break;
      case 'preview-open':preview();break;
      case 'backup':await backup();break;
      case 'to-ledger':{const job=toJob(q);await save();localStorage.setItem('oneul.quote.handoff.v1',JSON.stringify({version:1,quoteId:q.id,createdAt:Date.now(),job}));location.href='../../#jobs';break;}
    }
  });
});
function hideQuick(){$('#equipment-options').hidden=true;$('#equipment-search').setAttribute('aria-expanded','false');$('#equipment-search').removeAttribute('aria-activedescendant');quickIndex=-1;}
function renderQuick(){
  const term=$('#equipment-search').value;quickResults=searchCatalog(profile.catalog,term).slice(0,15);quickIndex=-1;
  if(!term.trim()){hideQuick();return;}
  $('#equipment-options').innerHTML=quickResults.map((i,n)=>`<button type="button" role="option" aria-selected="false" id="quick-${n}" data-quick="${esc(i.id)}"><span><strong>${esc(i.name)} · ${esc(i.model)}</strong><small>${esc(i.maker)} · ${esc(i.unit)}</small></span><b>${i.rate===''?'단가 미입력':money(i.rate)+'원/일'}</b></button>`).join('')||'<p class="hint">일치하는 모델이 없습니다. ‘장비 선택’에서 보유 모델을 등록해 주세요.</p>';
  $('#equipment-options').hidden=false;$('#equipment-search').setAttribute('aria-expanded','true');$('#equipment-search').removeAttribute('aria-activedescendant');
}
$('#equipment-search').addEventListener('keydown',e=>{
  if(e.isComposing)return;
  if(e.key==='Escape'){hideQuick();return;}
  if($('#equipment-options').hidden)return;
  if(['ArrowDown','ArrowUp'].includes(e.key)&&quickResults.length){e.preventDefault();quickIndex=(quickIndex+(e.key==='ArrowDown'?1:-1)+quickResults.length)%quickResults.length;$$('#equipment-options [role=option]').forEach((o,i)=>o.setAttribute('aria-selected',String(i===quickIndex)));$('#equipment-search').setAttribute('aria-activedescendant','quick-'+quickIndex);}
  if(e.key==='Enter'){e.preventDefault();const index=quickIndex<0?0:quickIndex;if(quickResults[index])safe(()=>{catalogTarget=null;addItem(quickResults[index]);$('#equipment-search').value='';hideQuick();});}
});
document.addEventListener('click',e=>{if(!e.target.closest('.quick-search'))hideQuick();});
function openModel(id){
  modelDraft=clone(profile.catalog.find(i=>i.id===id)||{id:uid(),name:'',model:'',maker:'',unit:'EA',rate:'',aliases:'',accessories:[]});modelDraft.accessories??=[];
  for(const key of ['name','model','maker','unit','rate','aliases'])$('#model-form').elements[key].value=modelDraft[key]??'';
  $('#model-form').elements.inactive.checked=!!modelDraft.inactive;renderRules();$('#model-dialog').showModal();
}
function renderRules(){
  $('#model-rules').innerHTML=modelDraft.accessories.map((r,i)=>`<tr>${['name','model','quantity'].map(k=>`<td><input aria-label="${i+1}번 준비물 ${k==='name'?'이름':k==='model'?'규격':'기준 수량'}" data-rule="${i}" data-key="${k}" ${k==='quantity'?'type="number" min="0.01" max="100000" step="0.01"':'maxlength="100"'} value="${esc(r[k])}"></td>`).join('')}<td><select aria-label="${i+1}번 준비물 수량 기준" data-rule="${i}" data-key="mode"><option value="unit" ${r.mode==='unit'?'selected':''}>장비 1개당</option><option value="event" ${r.mode==='event'?'selected':''}>행사당 고정</option></select></td><td><input aria-label="${i+1}번 준비물 단위" maxlength="20" data-rule="${i}" data-key="unit" value="${esc(r.unit)}"></td><td><button type="button" data-remove-rule="${i}" aria-label="${i+1}번 준비물 규칙 삭제">×</button></td></tr>`).join('');
}
$('#model-form').addEventListener('submit',e=>{e.preventDefault();safe(async()=>{
  const f=e.target;for(const k of ['name','model','maker','unit','rate','aliases'])modelDraft[k]=f.elements[k].value.trim();modelDraft.inactive=f.elements.inactive.checked;modelDraft.custom=true;
  if(!modelDraft.name||!modelDraft.model||!modelDraft.unit)throw new Error('장비명·모델명·단위를 입력해 주세요.');
  if(modelDraft.rate!==''&&(!Number.isInteger(+modelDraft.rate)||+modelDraft.rate<0||+modelDraft.rate>1e9))throw new Error('기본 단가는 0 이상의 정수로 입력해 주세요.');
  if(profile.catalog.some(i=>i.id!==modelDraft.id&&i.name.trim().toLowerCase()===modelDraft.name.toLowerCase()&&i.model.trim().toLowerCase()===modelDraft.model.toLowerCase()))throw new Error('동일한 장비명과 모델이 이미 등록돼 있어요. 기존 모델을 수정해 주세요.');
  validateRules(modelDraft.accessories);const next=clone(profile);const index=next.catalog.findIndex(i=>i.id===modelDraft.id);if(index<0)next.catalog.push(modelDraft);else next.catalog[index]=modelDraft;
  await put('state',next,'profile');profile=next;renderCatalog();renderSetup();$('#model-dialog').close();toast('모델과 추가 준비물 규칙을 저장했습니다.');
});});
function updatePackingStatus(){
  const rows=packingRows(q).filter(r=>r.included);$('#packing-count').textContent=`${rows.length}품목 · 출고 ${rows.filter(r=>r.out).length} · 반납 ${rows.filter(r=>r.returned).length}`;
  for(const node of $$('[data-pack-key][data-check]')){const row=packingRows(q).find(r=>r.key===node.dataset.packKey);if(row)node.checked=!!row[node.dataset.check];}
}
function renderPacking(){
  const rows=packingRows(q);$('#packing-empty').hidden=!!rows.length;
  const checkbox=(r,k,label)=>`<input type="checkbox" aria-label="${esc(r.name||'준비물')} ${label}" data-pack-key="${esc(r.key)}" data-check="${k}" ${r[k]?'checked':''} ${k!=='included'&&!r.included?'disabled':''}>`;
  const input=(r,k,label)=>`<input aria-label="현장 준비물 ${label}" maxlength="100" data-extra="${esc(r.id)}" data-key="${k}" ${k==='quantity'?'type="number" min="1" max="100000" step="1"':''} value="${esc(r[k]??'')}">`;
  $('#packing-rows').innerHTML=rows.map(r=>`<tr class="${r.included?'':'excluded'}"><td>${checkbox(r,'included','포함')}</td><td>${r.kind==='현장 추가'?input(r,'name','이름')+input(r,'model','모델'):`<strong>${esc(r.name)}</strong><small>${esc(r.model)}</small>`}</td><td>${r.kind==='현장 추가'?input(r,'quantity','수량')+input(r,'unit','단위'):`${Number.isFinite(r.quantity)?r.quantity:'—'} ${esc(r.unit)}`}</td><td>${checkbox(r,'out','출고 확인')}</td><td>${checkbox(r,'returned','반납 확인')}</td><td>${esc(r.kind)}<small>${esc(r.source)}</small></td><td>${r.kind==='현장 추가'?input(r,'note','비고'):esc(r.note)}</td><td>${r.kind==='현장 추가'?`<button type="button" data-remove-extra="${esc(r.id)}" aria-label="현장 준비물 삭제">×</button>`:''}</td></tr>`).join('');updatePackingStatus();
}
function renderDeviceOptions(){$('#io-devices').innerHTML=[...new Set(q.items.map(i=>i.model).filter(Boolean))].map(m=>`<option value="${esc(m)}"></option>`).join('');}
const ioHeadings={inputs:['CH','소스','마이크 / DI','패치','48V','스탠드','위치','비고'],outputs:['OUT','버스','용도 / 목적지','출력 장비','패치','커넥터','위치','비고']};
function renderIO(){
  for(const kind of ['inputs','outputs'])$('#io-'+kind).innerHTML=q.io[kind].map((row,i)=>`<tr>${IO_FIELDS[kind].map((key,col)=>`<td>${key==='phantom'?`<select aria-label="인풋 ${i+1} ${ioHeadings[kind][col]}" data-io="${kind}" data-row="${i}" data-key="${key}">${['','미정','OFF','ON'].map(v=>`<option value="${v}" ${row[key]===v?'selected':''}>${v||'선택'}</option>`).join('')}</select>`:`<input aria-label="${kind==='inputs'?'인풋':'아웃풋'} ${i+1} ${ioHeadings[kind][col]}" data-io="${kind}" data-row="${i}" data-key="${key}" maxlength="100" ${key==='device'?'list="io-devices"':''} value="${esc(row[key]||'')}">`}</td>`).join('')}<td><button type="button" data-remove-io="${kind}" data-row="${i}" aria-label="${kind==='inputs'?'인풋':'아웃풋'} ${i+1} 삭제">×</button></td></tr>`).join('');
  $$('[data-io-meta]').forEach(n=>n.value=q.io[n.dataset.ioMeta]||'');renderDeviceOptions();
}
function addIORows(kind,count){
  if(q.io[kind].length+count>128)throw new Error('각각 128채널까지 추가할 수 있어요.');
  let number=Math.max(0,...q.io[kind].map(r=>/^\d+$/.test(r.channel)?+r.channel:0));
  for(let i=0;i<count;i++)q.io[kind].push(blankIO(kind,number++));changed();renderIO();
}
function previewIO(){
  validateIO(q);const layout=ioLayout(q),td=(v,attrs='')=>`<td ${attrs}>${esc(v)}</td>`;
  const rows=layout.inputs.map((r,i)=>{let right='';const output=layout.outputs[i];
    if(output)right=td(output.channel,'class="io-divider"')+(output.span?td(output.name,`rowspan="${output.span}"`)+td(output.note,`rowspan="${output.span}"`):'');
    else if(i===layout.outputCount)right='<th colspan="3" class="io-divider io-check-head">Check List</th>';
    else right=td(layout.checklist[i-layout.outputCount-1]||'','colspan="3" class="io-divider io-check-line"');
    return `<tr>${td(r.channel)}${td(r.name)}${td(r.std)}${td(r.note)}${right}</tr>`;
  }).join('');
  $('#preview').innerHTML=`<div class="io-proof-scroll"><article class="io-proof"><header class="io-proof-header"><img src="../assets/oneul-logo-original.png" alt="사운드 오늘"><h3>${esc(q.project?q.project+' I/O Sheet':'I/O Sheet')}</h3></header><p class="io-proof-meta">일정: ${esc(q.start)}${q.end&&q.end!==q.start?' ~ '+esc(q.end):''}　 장소: ${esc(q.venue)}</p><p class="io-proof-contact">${esc(ioContact(q,profile.supplier))}</p><table class="io-proof-table"><colgroup>${[6,19,7,23,6,18,21].map(w=>`<col class="io-col-${w}">`).join('')}</colgroup><thead><tr><th colspan="7" class="io-sheet-title">I/O Sheet</th></tr><tr><th colspan="4">Input</th><th colspan="3" class="io-divider">Monitor Output</th></tr><tr>${['No.','Name','STD','Use/Note','No.','Name','Use/Note'].map((v,i)=>`<th ${i===4?'class="io-divider"':''}>${v}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table>${q.io.notes?`<p class="io-proof-notes"><b>현장 기술 메모</b><br>${esc(q.io.notes)}</p>`:''}</article></div>`;$('#preview-dialog').showModal();
}
async function exportBundle(){
  validate(q);validateIO(q);const packing=validatePacking(q);
  if(!profile.quote||!profile.gear)throw new Error('견적서와 장비리스트 원본 양식을 먼저 등록해 주세요.');
  await save();const z=new JSZip();z.file(filename('견적서'),await exportQuote(profile.quote.bytes,q,profile.supplier));z.file(filename('출고장비리스트'),await exportEquipment(profile.gear.bytes,q,profile.supplier,packing));z.file(filename('인풋아웃풋리스트'),await exportIO(q,profile.supplier));
  download(await z.generateAsync({type:'uint8array',compression:'DEFLATE'}),filename('현장서류').replace(/xlsx$/,'zip'),'application/zip');toast('견적서·출고 장비리스트·인풋아웃풋 3종을 만들었습니다.');
}
window.addEventListener('beforeunload',e=>{if(saveTimer&&status&&$('#save-state').textContent==='저장 중…'){e.preventDefault();e.returnValue='';}});
async function init(){
  try{db=await openDB();profile={...profile,...await read('state','profile')};profile.supplier={...DEFAULT_SUPPLIER,...profile.supplier};quotes=(await read('quotes')).map(normalizeQuote);const current=await read('state','current');q=normalizeQuote(clone(quotes.find(x=>x.id===current))||freshQuote(profile.fees));renderSetup();render();$('#setup').open=!(profile.quote&&profile.gear);}
  catch(e){status('저장소 연결 실패');toast('브라우저 저장소를 열 수 없어요. 일반 창에서 다시 열어 주세요.');$$('button').forEach(b=>b.disabled=true);console.error(e);}
}
init();
