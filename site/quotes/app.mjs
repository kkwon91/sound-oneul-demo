import {VERSION,DEFAULT_SUPPLIER,today,uid,clone,freshQuote,calculate,validate,toJob,mergeCatalog,money} from './core.mjs';
import {readWorkbook,exportQuote,exportEquipment} from './xlsx.mjs';
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let db,profile={supplier:{...DEFAULT_SUPPLIER},catalog:[],fees:[]},quotes=[],q,saveTimer,toastTimer,lastURL,busy=false;
const status=s=>$('#save-state').textContent=s;
const toast=s=>{$('#toast').textContent=s;$('#toast').hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>$('#toast').hidden=true,5000);};
const openDB=()=>new Promise((resolve,reject)=>{const r=indexedDB.open('sound-oneul-quotes',1);r.onupgradeneeded=()=>{r.result.createObjectStore('state');r.result.createObjectStore('quotes',{keyPath:'id'});};r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
const read=(store,key)=>new Promise((resolve,reject)=>{const r=db.transaction(store).objectStore(store)[key===undefined?'getAll':'get'](key);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});
const put=(store,data,key)=>new Promise((resolve,reject)=>{const tx=db.transaction(store,'readwrite');if(key===undefined)tx.objectStore(store).put(data);else tx.objectStore(store).put(data,key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);tx.onabort=()=>reject(tx.error);});
const saveProfile=()=>put('state',profile,'profile');
function meaningful(){return q&&!!(q.project||q.client||q.items.length||q.fees.some(f=>+f.quantity>0));}
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
  renderTotals();
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
  $$('[data-q]').forEach(n=>n.value=String(q[n.dataset.q]??''));renderRows();renderHistory();status(q.updatedAt?'저장된 견적':'새 견적');
  $('#archive').textContent=q.archived?'보관 해제':'보관';
}
function renderCatalog(){
  const term=$('#catalog-search').value.trim().toLowerCase();
  const items=profile.catalog.filter(i=>`${i.name} ${i.model} ${i.maker}`.toLowerCase().includes(term));
  $('#catalog-list').innerHTML=items.map(i=>`<div class="catalog-row"><div><strong>${esc(i.model)}</strong><small>${esc(i.name)} · ${esc(i.maker)} · ${esc(i.unit)}</small></div><label>1일 단가 (원)<input aria-label="${esc(i.model)} 기본 단가" type="number" min="0" step="1" data-price="${esc(i.id)}" value="${esc(i.rate)}"></label><button data-add="${esc(i.id)}">추가</button></div>`).join('')||'<p class="empty">등록한 장비가 없습니다.<br><small>원본 엑셀을 등록하거나 직접 입력해 주세요.</small></p>';
}
async function importTemplate(file,kind){
  if(!file)return;const bytes=await file.arrayBuffer();const book=await readWorkbook(bytes,kind);
  const next=clone(profile);next[kind==='quote'?'quote':'gear']={name:file.name,bytes};
  const other=kind==='quote'?next.gear:next.quote;
  const otherBook=other?await readWorkbook(other.bytes,kind==='quote'?'equipment':'quote'):null;
  const qb=kind==='quote'?book:otherBook,gb=kind==='equipment'?book:otherBook;
  next.catalog=mergeCatalog(qb?.items||[],gb?.items||[]);if(qb)next.fees=qb.fees;
  await put('state',next,'profile');profile=next;
  if(!q.fees.length&&profile.fees.length){q.fees=clone(profile.fees).map(f=>({...f,quantity:0,days:1}));if(meaningful())await save();}
  renderSetup();renderRows();$('#setup').open=!(profile.quote&&profile.gear);toast(`${kind==='quote'?'견적서 양식':'장비리스트'} 등록 완료 · 장비 ${profile.catalog.length}개`);
}
function addItem(item,quantity=1){
  if(q.items.length>=120)throw new Error('장비는 120개 품목까지 추가할 수 있어요.');
  const existing=q.items.find(i=>i.catalogId===item.id&&i.rate===item.rate);
  if(existing)existing.quantity=Number(existing.quantity)+quantity;
  else q.items.push({name:item.name,model:item.model,maker:item.maker,quantity,unit:item.unit,rate:item.rate,note:item.note||'',catalogId:item.id});
  changed();renderRows();
}
function download(bytes,name,mime){
  if(lastURL)URL.revokeObjectURL(lastURL);
  lastURL=URL.createObjectURL(new Blob([bytes],{type:mime}));const a=$('#last-download');a.href=lastURL;a.download=name;a.hidden=false;a.textContent=name+' 다시 받기';a.click();
}
const xlsxMime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const filename=(label)=>`${q.date}_${q.project}_${label}.xlsx`.replace(/[\\/:*?"<>|\x00-\x1f]/g,'_');
async function exportFile(kind){
  validate(q);const source=kind==='quote'?profile.quote:profile.gear;
  if(!source){$('#setup').open=true;throw new Error(`${kind==='quote'?'견적서':'장비리스트'} 원본 엑셀을 먼저 등록해 주세요.`);}
  if(!profile.supplier.name.trim()||!profile.supplier.representative.trim()||!profile.supplier.number.trim())throw new Error('공급자 상호·대표자·사업자등록번호를 확인해 주세요.');
  await save();const out=await (kind==='quote'?exportQuote:exportEquipment)(source.bytes,q,profile.supplier);
  download(out,filename(kind==='quote'?'견적서':'장비리스트'),xlsxMime);toast('엑셀 파일을 만들었습니다.');
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
  const base=freshQuote();for(const k of Object.keys(base))if(k in value)base[k]=value[k];base.archived=Boolean(value.archived);return base;
}
async function restore(file){
  if(!file)return;if(file.size>30*1024*1024)throw new Error('30MB 이하의 백업 파일을 선택해 주세요.');
  const z=await JSZip.loadAsync(await file.arrayBuffer()),entry=z.file('oneul-quotes.json');if(!entry)throw new Error('사운드 오늘 견적 백업 파일이 아니에요.');
  const data=JSON.parse(await entry.async('string'));if(data.version!==VERSION||!Array.isArray(data.quotes)||!Array.isArray(data.profile?.catalog)||!Array.isArray(data.profile?.fees))throw new Error('지원하지 않는 백업 형식입니다.');
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
  if(n.dataset.item!==undefined||n.dataset.fee!==undefined){const row=n.dataset.item!==undefined?q.items[+n.dataset.item]:q.fees[+n.dataset.fee];row[n.dataset.key]=n.value;changed();renderTotals();}
  if(n.dataset.supplier){profile.supplier[n.dataset.supplier]=n.value;saveProfile().catch(()=>toast('공급자 정보 저장 실패'));}
  if(n.dataset.price){const it=profile.catalog.find(i=>i.id===n.dataset.price);it.rate=n.value;saveProfile().catch(()=>toast('장비 단가 저장 실패'));}
  if(n.id==='history-search')renderHistory();if(n.id==='catalog-search')renderCatalog();
});
document.addEventListener('change',e=>{
  const n=e.target;if(n.id==='quote-template')safe(()=>importTemplate(n.files[0],'quote'));if(n.id==='gear-template')safe(()=>importTemplate(n.files[0],'equipment'));if(n.id==='restore')safe(()=>restore(n.files[0]));
});
document.addEventListener('click',e=>{
  const n=e.target.closest('button,[data-close]');if(!n)return;
  if(n.dataset.close){$('#'+n.dataset.close).close();return;}
  safe(async()=>{
    if(n.dataset.open){if(meaningful())await save();q=clone(quotes.find(x=>x.id===n.dataset.open));await put('state',q.id,'current');render();}
    if(n.dataset.restore){const x=quotes.find(x=>x.id===n.dataset.restore);x.archived=false;await put('quotes',x);if(q.id===x.id)q.archived=false;render();}
    if(n.dataset.add){addItem(profile.catalog.find(i=>i.id===n.dataset.add));toast('장비를 추가했습니다.');}
    if(n.dataset.removeItem!==undefined){q.items.splice(+n.dataset.removeItem,1);changed();renderRows();}
    if(n.dataset.removeFee!==undefined){q.fees.splice(+n.dataset.removeFee,1);changed();renderRows();}
    switch(n.id){
      case 'new':if(meaningful())await save();q=freshQuote(profile.fees);render();await put('state',null,'current');break;
      case 'save':await save();toast('견적을 저장했습니다.');break;
      case 'duplicate':await save();q={...clone(q),id:uid(),number:q.number+'-복사',date:today(),archived:false,updatedAt:''};await save();render();toast('복사한 견적을 열었습니다.');break;
      case 'archive':q.archived=!q.archived;await save();render();toast(q.archived?'보관함으로 옮겼습니다. 언제든 복원할 수 있어요.':'보관을 해제했습니다.');break;
      case 'catalog-open':renderCatalog();$('#catalog-dialog').showModal();break;
      case 'source-pack':if(!profile.catalog.length)throw new Error('원본 엑셀을 먼저 등록해 주세요.');for(const it of profile.catalog)addItem(it,Number(it.sourceQuantity)||1);toast('원본 장비 구성을 추가했습니다.');break;
      case 'custom-item':addItem({id:uid(),name:'',model:'',maker:'',unit:'EA',rate:'',note:''});break;
      case 'add-fee':if(q.fees.length>=40)throw new Error('비용은 40개 항목까지 추가할 수 있어요.');q.fees.push({name:'기타 비용',rate:'',days:1,quantity:1,daysUnit:'일',quantityUnit:'식'});changed();renderRows();break;
      case 'download-quote':await exportFile('quote');break;
      case 'download-gear':await exportFile('equipment');break;
      case 'preview-open':preview();break;
      case 'backup':await backup();break;
      case 'to-ledger':{const job=toJob(q);await save();localStorage.setItem('oneul.quote.handoff.v1',JSON.stringify({version:1,quoteId:q.id,createdAt:Date.now(),job}));location.href='../../#jobs';break;}
    }
  });
});
window.addEventListener('beforeunload',e=>{if(saveTimer&&status&&$('#save-state').textContent==='저장 중…'){e.preventDefault();e.returnValue='';}});
async function init(){
  try{db=await openDB();profile={...profile,...await read('state','profile')};profile.supplier={...DEFAULT_SUPPLIER,...profile.supplier};quotes=await read('quotes');const current=await read('state','current');q=clone(quotes.find(x=>x.id===current))||freshQuote(profile.fees);renderSetup();render();$('#setup').open=!(profile.quote&&profile.gear);}
  catch(e){status('저장소 연결 실패');toast('브라우저 저장소를 열 수 없어요. 일반 창에서 다시 열어 주세요.');$$('button').forEach(b=>b.disabled=true);console.error(e);}
}
init();
