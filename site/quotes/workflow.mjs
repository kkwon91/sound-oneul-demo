import {clone,uid} from './core.mjs';

const CHO=[...'ㄱㄲㄴㄷㄸㄹㅁㅂㅃㅅㅆㅇㅈㅉㅊㅋㅌㅍㅎ'];
const norm=s=>String(s??'').normalize('NFKC').toLowerCase().replace(/[\s\p{P}\p{S}]/gu,'');
export function initials(s){return [...String(s??'').normalize('NFC')].map(c=>{const n=c.charCodeAt(0)-0xac00;return n>=0&&n<11172?CHO[Math.floor(n/588)]:c;}).join('');}
export function categoryAliases(name){
  const n=String(name).toLowerCase();
  return [/mixer/.test(n)?'믹서 콘솔 믹싱':'',/wireless/.test(n)?'무선 마이크 무선마이크':'',/speaker/.test(n)?'스피커 메인스피커':'',/stand/.test(n)?'스탠드':'',/power/.test(n)?'전원 전기 파워':'',/signal/.test(n)?'시그널 신호 케이블':''].filter(Boolean).join(' ');
}
export function searchCatalog(catalog,term,{includeInactive=false}={}){
  const tokens=String(term??'').trim().split(/\s+/).filter(Boolean);
  return catalog.filter(i=>(includeInactive||!i.inactive)&&tokens.every(token=>{
    const fields=[i.name,i.model,i.maker,i.aliases,categoryAliases(i.name)];
    // Match each character as typed, or a Hangul initial against a syllable.
    const needle=norm(token);
    return fields.some(field=>{const hay=norm(field),cho=norm(initials(field));if(hay.includes(needle)||cho.includes(needle))return true;
      for(let p=0;p<=hay.length-needle.length;p++)if([...needle].every((c,k)=>c===hay[p+k]||c===norm(initials(hay[p+k]))))return true;return false;});
  }));
}
export function normalizeQuote(q){
  q.items=(q.items||[]).map(i=>({...i,lineId:i.lineId||uid(),accessories:Array.isArray(i.accessories)?i.accessories:[]}));
  q.packingExtras=Array.isArray(q.packingExtras)?q.packingExtras:[];
  q.packingChecks=q.packingChecks&&typeof q.packingChecks==='object'?q.packingChecks:{};
  q.packingExcluded=Array.isArray(q.packingExcluded)?q.packingExcluded:[];
  q.io={inputs:[],outputs:[],notes:'',checklist:'',contactName:'',contactPhone:'',contactEmail:'',...q.io};
  return q;
}
export function itemFromCatalog(item,quantity=1){return {lineId:uid(),catalogId:item.id,name:item.name,model:item.model,maker:item.maker||'',quantity,unit:item.unit||'EA',rate:item.rate,note:item.note||'',accessories:clone(item.accessories||[])};}
export function reconcileCatalog(imported,existing){
  const key=i=>norm(i.name)+'|'+norm(i.model);
  const result=imported.map(i=>{const old=existing.find(x=>key(x)===key(i));return old?{...i,...(old.custom?old:{}),id:old.id,rate:old.rate,aliases:old.aliases||'',accessories:old.accessories||[],inactive:!!old.inactive,sourceQuantity:i.sourceQuantity}:{...i,aliases:'',accessories:[]};});
  for(const old of existing)if(!result.some(i=>key(i)===key(old)))result.push(old);
  return result;
}
export function validateRules(rules){
  for(const r of rules){if(!String(r.name||'').trim())throw new Error('추가 준비물 이름을 입력해 주세요.');
    if(!Number.isFinite(+r.quantity)||+r.quantity<=0||+r.quantity>100000)throw new Error('준비물 기준 수량은 0보다 큰 숫자로 입력해 주세요.');
    if(!['unit','event'].includes(r.mode))throw new Error('준비물 수량 기준을 확인해 주세요.');}
}
export function packingRows(q){
  const rows=[],auto=new Map();
  for(const i of q.items){
    rows.push({key:'item:'+i.lineId,name:i.name,model:i.model,maker:i.maker,unit:i.unit||'EA',quantity:Number(i.quantity),kind:'견적 장비',source:'견적 포함',note:i.note||''});
    for(const r of i.accessories||[]){
      const key='auto:'+norm(r.name)+'|'+norm(r.model)+'|'+norm(r.unit||'EA');
      if(!auto.has(key))auto.set(key,{key,name:r.name,model:r.model||'',unit:r.unit||'EA',quantity:0,kind:'연결 준비물',sources:[],parts:new Map(),note:r.note||''});
      const row=auto.get(key),ruleKey=(i.catalogId||i.lineId)+':'+r.id;
      const part=row.parts.get(ruleKey)||{mode:r.mode,quantity:+r.quantity,parentQuantity:0};part.parentQuantity+=Number(i.quantity);row.parts.set(ruleKey,part);
      if(!row.sources.includes(i.model))row.sources.push(i.model);
    }
  }
  for(const row of auto.values()){
    row.quantity=[...row.parts.values()].reduce((n,r)=>n+(r.mode==='event'?Math.ceil(r.quantity):Math.ceil(r.quantity*r.parentQuantity)),0);
    row.source=row.sources.join(', ');delete row.parts;delete row.sources;rows.push(row);
  }
  for(const i of q.packingExtras||[])rows.push({...i,key:'extra:'+i.id,quantity:Number(i.quantity),kind:'현장 추가',source:'직접 추가'});
  return rows.map(row=>{const signature=JSON.stringify([row.name,row.model,row.unit,row.quantity,row.source]);const check=q.packingChecks?.[row.key];return {...row,signature,included:!(q.packingExcluded||[]).includes(row.key),out:check?.signature===signature&&!!check.out,returned:check?.signature===signature&&!!check.returned};});
}
export function validatePacking(q){
  const rows=packingRows(q).filter(r=>r.included);
  if(rows.length>300)throw new Error('출고 장비는 300개 품목까지 작성할 수 있어요.');
  for(const r of rows)if(!r.name?.trim()||!Number.isInteger(r.quantity)||r.quantity<=0||r.quantity>100000)throw new Error('출고 장비의 이름과 수량(1개 이상 정수)을 확인해 주세요.');
  return rows;
}
export const IO_FIELDS={inputs:['channel','source','device','patch','phantom','stand','position','note'],outputs:['channel','bus','destination','device','patch','connector','position','note']};
export function blankIO(kind,index){return {id:uid(),...Object.fromEntries(IO_FIELDS[kind].map(k=>[k,k==='channel'?String(index+1):'']))};}
export function validateIO(q){
  for(const kind of ['inputs','outputs']){
    const rows=q.io?.[kind]||[];if(rows.length>128)throw new Error('인풋과 아웃풋은 각각 128채널까지 작성할 수 있어요.');
    const used=new Set();for(const r of rows){const channel=String(r.channel||'').trim().toLowerCase();if(!channel)throw new Error('인풋·아웃풋 채널 번호를 입력해 주세요.');if(used.has(channel))throw new Error(`${kind==='inputs'?'인풋':'아웃풋'} 채널 ${r.channel}이 중복됩니다.`);used.add(channel);}
  }
  return q.io;
}

// One presentation model keeps the browser proof and Excel sheet consistent.
export function ioLayout(q){
  const io=q.io||{},checklist=String(io.checklist||'').replace(/\r/g,'').split('\n');
  const inputCount=Math.max(32,io.inputs?.length||0),outputCount=Math.max(24,io.outputs?.length||0);
  const count=Math.max(inputCount,outputCount+1+Math.max(7,checklist.length));
  const detail=(r,fields)=>fields.filter(([key])=>r[key]).map(([key,label])=>label?`${label}: ${r[key]}`:r[key]).join('\n');
  const inputs=Array.from({length:count},(_,i)=>{const r=io.inputs?.[i];return r?{channel:r.channel,name:r.source||'',std:r.stand||'',note:detail(r,[['device','장비'],['patch','패치'],['phantom','48V'],['position','위치'],['note','']])}:{channel:i<inputCount?String(i+1):'',name:'',std:'',note:''};});
  const outputs=Array.from({length:outputCount},(_,i)=>{const r=io.outputs?.[i];return r?{channel:r.channel,name:r.bus||r.destination||'',note:detail(r,[...(r.bus?[['destination','']]:[]),['device','장비'],['patch','패치'],['connector','연결'],['position','위치'],['note','']]),span:1}:{channel:String(i+1),name:'',note:'',span:1};});
  // Only identical adjacent populated labels/notes merge; individual channel numbers remain.
  for(let i=0;i<outputs.length;){let end=i+1;while(outputs[i].name&&end<outputs.length&&outputs[end].name===outputs[i].name&&outputs[end].note===outputs[i].note)end++;outputs[i].span=end-i;for(let j=i+1;j<end;j++)outputs[j].span=0;i=end;}
  return {inputs,outputs,inputCount,outputCount,count,checklist};
}
export function ioContact(q,s){return [q.io?.contactName||q.issuer||s.representative,q.io?.contactPhone||s.phone,q.io?.contactEmail||s.email].filter(Boolean).join(' / ');}
