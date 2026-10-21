export const VERSION = 2;
export const money = n => Number(n).toLocaleString('ko-KR');
export const today = () => new Date(Date.now()-new Date().getTimezoneOffset()*60000).toISOString().slice(0,10);
export const uid = () => crypto.randomUUID();
export const clone = x => structuredClone(x);
export const DEFAULT_SUPPLIER = {name:'사운드 오늘', representative:'권지원', number:'309-14-63535', phone:'010-4985-3618', email:'sound_oneul@naver.com', address:'인천광역시 강화군 화도면 해안남로 1114번길 18'};
export function freshQuote(fees=[]){
  return {id:uid(),number:`ON-${today().replaceAll('-','')}-${Math.random().toString(36).slice(2,6).toUpperCase()}`,date:today(),client:'',contact:'',project:'',venue:'',start:'',end:'',rentalDays:1,setupDays:0,setupPercent:20,discount:0,vatOn:true,items:[],fees:fees.map(f=>({...f,quantity:0,days:1})),notes:'',issuer:'권지원',updatedAt:''};
}
function numeric(v,label,{integer=false,max=1e12}={}){
  if(v===''||v===null||v===undefined||!Number.isFinite(Number(v))||Number(v)<0||Number(v)>max||(integer&&!Number.isInteger(Number(v)))) throw new Error(`${label}에 올바른 숫자를 입력해 주세요.`);
  return Number(v);
}
export function calculate(q){
  const lines=q.items.map((item,i)=>Math.round(numeric(item.quantity,`${i+1}번 장비 수량`,{integer:true,max:100000})*numeric(item.rate,`${i+1}번 장비 단가`,{integer:true,max:1e9})));
  const daily=lines.reduce((a,b)=>a+b,0);
  const rental=daily*numeric(q.rentalDays,'임대 일수',{integer:true,max:3660});
  const setupRate=Math.floor(daily*numeric(q.setupPercent,'설치일 청구 비율',{max:100})/100);
  const setup=setupRate*numeric(q.setupDays,'설치 일수',{integer:true,max:3660});
  const feeLines=q.fees.map((f,i)=>Math.round(numeric(f.rate,`${i+1}번 비용 단가`,{integer:true,max:1e9})*numeric(f.quantity,`${i+1}번 비용 수량`,{integer:true,max:100000})*numeric(f.days,`${i+1}번 비용 횟수`,{max:3660})));
  const fees=feeLines.reduce((a,b)=>a+b,0), beforeDiscount=rental+setup+fees;
  const discount=numeric(q.discount,'네고 할인액',{integer:true});
  if(discount>beforeDiscount)throw new Error('네고 할인액은 할인 전 공급가액보다 클 수 없어요.');
  const supply=beforeDiscount-discount, vat=q.vatOn?Math.floor(supply/10):0;
  if(!Number.isSafeInteger(supply+vat))throw new Error('계산 가능한 금액 범위를 초과했어요.');
  return {lines,daily,rental,setupRate,setup,feeLines,fees,beforeDiscount,discount,supply,vat,total:supply+vat};
}
export function validate(q){
  if(!q.client?.trim())throw new Error('거래처를 입력해 주세요.');
  if(!q.project?.trim())throw new Error('행사명을 입력해 주세요.');
  for(const [key,name] of [['date','견적일'],['start','행사 시작일'],['end','행사 종료일']]){
    if(!/^\d{4}-\d{2}-\d{2}$/.test(q[key])||Number.isNaN(Date.parse(q[key])))throw new Error(`${name}을 입력해 주세요.`);
  }
  if(q.end<q.start)throw new Error('행사 종료일은 시작일보다 빠를 수 없어요.');
  if(!q.items.length&&!q.fees.some(f=>+f.quantity>0))throw new Error('장비 또는 인건비·운송비를 추가해 주세요.');
  if(q.items.length>120||q.fees.length>40)throw new Error('장비는 120개 품목, 부대비용은 40개 항목까지 작성할 수 있어요.');
  if(q.items.some(i=>!i.model?.trim()||!i.name?.trim()||Number(i.quantity)<=0))throw new Error('장비명·모델명과 1개 이상의 수량을 입력해 주세요.');
  if(q.fees.some(f=>!f.name?.trim()))throw new Error('인건비·운송비 항목명을 입력해 주세요.');
  return calculate(q);
}
export function toJob(q){
  const c=validate(q);
  return {title:q.project,client:q.client,contact:q.contact,venue:q.venue,start:q.start,end:q.end,supply:c.supply,vatOn:q.vatOn,gear:[...new Set(q.items.map(i=>i.model))],memo:`견적번호: ${q.number}\n${q.notes||''}`.trim()};
}
export function mergeCatalog(quoteItems,equipment){
  const key=x=>`${x.name.trim().toLowerCase()}|${x.model.trim().toLowerCase()}`;
  const prices=new Map(quoteItems.map(i=>[key(i),i.rate]));
  const list=equipment.length?equipment:quoteItems;
  const result=list.map(i=>({...i,id:uid(),rate:prices.has(key(i))?prices.get(key(i)):'',sourceQuantity:i.quantity}));
  for(const i of quoteItems) if(!result.some(r=>key(r)===key(i))) result.push({...i,id:uid(),sourceQuantity:i.quantity});
  return result;
}
