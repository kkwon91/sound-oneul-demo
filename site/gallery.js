(() => {
 const grid=document.querySelector('#photos'), dialog=document.querySelector('#photo-viewer');
 if(!grid || !dialog || typeof dialog.showModal!=='function') return;
 const cards=[...grid.querySelectorAll('a[data-photo]')];
 if(!cards.length) return;
 const filters=document.querySelector('#photo-filters');
 let visible=cards,current=0,opener=null;
 function updateViewer(){
  const card=visible[current],source=card.querySelector('img');
  const img=document.querySelector('#viewer-image');img.src=card.href;img.alt=source.alt;
  document.querySelector('#viewer-caption').textContent=card.dataset.title;
  document.querySelector('#viewer-count').textContent=(current+1)+' / '+visible.length+' · '+card.dataset.category;
  document.querySelector('#download-photo').href=card.href;
 }
 function move(step){current=(current+step+visible.length)%visible.length;updateViewer();}
 grid.addEventListener('click',e=>{
  const card=e.target.closest('a[data-photo]');
  if(!card || e.button!==0 || e.ctrlKey || e.metaKey || e.shiftKey || e.altKey) return;
  e.preventDefault();opener=card;current=visible.indexOf(card);updateViewer();dialog.showModal();
 });
 filters.hidden=false;
 filters.addEventListener('click',e=>{
  const button=e.target.closest('button[data-category]');if(!button)return;
  const category=button.dataset.category;
  filters.querySelectorAll('button').forEach(b=>{b.classList.toggle('primary',b===button);b.setAttribute('aria-pressed',String(b===button));});
  cards.forEach(card=>{card.hidden=category!=='전체' && card.dataset.category!==category;});
  visible=cards.filter(card=>!card.hidden);
  document.querySelector('#photo-count').textContent=visible.length+'장의 사진';
 });
 document.querySelector('#close-viewer').addEventListener('click',()=>dialog.close());
 document.querySelector('#previous-photo').addEventListener('click',()=>move(-1));
 document.querySelector('#next-photo').addEventListener('click',()=>move(1));
 dialog.addEventListener('close',()=>opener?.focus());
 dialog.addEventListener('click',e=>{if(e.target===dialog)dialog.close();});
 dialog.addEventListener('keydown',e=>{
  if(e.key==='ArrowLeft'){e.preventDefault();move(-1);}
  if(e.key==='ArrowRight'){e.preventDefault();move(1);}
 });
})();
