(() => {
 const grid=document.querySelector('#photos'), dialog=document.querySelector('#photo-viewer');
 let photos=[],visible=[],current=0;
 function show(category){
  visible=photos.filter(p=>category==='전체'||p.category===category); grid.replaceChildren();
  visible.forEach((photo,index)=>{
   const card=document.createElement('button'); card.className='photo-card'; card.type='button';
   card.setAttribute('aria-label',photo.title+' 확대 보기');
   const img=new Image(); img.src=photo.src; img.alt=photo.alt; img.loading='lazy'; img.className='photo';
   const category=document.createElement('span'); category.className='tag'; category.textContent=photo.category;
   const title=document.createElement('span'); title.className='photo-title'; title.textContent=photo.title+' ↗';
   card.append(img,category,title);
   card.addEventListener('click',()=>{current=index;updateViewer();dialog.showModal();}); grid.append(card);
  });
  document.querySelector('#photo-count').textContent=visible.length+'장의 사진';
 }
 function updateViewer(){
  const photo=visible[current], img=document.querySelector('#viewer-image'); img.src=photo.src;img.alt=photo.alt;
  document.querySelector('#viewer-caption').textContent=photo.title;
  document.querySelector('#viewer-count').textContent=(current+1)+' / '+visible.length+' · '+photo.category;
  document.querySelector('#download-photo').href=photo.src;
 }
 function move(step){current=(current+step+visible.length)%visible.length;updateViewer();}
 document.querySelector('#photo-filters').addEventListener('click',e=>{
  const button=e.target.closest('button[data-category]');if(!button)return;
  document.querySelectorAll('#photo-filters button').forEach(b=>{b.classList.toggle('primary',b===button);b.setAttribute('aria-pressed',String(b===button));});
  show(button.dataset.category);
 });
 document.querySelector('#close-viewer').addEventListener('click',()=>dialog.close());
 document.querySelector('#previous-photo').addEventListener('click',()=>move(-1));
 document.querySelector('#next-photo').addEventListener('click',()=>move(1));
 dialog.addEventListener('click',e=>{if(e.target===dialog)dialog.close();});
 dialog.addEventListener('keydown',e=>{
  if(e.key==='ArrowLeft'){e.preventDefault();move(-1);}
  if(e.key==='ArrowRight'){e.preventDefault();move(1);}
 });
 fetch('gallery.json').then(r=>{if(!r.ok)throw Error();return r.json();}).then(data=>{photos=data;show('전체');})
 .catch(()=>{document.querySelector('#gallery-error').textContent='사진을 불러오지 못했습니다. 새로고침해 주세요.';});
})();

