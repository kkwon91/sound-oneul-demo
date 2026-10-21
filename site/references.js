(() => {
  const filters=document.querySelector('#filters'),records=document.querySelector('#records');
  if(!filters || !records) return;
  filters.hidden=false;
  filters.addEventListener('click',event=>{
    const button=event.target.closest('button[data-type]');if(!button)return;
    filters.querySelectorAll('button').forEach(item=>{item.classList.toggle('primary',item===button);item.setAttribute('aria-pressed',String(item===button));});
    records.querySelectorAll('article[data-type]').forEach(record=>{record.hidden=button.dataset.type!=='전체'&&record.dataset.type!==button.dataset.type;});
  });
})();
