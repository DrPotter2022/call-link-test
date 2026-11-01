document.addEventListener('DOMContentLoaded',()=>{
  const yearToggle=document.getElementById('yearGroupToggle');
  if(yearToggle){
    yearToggle.addEventListener('click',()=>{
      const boxes=[...document.querySelectorAll('#years input[type="checkbox"]')];
      const allChecked=boxes.length>0 && boxes.every(cb=>cb.checked);
      boxes.forEach(cb=>cb.checked=!allChecked);
      if(typeof render==='function') render();
    });
  }
});
