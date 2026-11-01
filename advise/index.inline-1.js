document.querySelectorAll('.accordion-header').forEach(header => {
    header.addEventListener('click', () => {
      const accordion = header.parentElement;
      accordion.classList.toggle('active');
    });
  });

  function toggleMenu() {
    const menu = document.getElementById('menuDropdown');
    menu.style.display = (menu.style.display === 'block') ? 'none' : 'block';
  }

  window.addEventListener('click', function(e) {
    const button = document.querySelector('.menu-button');
    const menu = document.getElementById('menuDropdown');
    if (!button.contains(e.target) && !menu.contains(e.target)) {
      menu.style.display = 'none';
    }
  });
