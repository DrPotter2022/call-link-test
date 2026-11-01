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
document.addEventListener('DOMContentLoaded', function () {
  const headers = document.querySelectorAll('.accordion-header');
  headers.forEach(header => {
    header.addEventListener('click', () => {
      const accordion = header.parentElement;
      accordion.classList.toggle('active');
    });
  });
});
