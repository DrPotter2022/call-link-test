// スクロール出現（reduced motion 尊重）
  const prefersReduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const targets = document.querySelectorAll('.reveal');
  if (prefersReduced || !('IntersectionObserver' in window)) {
    targets.forEach(el => el.classList.add('on'));
  } else {
    const io = new IntersectionObserver(entries => {
      entries.forEach(e => {
        if (e.isIntersecting) { e.target.classList.add('on'); io.unobserve(e.target); }
      });
    }, { threshold: 0.15 });
    targets.forEach(el => io.observe(el));
  }
