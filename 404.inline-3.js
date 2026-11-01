(function () {
    var root = location.protocol === 'file:' ? '' : '/';

    /* ヘッダー・フッターは共通スクリプトから組み立てる */
    var s = document.createElement('script');
    s.src = root + 'assets/js/site.js';
    s.onload = function () {
      WIZE.load('site').then(function (site) { WIZE.renderChrome(site, ''); });
    };
    document.body.appendChild(s);

    /* 数字を押したら、勢いよく回してから読み込み直す */
    var btn = document.getElementById('nf-reload');
    btn.addEventListener('click', function () {
      var reduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
      if (reduced) return location.reload();
      btn.classList.add('is-boost');
      setTimeout(function () { location.reload(); }, 900);
    });
  })();
