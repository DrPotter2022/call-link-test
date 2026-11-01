/* 管理者としてログインしている場合だけ、編集モードを読み込む。
     一般の閲覧者には余計なスクリプトを配らないため。 */
  (function () {
    try {
      var s = JSON.parse(localStorage.getItem('wize_admin_session') || 'null');
      if (!s || !s.until || Date.now() > s.until) return;
      var el = document.createElement('script');
      el.src = '../assets/js/admin-inline.js';
      document.head.appendChild(el);
    } catch (e) { /* 読み込めなければ通常表示のまま */ }
  })();
