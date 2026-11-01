/* ペライチ時代のURLからの流入を、新しい場所へ寄せる。
     apex（wize.uno）を GitHub Pages へ向けると、旧URLはすべてここへ来る。 */
  (function () {
    var map = {
      '/.note': '/note/',
      '/.note/': '/note/',
      '/note': '/note/',
      '/seminar': '/seminar/',
      '/recruit': '/',
      /* 法務ページはペライチ側に残しているので、そのまま送る */
      '/privacy_policies/view/xcfnirpkrd8': 'https://peraichi.com/privacy_policies/view/xcfnirpkrd8',
      '/terms_for_customer/view/xcfnirpkrd8': 'https://peraichi.com/terms_for_customer/view/xcfnirpkrd8',
      '/commercial_transactions/view/xcfnirpkrd8': 'https://peraichi.com/commercial_transactions/view/xcfnirpkrd8',
    };
    var path = location.pathname.replace(/\/+$/, '') || '/';
    var to = map[location.pathname] || map[path] || map[path.toLowerCase()];
    if (to && to !== location.pathname) {
      location.replace(/^https?:/.test(to) ? to : to + location.hash);
    }
  })();
