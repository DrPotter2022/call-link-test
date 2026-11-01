(function routeMobileToMobilePage() {
      try {
        const url = new URL(window.location.href);
        if (url.searchParams.get("view") === "desktop") return;
        const ua = navigator.userAgent || "";
        const mobileUa = /Android|iPhone|iPad|iPod|Mobile|Windows Phone|webOS/i.test(ua);
        const touch = navigator.maxTouchPoints > 0;
        const narrow = Math.min(window.screen.width || 9999, window.innerWidth || 9999) <= 900;
        if (mobileUa || (touch && narrow)) {
          const ok = window.confirm("スマホ向けページ（index_mobile.html）へ切り替えますか？");
          if (!ok) return;
          const next = new URL("index_mobile.html", url);
          url.searchParams.forEach((v, k) => next.searchParams.set(k, v));
          next.hash = url.hash;
          window.location.replace(next.toString());
        }
      } catch {}
    })();
  