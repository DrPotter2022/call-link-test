(function promptDesktopOnMobilePage() {
      try {
        const url = new URL(window.location.href);
        if (url.searchParams.get("view") === "mobile") return;
        const ua = navigator.userAgent || "";
        const mobileUa = /Android|iPhone|iPad|iPod|Mobile|Windows Phone|webOS/i.test(ua);
        const touch = navigator.maxTouchPoints > 0;
        const narrow = Math.min(window.screen.width || 9999, window.innerWidth || 9999) <= 900;
        const isDesktop = !mobileUa && !(touch && narrow);
        if (!isDesktop) return;
        if (sessionStorage.getItem("mobile_to_desktop_prompted") === "1") return;
        sessionStorage.setItem("mobile_to_desktop_prompted", "1");
        const ok = window.confirm("PC表示向けのページへ切り替えますか？\n「OK」で index.html を開きます。");
        if (!ok) return;
        const next = new URL("index.html", url);
        url.searchParams.forEach((v, k) => next.searchParams.set(k, v));
        next.hash = url.hash;
        window.location.replace(next.toString());
      } catch {}
    })();
  