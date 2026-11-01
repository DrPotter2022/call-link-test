const DATE_MIN = new Date("2025-12-01");
    const DATE_MAX = new Date("2026-03-31");
    let currentDate = new Date("2025-12-01");
    let selectedZoneId = null;
    let selectedZoneIds = [];
    let rowSort = { key: "zone", dir: 1 };
    let dailyOps = {};
    let dailyOpsByDateFingerprint = {};
    let snowfallByDate = {};
    let snowDepthByDate = {};
    let snowChart = null;
    let chartBlinkOn = true;
    let visibleZoneIds = [];
    let zoneLayers = new Map();
    let blinkOn = true;
    let map = null;
    let zoneAlias = {};
    let zoneAliasLoose = {};
    let knownZoneIds = new Set();
    let contractorByFingerprint = {};
    let contractorEval = {};
    let zoneEvalByZone = {};
    let zoneCostByZone = {};
    let showCostDetails = false;
    const mojibakeAreaFallback = [
      "後潟工区",
      "内真部工区",
      "奥内1工区",
      "奥内2工区",
      "飛鳥工区"
    ];
    const mojibakeNameMap = {
      "蠕梧ス溷キ・蛹コ": "後潟工区",
      "蜀・悄驛ィ蟾・蛹コ": "内真部工区",
      "螂・蜀・キ・蛹コ1": "奥内1工区",
      "螂・蜀・キ・蛹コ2": "奥内2工区",
      "鬟幃ウ・蟾・蛹コ": "飛鳥工区",
      "蠕梧ｽ溷ｷ･蛹ｺ": "後潟工区",
      "蜀・悄驛ｨ蟾･蛹ｺ": "内真部工区",
      "螂･蜀・ｷ･蛹ｺ1": "奥内1工区",
      "螂･蜀・ｷ･蛹ｺ2": "奥内2工区",
      "鬟幃ｳ･蟾･蛹ｺ": "飛鳥工区"
    };

    let contractors = {};

    const statusEl = document.getElementById("status");
    const rowsEl = document.getElementById("rows");
    const detailEl = document.getElementById("detail");
    const calendarEl = document.getElementById("calendar");
    const dateLabelEl = document.getElementById("currentDateLabel");
    const currentDateBadgeEl = document.getElementById("currentDateBadge");
    const calendarPickerEl = document.getElementById("calendarPicker");
    const calendarModalEl = document.getElementById("calendarModal");
    const closeCalendarBtnEl = document.getElementById("closeCalendarBtn");
    const snowPanelEl = document.getElementById("snowPanel");
    const noteModalEl = document.getElementById("noteModal");
    const closeNoteBtnEl = document.getElementById("closeNoteBtn");
    const toggleCalendarBtnEl = document.getElementById("toggleCalendarBtn");
    const toggleSnowBtnEl = document.getElementById("toggleSnowBtn");
    const toggleNoteBtnEl = document.getElementById("toggleNoteBtn");
    const commonSnowfallLabelEl = document.getElementById("commonSnowfallLabel");
    const commonSnowDepthLabelEl = document.getElementById("commonSnowDepthLabel");
    const snowChartCanvas = document.getElementById("snowChart");
    const detailTableEl = document.getElementById("detailTable");
    const toggleCostColsBtnEl = document.getElementById("toggleCostColsBtn");
    const selectedZonesCardsEl = document.getElementById("selectedZonesCards");
    const zonePrimaryEl = document.getElementById("zonePrimary");
    const zoneBranchEl = document.getElementById("zoneBranch");
    const selectedZoneTagsEl = document.getElementById("selectedZoneTags");

    const PRIMARY_ORDER = ["A","B","C","D","E","F","G","H","I","J","K","L","その他"];

    function fmtDate(date) {
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    }

    function clampDate(date) {
      if (date < DATE_MIN) return new Date(DATE_MIN);
      if (date > DATE_MAX) return new Date(DATE_MAX);
      return date;
    }

    function normalizeZoneId(name) {
      const src = String(name || "")
        .replace(/\u3000/g, " "); // full-width space -> half-width space

      const base = src
        .normalize("NFKC")
        .trim()
        .replace(/\s+/g, "")
        .replace(/[‐‑‒–—―ー−－ｰ]/g, "-")
        .replace(/[・･]/g, "・")
        .replace(/[一壱]/g, "1")
        .replace(/[二弐]/g, "2")
        .replace(/[三参]/g, "3")
        .replace(/[四]/g, "4")
        .replace(/[五]/g, "5")
        .replace(/[六]/g, "6")
        .replace(/[七]/g, "7")
        .replace(/[八]/g, "8")
        .replace(/[九]/g, "9")
        .replace(/[〇零]/g, "0")
        .replace(/A\s*-\s*/i, "A-")
        .replace(/B\s*-\s*/i, "B-")
        .replace(/工区(\d)$/u, "$1工区")
        .replace(/^奥内工区(\d)$/u, "奥内$1工区")
        .replace(/^3本木工区$/u, "三本木工区");

      // Normalize block IDs regardless of full/half width mix: A-1, B-2-1, C-11-2A...
      const blockMatch = base.match(/^([A-L])\s*[-]?\s*(\d+)(?:\s*[-]\s*([0-9A-Z]+))?(?:\s*[-]\s*([0-9A-Z]+))?$/i);
      let normalized = base;
      if (blockMatch) {
        const parts = [blockMatch[1].toUpperCase(), blockMatch[2], blockMatch[3], blockMatch[4]].filter(Boolean);
        normalized = `${parts[0]}-${parts.slice(1).join("-")}`;
      }

      return mojibakeNameMap[normalized] || normalized;
    }

    function looseZoneKey(name) {
      return String(name || "")
        .normalize("NFKC")
        .replace(/\s+/g, "")
        .replace(/[・･]/g, "")
        .replace(/[‐‑‒–—―ー−－ｰ]/g, "")
        .replace(/[\uFFFD]/g, "")
        .replace(/[^\p{L}\p{N}]/gu, "")
        .toUpperCase();
    }

    function canonicalZoneId(zoneId) {
      const n = normalizeZoneId(zoneId);
      if (zoneAlias[n]) return zoneAlias[n];
      if (/^[A-L]-\d/i.test(n)) return n;
      const lk = looseZoneKey(zoneId);
      if (lk && zoneAliasLoose[lk]) return zoneAliasLoose[lk];
      return n;
    }

    function normalizeDailyOps(rawOps) {
      const out = {};
      for (const [dateKey, zones] of Object.entries(rawOps || {})) {
        out[dateKey] = {};
        for (const [zoneKey, op] of Object.entries(zones || {})) {
          out[dateKey][normalizeZoneId(zoneKey)] = {
            command: !!op.command,
            commandCount: Number(op.commandCount || 0),
            continued: !!op.continued,
            executed: !!op.executed
          };
        }
      }
      return out;
    }

    function canonicalizeDailyOps(rawOps) {
      const out = {};
      for (const [dateKey, zones] of Object.entries(rawOps || {})) {
        out[dateKey] = out[dateKey] || {};
        for (const [zoneKey, op] of Object.entries(zones || {})) {
          const id = canonicalZoneId(zoneKey);
          const prev = out[dateKey][id] || { command: false, commandCount: 0, continued: false, executed: false };
          out[dateKey][id] = {
            command: prev.command || !!op.command,
            commandCount: Math.max(Number(prev.commandCount || 0), Number(op.commandCount || 0)),
            continued: prev.continued || !!op.continued,
            executed: prev.executed || !!op.executed
          };
        }
      }
      return out;
    }

    function canonicalizeContractors(rawContractors) {
      const out = {};
      for (const [zoneKey, contractor] of Object.entries(rawContractors || {})) {
        out[canonicalZoneId(zoneKey)] = contractor;
      }
      return out;
    }

    function canonicalizeContractorEval(rawEval) {
      const out = {};
      for (const [name, v] of Object.entries(rawEval || {})) {
        const score = Number(v?.score ?? v?.評価点 ?? NaN);
        const grade = String(v?.grade ?? v?.評価 ?? "").trim();
        out[String(name)] = {
          score: Number.isFinite(score) ? score : null,
          grade: grade || null
        };
      }
      return out;
    }

    function canonicalizeZoneEval(rawEval) {
      const src = rawEval?.byZone || rawEval || {};
      const out = {};
      for (const [zoneName, v] of Object.entries(src)) {
        const key = normalizeZoneId(zoneName);
        const score = Number(v?.score ?? v?.評点 ?? NaN);
        const grade = String(v?.grade ?? v?.評価 ?? "").trim();
        out[key] = {
          score: Number.isFinite(score) ? score : null,
          grade: grade || null
        };
      }
      return out;
    }

    function canonicalizeZoneCost(rawCost) {
      const src = rawCost?.rows || [];
      const out = {};
      src.forEach((r) => {
        const key = canonicalZoneId(r?.key || "");
        if (!key) return;
        out[key] = {
          d: r?.d ?? "",
          e: r?.e ?? "",
          f: r?.f ?? "",
          g: r?.g ?? "",
          h: r?.h ?? "",
          i: r?.i ?? "",
          j: r?.j ?? "",
          k: r?.k ?? "",
          l: r?.l ?? "",
          m: r?.m ?? "",
          n: r?.n ?? "",
          o: r?.o ?? ""
        };
      });
      return out;
    }

    function fmtCost(v) {
      const n = Number(v);
      if (Number.isFinite(n)) return n.toLocaleString("ja-JP");
      return String(v || "-");
    }

    function buildCostHtml(cost) {
      return cost
        ? `<div style="margin-top:8px;">
             <div>除排雪延長: <strong>${fmtCost(cost.d)}</strong></div>
             <div>①当初契約額: <strong>${fmtCost(cost.e)}</strong></div>
             <details style="margin-top:4px;">
               <summary style="cursor:pointer; user-select:none;">＋ 詳細内訳を表示</summary>
               <div style="margin-top:6px;">
                 <div>②工区延長: ${fmtCost(cost.f)}</div>
                 <div>③1/25~2/5降雪量見合(147cm): ${fmtCost(cost.g)}</div>
                 <div>④2/6~2/21降雪量見合い(72cm): ${fmtCost(cost.h)}</div>
                 <div>⑤2/6~2/21時間単価見合 合計: ${fmtCost(cost.i)}</div>
                 <div>市負担: ${fmtCost(cost.j)}</div>
                 <div>⑤県負担: ${fmtCost(cost.k)}</div>
                 <div>⑥2/22~3/14降雪量見合(17cm): ${fmtCost(cost.l)}</div>
                 <div>⑦応援除雪: ${fmtCost(cost.m)}</div>
                 <div>⑧市民雪寄せ場排雪業務: ${fmtCost(cost.n)}</div>
                 <div>①~⑧変更後委託料: ${fmtCost(cost.o)}</div>
               </div>
             </details>
           </div>`
        : `<div style="margin-top:8px;">除排雪延長: -<br>①当初契約額: -</div>`;
    }

    function buildZoneStatusHtml(d, timing) {
      const isPeak = timing.daysToGreenClear !== null && timing.seasonMaxClearDays !== null && timing.daysToGreenClear === timing.seasonMaxClearDays;
      const clearText = timing.daysToGreenClear ?? "-";
      return (
        `${d.command ? '<span class="chip red">指令日</span>' : ""}` +
        `${d.continued ? '<span class="chip green">指令継続</span>' : ""}` +
        `${d.commandCount ? `<span>指令回数: <strong>${d.commandCount}</strong></span>` : ""}` +
        `${d.executed ? "<br><b>除排雪実施日</b>" : ""}` +
        `<br>実施状況: ${d.command ? "指令" : "-"} / ${d.continued ? `継続(${timing.continuedDays}日)` : "-"} / ${d.executed ? `実施(${timing.daysToExecution}日)` : "-"} / ${isPeak ? `<span class="alert-red">解除(${clearText}日)</span>` : `解除(${clearText}日)`}` +
        `<br>当期の解除に要した最大日数: ${timing.seasonMaxClearDays ?? "-"}日`
      );
    }

    function buildHoverInfoHtml(zoneId) {
      const d = getDayOps(fmtDate(currentDate), zoneId);
      const timing = buildZoneTiming(zoneId, currentDate);
      const contractorName = contractors[zoneId] || contractorByFingerprint[zoneFingerprint(zoneId)] || "未設定";
      const zoneEval = zoneEvalByZone[normalizeZoneId(zoneId)] || null;
      const cost = zoneCostByZone[normalizeZoneId(zoneId)] || null;
      return (
        `<strong style="font-size:34px;">${zoneId}</strong><br>` +
        `業者: ${contractorName}<br>` +
        `評点・評価: ${zoneEval ? `${zoneEval.score ?? "-"}点・${zoneEval.grade ?? "-"}` : "未設定"}<br>` +
        buildZoneStatusHtml(d, timing) +
        `<br>除排雪延長: <strong>${cost ? fmtCost(cost.d) : "-"}</strong>` +
        `<br>①当初契約額: <strong>${cost ? fmtCost(cost.e) : "-"}</strong>`
      );
    }

    function zoneFingerprint(zoneId) {
      return String(zoneId || "")
        .normalize("NFKC")
        .toUpperCase()
        .replace(/[\u3000\s\-‐‑‒–—―ーｰ・･]/g, "");
    }

    function buildDailyOpsFingerprintIndex() {
      const idx = {};
      for (const [dateKey, zones] of Object.entries(dailyOps || {})) {
        idx[dateKey] = {};
        for (const [zoneKey, op] of Object.entries(zones || {})) {
          const fp = zoneFingerprint(zoneKey);
          const prev = idx[dateKey][fp] || { command: false, commandCount: 0, continued: false, executed: false };
          idx[dateKey][fp] = {
            command: prev.command || !!op.command,
            commandCount: Math.max(Number(prev.commandCount || 0), Number(op.commandCount || 0)),
            continued: prev.continued || !!op.continued,
            executed: prev.executed || !!op.executed
          };
        }
      }
      dailyOpsByDateFingerprint = idx;
    }

    function buildContractorFingerprintIndex() {
      const idx = {};
      for (const [zoneKey, contractor] of Object.entries(contractors || {})) {
        idx[zoneFingerprint(zoneKey)] = contractor;
      }
      contractorByFingerprint = idx;
    }

    function buildKnownZoneIds() {
      const set = new Set();
      Object.values(dailyOps || {}).forEach((zones) => {
        Object.keys(zones || {}).forEach((z) => set.add(canonicalZoneId(z)));
      });
      Object.keys(contractors || {}).forEach((z) => set.add(canonicalZoneId(z)));
      Object.keys(zoneAlias || {}).forEach((z) => set.add(canonicalZoneId(z)));
      Object.values(zoneAlias || {}).forEach((z) => set.add(canonicalZoneId(z)));
      knownZoneIds = set;
    }

    function repairBrokenZoneId(id, alreadyUsedIds) {
      const s = String(id || "");
      if (!s) return s;
      if (knownZoneIds.has(s)) return s;

      // Keep only block-pattern characters and normalize full-width forms.
      const normalized = s.normalize("NFKC").replace(/[^A-Z0-9-]/g, "");
      if (knownZoneIds.has(normalized)) return normalized;

      // Try exact block-number prefix matching (e.g. "B-2-" -> "B-2-1"/"B-2-2").
      const prefixMatch = normalized.match(/^([A-L]-\d+-)/);
      if (prefixMatch) {
        const prefix = prefixMatch[1];
        const candidates = Array.from(knownZoneIds)
          .filter((z) => z.startsWith(prefix))
          .sort((a, b) => a.localeCompare(b, "ja"));
        if (candidates.length === 1) return candidates[0];
        const unused = candidates.find((c) => !alreadyUsedIds.has(c));
        if (unused) return unused;
      }

      // Try "A-12" -> closest exact known id if available.
      const baseMatch = normalized.match(/^([A-L]-\d+)$/);
      if (baseMatch && knownZoneIds.has(baseMatch[1])) return baseMatch[1];

      return s;
    }

    function decodeKmlTextFromBuffer(buffer) {
      const decoders = ["utf-8", "shift-jis"];
      let bestText = "";
      let bestScore = Number.POSITIVE_INFINITY;

      decoders.forEach((enc) => {
        try {
          const text = new TextDecoder(enc).decode(buffer);
          const mojibakeHits = (text.match(/[ 蜀螂蛹鬟蠕]/g) || []).length;
          if (mojibakeHits < bestScore) {
            bestScore = mojibakeHits;
            bestText = text;
          }
        } catch {
          // Ignore unsupported decoder and continue fallback candidates.
        }
      });

      return bestText || new TextDecoder("utf-8").decode(buffer);
    }

    function decodeTextFromBuffer(buffer, typeHint = "text") {
      let utf8 = "";
      let utf8Valid = false;
      try {
        utf8 = new TextDecoder("utf-8", { fatal: true }).decode(buffer);
        utf8Valid = true;
      } catch {
        utf8 = new TextDecoder("utf-8").decode(buffer);
      }

      const utf8ReplacementCount = (utf8.match(/ /g) || []).length;
      const utf8MojibakeHits = (utf8.match(/[蜀螂蛹鬟蠕繧縺譌譯驛]/g) || []).length;

      try {
        const sjis = new TextDecoder("shift-jis").decode(buffer);
        const sjisReplacementCount = (sjis.match(/ /g) || []).length;
        const sjisMojibakeHits = (sjis.match(/[蜀螂蛹鬟蠕繧縺譌譯驛]/g) || []).length;

        if (typeHint === "json") {
          let utf8JsonOk = false;
          let sjisJsonOk = false;
          try { JSON.parse(utf8); utf8JsonOk = true; } catch {}
          try { JSON.parse(sjis); sjisJsonOk = true; } catch {}

          if (utf8JsonOk && !sjisJsonOk) return utf8;
          if (!utf8JsonOk && sjisJsonOk) return sjis;
          if (!utf8JsonOk && !sjisJsonOk) return utf8;
          const utf8Score = utf8ReplacementCount + utf8MojibakeHits * 3;
          const sjisScore = sjisReplacementCount + sjisMojibakeHits * 3;
          return sjisScore < utf8Score ? sjis : utf8;
        } else if (typeHint === "kml") {
          const utf8Score = utf8ReplacementCount + utf8MojibakeHits * 3;
          const sjisScore = sjisReplacementCount + sjisMojibakeHits * 3;
          const utf8HasKml = /<kml[\s>]/i.test(utf8);
          const sjisHasKml = /<kml[\s>]/i.test(sjis);
          if (!utf8HasKml && sjisHasKml) return sjis;
          if (utf8HasKml && !sjisHasKml) return utf8;
          return sjisScore < utf8Score ? sjis : utf8;
        } else {
          if (!utf8Valid) return sjis;
          const utf8Score = utf8ReplacementCount + utf8MojibakeHits * 3;
          const sjisScore = sjisReplacementCount + sjisMojibakeHits * 3;
          return sjisScore < utf8Score ? sjis : utf8;
        }

        if (!utf8Valid) return sjis;
        return utf8;
      } catch {
        return utf8;
      }
    }

    async function fetchJsonWithEncoding(url) {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
      const buffer = await res.arrayBuffer();
      const text = decodeTextFromBuffer(buffer, "json");
      return JSON.parse(text);
    }

    async function fetchTextWithEncoding(url, typeHint = "text") {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}: ${url}`);
      const buffer = await res.arrayBuffer();
      return decodeTextFromBuffer(buffer, typeHint);
    }

    function parseCoords(coordText) {
      return coordText
        .trim()
        .split(/\s+/)
        .map((v) => {
          const p = v.split(",");
          return { lng: Number(p[0]), lat: Number(p[1]) };
        })
        .filter((p) => !Number.isNaN(p.lng) && !Number.isNaN(p.lat));
    }

    function parseKmlText(xmlText) {
      // Repair common malformed name tags observed in exported KML.
      const fixedText = String(xmlText || "")
        .replace(/<name>([^<]*?)\/name>/g, "<name>$1</name>")
        .replace(/<name>([^<]*?)<\/name>/g, "<name>$1</name>");

      const zones = [];
      const used = new Set();
      const usedZoneIds = new Set();

      // Primary path: XML parser when markup is valid enough.
      const dom = new DOMParser().parseFromString(fixedText, "application/xml");
      const hasParserError = dom.getElementsByTagName("parsererror").length > 0;

      if (!hasParserError) {
        const placemarks = Array.from(dom.getElementsByTagName("Placemark"));
        placemarks.forEach((pm, idx) => {
          const rawName = pm.getElementsByTagName("name")[0]?.textContent || "";
          let id = normalizeZoneId(rawName);
          if (!id || /[蜀螂蛹鬟蠕]/.test(id)) {
            if (idx < mojibakeAreaFallback.length) id = mojibakeAreaFallback[idx];
          }
          id = canonicalZoneId(id);
          id = repairBrokenZoneId(id, usedZoneIds);
          if (!id) return;

          const coordsEl = pm.getElementsByTagName("coordinates")[0];
          if (!coordsEl) return;
          const points = parseCoords(coordsEl.textContent || "");
          if (points.length < 3) return;
          const key = `${id}::${points.length}::${points[0].lat},${points[0].lng}`;
          if (used.has(key)) return;
          used.add(key);
          usedZoneIds.add(id);
          zones.push({ id, points });
        });
      }

      // Fallback path: regex-based extraction for malformed KML.
      if (zones.length < 20) {
        const pmBlocks = fixedText.match(/<Placemark\b[\s\S]*?<\/Placemark>/g) || [];
        pmBlocks.forEach((block, idx) => {
          const nameMatch = block.match(/<name>([\s\S]*?)<\/name>/);
          let rawName = nameMatch ? nameMatch[1].replace(/<!\[CDATA\[|\]\]>/g, "").trim() : "";
          let id = normalizeZoneId(rawName);
          if (!id || /[蜀螂蛹鬟蠕]/.test(id)) {
            if (idx < mojibakeAreaFallback.length) id = mojibakeAreaFallback[idx];
          }
          id = canonicalZoneId(id);
          id = repairBrokenZoneId(id, usedZoneIds);
          if (!id) return;
          const coordMatch = block.match(/<coordinates>([\s\S]*?)<\/coordinates>/);
          if (!coordMatch) return;
          const points = parseCoords(coordMatch[1] || "");
          if (points.length < 3) return;
          const key = `${id}::${points.length}::${points[0].lat},${points[0].lng}`;
          if (used.has(key)) return;
          used.add(key);
          usedZoneIds.add(id);
          zones.push({ id, points });
        });
      }

      return zones;
    }

    function getDayOps(dateStr, zoneId) {
      const exact = dailyOps[dateStr]?.[zoneId];
      if (exact) return exact;
      const fp = zoneFingerprint(zoneId);
      const byFp = dailyOpsByDateFingerprint[dateStr]?.[fp];
      if (byFp) return byFp;
      return {
        command: false,
        commandCount: 0,
        continued: false,
        executed: false
      };
    }

    function diffDays(fromDate, toDate) {
      const ms = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate()).getTime() -
        new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate()).getTime();
      return Math.max(0, Math.floor(ms / 86400000));
    }

    function findLatestCommandDate(zoneId, targetDate) {
      let cur = new Date(DATE_MIN);
      let latest = null;
      while (cur <= targetDate) {
        const d = getDayOps(fmtDate(cur), zoneId);
        if (d.command) latest = new Date(cur);
        cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
      }
      return latest;
    }

    function computeClearDaysFromCommand(zoneId, commandDate) {
      let cur = new Date(commandDate.getFullYear(), commandDate.getMonth(), commandDate.getDate() + 1);
      let sawContinued = false;
      while (cur <= DATE_MAX) {
        const d = getDayOps(fmtDate(cur), zoneId);
        if (d.continued) sawContinued = true;
        if (!d.continued && (sawContinued || cur > commandDate)) {
          return diffDays(commandDate, cur);
        }
        cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
      }
      return null;
    }

    function buildZoneTiming(zoneId, targetDate) {
      const latestCommand = findLatestCommandDate(zoneId, targetDate);
      const seasonMaxClearDays = getSeasonMaxClearDays(zoneId);
      if (!latestCommand) return { continuedDays: 0, daysToExecution: null, daysToGreenClear: null, seasonMaxClearDays };
      const clearDays = computeClearDaysFromCommand(zoneId, latestCommand);
      return {
        continuedDays: diffDays(latestCommand, targetDate),
        daysToExecution: diffDays(latestCommand, targetDate),
        daysToGreenClear: clearDays,
        seasonMaxClearDays
      };
    }

    function getSeasonMaxClearDays(zoneId) {
      let max = null;
      let cur = new Date(DATE_MIN);
      while (cur <= DATE_MAX) {
        const d = getDayOps(fmtDate(cur), zoneId);
        if (d.command) {
          const span = computeClearDaysFromCommand(zoneId, cur);
          if (span !== null && (max === null || span > max)) max = span;
        }
        cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
      }
      return max;
    }

    function enableDetailTableColumnResize() {
      const headers = detailTableEl.querySelectorAll("thead th");
      headers.forEach((th) => {
        if (th.querySelector(".col-resizer")) return;
        const resizer = document.createElement("span");
        resizer.className = "col-resizer";
        th.appendChild(resizer);

        let startX = 0;
        let startWidth = 0;

        const onMouseMove = (ev) => {
          const next = Math.max(56, startWidth + (ev.clientX - startX));
          th.style.width = `${next}px`;
        };
        const onMouseUp = () => {
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup", onMouseUp);
          document.body.style.userSelect = "";
          document.body.style.cursor = "";
        };

        resizer.addEventListener("mousedown", (ev) => {
          ev.preventDefault();
          startX = ev.clientX;
          startWidth = th.getBoundingClientRect().width;
          document.body.style.userSelect = "none";
          document.body.style.cursor = "col-resize";
          document.addEventListener("mousemove", onMouseMove);
          document.addEventListener("mouseup", onMouseUp);
        });
      });
    }

    function enableDetailTableSort() {
      const keys = ["zone","eval","d","e","f","g","h","i","j","k","l","m","n","o","command","count","continued","executed","clear","seasonClear"];
      const headers = detailTableEl.querySelectorAll("thead th");
      headers.forEach((th, i) => {
        const key = keys[i];
        if (!key) return;
        th.classList.add("sortable");
        const baseText = th.textContent;
        const paint = () => {
          const mark = rowSort.key === key ? (rowSort.dir === 1 ? " ▲" : " ▼") : "";
          th.textContent = baseText + mark;
          const resizer = document.createElement("span");
          resizer.className = "col-resizer";
          th.appendChild(resizer);
        };
        paint();
        th.addEventListener("click", (ev) => {
          if (ev.target.classList.contains("col-resizer")) return;
          if (rowSort.key === key) rowSort.dir *= -1;
          else rowSort = { key, dir: 1 };
          headers.forEach((h, idx) => {
            const t = h.childNodes[0] && h.childNodes[0].nodeType === Node.TEXT_NODE ? h.childNodes[0].nodeValue : h.textContent;
            const plain = (t || h.textContent).replace(/\s[▲▼]$/, "");
            h.innerHTML = "";
            h.textContent = plain;
            const rz = document.createElement("span");
            rz.className = "col-resizer";
            h.appendChild(rz);
          });
          const mark = rowSort.dir === 1 ? " ▲" : " ▼";
          headers[i].childNodes[0].nodeValue = headers[i].childNodes[0].nodeValue + mark;
          refreshRows();
          enableDetailTableColumnResize();
        });
      });
    }

    function enableDetailTableDragScroll() {
      const wrap = detailTableEl.closest(".table-wrap.wide");
      if (!wrap) return;
      let isDown = false;
      let sx = 0;
      let sy = 0;
      let sl = 0;
      let st = 0;
      wrap.addEventListener("mousedown", (e) => {
        isDown = true;
        wrap.classList.add("dragging");
        sx = e.pageX;
        sy = e.pageY;
        sl = wrap.scrollLeft;
        st = wrap.scrollTop;
      });
      window.addEventListener("mouseup", () => {
        isDown = false;
        wrap.classList.remove("dragging");
      });
      wrap.addEventListener("mousemove", (e) => {
        if (!isDown) return;
        e.preventDefault();
        wrap.scrollLeft = sl - (e.pageX - sx);
        wrap.scrollTop = st - (e.pageY - sy);
      });
    }

    function compareZoneIds(a, b) {
      const na = normalizeZoneId(a);
      const nb = normalizeZoneId(b);
      const ma = na.match(/^([A-L])-(\d+)(?:-(\d+))?(?:-([A-Z]))?$/);
      const mb = nb.match(/^([A-L])-(\d+)(?:-(\d+))?(?:-([A-Z]))?$/);
      if (ma && mb) {
        const blockCmp = ma[1].localeCompare(mb[1], "ja");
        if (blockCmp !== 0) return blockCmp;
        const n1 = Number(ma[2]) - Number(mb[2]);
        if (n1 !== 0) return n1;
        const n2 = Number(ma[3] || 0) - Number(mb[3] || 0);
        if (n2 !== 0) return n2;
        return (ma[4] || "").localeCompare(mb[4] || "", "ja");
      }
      if (ma) return -1;
      if (mb) return 1;
      return na.localeCompare(nb, "ja");
    }

    function isSelectedZone(zoneId) {
      return selectedZoneIds.includes(zoneId);
    }

    function toggleZoneSelection(zoneId) {
      if (isSelectedZone(zoneId)) {
        selectedZoneIds = selectedZoneIds.filter((z) => z !== zoneId);
      } else {
        selectedZoneIds = [...selectedZoneIds, zoneId];
      }
      selectedZoneId = selectedZoneIds.length ? selectedZoneIds[selectedZoneIds.length - 1] : null;
    }

    function getRowSortValue(row, key) {
      if (key === "zone") return row.zoneId;
      if (key === "eval") return `${row.evalScore ?? ""}:${row.evalGrade ?? ""}`;
      if (key === "command") return row.d.command ? 1 : 0;
      if (key === "count") return Number(row.d.commandCount || 0);
      if (key === "continued") return row.d.continued ? row.timing.continuedDays : -1;
      if (key === "executed") return row.d.executed ? row.timing.daysToExecution : -1;
      if (key === "clear") return row.timing.daysToGreenClear ?? -1;
      if (key === "seasonClear") return row.timing.seasonMaxClearDays ?? -1;
      if (["d","e","f","g","h","i","j","k","l","m","n","o"].includes(key)) return Number(row.cost?.[key] || 0);
      return "";
    }

    function compareRows(a, b) {
      const va = getRowSortValue(a, rowSort.key);
      const vb = getRowSortValue(b, rowSort.key);
      if (rowSort.key === "zone") return compareZoneIds(a.zoneId, b.zoneId) * rowSort.dir;
      if (typeof va === "number" && typeof vb === "number") {
        if (va === vb) return compareZoneIds(a.zoneId, b.zoneId);
        return (va - vb) * rowSort.dir;
      }
      const c = String(va).localeCompare(String(vb), "ja", { numeric: true });
      return c === 0 ? compareZoneIds(a.zoneId, b.zoneId) : c * rowSort.dir;
    }

    function clearLayers() {
      zoneLayers.forEach((layer) => map.removeLayer(layer));
      zoneLayers = new Map();
    }

    function styleForZone(zoneId) {
      const d = getDayOps(fmtDate(currentDate), zoneId);
      let fillColor = "#5ca4e6";
      let fillOpacity = 0.22;
      if (d.command) {
        fillColor = "#e55757";
        fillOpacity = 0.35;
      } else if (d.continued) {
        fillColor = "#6dbc6f";
        fillOpacity = 0.35;
      }
      if (d.executed) fillOpacity = blinkOn ? 0.62 : 0.14;
      return {
        color: isSelectedZone(zoneId) ? "#ff7a00" : "#28435e",
        weight: isSelectedZone(zoneId) ? 4 : 2,
        fillColor,
        fillOpacity
      };
    }

    function applyZoneVisuals() {
      zoneLayers.forEach((layer, zoneId) => {
        layer.setStyle(styleForZone(zoneId));
      });
    }

    function renderSelectedZoneCards() {
      selectedZonesCardsEl.innerHTML = "";
      if (selectedZoneIds.length < 1) return;
      selectedZoneIds.forEach((zoneId, index) => {
        const d = getDayOps(fmtDate(currentDate), zoneId);
        const timing = buildZoneTiming(zoneId, currentDate);
        const contractorName = contractors[zoneId] || contractorByFingerprint[zoneFingerprint(zoneId)] || "未設定";
        const zoneEval = zoneEvalByZone[normalizeZoneId(zoneId)] || null;
        const cost = zoneCostByZone[normalizeZoneId(zoneId)] || null;
        const card = document.createElement("div");
        card.className = "zone-mini-card";
        card.draggable = true;
        card.dataset.zoneId = zoneId;
        card.innerHTML =
          `<button class="zone-mini-close" type="button" title="選択解除">×</button>` +
          `<strong style="font-size:28px;">${zoneId}</strong><br>` +
          `業者: ${contractorName}<br>` +
          `評点・評価: ${zoneEval ? `${zoneEval.score ?? "-"}点・${zoneEval.grade ?? "-"}` : "未設定"}<br>` +
          buildZoneStatusHtml(d, timing) +
          buildCostHtml(cost);
        const closeBtn = card.querySelector(".zone-mini-close");
        closeBtn.addEventListener("click", (ev) => {
          ev.stopPropagation();
          selectedZoneIds = selectedZoneIds.filter((z) => z !== zoneId);
          selectedZoneId = selectedZoneIds.length ? selectedZoneIds[selectedZoneIds.length - 1] : null;
          refreshAll();
        });
        card.addEventListener("dragstart", () => {
          card.classList.add("dragging");
          card.dataset.dragIndex = String(index);
        });
        card.addEventListener("dragend", () => {
          card.classList.remove("dragging");
        });
        card.addEventListener("dragover", (ev) => ev.preventDefault());
        card.addEventListener("drop", (ev) => {
          ev.preventDefault();
          const from = Number(document.querySelector(".zone-mini-card.dragging")?.dataset.dragIndex ?? -1);
          const to = Number(card.dataset.dragIndex ?? index);
          if (from < 0 || to < 0 || from === to) return;
          const next = selectedZoneIds.slice();
          const [moved] = next.splice(from, 1);
          next.splice(to, 0, moved);
          selectedZoneIds = next;
          selectedZoneId = selectedZoneIds.length ? selectedZoneIds[selectedZoneIds.length - 1] : null;
          refreshAll();
        });
        selectedZonesCardsEl.appendChild(card);
      });
    }

    function selectZone(zoneId, opts = {}) {
      const shouldToggle = opts.toggle !== false;
      if (shouldToggle) toggleZoneSelection(zoneId);
      else selectedZoneId = zoneId;
      if (!selectedZoneId) {
        detailEl.textContent = "";
        refreshRows();
        applyZoneVisuals();
        renderSelectedZoneCards();
        refreshSelectedZoneTags();
        return;
      }
      zoneId = selectedZoneId;
      syncPickerToZone(zoneId);
      detailEl.innerHTML = "";
      refreshRows();
      applyZoneVisuals();
      renderSelectedZoneCards();
      refreshSelectedZoneTags();
      const selectedRow = rowsEl.querySelector("tr.selected-row");
      if (selectedRow) {
        const tableScroller = rowsEl.closest(".table-wrap");
        if (tableScroller) {
          const targetTop = selectedRow.offsetTop - tableScroller.clientHeight / 2 + selectedRow.clientHeight / 2;
          const maxTop = Math.max(0, tableScroller.scrollHeight - tableScroller.clientHeight);
          tableScroller.scrollTo({
            top: Math.max(0, Math.min(maxTop, targetTop)),
            behavior: "smooth"
          });
        }
      }
    }

    function refreshRows() {
      const dateStr = fmtDate(currentDate);
      rowsEl.innerHTML = "";
      const headerRow = document.getElementById("detailHeaderRow");
      const zones = selectedZoneIds.slice();

      if (headerRow) {
        headerRow.innerHTML = "<th>項目</th>" + zones.map((z) => `<th>${z}</th>`).join("");
      }

      if (!zones.length) {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>工区を選択してください（複数選択可）</td>`;
        rowsEl.appendChild(tr);
        return;
      }

      const zoneData = zones.map((zoneId) => {
        const d = getDayOps(dateStr, zoneId);
        const timing = buildZoneTiming(zoneId, currentDate);
        const zoneEval = zoneEvalByZone[normalizeZoneId(zoneId)] || null;
        const cost = zoneCostByZone[normalizeZoneId(zoneId)] || null;
        return { zoneId, d, timing, zoneEval, cost };
      });

      const metrics = [
        { label: "評点・評価", get: (x) => x.zoneEval ? `${x.zoneEval.score ?? "-"}・${x.zoneEval.grade ?? "-"}` : "" },
        { label: "D 除排雪延長", get: (x) => x.cost ? fmtCost(x.cost.d) : "" },
        { label: "E ①当初契約額", get: (x) => x.cost ? fmtCost(x.cost.e) : "" },
        { label: "F ②工区延長", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.f) : "" },
        { label: "G ③1/25~2/5降雪量見合(147cm)", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.g) : "" },
        { label: "H ④2/6~2/21降雪量見合い(72cm)", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.h) : "" },
        { label: "I ⑤2/6~2/21時間単価見合 合計", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.i) : "" },
        { label: "J 市負担", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.j) : "" },
        { label: "K ⑤県負担", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.k) : "" },
        { label: "L ⑥2/22~3/14降雪量見合(17cm)", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.l) : "" },
        { label: "M ⑦応援除雪", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.m) : "" },
        { label: "N ⑧市民雪寄せ場排雪業務", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.n) : "" },
        { label: "O ⑧変更後委託料", costDetail: true, get: (x) => x.cost ? fmtCost(x.cost.o) : "" },
        { label: "指令", get: (x) => x.d.command ? "赤" : "" },
        { label: "回数", get: (x) => String(x.d.commandCount || "") },
        { label: "継続", get: (x) => x.d.continued ? `緑（${x.timing.continuedDays}日）` : "" },
        { label: "実施", get: (x) => x.d.executed ? `実施（${x.timing.daysToExecution}日）` : "" },
        { label: "解除日数", get: (x) => x.timing.daysToGreenClear !== null ? `${x.timing.daysToGreenClear}日` : "" },
        { label: "当期の解除に要した最大日数", get: (x) => x.timing.seasonMaxClearDays !== null ? `${x.timing.seasonMaxClearDays}日` : "" }
      ];

      metrics.forEach((m) => {
        if (m.costDetail && !showCostDetails) return;
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${m.label}</td>` + zoneData.map((z) => `<td>${m.get(z)}</td>`).join("");
        rowsEl.appendChild(tr);
      });
    }

    function refreshCalendar() {
      const dateStr = fmtDate(currentDate);
      const y = currentDate.getFullYear();
      const m = currentDate.getMonth();
      const first = new Date(y, m, 1);
      const last = new Date(y, m + 1, 0);
      const startWeek = first.getDay();

      calendarEl.innerHTML = "";
      ["日", "月", "火", "水", "木", "金", "土"].forEach((w) => {
        const el = document.createElement("div");
        el.className = "day-head";
        el.textContent = w;
        calendarEl.appendChild(el);
      });

      for (let i = 0; i < startWeek; i += 1) {
        const b = document.createElement("div");
        b.className = "day-cell";
        calendarEl.appendChild(b);
      }

      for (let d = 1; d <= last.getDate(); d += 1) {
        const dt = new Date(y, m, d);
        const key = fmtDate(dt);
        const hasEvent = !!dailyOps[key] && Object.keys(dailyOps[key]).length > 0;
        const c = document.createElement("div");
        c.className = `day-cell${key === dateStr ? " today" : ""}${hasEvent ? " has-event" : ""}`;
        c.textContent = String(d);
        c.addEventListener("click", () => {
          currentDate = dt;
          refreshAll();
        });
        calendarEl.appendChild(c);
      }
      if (calendarPickerEl) calendarPickerEl.value = dateStr;
    }

    function extractPrimaryLabel(zoneId) {
      const n = normalizeZoneId(zoneId);
      const m = n.match(/^([A-L])-(\d+)/);
      if (m) return m[1];
      return "その他";
    }

    function extractBranchLabel(zoneId) {
      const primary = extractPrimaryLabel(zoneId);
      if (primary === "その他") return String(zoneId);
      const n = normalizeZoneId(zoneId);
      const m = n.match(/^([A-L])-(\d+)(?:-(\d+))?/);
      if (m) return m[3] ? `${m[2]}-${m[3]}` : m[2];
      const d = n.match(/(\d+)/);
      return d ? d[1] : "1";
    }

    function buildZonePicker() {
      if (!zonePrimaryEl || !zoneBranchEl) return;
      const groups = {};
      visibleZoneIds.forEach((z) => {
        const p = extractPrimaryLabel(z);
        const b = extractBranchLabel(z);
        if (!groups[p]) groups[p] = new Set();
        groups[p].add(b);
      });
      const ordered = PRIMARY_ORDER.filter((p) => groups[p]);
      zonePrimaryEl.innerHTML = ordered.map((p) => `<option value="${p}">${p}</option>`).join("");
      const applyBranchSelection = () => {
        if (!zoneBranchEl.value) return;
        const zone = findZoneByPicker(zonePrimaryEl.value, zoneBranchEl.value);
        if (!zone) return;
        if (!selectedZoneIds.includes(zone)) selectedZoneIds.push(zone);
        selectedZoneId = zone;
        refreshAll();
      };
      const syncBranch = () => {
        const p = zonePrimaryEl.value;
        const arr = Array.from(groups[p] || []).sort((a, b) => String(a).localeCompare(String(b), "ja", { numeric: true }));
        zoneBranchEl.innerHTML = `<option value="">枝番を選択</option>` + arr.map((b) => `<option value="${b}">${b}</option>`).join("");
        zoneBranchEl.value = "";
      };
      zonePrimaryEl.onchange = syncBranch;
      syncBranch();
      zoneBranchEl.onchange = () => {
        applyBranchSelection();
      };
    }

    function findZoneByPicker(primary, branch) {
      const candidates = visibleZoneIds.filter((z) => extractPrimaryLabel(z) === primary);
      const byBranch = candidates.find((z) => extractBranchLabel(z) === branch);
      return byBranch || candidates[0] || null;
    }

    function syncPickerToZone(zoneId) {
      if (!zonePrimaryEl || !zoneBranchEl || !zoneId) return;
      const primary = extractPrimaryLabel(zoneId);
      const branch = extractBranchLabel(zoneId);
      if (zonePrimaryEl.value !== primary) {
        zonePrimaryEl.value = primary;
        if (typeof zonePrimaryEl.onchange === "function") zonePrimaryEl.onchange();
      }
      const hasBranch = Array.from(zoneBranchEl.options || []).some((opt) => opt.value === branch);
      if (hasBranch) zoneBranchEl.value = branch;
    }

    function refreshSelectedZoneTags() {
      if (!selectedZoneTagsEl) return;
      selectedZoneTagsEl.innerHTML = selectedZoneIds
        .map((z) => `<span class="selected-zone-tag">${z} <button type="button" data-zone-remove="${z}" style="border:0;background:transparent;cursor:pointer;font-size:14px;line-height:1;">×</button></span>`)
        .join("");
      selectedZoneTagsEl.querySelectorAll("button[data-zone-remove]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const zone = btn.getAttribute("data-zone-remove");
          selectedZoneIds = selectedZoneIds.filter((z) => z !== zone);
          selectedZoneId = selectedZoneIds.length ? selectedZoneIds[selectedZoneIds.length - 1] : null;
          refreshAll();
        });
      });
    }

    function buildDateRangeKeys() {
      const out = [];
      let cur = new Date(DATE_MIN);
      while (cur <= DATE_MAX) {
        out.push(fmtDate(cur));
        cur = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate() + 1);
      }
      return out;
    }

    function buildOrUpdateSnowChart() {
      const dates = buildDateRangeKeys();
      const snowfallSeries = dates.map((d) => Number(snowfallByDate[d] ?? 0));
      const depthSeries = dates.map((d) => Number(snowDepthByDate[d] ?? 0));
      const current = fmtDate(currentDate);
      const labels = dates.map((d) => d.slice(5));
      const barColors = dates.map((d) => d === current ? "rgba(21,94,174,0.82)" : "rgba(21,94,174,0.45)");
      const pointSizes = dates.map((d) => d === current ? (chartBlinkOn ? 8 : 3) : 2);

      if (!snowChart) {
        snowChart = new Chart(snowChartCanvas, {
          type: "bar",
          data: {
            labels,
            datasets: [
              {
                type: "bar",
                label: "降雪量(cm)",
                data: snowfallSeries,
                backgroundColor: barColors,
                borderColor: "rgba(21,94,174,1)",
                borderWidth: 1,
                yAxisID: "y"
              },
              {
                type: "line",
                label: "積雪量(cm)",
                data: depthSeries,
                borderColor: "rgba(220,88,61,1)",
                backgroundColor: "rgba(220,88,61,0.16)",
                tension: 0.25,
                pointRadius: pointSizes,
                pointHoverRadius: 6,
                pointBackgroundColor: "rgba(220,88,61,1)",
                yAxisID: "y1"
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: "index", intersect: false },
            plugins: {
              legend: {
                position: "top",
                labels: { font: { size: 14 } }
              },
              tooltip: { enabled: true }
            },
            scales: {
              x: {
                ticks: { maxRotation: 0, autoSkip: true, maxTicksLimit: 10, font: { size: 12 } }
              },
              y: {
                title: { display: true, text: "降雪量(cm)", font: { size: 13 } },
                beginAtZero: true,
                ticks: { font: { size: 12 } }
              },
              y1: {
                position: "right",
                title: { display: true, text: "積雪量(cm)", font: { size: 13 } },
                beginAtZero: true,
                grid: { drawOnChartArea: false },
                ticks: { font: { size: 12 } }
              }
            }
          }
        });
      } else {
        snowChart.data.labels = labels;
        snowChart.data.datasets[0].data = snowfallSeries;
        snowChart.data.datasets[0].backgroundColor = barColors;
        snowChart.data.datasets[1].data = depthSeries;
        snowChart.data.datasets[1].pointRadius = pointSizes;
        snowChart.update();
      }
    }

    function updateSnowChartHighlight() {
      if (!snowChart) return;
      const dates = buildDateRangeKeys();
      const current = fmtDate(currentDate);
      snowChart.data.datasets[0].backgroundColor = dates.map((d) =>
        d === current ? "rgba(21,94,174,0.82)" : "rgba(21,94,174,0.45)"
      );
      snowChart.data.datasets[1].pointRadius = dates.map((d) => (d === current ? (chartBlinkOn ? 8 : 3) : 2));
      snowChart.data.datasets[1].pointHoverRadius = dates.map((d) => (d === current ? 10 : 6));
      snowChart.update();
    }

    function refreshAll() {
      const dateStr = fmtDate(currentDate);
      dateLabelEl.textContent = `日付: ${dateStr}`;
      if (currentDateBadgeEl) currentDateBadgeEl.textContent = dateStr;
      commonSnowfallLabelEl.textContent = `降雪量: ${Number(snowfallByDate[dateStr] ?? 0)} cm`;
      commonSnowDepthLabelEl.textContent = `積雪量: ${Number(snowDepthByDate[dateStr] ?? 0)} cm`;

      refreshCalendar();
      refreshRows();
      if (selectedZoneId) selectZone(selectedZoneId, { toggle: false });
      else {
        detailEl.textContent = "";
        renderSelectedZoneCards();
      }

      refreshSelectedZoneTags();
      applyZoneVisuals();
      updateSnowChartHighlight();
    }

    function renderZones(zones) {
      clearLayers();
      visibleZoneIds = zones.map((z) => z.id);
      const latlngsAll = [];

      zones.forEach((z) => {
        const latlngs = z.points.map((p) => [p.lat, p.lng]);
        const layer = L.polygon(latlngs, styleForZone(z.id));
        layer.on("click", () => selectZone(z.id));
        layer.on("mouseover", () => {
          layer.bindTooltip(buildHoverInfoHtml(z.id), {
            sticky: true,
            direction: "top",
            opacity: 0.96
          });
          layer.openTooltip();
        });
        layer.on("mouseout", () => {
          layer.closeTooltip();
        });
        layer.addTo(map);
        zoneLayers.set(z.id, layer);
        latlngs.forEach((ll) => latlngsAll.push(ll));
      });

      if (latlngsAll.length > 0) {
        map.fitBounds(L.latLngBounds(latlngsAll), { padding: [8, 8] });
      }
      map.setView([40.8222, 140.7474], 12);
      if (visibleZoneIds.length > 0) {
        selectedZoneId = null;
        selectedZoneIds = [];
      }
      buildZonePicker();
      refreshAll();
      statusEl.textContent = `KML読込成功: ${zones.length}工区を表示中`;
    }

    async function loadFixedKml() {
      try {
        const kmlText = await fetchTextWithEncoding("./data/aomori_latest.kml", "kml");
        const zones = parseKmlText(kmlText);
        renderZones(zones);
      } catch (e) {
        statusEl.textContent = `KML読込失敗: ${e.message}`;
      }
    }

    async function loadDailyData() {
      try {
        const [rawOps, rawContractors, rawAlias, rawEval, rawZoneEval, rawZoneCost, metricRaw, snowfallRaw, depthRaw] = await Promise.all([
          fetchJsonWithEncoding("./data/daily_ops_from_xlsx.json"),
          fetchJsonWithEncoding("./data/contractors_from_xlsx.json"),
          fetchJsonWithEncoding("./data/zone_alias_auto.json"),
          fetchJsonWithEncoding("./data/contractor_eval.json").catch(() => ({})),
          fetchJsonWithEncoding("./data/zone_score_eval.json").catch(() => ({})),
          fetchJsonWithEncoding("./data/zone_cost_detail.json").catch(() => ({})),
          fetchJsonWithEncoding("./data/daily_snow_metrics_from_jma.json").catch(() => null),
          fetchJsonWithEncoding("./data/daily_snowfall_from_jma.json").catch(() => null),
          fetchJsonWithEncoding("./data/daily_snowdepth_from_jma.json").catch(() => null)
        ]);

        if (rawAlias) {
          const normalizedAlias = {};
          const normalizedAliasLoose = {};
          for (const [k, v] of Object.entries(rawAlias || {})) {
            const nk = normalizeZoneId(k);
            const nv = normalizeZoneId(v);
            normalizedAlias[nk] = nv;
            const lk = looseZoneKey(k);
            if (lk && !normalizedAliasLoose[lk]) normalizedAliasLoose[lk] = nv;
          }
          zoneAlias = normalizedAlias;
          zoneAliasLoose = normalizedAliasLoose;
        }
        dailyOps = canonicalizeDailyOps(normalizeDailyOps(rawOps || {}));
        contractors = canonicalizeContractors(rawContractors || {});
        contractorEval = canonicalizeContractorEval(rawEval || {});
        zoneEvalByZone = canonicalizeZoneEval(rawZoneEval || {});
        zoneCostByZone = canonicalizeZoneCost(rawZoneCost || {});
        buildKnownZoneIds();
        buildDailyOpsFingerprintIndex();
        buildContractorFingerprintIndex();

        if (metricRaw) {
          const metric = metricRaw;
          for (const [d, v] of Object.entries(metric)) {
            snowfallByDate[d] = Number(v.snowfallCm ?? 0);
            snowDepthByDate[d] = Number(v.snowDepthCm ?? 0);
          }
          statusEl.textContent = "気象庁データ（降雪量/積雪量）を読込済み";
        } else {
          if (snowfallRaw) snowfallByDate = snowfallRaw;
          if (depthRaw) snowDepthByDate = depthRaw;
          statusEl.textContent = "気象庁データ（分割JSON）を読込済み";
        }
      } catch {
        statusEl.textContent = "気象データ読込に失敗（表示は0として継続）";
      }

      buildOrUpdateSnowChart();
      refreshAll();
    }

    function initMap() {
      map = L.map("map", { zoomControl: true }).setView([40.8222, 140.7474], 12);
      L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
        maxZoom: 19,
        attribution: "&copy; OpenStreetMap contributors"
      }).addTo(map);
    }

    document.getElementById("prevDayBtn").addEventListener("click", () => {
      currentDate = clampDate(new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - 1));
      refreshAll();
    });
    document.getElementById("nextDayBtn").addEventListener("click", () => {
      currentDate = clampDate(new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() + 1));
      refreshAll();
    });
    document.getElementById("prevMonthBtn").addEventListener("click", () => {
      currentDate = clampDate(new Date(currentDate.getFullYear(), currentDate.getMonth() - 1, currentDate.getDate()));
      refreshAll();
    });
    document.getElementById("nextMonthBtn").addEventListener("click", () => {
      currentDate = clampDate(new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, currentDate.getDate()));
      refreshAll();
    });
    if (calendarPickerEl) {
      calendarPickerEl.min = "2025-12-01";
      calendarPickerEl.max = "2026-03-31";
      calendarPickerEl.addEventListener("change", () => {
        if (!calendarPickerEl.value) return;
        currentDate = clampDate(new Date(calendarPickerEl.value));
        refreshAll();
      });
    }
    toggleCalendarBtnEl?.addEventListener("click", () => calendarModalEl.classList.add("open"));
    toggleSnowBtnEl?.addEventListener("click", () => snowPanelEl.classList.toggle("open"));
    toggleNoteBtnEl?.addEventListener("click", () => noteModalEl.classList.add("open"));
    closeCalendarBtnEl?.addEventListener("click", () => calendarModalEl.classList.remove("open"));
    closeNoteBtnEl?.addEventListener("click", () => noteModalEl.classList.remove("open"));
    calendarModalEl?.addEventListener("click", (ev) => {
      if (ev.target === calendarModalEl) calendarModalEl.classList.remove("open");
    });
    noteModalEl?.addEventListener("click", (ev) => {
      if (ev.target === noteModalEl) noteModalEl.classList.remove("open");
    });
    toggleCostColsBtnEl.addEventListener("click", () => {
      showCostDetails = !showCostDetails;
      toggleCostColsBtnEl.textContent = showCostDetails ? "－ F〜Oを隠す" : "＋ F〜Oを表示";
      refreshRows();
    });

    setInterval(() => {
      blinkOn = !blinkOn;
      chartBlinkOn = !chartBlinkOn;
      applyZoneVisuals();
      updateSnowChartHighlight();
    }, 550);

    (function initConsent() {
      const overlay = document.getElementById("consentOverlay");
      const agreeBtn = document.getElementById("agreeBtn");
      const ua = navigator.userAgent || "";
      const isBrave = /Brave/i.test(ua);
      let started = false;

      async function startApp() {
        if (started) return;
        started = true;
        initMap();
        await loadDailyData();
        await loadFixedKml();
        if (map) {
          setTimeout(() => map.invalidateSize(true), 80);
          setTimeout(() => map.invalidateSize(true), 300);
        }
      }

      // Show consent every time page is opened. App starts only after agreement.
      document.body.classList.add("preconsent");
      overlay.classList.add("open");
      overlay.style.position = "fixed";
      overlay.style.inset = "0";
      overlay.style.pointerEvents = "auto";
      overlay.style.touchAction = "auto";
      overlay.style.zIndex = "2147483647";

      function acceptConsent() {
        document.body.classList.remove("preconsent");
        overlay.classList.remove("open");
        overlay.style.display = "none";
        const emergency = document.getElementById("consentOverlayEmergency");
        if (emergency) emergency.remove();
        startApp();
      }

      // Brave: skip overlay rendering path and show emergency consent card immediately.
      if (isBrave) {
        document.body.classList.add("preconsent");
        overlay.classList.remove("open");
        overlay.style.display = "none";
        const emergency = document.createElement("div");
        emergency.id = "consentOverlayEmergency";
        emergency.style.cssText = "position:fixed;inset:0;z-index:2147483647;background:rgba(12,26,44,.72);display:flex;align-items:center;justify-content:center;padding:16px;";
        emergency.innerHTML = "<div style='width:min(92vw,680px);background:#fff;border:1px solid #d4deea;border-radius:14px;padding:16px;'><h2 style='margin:0 0 8px;'>ご利用にあたっての了承事項</h2><p style='margin:0 0 12px;'>内容をご確認のうえ、了承しましたを押してください。</p><div style='display:flex;justify-content:flex-end;'><button id='agreeBtnEmergency' type='button' style='min-height:38px;padding:6px 12px;border:1px solid #b8cbe0;border-radius:6px;background:#fff;'>了承しました</button></div></div>";
        document.body.appendChild(emergency);
        const eb = document.getElementById("agreeBtnEmergency");
        if (eb) {
          eb.addEventListener("click", () => acceptConsent());
          eb.addEventListener("touchend", (ev) => {
            ev.preventDefault();
            acceptConsent();
          }, { passive: false });
        }
        return;
      }

      agreeBtn.addEventListener("click", () => { acceptConsent(); });
      agreeBtn.addEventListener("touchend", (ev) => {
        ev.preventDefault();
        acceptConsent();
      }, { passive: false });

      // Fallback for browsers where overlay rendering/touch hit testing is unstable.
      setTimeout(() => {
        if (!document.body.classList.contains("preconsent")) return;
        const rect = overlay.getBoundingClientRect();
        const hidden = rect.width < 120 || rect.height < 120 || getComputedStyle(overlay).display === "none";
        if (!hidden) return;
        overlay.innerHTML = `
          <div class="consent-card" style="max-width:92vw;">
            <h2>ご利用にあたっての了承事項</h2>
            <p>表示不具合時のフォールバック画面です。内容に同意の上で続行してください。</p>
            <div class="agree-row"><button id="agreeBtnFallback" type="button">了承しました</button></div>
          </div>
        `;
        const fb = document.getElementById("agreeBtnFallback");
        if (fb) {
          fb.addEventListener("click", () => acceptConsent());
          fb.addEventListener("touchend", (ev) => {
            ev.preventDefault();
            acceptConsent();
          }, { passive: false });
        }
      }, 450);

      // Brave/iOS fallback: inject a plain fixed consent card if overlay rendering is still unavailable.
      setTimeout(() => {
        if (!document.body.classList.contains("preconsent")) return;
        const rect = overlay.getBoundingClientRect();
        const hidden = rect.width < 120 || rect.height < 120 || getComputedStyle(overlay).display === "none";
        if (!hidden) return;
        const emergency = document.createElement("div");
        emergency.id = "consentOverlayEmergency";
        emergency.style.cssText = "position:fixed;inset:0;z-index:2147483647;background:rgba(12,26,44,.72);display:flex;align-items:center;justify-content:center;padding:16px;";
        emergency.innerHTML = "<div style='width:min(92vw,680px);background:#fff;border:1px solid #d4deea;border-radius:14px;padding:16px;'><h2 style='margin:0 0 8px;'>ご利用にあたっての了承事項</h2><p style='margin:0 0 12px;'>内容をご確認のうえ、了承しましたを押してください。</p><div style='display:flex;justify-content:flex-end;'><button id='agreeBtnEmergency' type='button' style='min-height:38px;padding:6px 12px;border:1px solid #b8cbe0;border-radius:6px;background:#fff;'>了承しました</button></div></div>";
        document.body.appendChild(emergency);
        const eb = document.getElementById("agreeBtnEmergency");
        if (eb) {
          eb.addEventListener("click", () => acceptConsent());
          eb.addEventListener("touchend", (ev) => {
            ev.preventDefault();
            acceptConsent();
          }, { passive: false });
        }
      }, 900);
    })();
  