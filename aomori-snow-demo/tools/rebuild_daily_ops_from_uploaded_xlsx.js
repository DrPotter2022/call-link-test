const fs = require("fs");
const path = require("path");
const XLSX = require("../node_modules/xlsx");

const ROOT = path.resolve(__dirname, "..");
const INPUT_XLSX = "C:/Users/with_/Downloads/aomori_pages2-11_bold_colors_corrected.xlsx";
const OUT_JSON = path.join(ROOT, "data", "daily_ops_from_xlsx.json");

const SEASON_START = "2025-12-01";
const SEASON_END = "2026-03-31";

const RED_RGB = new Set(["E6B8B7"]);
const GREEN_RGB = new Set(["D8E4BC", "D7E4BD"]);

function dateRange(start, end) {
  const out = [];
  const cur = new Date(`${start}T00:00:00+09:00`);
  const to = new Date(`${end}T00:00:00+09:00`);
  while (cur <= to) {
    const y = cur.getFullYear();
    const m = String(cur.getMonth() + 1).padStart(2, "0");
    const d = String(cur.getDate()).padStart(2, "0");
    out.push(`${y}-${m}-${d}`);
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

function toHalfWidth(str) {
  const s = String(str || "");
  let out = "";
  for (const ch of s) {
    const code = ch.charCodeAt(0);
    if (code >= 0xff01 && code <= 0xff5e) {
      out += String.fromCharCode(code - 0xfee0);
    } else if (code === 0x3000) {
      out += " ";
    } else if (code === 0xff0d || code === 0x2015 || code === 0x30fc) {
      out += "-";
    } else {
      out += ch;
    }
  }
  return out;
}

function normalizeZoneId(s) {
  return toHalfWidth(s).toUpperCase().replace(/\s+/g, "");
}

function parseDateCell(v) {
  const m = String(v || "").match(/^(\d{1,2})\/(\d{1,2})$/);
  if (!m) return null;
  const mm = Number(m[1]);
  const dd = Number(m[2]);
  const yyyy = mm >= 12 ? 2025 : 2026;
  return `${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
}

function main() {
  if (!fs.existsSync(INPUT_XLSX)) {
    throw new Error(`Input xlsx not found: ${INPUT_XLSX}`);
  }

  const previous = JSON.parse(fs.readFileSync(OUT_JSON, "utf8"));
  const previousExecuted = {};
  for (const [date, zoneMap] of Object.entries(previous || {})) {
    if (!zoneMap || typeof zoneMap !== "object") continue;
    for (const [zoneKey, d] of Object.entries(zoneMap)) {
      if (!d || typeof d !== "object") continue;
      const zoneId = normalizeZoneId(d.zoneId || zoneKey);
      previousExecuted[`${date}__${zoneId}`] = !!d.executed;
    }
  }

  const wb = XLSX.readFile(INPUT_XLSX, { cellStyles: true, cellDates: false, raw: true });
  const out = {};
  for (const d of dateRange(SEASON_START, SEASON_END)) out[d] = {};

  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    const dateCols = {};
    for (let c = 3; c < 80; c++) {
      const addr = XLSX.utils.encode_cell({ r: 3, c });
      const dateKey = parseDateCell(ws[addr] && ws[addr].v);
      if (dateKey) dateCols[c] = dateKey;
    }

    for (let r = 6; r < 1000; r++) {
      const zoneAddr = XLSX.utils.encode_cell({ r, c: 0 });
      const contractorAddr = XLSX.utils.encode_cell({ r, c: 1 });
      const cntAddr = XLSX.utils.encode_cell({ r, c: 2 });
      const zoneRaw = ws[zoneAddr] && ws[zoneAddr].v;
      if (!zoneRaw) continue;

      const zoneId = normalizeZoneId(zoneRaw);
      // Keep both block IDs (A-1, B-2-1, ...) and Japanese district names (..工区)
      const isBlockId = /^[A-L]-/i.test(zoneId);
      const isDistrictName = /工区/u.test(zoneId);
      if (!isBlockId && !isDistrictName) continue;
      const contractor = (ws[contractorAddr] && ws[contractorAddr].v) || "";
      const commandCount = Number((ws[cntAddr] && ws[cntAddr].v) || 0);

      for (const [cStr, dateKey] of Object.entries(dateCols)) {
        const c = Number(cStr);
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell || cell.v === undefined || cell.v === "") continue;

        const rgb = (((cell.s || {}).fgColor || {}).rgb || "").toUpperCase();
        const command = RED_RGB.has(rgb);
        const continued = GREEN_RGB.has(rgb);
        if (!command && !continued) continue;

        const executed = previousExecuted[`${dateKey}__${zoneId}`] || false;
        out[dateKey][zoneId] = {
          zoneId,
          contractor,
          command,
          commandCount,
          continued,
          executed,
        };
      }
    }
  }

  fs.writeFileSync(OUT_JSON, `${JSON.stringify(out, null, 2)}\n`, "utf8");

  const zone = "G-6-1";
  const dates = Object.keys(out).sort();
  const diffDays = (a, b) =>
    Math.round((new Date(`${b}T00:00:00+09:00`) - new Date(`${a}T00:00:00+09:00`)) / 86400000);
  const clearFrom = (commandDate) => {
    let sawContinued = false;
    for (const cur of dates) {
      if (cur < commandDate) continue;
      const d = out[cur][zone];
      if (!d) continue;
      if (d.continued) sawContinued = true;
      if (!d.continued && (sawContinued || cur > commandDate)) return diffDays(commandDate, cur);
    }
    return null;
  };
  let max = 0;
  for (const cur of dates) {
    const d = out[cur][zone];
    if (d && d.command) {
      const span = clearFrom(cur);
      if (span != null && span > max) max = span;
    }
  }

  console.log(
    JSON.stringify(
      {
        output: OUT_JSON,
        sheets: wb.SheetNames.length,
        g_6_1_max_clear_days: max,
      },
      null,
      2
    )
  );
}

main();
