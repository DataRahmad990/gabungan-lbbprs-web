// Gabungan generik form konven (LBBPRK-XXXX). Port dari core/form_konven.py.
// Label kolom dari merged-cell header (ws['!merges']). Nama via 0016.
import * as S from "./sandiKonven.js";
import { detectBank, discoverBranches } from "./bank.js";
import { buildNameMap, lookupName } from "./pihakLawan.js";
import * as H from "./helpers.js";

function readSheet(bytes, XLSX) {
  const wb = XLSX.read(bytes, { type: "array", raw: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  const merges = (ws["!merges"] || []).map(m => ({ r: m.s.r, c0: m.s.c, c1: m.e.c }));
  const ncols = ws["!ref"] ? XLSX.utils.decode_range(ws["!ref"]).e.c + 1 : 0;
  return { aoa, merges, ncols };
}

function isDigits(v, max) {
  let s = String(v == null ? "" : v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return new RegExp(`^\\d{1,${max}}$`).test(s);
}

function findDataStart(aoa) {
  for (let r = 14; r < Math.min(60, aoa.length); r++) {
    if (isDigits(H.cell(aoa, r, 4), 3)) return r;
  }
  return 18;
}

function txt(v) { return String(v == null ? "" : v).trim(); }
function isNumLabel(s) { return s !== "" && /^[\d.]+$/.test(s); }

function headerLabels(aoa, merges, ncols, ds) {
  let bestHr = ds - 1, bestCnt = -1;
  for (let r = Math.max(0, ds - 4); r < ds; r++) {
    let cnt = 0;
    for (let c = 3; c < ncols; c++) { const v = txt(H.cell(aoa, r, c)); if (v && !isNumLabel(v)) cnt++; }
    if (cnt >= bestCnt) { bestCnt = cnt; bestHr = r; }
  }
  const hr = bestHr;
  const parent = {};
  for (const m of merges) {
    if (m.r >= hr && m.r < ds && (m.c1 - m.c0 + 1) < ncols * 0.6) {
      const v = txt(H.cell(aoa, m.r, m.c0));
      if (v) for (let c = m.c0; c <= m.c1; c++) if (!(c in parent)) parent[c] = v;
    }
  }
  const labels = {};
  for (let c = 3; c < ncols; c++) {
    let base = txt(H.cell(aoa, hr, c)) || parent[c] || "";
    if (isNumLabel(base)) base = parent[c] || "";
    const sub = hr + 1 < ds ? txt(H.cell(aoa, hr + 1, c)) : "";
    const label = (sub && sub !== base) ? `${base} ${sub}`.trim() : base;
    if (label) labels[c] = label;
  }
  return labels;
}

function sheet(XLSX, cols, data) {
  const aoa = [cols, ...data.map(r => cols.map(c => { const v = r[c]; return v === undefined ? "" : v; }))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const fill = { fill: { fgColor: { rgb: "1F4E79" } }, font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 } };
  for (let c = 0; c < cols.length; c++) { const a = XLSX.utils.encode_cell({ r: 0, c }); if (ws[a]) ws[a].s = fill; }
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}

export function makeFormProcessor(formCode, title, namePrefix) {
  return function (files, period, XLSX) {
    const prefix = `LBBPRK-${formCode}-`;
    const branches = discoverBranches(files, prefix);
    if (!branches.length) return null;  // form ga ada -> skip

    const nameMap = buildNameMap(files, XLSX);
    const rawBranch = {};
    const labelsByCol = {};
    const nonempty = new Set();
    for (const code of branches) {
      const name = Object.keys(files).find(n => { const b = n.split("/").pop(); return b.startsWith(prefix) && (b.endsWith(`-${code}.xls`) || b.endsWith(`-${code}_part1.xls`)); });
      if (!name) continue;
      const { aoa, merges, ncols } = readSheet(files[name], XLSX);
      const ds = findDataStart(aoa);
      const labels = headerLabels(aoa, merges, ncols, ds);
      const used = Object.keys(labels).map(Number).sort((a, b) => a - b);
      for (const c of used) if (!(c in labelsByCol)) labelsByCol[c] = labels[c];
      const rows = [];
      for (let r = ds; r < aoa.length; r++) {
        const key = used.length ? txt(H.cell(aoa, r, used[0])) : "";
        const cells = {};
        let any = false;
        for (const c of used) { const v = H.cell(aoa, r, c); cells[c] = v; if (txt(v)) { any = true; nonempty.add(c); } }
        if (key.toUpperCase() === "JUMLAH" || !any) continue;
        rows.push(cells);
      }
      rawBranch[code] = rows;
    }

    const kept = Object.keys(labelsByCol).map(Number).sort((a, b) => a - b).filter(c => nonempty.has(c));
    const seen = {}, colLabel = {};
    for (const c of kept) {
      let lab = labelsByCol[c];
      if (seen[lab]) { seen[lab]++; lab = `${labelsByCol[c]} #${seen[labelsByCol[c]]}`; } else seen[lab] = 1;
      colLabel[c] = lab;
    }
    const idCol = kept.find(c => labelsByCol[c].includes("ID Pihak Lawan"));
    const hasName = !!(Object.keys(nameMap).length && idCol !== undefined);
    const colsOrder = ["Cabang", ...kept.map(c => colLabel[c]), ...(hasName ? ["Nama"] : [])];

    function toRow(code, cells) {
      const row = { "Cabang": code };
      for (const c of kept) {
        let v = cells[c] === undefined ? "" : cells[c];
        const base = labelsByCol[c];
        if (S.TRANSLATE_MAP[base] && v !== null && v !== "" && v !== 0) v = S.translate(v, S.TRANSLATE_MAP[base]);
        row[colLabel[c]] = v;
      }
      if (hasName) row["Nama"] = lookupName(nameMap, cells[idCol]);
      return row;
    }
    const perBranch = {};
    let all = [];
    for (const code of Object.keys(rawBranch).sort()) { perBranch[code] = rawBranch[code].map(c => toRow(code, c)); all = all.concat(perBranch[code]); }

    const wb = XLSX.utils.book_new();
    const tag = detectBank(files, XLSX).tag;
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[`${title} (LBBPRK-${formCode}) - ${tag}`, period.periodeLabel], [], ["Total baris", all.length]]), "RINGKASAN");
    XLSX.utils.book_append_sheet(wb, sheet(XLSX, colsOrder, all), "SEMUA CABANG");
    for (const code of Object.keys(perBranch).sort()) XLSX.utils.book_append_sheet(wb, sheet(XLSX, colsOrder, perBranch[code]), `Cabang ${code}`);

    const summary = { jumlah_baris: all.length };
    for (const col of colsOrder) {
      if (["Baki Debet", "Nominal", "Jumlah"].some(k => col.includes(k))) {
        let tot = 0; for (const r of all) { const v = parseFloat(r[col]); if (!isNaN(v)) tot += v; }
        if (tot) summary[`total[${col}]`] = Math.round(tot * 100) / 100;
      }
    }
    return { filename: `${namePrefix}_${tag}_${period.periodeLabel}.xlsx`, data: XLSX.write(wb, { type: "array", bookType: "xlsx" }), summary };
  };
}
