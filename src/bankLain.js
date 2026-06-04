// Processor ABP (Form 24.00) + Penempatan (Form 05.00) - struktur mirip.
// Port dari core/abp.py & core/penempatan.py.
import * as abpData from "./abpData.js";
import * as penData from "./penempatanData.js";
import { translateSandi } from "./sandi.js";
import * as H from "./helpers.js";
import { detectBank, discoverBranches } from "./bank.js";

function resolveSandiMap(data) {
  const map = {};
  for (const [col, dictName] of Object.entries(data.SANDI_MAP_NAMES)) {
    map[col] = data[dictName];
  }
  return map;
}

function readBranch(aoa, branchCode, data, requireNamaCol) {
  const sandiMap = resolveSandiMap(data);
  const rows = [];
  for (let r = 17; r < aoa.length; r++) {
    const idVal = String(H.cell(aoa, r, 3)).trim();
    if (!idVal || idVal.toUpperCase() === "JUMLAH") continue;
    if (requireNamaCol !== null) {
      const nama = String(H.cell(aoa, r, requireNamaCol)).trim();
      if (!nama) continue;
    }
    const row = { "Cabang": branchCode, "Baris Asli": r + 1 };
    for (const [colName, colIdx] of Object.entries(data.COL_MAP)) {
      let v = H.cell(aoa, r, colIdx);
      if (typeof v === "number" && Number.isInteger(v)) v = v; // keep
      row[colName] = v;
    }
    for (const [colName, dict] of Object.entries(sandiMap)) {
      const v = row[colName];
      if (v !== null && v !== "" && v !== 0) row[colName] = translateSandi(v, dict);
    }
    rows.push(row);
  }
  return rows;
}

function headerStyledSheet(XLSX, cols, rows, color) {
  const aoa = [cols, ...rows.map(r => cols.map(c => { const v = r[c]; return v === undefined ? "" : v; }))];
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  const fill = { fill: { fgColor: { rgb: color } }, font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 } };
  for (let c = 0; c < cols.length; c++) {
    const a = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[a]) ws[a].s = fill;
  }
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}

function makeProcessor({ prefix, data, requireNamaCol, color, sheetRingkasan, title, namePrefix }) {
  return function (files, period, XLSX) {
    const wb = XLSX.utils.book_new();
    const allData = [];
    const perBranch = {};
    for (const code of discoverBranches(files, prefix)) {  // dinamis
      const name = Object.keys(files).find(n => {
        const b = n.split("/").pop();
        return b.startsWith(prefix) && (b.endsWith(`-${code}.xls`) || b.endsWith(`-${code}_part1.xls`));
      });
      if (!name) continue;
      const wbx = XLSX.read(files[name], { type: "array", cellDates: true });
      const aoa = XLSX.utils.sheet_to_json(wbx.Sheets[wbx.SheetNames[0]], { header: 1, raw: true, defval: null });
      const rows = readBranch(aoa, code, data, requireNamaCol);
      perBranch[code] = rows;
      allData.push(...rows);
    }

    // RINGKASAN sederhana
    const totalNom = allData.reduce((s, r) => s + (parseFloat(r["Jumlah"]) || 0), 0);
    const ring = [[`${title} - BPRS SURIYAH (${period.periodeLabel})`],
      ["Total rekening", allData.length], ["Total nominal", Math.round(totalNom)]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ring), sheetRingkasan);

    XLSX.utils.book_append_sheet(wb, headerStyledSheet(XLSX, data.OUTPUT_COLS, allData, color), "SEMUA CABANG");
    for (const code of Object.keys(perBranch).sort()) {
      XLSX.utils.book_append_sheet(wb, headerStyledSheet(XLSX, data.OUTPUT_COLS, perBranch[code], color), `Cabang ${code}`);
    }

    const totalBlk = allData.reduce((s, r) => s + (parseFloat(r["Nominal Diblokir"]) || 0), 0);
    const summary = { jumlah_rekening: allData.length, total_nominal: Math.round(totalNom * 100) / 100 };
    if (data.OUTPUT_COLS.includes("Nominal Diblokir")) summary.total_diblokir = Math.round(totalBlk * 100) / 100;

    const tag = detectBank(files, XLSX).tag;
    const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    return { filename: `${namePrefix}_${tag}_${period.periodeLabel}.xlsx`, data: out, summary };
  };
}

export const processAbp = makeProcessor({
  prefix: "LBBPRS-KC2400-", data: abpData, requireNamaCol: null, color: "C0392B",
  sheetRingkasan: "RINGKASAN ABP", title: "RINGKASAN LIABILITAS KEPADA BANK LAIN",
  namePrefix: "GABUNGAN_ABP_LIABILITAS_BANK_LAIN",
});

export const processPenempatan = makeProcessor({
  prefix: "LBBPRS-KC0500-", data: penData, requireNamaCol: 5, color: "1F4E79",
  sheetRingkasan: "RINGKASAN", title: "RINGKASAN PENEMPATAN PADA BANK LAIN",
  namePrefix: "GABUNGAN_PENEMPATAN_BANK_LAIN",
});
