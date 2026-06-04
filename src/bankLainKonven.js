// Penempatan pada Bank Lain BPR Konvensional (LBBPRK-0500). Port dari penempatan_konven.py.
// Nominal = kolom 26 (verified = anchor neraca).
import * as S from "./sandiKonven.js";
import { detectBank, discoverBranches } from "./bank.js";
import * as H from "./helpers.js";

const COLS = {
  "Sandi Bank": 5, "Lokasi Bank": 6, "Jenis": 15, "Hubungan dengan Bank": 16,
  "Tanggal Mulai": 17, "Tanggal Jatuh Tempo": 18, "Kualitas": 21, "Nominal": 26,
};
const OUTPUT = ["Cabang", "Sandi Bank", "Lokasi Bank", "Jenis", "Hubungan dengan Bank",
  "Tanggal Mulai", "Tanggal Jatuh Tempo", "Kualitas", "Nominal"];

function read(aoa, cabang) {
  const rows = [];
  for (let r = 16; r < aoa.length; r++) {
    const sb = String(H.cell(aoa, r, 5)).trim();
    if (!sb || sb.toUpperCase() === "JUMLAH") continue;
    if (typeof H.cell(aoa, r, 26) !== "number") continue;  // nominal harus angka
    const row = { "Cabang": cabang };
    for (const [name, idx] of Object.entries(COLS)) row[name] = H.cell(aoa, r, idx);
    row["Hubungan dengan Bank"] = S.translate(row["Hubungan dengan Bank"], S.HUBUNGAN_BANK_LABEL);
    row["Kualitas"] = S.translate(row["Kualitas"], S.KUALITAS_LABEL);
    rows.push(row);
  }
  return rows;
}

function sheet(XLSX, data) {
  const aoa = [OUTPUT, ...data.map(r => OUTPUT.map(c => { const v = r[c]; return v === undefined ? "" : v; }))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const fill = { fill: { fgColor: { rgb: "1F4E79" } }, font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 } };
  for (let c = 0; c < OUTPUT.length; c++) { const a = XLSX.utils.encode_cell({ r: 0, c }); if (ws[a]) ws[a].s = fill; }
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}

export function processPenempatanKonven(files, period, XLSX) {
  const wb = XLSX.utils.book_new();
  const perBranch = {};
  let all = [];
  for (const code of discoverBranches(files, "LBBPRK-0500-")) {
    const name = Object.keys(files).find(n => { const b = n.split("/").pop(); return b.startsWith("LBBPRK-0500-") && (b.endsWith(`-${code}.xls`) || b.endsWith(`-${code}_part1.xls`)); });
    if (!name) continue;
    const wx = XLSX.read(files[name], { type: "array", cellDates: true });
    const aoa = XLSX.utils.sheet_to_json(wx.Sheets[wx.SheetNames[0]], { header: 1, raw: true, defval: null });
    const rows = read(aoa, code);
    perBranch[code] = rows; all = all.concat(rows);
  }
  const total = all.reduce((s, r) => s + (parseFloat(r["Nominal"]) || 0), 0);
  const ring = [["PENEMPATAN PADA BANK LAIN (LBBPRK)", period.periodeLabel], [],
    ["Total rekening", all.length], ["Total nominal", Math.round(total)]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ring), "RINGKASAN");
  XLSX.utils.book_append_sheet(wb, sheet(XLSX, all), "SEMUA CABANG");
  for (const code of Object.keys(perBranch).sort()) XLSX.utils.book_append_sheet(wb, sheet(XLSX, perBranch[code]), `Cabang ${code}`);

  const tag = detectBank(files, XLSX).tag;
  return {
    filename: `GABUNGAN_PENEMPATAN_BANK_LAIN_${tag}_${period.periodeLabel}.xlsx`,
    data: XLSX.write(wb, { type: "array", bookType: "xlsx" }),
    summary: { jumlah_rekening: all.length, total_nominal: Math.round(total * 100) / 100 },
  };
}
