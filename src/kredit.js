// Processor Kredit BPR Konvensional (LBBPRK-0600). Port dari core/kredit.py.
import * as S from "./sandiKonven.js";
import { detectBank, discoverBranches } from "./bank.js";
import { buildNameMap, lookupName } from "./pihakLawan.js";
import * as H from "./helpers.js";

const KREDIT_COLS = {
  "ID Pihak Lawan": 5, "No. Identitas": 6, "No. Rekening": 16, "Jenis": 17,
  "Status Restrukturisasi": 18, "Jenis Penggunaan": 19, "Hubungan dengan Bank": 20,
  "Sumber Dana Pelunasan": 23, "Periode Pembayaran": 24, "Tanggal Mulai": 29, "Tanggal Jatuh Tempo": 31, "Kualitas": 33, "Tgl Mulai Macet": 34,
  "Hari Tunggakan Pokok": 35, "Hari Tunggakan Bunga": 36, "Tunggakan Pokok": 37,
  "Tunggakan Bunga": 38, "Sektor Ekonomi": 41, "Kategori Usaha": 42, "Lokasi Penggunaan": 43,
  "Suku Bunga (%)": 44, "Cara Perhitungan Bunga": 45, "Plafon": 51, "Baki Debet": 53,
  "CKPN": 59, "Status BMPK": 63, "Tgl Akad Awal": 67, "Tgl Akad Akhir": 68,
};
const OUTPUT_COLS = [
  "Cabang", "Kol Label", "NPL Flag", "Kualitas", "Nama", "ID Pihak Lawan", "No. Identitas", "No. Rekening",
  "Jenis", "Jenis Penggunaan", "Status Restrukturisasi", "Hubungan dengan Bank",
  "Plafon", "Baki Debet", "Tunggakan Pokok", "Tunggakan Bunga",
  "Hari Tunggakan Pokok", "Hari Tunggakan Bunga", "Tgl Mulai Macet",
  "CKPN", "Suku Bunga (%)", "Cara Perhitungan Bunga", "Periode Pembayaran", "Sektor Ekonomi", "Kategori Usaha",
  "Lokasi Penggunaan", "Sumber Dana Pelunasan", "Status BMPK", "Tgl Akad Awal", "Tgl Akad Akhir",
  "Tanggal Mulai", "Tanggal Jatuh Tempo", "Jangka Waktu (Bulan)", "Jangka Waktu", "File Sumber",
];
const MONEY = new Set(["Plafon", "Baki Debet", "Tunggakan Pokok", "Tunggakan Bunga", "CKPN"]);
const KOL = { 1: "1-Lancar", 2: "2-DPK", 3: "3-Kurang Lancar", 4: "4-Diragukan", 5: "5-Macet" };

function kolInt(v) { const n = parseInt(parseFloat(String(v).split(" - ")[0]), 10); return isNaN(n) ? 0 : n; }
function money(r, k) { const v = parseFloat(r[k]); return isNaN(v) ? 0 : v; }

function readBranch(aoa, cabang, nameMap) {
  const out = [];
  for (let r = 17; r < aoa.length; r++) {
    const idv = String(H.cell(aoa, r, 5)).trim();
    const rek = String(H.cell(aoa, r, 16)).trim();
    if (!idv || idv.toUpperCase() === "JUMLAH" || !rek) continue;
    const row = { "Cabang": cabang, "File Sumber": `0600-${cabang}` };
    for (const [name, idx] of Object.entries(KREDIT_COLS)) row[name] = H.cell(aoa, r, idx);
    row["Nama"] = lookupName(nameMap || {}, row["ID Pihak Lawan"]);
    for (const [name, mp] of Object.entries(S.TRANSLATE_MAP)) {
      if (name in row && row[name] !== null && row[name] !== "" && row[name] !== 0) row[name] = S.translate(row[name], mp);
    }
    const kol = kolInt(row["Kualitas"]);
    row._kol = kol;
    row["Kol Label"] = KOL[kol] || String(row["Kualitas"]);
    row["NPL Flag"] = kol >= 3 ? "YA" : "TIDAK";
    // Jangka waktu kredit: Tanggal Mulai -> Tanggal Jatuh Tempo (tenor; bukan Akad Awal/Akhir
    // yang utk kredit non-restruktur sering sama -> 0)
    const jw = H.jangkaWaktu(row["Tanggal Mulai"], row["Tanggal Jatuh Tempo"]);
    row["Jangka Waktu (Bulan)"] = jw.bulan;
    row["Jangka Waktu"] = jw.teks;
    out.push(row);
  }
  return out;
}

function sheet(XLSX, data) {
  const aoa = [OUTPUT_COLS, ...data.map(r => OUTPUT_COLS.map(c => { const v = r[c]; return v === undefined ? "" : v; }))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const fill = { fill: { fgColor: { rgb: "1F4E79" } }, font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 } };
  for (let c = 0; c < OUTPUT_COLS.length; c++) { const a = XLSX.utils.encode_cell({ r: 0, c }); if (ws[a]) ws[a].s = fill; }
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}

export function processKredit(files, period, XLSX) {
  const wb = XLSX.utils.book_new();
  const nameMap = buildNameMap(files, XLSX);
  let all = [];
  for (const code of discoverBranches(files, "LBBPRK-0600-")) {
    const name = Object.keys(files).find(n => { const b = n.split("/").pop(); return b.startsWith("LBBPRK-0600-") && (b.endsWith(`-${code}.xls`) || b.endsWith(`-${code}_part1.xls`)); });
    if (!name) continue;
    const wx = XLSX.read(files[name], { type: "array", cellDates: true });
    const aoa = XLSX.utils.sheet_to_json(wx.Sheets[wx.SheetNames[0]], { header: 1, raw: true, defval: null });
    all = all.concat(readBranch(aoa, code, nameMap));
  }

  const totalBd = all.reduce((s, r) => s + money(r, "Baki Debet"), 0);
  // RINGKASAN
  const ring = [["RINGKASAN KREDIT (LBBPRK)", period.periodeLabel], [], ["Kolektibilitas", "Jumlah Rek", "Baki Debet"]];
  for (const k of [1, 2, 3, 4, 5]) {
    const sub = all.filter(r => r._kol === k);
    ring.push([KOL[k], sub.length, Math.round(sub.reduce((s, r) => s + money(r, "Baki Debet"), 0))]);
  }
  const nplBd = all.filter(r => r._kol >= 3).reduce((s, r) => s + money(r, "Baki Debet"), 0);
  ring.push([], ["TOTAL", all.length, Math.round(totalBd)]);
  ring.push([`NPL: ${(totalBd ? nplBd / totalBd * 100 : 0).toFixed(2)}% (Rp ${Math.round(nplBd).toLocaleString("id-ID")})`]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ring), "RINGKASAN");

  const sorted = all.slice().sort((a, b) => (b._kol - a._kol) || (money(b, "Baki Debet") - money(a, "Baki Debet")));
  XLSX.utils.book_append_sheet(wb, sheet(XLSX, sorted), "SEMUA KREDIT");
  for (const k of [1, 2, 3, 4, 5]) {
    const sub = sorted.filter(r => r._kol === k);
    if (sub.length) XLSX.utils.book_append_sheet(wb, sheet(XLSX, sub), `Kol ${k} - ${KOL[k].split("-")[1]}`.slice(0, 31));
  }

  const tag = detectBank(files, XLSX).tag;
  const data = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return {
    filename: `GABUNGAN_KREDIT_${tag}_${period.periodeLabel}.xlsx`,
    data,
    summary: (() => {
      const sind = all.filter(r => String(r["Jenis"] || "").startsWith("01"));
      return {
        jumlah_rekening: all.length,
        total_baki_debet: Math.round(totalBd * 100) / 100,
        total_plafon: Math.round(all.reduce((s, r) => s + money(r, "Plafon"), 0) * 100) / 100,
        jumlah_npl: all.filter(r => r._kol >= 3).length,
        sindikasi_jml: sind.length,
        sindikasi_baki_debet: Math.round(sind.reduce((s, r) => s + money(r, "Baki Debet"), 0) * 100) / 100,
      };
    })(),
  };
}
