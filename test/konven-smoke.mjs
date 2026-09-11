// Smoke test jalur konven: pastikan perubahan formKonven/bankNames tidak bikin regresi.
import fs from "fs"; import path from "path";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { detectPeriod } from "../src/period.js";
import { processKredit } from "../src/kredit.js";
import { makeFormProcessor } from "../src/formKonven.js";

const ZIP = process.argv[2];
// Tanpa angka bank di repo publik: yang diuji konsistensi internal.
// Acuan opsional lewat env EXPECT_ROWS / EXPECT_BAKI.
async function unzip(buf) {
  const zip = await JSZip.loadAsync(buf); const files = {};
  for (const n of Object.keys(zip.files)) { if (zip.files[n].dir) continue; files[n] = await zip.files[n].async("uint8array"); }
  return files;
}
const files = await unzip(fs.readFileSync(ZIP));
const sample = Object.keys(files).find(n => n.split("/").pop().startsWith("LBBPRK-0600-"));
const period = detectPeriod(path.basename(sample));
let fail = 0;
const kr = processKredit(files, period, XLSX);
const kwb = XLSX.read(kr.data, { type: "array" });
const ka = XLSX.utils.sheet_to_json(kwb.Sheets["SEMUA KREDIT"], { header: 1, defval: null });
const kh = ka[0]; const iRek = kh.indexOf("No. Rekening"), iBd = kh.indexOf("Baki Debet");
const krows = ka.slice(1).filter(r => r && String(r[iRek] ?? "").trim());
const sheetBaki = krows.reduce((s, r) => { const v = parseFloat(r[iBd]); return s + (isNaN(v) ? 0 : v); }, 0);
const okRows = krows.length === kr.summary.jumlah_rekening;
const okBaki = Math.round(sheetBaki) === Math.round(kr.summary.total_baki_debet);
const uniq = new Set(krows.map(r => String(r[iRek]).trim())).size;
const okDup = uniq === krows.length;
console.log(`${okRows ? "OK   " : "GAGAL"} Baris sheet == summary (${krows.length})`);
console.log(`${okBaki ? "OK   " : "GAGAL"} Baki sheet == summary`);
console.log(`${okDup ? "OK   " : "GAGAL"} Tidak ada rekening dobel (unik ${uniq})`);
if (!okRows) fail++; if (!okBaki) fail++; if (!okDup) fail++;
if (process.env.EXPECT_ROWS && krows.length !== Number(process.env.EXPECT_ROWS)) { console.log("GAGAL: EXPECT_ROWS tidak cocok"); fail++; }
if (process.env.EXPECT_BAKI && Math.round(sheetBaki) !== Number(process.env.EXPECT_BAKI)) { console.log("GAGAL: EXPECT_BAKI tidak cocok"); fail++; }

for (const [code, title, nm] of [["1100", "Daftar Tabungan", "GABUNGAN_TABUNGAN"], ["1200", "Daftar Deposito", "GABUNGAN_DEPOSITO"], ["0601", "Daftar Agunan", "GABUNGAN_AGUNAN"]]) {
  const r = makeFormProcessor(code, title, nm)(files, period, XLSX);
  const ok = r && r.summary.jumlah_baris > 0;
  if (!ok) fail++;
  console.log(`${ok ? "OK   " : "GAGAL"} Form ${code} ${title.padEnd(16)} baris=${r ? r.summary.jumlah_baris : "-"}`);
}
console.log(fail === 0 ? "\nKONVEN SMOKE LULUS" : `\n${fail} GAGAL`);
process.exit(fail === 0 ? 0 : 1);
