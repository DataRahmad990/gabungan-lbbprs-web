// Test: form Sindikasi harus punya kolom nama bank peserta, terisi untuk sandi
// yang namanya tersedia di ZIP, dan KOSONG (bukan tebakan) untuk yang tidak ada.
import fs from "fs"; import path from "path";
import XLSX from "xlsx-js-style";
import JSZip from "jszip";
import { detectPeriod } from "../src/period.js";
import { makeFormProcessor } from "../src/formKonven.js";
import { TRANSLATE_MAP as SYR_TMAP, translateSandi as syrTranslate } from "../src/sandi.js";

const ZIP = process.argv[2];
async function unzip(buf) {
  const zip = await JSZip.loadAsync(buf); const files = {};
  for (const n of Object.keys(zip.files)) { if (zip.files[n].dir) continue; files[n] = await zip.files[n].async("uint8array"); }
  return files;
}
const files = await unzip(fs.readFileSync(ZIP));
const period = detectPeriod(path.basename(ZIP));
const cfg = { reportPrefix: "LBBPRS", translateMap: SYR_TMAP, translate: syrTranslate,
  robustDataStart: true, fillDown: ["Nomor Rekening"], bankNameCols: ["Sandi Bank Peserta"] };
const res = makeFormProcessor("KC4200", "Daftar Pembiayaan Sindikasi", "GABUNGAN_SINDIKASI", cfg)(files, period, XLSX);
if (!res) { console.log("GAGAL: form sindikasi tidak terbaca"); process.exit(1); }

const wb = XLSX.read(res.data, { type: "array" });
const aoa = XLSX.utils.sheet_to_json(wb.Sheets["SEMUA CABANG"], { header: 1, defval: null });
const hdr = aoa[0];
const iSandi = hdr.findIndex(h => String(h).includes("Sandi Bank Peserta"));
const iNama = hdr.findIndex(h => String(h).includes("Nama Bank Peserta"));
let fail = 0;
console.log(`kolom sandi  : ${iSandi >= 0 ? hdr[iSandi] : "TIDAK ADA"}`);
console.log(`kolom nama   : ${iNama >= 0 ? hdr[iNama] : "TIDAK ADA"}`);
if (iNama < 0) { console.log("GAGAL: kolom nama bank peserta tidak dibuat"); process.exit(1); }
if (iNama !== iSandi + 1) { console.log("GAGAL: kolom nama tidak bersebelahan dengan sandi"); fail++; }

const rows = aoa.slice(1).filter(r => r && String(r[iSandi] ?? "").trim());
const terisi = rows.filter(r => String(r[iNama] ?? "").trim()).length;
console.log(`baris        : ${rows.length}, nama terisi: ${terisi}`);
console.log(`summary      : ${res.summary.sandi_bank_bernama}/${res.summary.sandi_bank_unik} sandi bernama, ${res.summary.sandi_bank_belum_bernama} belum`);
if (terisi === 0) { console.log("GAGAL: tidak satu pun nama berhasil diisi"); fail++; }

// Tidak boleh ada nama untuk sandi yang tidak dikenal (anti-ngarang).
const { buildBankNameMap, lookupBankName } = await import("../src/bankNames.js");
const map = buildBankNameMap(files, period, XLSX);
let ngarang = 0;
for (const r of rows) {
  const nm = String(r[iNama] ?? "").trim();
  if (nm && nm !== lookupBankName(map, r[iSandi])) ngarang++;
}
console.log(`${ngarang === 0 ? "OK   " : "GAGAL"} nama selalu berasal dari peta (tidak ada tebakan): ${ngarang} pelanggaran`);
if (ngarang) fail++;
console.log("\ncontoh:");
for (const r of rows.slice(0, 6)) console.log(`   ${r[iSandi]} -> ${String(r[iNama] ?? "") || "(kosong)"}`);
console.log(fail === 0 ? "\nSEMUA TEST LULUS" : `\n${fail} TEST GAGAL`);
process.exit(fail === 0 ? 0 : 1);
