// Peta "sandi bank" -> "nama bank", dibangun dari isi ZIP itu sendiri.
//
// Kenapa perlu: form Sindikasi (konven 0602 / syariah KC4200) cuma menyimpan
// "Sandi Bank Peserta Sindikasi" berupa kode 6 digit, tanpa nama. Nama bank tidak
// ada di form itu, jadi harus diambil dari form lain di ZIP yang sama.
//
// Sumber yang dipakai (semuanya dari data pelapor, tidak ada daftar karangan):
//   1. Bank pelapor sendiri  : kode bank dari nama file ZIP + nama dari GB0200.
//   2. Form Penempatan pada Bank Lain syariah (KC0500): kolom 4 = Sandi Bank,
//      kolom 5 = Nama Bank. Ini satu-satunya form labul yang memasangkan keduanya.
//
// Kode yang tidak ketemu SENGAJA dibiarkan kosong, bukan ditebak.
import * as H from "./helpers.js";
import { detectBank } from "./bank.js";
import { BANK_REGISTRY } from "./bankRegistry.js";

export function normSandi(v) {
  let s = String(v == null ? "" : v).trim();
  if (s.endsWith(".0")) s = s.slice(0, -2);
  return s;
}

export function buildBankNameMap(files, period, XLSX) {
  const bySandi = {};
  const sumber = {};

  // 1. Bank pelapor sendiri.
  if (period && period.kodeBank) {
    try {
      const { nama } = detectBank(files, XLSX);
      if (nama && nama !== "BANK") {
        bySandi[normSandi(period.kodeBank)] = nama;
        sumber[normSandi(period.kodeBank)] = "bank pelapor";
      }
    } catch { /* abaikan */ }
  }

  // 2. Form Penempatan pada Bank Lain syariah (KC0500): sandi kol 4, nama kol 5.
  for (const name of Object.keys(files)) {
    const b = name.split("/").pop();
    if (!b.startsWith("LBBPRS-KC0500-") || !b.endsWith(".xls")) continue;
    try {
      const wb = XLSX.read(files[name], { type: "array", raw: true });
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
      for (let r = 17; r < aoa.length; r++) {
        const sandi = normSandi(H.cell(aoa, r, 4));
        const nama = String(H.cell(aoa, r, 5) ?? "").trim();
        if (!sandi || !nama || sandi.toUpperCase() === "JUMLAH") continue;
        if (!(sandi in bySandi)) { bySandi[sandi] = nama; sumber[sandi] = "form penempatan (KC0500)"; }
      }
    } catch { /* lanjut file berikutnya */ }
  }

  // 3. Daftar manual (bankRegistry.js) - menang atas deteksi otomatis.
  for (const [sandi, nama] of Object.entries(BANK_REGISTRY || {})) {
    const k = normSandi(sandi);
    if (k && String(nama || "").trim()) { bySandi[k] = String(nama).trim(); sumber[k] = "daftar manual"; }
  }

  return { bySandi, sumber };
}

export function lookupBankName(map, sandi) {
  if (!map || !map.bySandi) return "";
  return map.bySandi[normSandi(sandi)] || "";
}

export function bankNameCount(map) {
  return map && map.bySandi ? Object.keys(map.bySandi).length : 0;
}
