// Enrich helpers (port dari core/enrich.py)
import { KODE_DATI2 } from "./kodeDati2.js";
import { HT_CODES, NPF_KOL, KETERANGAN_PENGIKATAN } from "./enrichData.js";

export function kode(val) {
  if (val === null || val === undefined) return "";
  let s = String(val).trim();
  if (s.includes(" - ")) return s.split(" - ")[0].trim();
  return s;
}

export function formatSandiDati2(raw) {
  if (raw === null || raw === undefined || raw === "") return "";
  const n = Number(raw);
  if (!isNaN(n) && isFinite(n)) return String(Math.trunc(n)).padStart(4, "0");
  const s = String(raw).trim();
  return /^\d+$/.test(s) ? s.padStart(4, "0") : s;
}

export function translateDati2(raw) {
  const sandi = formatSandiDati2(raw);
  if (!sandi) return "";
  if (sandi in KODE_DATI2) return `${sandi} - ${KODE_DATI2[sandi]}`;
  return sandi;
}

export function htFlag(jenisPengikatan) {
  return HT_CODES.has(kode(jenisPengikatan)) ? "YA" : "TIDAK";
}

export function npfFlag(kolVal) {
  return NPF_KOL.has(kode(kolVal)) ? "YA" : "TIDAK";
}

// (posisi - tgl mulai macet) dalam hari. null kalau gak ada tanggal macet.
export function lamaMacetHari(tglMacet, posisi) {
  if (!tglMacet) return null;
  const MS = 86400000;
  return Math.round((posisi.getTime() - tglMacet.getTime()) / MS);
}

export function keteranganPengikatan(jenisPengikatan) {
  return KETERANGAN_PENGIKATAN[kode(jenisPengikatan)] || "";
}
