// Deteksi periode + kode bank dari nama file ZIP LBBPRS (port dari core/period.py)
export const BULAN_ID = {1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MEI",6:"JUN",
  7:"JUL",8:"AGU",9:"SEP",10:"OKT",11:"NOV",12:"DES"};

export function detectPeriod(namaFile) {
  const base = String(namaFile).split("/").pop();
  const m = base.match(/-(\d{8})-(\d{6})-/);
  if (!m) {
    throw new Error(`Nama file '${base}' nggak sesuai pola LBBPRS. ` +
      "Harus ada '-YYYYMMDD-KODEBANK-' (contoh: ...-20260228-620086-...).");
  }
  const ymd = m[1], kodeBank = m[2];
  const y = +ymd.slice(0,4), mo = +ymd.slice(4,6), d = +ymd.slice(6,8);
  const tanggal = new Date(Date.UTC(y, mo-1, d));
  if (isNaN(tanggal) || tanggal.getUTCMonth() !== mo-1) {
    throw new Error(`Tanggal '${ymd}' di nama file '${base}' nggak valid.`);
  }
  const jm = base.match(/^LBBPRS-LBBPRS\d+-([A-Z])-/);
  const jenis = jm ? jm[1] : "?";
  return {
    tanggalPosisi: tanggal,
    periodeLabel: `${BULAN_ID[mo]}${y}`,
    kodeBank, jenisLapor: jenis,
  };
}
