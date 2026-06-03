# Gabung Data Bulanan LBBPRS (versi web)

Web app yang jalan **sepenuhnya di browser**. Upload 1 ZIP laporan bulanan LBBPRS,
langsung jadi beberapa file Excel gabungan siap pakai (pembiayaan, ABP, penempatan,
neraca + laba rugi tren).

## Penting soal data
**Nggak ada data yang dikirim ke server mana pun.** Semua pemrosesan (baca .xls,
gabung, bikin .xlsx) terjadi di browser pemakai. GitHub Pages cuma nyajiin halaman
statis. Library (SheetJS/JSZip) di-vendor lokal di `vendor/`, jadi nol fetch ke CDN.
Repo ini berisi **kode doang, nol data pemeriksaan**.

## Cara pakai
Buka link-nya, masukin PIN, drag ZIP, download hasil. Selesai. Nggak ada yang disimpen
(riwayat cuma di localStorage browser, bisa dihapus kapan aja).

## Struktur
- `index.html` - halaman + PIN gate
- `src/` - logika (port dari versi Python): period, sandi SEOJK 17, processor pembiayaan/abp/penempatan/neraca
- `vendor/` - SheetJS (xlsx-js-style) + JSZip, lokal
- `test/` - test Node (verifikasi hasil sama dengan versi Python/golden)

## Catatan
PIN disimpan sebagai hash SHA-256 (bukan teks asli). Ini gembok ringan biar nggak
sembarang orang masuk, bukan pengaman tingkat bank. Karena data nggak pernah disimpen
di server, risikonya kecil.
