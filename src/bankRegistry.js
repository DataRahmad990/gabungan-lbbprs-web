// Daftar sandi bank -> nama bank, diisi MANUAL dari sumber resmi.
//
// Kenapa file ini ada: form Sindikasi cuma menyimpan sandi bank peserta. Nama bank
// bisa diambil otomatis dari form Penempatan (KC0500), tapi itu cuma menutup bank
// yang kebetulan jadi tempat penempatan dana. Peserta sindikasi lain tidak tertutup.
//
// ATURAN: isi HANYA dari daftar sandi resmi (mis. daftar LJK OJK / data APOLO).
// Jangan menebak dari ingatan. Sandi yang tidak ada di sini akan dikosongkan di
// output, dan itu memang perilaku yang diinginkan.
//
// Format: "sandi": "Nama Bank"
// Contoh:
//   "620086": "PT Bank Perekonomian Rakyat Syariah Suriyah",
//
// Isian di sini menang atas hasil deteksi otomatis, jadi bisa dipakai untuk
// membetulkan nama yang di form penempatan tertulis singkat/tidak konsisten.
export const BANK_REGISTRY = {
};
