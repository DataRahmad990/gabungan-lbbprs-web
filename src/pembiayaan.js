// Processor Pembiayaan (port dari core/pembiayaan.py) - jalan di browser & Node.
import { FORMS, OUTPUT_COLS, BH_EXTRA_COLS, IJR_EXTRA_COLS } from "./colmaps.js";
import { TRANSLATE_MAP, translateSandi } from "./sandi.js";
import { REF_PENGIKATAN_ROWS, REF_AGUNAN_ROWS } from "./enrichData.js";
import * as E from "./enrich.js";
import * as H from "./helpers.js";
import { detectBank, discoverBranches } from "./bank.js";

function findBranchAoa(files, formCode, code, XLSX) {
  const prefix = `LBBPRS-${formCode}-`;
  const name = Object.keys(files).find(n => {
    const b = n.split("/").pop();
    return b.startsWith(prefix) && (b.endsWith(`-${code}.xls`) || b.endsWith(`-${code}_part1.xls`));
  });
  if (!name) return null;
  const wb = XLSX.read(files[name], { type: "array", cellDates: true });
  const sh = wb.Sheets[wb.SheetNames[0]];
  const aoa = XLSX.utils.sheet_to_json(sh, { header: 1, raw: true, defval: null });
  return { aoa, name };
}

function readFormData(files, formCode, colMap, posisi, XLSX) {
  const out = [];
  for (const code of discoverBranches(files, `LBBPRS-${formCode}-`)) {  // dinamis
    const found = findBranchAoa(files, formCode, code, XLSX);
    if (!found) continue;
    const { aoa } = found;
    const start = H.findDataStartRow(aoa);
    for (let r = start; r < aoa.length; r++) {
      const idVal = H.cell(aoa, r, 3);
      const idStr = String(idVal).trim();
      if (!idStr || idStr.toUpperCase() === "JUMLAH") continue;
      const nama = String(H.cell(aoa, r, 4)).trim();
      if (!nama) continue;

      const row = { "Cabang": code, "File Sumber": `${formCode}-${code}`, "Baris Asli": `Baris ${r+1}` };
      for (const [colName, colIdx] of Object.entries(colMap)) {
        if (colIdx === null) { row[colName] = ""; continue; }
        const v = H.cell(aoa, r, colIdx);
        row[colName] = v;
      }

      // Baki Debet Ijarah = Plafon - Akum Penyusutan
      if ((row["Baki Debet"] === null || row["Baki Debet"] === "" || row["Baki Debet"] === 0) &&
          ("_akum_penyusutan" in row)) {
        const plafon = parseFloat(row["Plafon"]) || 0;
        const akum = parseFloat(row["_akum_penyusutan"]) || 0;
        row["Baki Debet"] = plafon - akum;
      }

      // Translate sandi
      for (const [colName, dict] of Object.entries(TRANSLATE_MAP)) {
        if (colName in row) {
          const v = row[colName];
          if (v !== null && v !== "" && v !== 0 && v !== 0.0) {
            row[colName] = translateSandi(v, dict);
          }
        }
      }

      // Overdue
      const jtDate = H.parseDate(row["Tanggal Jatuh Tempo"]);
      const od = H.calcOverdueDays(jtDate, posisi);
      row["Hari Overdue"] = od > 0 ? od : 0;

      // Jangka waktu pembiayaan: Tanggal Mulai -> Tanggal Jatuh Tempo
      const jw = H.jangkaWaktu(row["Tanggal Mulai"], row["Tanggal Jatuh Tempo"]);
      row["Jangka Waktu (Bulan)"] = jw.bulan;
      row["Jangka Waktu"] = jw.teks;

      // Kol
      let kolRaw = row["Kualitas"];
      let kolVal = (typeof kolRaw === "string" && kolRaw.includes(" - ")) ? kolRaw.split(" - ")[0] : kolRaw;
      row["Kol Label"] = H.kolLabel(kolVal);
      row["Flag Anomali"] = H.detectAnomaly(kolVal, od, jtDate);
      row["_kol_raw"] = kolVal;

      // Enrichment
      row["Lokasi Penggunaan"] = E.translateDati2(row["Lokasi Penggunaan"]);
      row["Jenis Piutang"] = row["Jenis Piutang"] || row["Jenis Pembiayaan"] || "";
      const peng = row["Jenis Pengikatan"] || "";
      row["HT Flag"] = E.htFlag(peng);
      row["NPF Flag"] = E.npfFlag(kolVal);
      const tm = H.parseDate(row["Tgl Mulai Macet"]);
      row["Lama Macet (Hari)"] = E.lamaMacetHari(tm, posisi);
      row["Keterangan Pengikatan"] = E.keteranganPengikatan(peng);

      out.push(row);
    }
  }
  return out;
}

function aliasCombined(row, akad) {
  const r = { ...row, "Akad": akad };
  if ("Sifat Pembiayaan" in r && !("Sifat Piutang" in r)) r["Sifat Piutang"] = r["Sifat Pembiayaan"];
  if ("Status Pembiayaan" in r && !("Status Piutang" in r)) r["Status Piutang"] = r["Status Pembiayaan"];
  if ("Tunggakan Pokok Jumlah" in r && !("Tunggakan Pokok" in r)) r["Tunggakan Pokok"] = r["Tunggakan Pokok Jumlah"];
  if ("Tunggakan BH Jumlah" in r && !("Tunggakan Margin" in r)) r["Tunggakan Margin"] = r["Tunggakan BH Jumlah"];
  if ("Tunggakan Imbalan" in r && !("Tunggakan Margin" in r)) r["Tunggakan Margin"] = r["Tunggakan Imbalan"];
  return r;
}

function colsFor(extraCols, sample) {
  const cols = OUTPUT_COLS.slice();
  if (extraCols) {
    const idx = cols.indexOf("Sektor Ekonomi") + 1;
    for (const ec of [...extraCols].reverse()) {
      if (sample && ec in sample) cols.splice(idx, 0, ec);
    }
  }
  return cols;
}

function sheetFromRows(XLSX, rows, cols) {
  const aoa = [cols];
  for (const row of rows) {
    aoa.push(cols.map(c => {
      let v = row[c];
      if (v === undefined) v = "";
      return v;
    }));
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
  // header style
  const fill = { fill: { fgColor: { rgb: "1F4E79" } }, font: { color: { rgb: "FFFFFF" }, bold: true, sz: 10 } };
  for (let c = 0; c < cols.length; c++) {
    const addr = XLSX.utils.encode_cell({ r: 0, c });
    if (ws[addr]) ws[addr].s = fill;
  }
  ws["!freeze"] = { xSplit: 0, ySplit: 1 };
  return ws;
}

function refSheet(XLSX, rows) {
  const ws = XLSX.utils.aoa_to_sheet(rows.map(r => r.map(v => v === null ? "" : v)));
  ws["!cols"] = [{ wch: 10 }, { wch: 20 }, { wch: 40 }, { wch: 80 }];
  return ws;
}

export function processPembiayaan(files, period, XLSX) {
  const posisi = period.tanggalPosisi;
  const wb = XLSX.utils.book_new();
  const allStats = {};
  const allCombined = [];
  const akadSheets = []; // {name, rows, cols}

  for (const [formCode, formName, colMap] of FORMS) {
    const data = readFormData(files, formCode, colMap, posisi, XLSX);
    if (!data.length) continue;

    if (formName === "Bagi Hasil") {
      const mud = [], mus = [];
      for (const row of data) {
        const af = String(row["Jenis Akad"] || "");
        const code = af.includes(" - ") ? af.split(" - ")[0].trim() : af;
        (code === "20" ? mud : mus).push(row);
      }
      if (mud.length) akadSheets.push({ name: "Mudarabah", rows: mud, cols: colsFor(BH_EXTRA_COLS, mud[0]) });
      if (mus.length) akadSheets.push({ name: "Musyarakah", rows: mus, cols: colsFor(BH_EXTRA_COLS, mus[0]) });
      for (const row of mud) allCombined.push(aliasCombined(row, "Mudarabah"));
      for (const row of mus) allCombined.push(aliasCombined(row, "Musyarakah"));
      continue;
    }

    let extra = null;
    if (formName === "Ijarah") extra = IJR_EXTRA_COLS;
    akadSheets.push({ name: formName, rows: data, cols: colsFor(extra, data[0]) });
    for (const row of data) allCombined.push(aliasCombined(row, formName));
  }

  // Sort SEMUA AKAD: anomali dulu, kol desc, overdue desc
  allCombined.sort((a, b) => {
    const aa = a["Flag Anomali"] ? 0 : 1, ab = b["Flag Anomali"] ? 0 : 1;
    if (aa !== ab) return aa - ab;
    const ka = -(parseInt(parseFloat(a["_kol_raw"]), 10) || 0);
    const kb = -(parseInt(parseFloat(b["_kol_raw"]), 10) || 0);
    if (ka !== kb) return ka - kb;
    return -(a["Hari Overdue"] || 0) - (-(b["Hari Overdue"] || 0));
  });

  // RINGKASAN (sederhana)
  const ringkasan = [["RINGKASAN PEMBIAYAAN - BPRS SURIYAH", period.periodeLabel],
    ["Total rekening", allCombined.length]];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ringkasan), "RINGKASAN");

  // SEMUA AKAD
  XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, allCombined, OUTPUT_COLS), "SEMUA AKAD");
  // per akad
  for (const s of akadSheets) {
    XLSX.utils.book_append_sheet(wb, sheetFromRows(XLSX, s.rows, s.cols), s.name.slice(0, 31));
  }
  // REF sheets
  XLSX.utils.book_append_sheet(wb, refSheet(XLSX, REF_PENGIKATAN_ROWS), "REF PENGIKATAN");
  XLSX.utils.book_append_sheet(wb, refSheet(XLSX, REF_AGUNAN_ROWS), "REF AGUNAN");

  // Summary
  let totalBaki = 0, npf = 0, anomali = 0;
  for (const row of allCombined) {
    totalBaki += parseFloat(row["Baki Debet"]) || 0;
    if (row["NPF Flag"] === "YA") npf++;
    if (row["Flag Anomali"]) anomali++;
  }
  const summary = {
    jumlah_rekening: allCombined.length,
    total_baki_debet: Math.round(totalBaki * 100) / 100,
    jumlah_npf: npf,
    jumlah_anomali: anomali,
  };

  const tag = detectBank(files, XLSX).tag;
  const data = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  return {
    filename: `GABUNGAN_PEMBIAYAAN_${tag}_${period.periodeLabel}.xlsx`,
    data, summary,
  };
}
