#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
=============================================================================
 BUILD DATA — Monitoring PDP sampai Juni 2026, DIV SAK — PT PLN (Persero)
=============================================================================
Script ini membaca spreadsheet sumber ("data pdp sampai juni") dan mengubahnya
menjadi data/pdp-data.json, format data yang dipakai oleh index.html.

Sheet yang dibaca (harus persis nama ini di dalam spreadsheet):
    Rekap Settlement, D1, D2, D3, D4, Cluster Proyek 1-9,
    Biaya Ditangguhkan, SAP-DIVAKT

Cara pakai:
    1. Pastikan file spreadsheet sudah diexport/didownload sebagai .xlsx
       dengan nama "sheet.xlsx" di folder yang sama dengan script ini, ATAU
    2. Set environment variable SHEET_ID (ID Google Spreadsheet) dan jalankan
       script ini — ia akan otomatis mengunduh versi terbaru dari Google
       Sheets (spreadsheet harus di-share "Anyone with the link can view").

Jalankan:
    python3 build_data.py

Output:
    data/pdp-data.json

Catatan penting bila menambah bulan baru di spreadsheet (mis. Juli 2026):
    - Tambahkan entri baru ke daftar BULAN di bawah.
    - Tambahkan kolom "Saldo PDP <Bulan> <Tahun>" (dan "Progress Fisik (%)
      ... <Bulan> <Tahun>" khusus sheet D3) mengikuti pola nama kolom yang
      SAMA dengan bulan-bulan sebelumnya — script ini mencari kolom
      berdasarkan **teks header**, bukan posisi kolom, jadi urutan kolom
      boleh berantakan asalkan teks headernya konsisten dengan pola yang ada.
=============================================================================
"""
import io
import json
import os
import re
import sys
import urllib.request
import datetime

try:
    import openpyxl
except ImportError:
    sys.exit("Perlu paket 'openpyxl'. Jalankan: pip install openpyxl")

SHEET_ID = os.environ.get("SHEET_ID", "15RlWr9D1QsWxqBKnzFlz0bcFhzkxxM_xLvxYjpR-SzQ")
LOCAL_XLSX = os.environ.get("LOCAL_XLSX", "sheet.xlsx")
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "pdp-data.json")

# 10 bulan yang saat ini didukung struktur spreadsheet (Sep 2025 - Jun 2026).
BULAN = [
    ("September", "2025"), ("Oktober", "2025"), ("November", "2025"), ("Desember", "2025"),
    ("Januari", "2026"), ("Februari", "2026"), ("Maret", "2026"), ("April", "2026"),
    ("Mei", "2026"), ("Juni", "2026"),
]
PERIODE_LABEL = "Juni 2026"  # bulan snapshot yang ditampilkan di beranda dashboard

# Peta 11 UIP: key dipakai di sheet D1-D4 (kolom Unit), st = singkatan yang
# dipakai di tabel pivot "Rekap Settlement", full = nama "UIP ..." lengkap
# yang dipakai di sheet Cluster & SAP-DIVAKT.
UIPS = [
    {"key": "SUMBAGUT",     "st": "SBU",   "rekap": "Sumbagut",      "pos": [8.05, 23.09],  "label": "UIP Sumbagut"},
    {"key": "SUMBAGTENG",   "st": "SBT",   "rekap": "Sumbagteng",    "pos": [14.26, 36.48], "label": "UIP Sumbagteng"},
    {"key": "SUMBAGSEL",    "st": "SBS",   "rekap": "Sumbagsel",     "pos": [21.32, 51.66], "label": "UIP Sumbagsel"},
    {"key": "KALBAGBAR",    "st": "KLB",   "rekap": "Kalbagbar",     "pos": [31.17, 38.8],  "label": "UIP Kalbagbar"},
    {"key": "KALBAGTIM",    "st": "KLT",   "rekap": "Kalbagtim",     "pos": [47.23, 44.07], "label": "UIP Kalbagtim"},
    {"key": "SULAWESI",     "st": "SUL",   "rekap": "Sulawesi",      "pos": [52.8, 61.72],  "label": "UIP Sulawesi"},
    {"key": "JBB",          "st": "JBB",   "rekap": "JBB",           "pos": [25.82, 66.43], "label": "UIP JBB"},
    {"key": "JBT",          "st": "JBT",   "rekap": "JBT",           "pos": [33.53, 70.02], "label": "UIP JBT"},
    {"key": "JBTB",         "st": "JBTB",  "rekap": "JBTB",          "pos": [38.56, 71.14], "label": "UIP JBTB"},
    {"key": "MALUKU PAPUA", "st": "MPA",   "rekap": "Maluku Papua",  "pos": [84.06, 51.21], "label": "UIP Maluku Papua"},
    {"key": "NUSRA",        "st": "NUSRA", "rekap": "Nusra",         "pos": [51.94, 77.67], "label": "UIP Nusra"},
]
BY_KEY = {u["key"]: u for u in UIPS}
BY_ST = {u["st"]: u for u in UIPS}
BY_REKAP = {u["rekap"].strip().lower(): u for u in UIPS}

# Peta SVG Indonesia (statis, presentasional) dan viewBox-nya — dipakai untuk
# menggambar bubble map di halaman Beranda. Tidak berasal dari spreadsheet.
MAP_VIEWBOX = {"w": 1000, "h": 480}
MAP_PATH_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "map_path.txt")


# =============================================================================
# HELPERS
# =============================================================================
def clean(v):
    """Bersihkan satu nilai sel: None tetap None, tanggal -> 'dd Mon yyyy',
    angka bulat -> string tanpa '.0', teks -> whitespace dirapikan."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.strftime("%d %b %Y")
    if isinstance(v, (int, float)):
        return str(int(v)) if float(v).is_integer() else str(v)
    s = str(v).replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s if s else None


def rawtext(v):
    """Untuk nilai yang ingin dipertahankan APA ADANYA dengan tipe aslinya
    (mis. rincian Lot I/II/III yang tetap teks, atau angka yang tetap angka),
    hanya membuang spasi di ujung teks."""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.replace("\xa0", " ").strip()
        return s if s else None
    return v


def text_raw(v):
    """Untuk kolom nama/judul yang HARUS menjaga isi APA ADANYA — tanpa
    strip, tanpa normalisasi spasi/nbsp sama sekali — angka & tanggal tetap
    dikonversi ke teks. Tanggal diformat 'dd Mon yyyy'."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.strftime("%d %b %Y")
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v)
    return s if s != "" else None


def text_strip(v):
    """Untuk kolom teks deskriptif (keterangan, status, uraian, dsb.):
    spasi/baris-baru DI TENGAH teks dijaga apa adanya, tapi spasi di awal
    & akhir dibuang. Tanggal diformat 'dd Mon yyyy'."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.strftime("%d %b %Y")
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v).replace("\xa0", " ").strip()
    return s if s != "" else None


def text_plain(v):
    """Sama seperti text_strip, tapi tanggal TIDAK diformat ulang (dibiarkan
    representasi default Python) — dipakai utk kolom yang secara empiris
    tidak diformat oleh dashboard asli (mis. Biaya Ditangguhkan/target COD)."""
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return str(v)
    if isinstance(v, (int, float)):
        return str(v)
    s = str(v).replace("\xa0", " ").strip()
    return s if s != "" else None


def cap(v, n):
    """Potong string ke n karakter + '…' bila melebihi; selain itu biarkan."""
    if isinstance(v, str) and len(v) > n:
        return v[:n] + "…"
    return v


def fixcase(s):
    return s[0].upper() + s[1:] if s else s


def numf(v):
    """Parse angka dengan aman; teks yang tidak valid -> 0.0."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(".", "").replace(",", ".").strip())
    except (TypeError, ValueError):
        return 0.0


def numf_id(v):
    """Parse angka format Indonesia (titik = pemisah ribuan) dari sel teks,
    dipakai khusus untuk sheet Cluster Proyek 1-9."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s == "-":
        return 0.0
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def frac(v):
    """Normalisasi progress fisik (D3) ke PECAHAN 0..1 (bukan persen), karena
    itulah format yang dipakai index.html (dikali 100 saat ditampilkan).
    - Bukan angka -> None (data tidak valid/tidak bisa dibaca).
    - x > 1  -> dibagi 100 (nilai dientry sebagai persen, mis. 98.16 = 98.16%).
    - Hasil dibatasi ke [0, 1] untuk melindungi dari salah ketik ekstrem."""
    if not isinstance(v, (int, float)):
        return None
    x = float(v)
    if x > 1:
        x /= 100.0
    return round(max(0.0, min(x, 1.0)), 4)


def pct_num(v):
    """Progress dari sheet Cluster (angka 0..1 atau teks '93,35%') -> persen 0..100."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return round(float(v) * 100, 2)
    s = str(v).strip().replace("%", "").replace(",", ".")
    if not s:
        return None
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def durasi_bulan(v):
    """'X tahun, Y bulan' -> jumlah bulan. Selain itu (termasuk '-' atau 0) -> None."""
    if not isinstance(v, str):
        return None
    m = re.search(r"(\d+)\s*tahun.*?(\d+)\s*bulan", v, re.I)
    return int(m.group(1)) * 12 + int(m.group(2)) if m else None


def fungsi_of(v):
    """Normalisasi kolom Fungsi: 'KIT'/'Bagian Proyek KIT' -> 'KIT', dst.
    Sel kosong -> None."""
    if v is None:
        return None
    s = str(v).upper()
    if "KIT" in s:
        return "KIT"
    if "TL" in s:
        return "TL"
    if "GI" in s:
        return "GI"
    return None


def header_row(ws, row_no=3, max_col=80):
    out = []
    for c in range(1, max_col + 1):
        v = ws.cell(row=row_no, column=c).value
        out.append(re.sub(r"\s+", " ", str(v)).strip().lower() if v is not None else "")
    return out


def find_col(headers, *must_contain, start=0):
    for i in range(start, len(headers)):
        h = headers[i]
        if all(m.lower() in h for m in must_contain):
            return i
    raise RuntimeError(f"Kolom tidak ditemukan untuk kata kunci {must_contain}")


def find_col_opt(headers, *must_contain, start=0):
    try:
        return find_col(headers, *must_contain, start=start)
    except RuntimeError:
        return None


# =============================================================================
# LOAD WORKBOOK
# =============================================================================
def load_workbook():
    if os.path.exists(LOCAL_XLSX):
        print(f"[build_data] Membaca file lokal: {LOCAL_XLSX}")
        return openpyxl.load_workbook(LOCAL_XLSX, data_only=True)

    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
    print("[build_data] Mengunduh spreadsheet dari Google Sheets...")
    try:
        with urllib.request.urlopen(url, timeout=60) as resp:
            raw = resp.read()
    except Exception as e:
        sys.exit(
            "Gagal mengunduh spreadsheet. Pastikan sharing spreadsheet sudah "
            "\"Anyone with the link can view\", atau letakkan file export-nya "
            f"sebagai '{LOCAL_XLSX}' di folder ini.\nDetail error: {e}"
        )
    return openpyxl.load_workbook(io.BytesIO(raw), data_only=True)


# =============================================================================
# PARSERS PER SHEET
# =============================================================================
def parse_sap(wb):
    ws = wb["SAP-DIVAKT"]
    h = header_row(ws)
    c_unit = find_col(h, "unit")
    c_atbm = find_col(h, "atbm")
    c_bd = find_col(h, "biaya ditangguhkan")
    c_konst = find_col(h, "pdp konstruksi")
    c_mat = find_col(h, "pdp material")
    c_muka = find_col(h, "pembayaran dimuka")
    c_total = find_col(h, "total pdp")
    c_reg = find_col(h, "regional")

    rows = []
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
        unit = row[c_unit]
        if not unit or str(unit).strip().upper() == "TOTAL":
            continue
        atbm, bd, konst, mat, muka = (row[c_atbm] or 0, row[c_bd] or 0, row[c_konst] or 0,
                                       row[c_mat] or 0, row[c_muka] or 0)
        total = row[c_total] if row[c_total] is not None else (atbm + konst + mat + muka)
        reg = row[c_reg]
        rows.append({
            "co_code": text_plain(row[0]),
            "unit": clean(unit), "atbm": float(atbm), "biaya_ditangguhkan": float(bd),
            "pdp_konstruksi": float(konst), "pdp_material": float(mat),
            "pdp_pembayaran_dimuka": float(muka), "total_pdp": float(total),
            "regional": clean(reg),
        })

    totals = {
        "atbm": sum(r["atbm"] for r in rows), "biaya_ditangguhkan": sum(r["biaya_ditangguhkan"] for r in rows),
        "pdp_konstruksi": sum(r["pdp_konstruksi"] for r in rows), "pdp_material": sum(r["pdp_material"] for r in rows),
        "pdp_pembayaran_dimuka": sum(r["pdp_pembayaran_dimuka"] for r in rows),
        "total_pdp": sum(r["total_pdp"] for r in rows),
    }
    by_regional = {}
    for r in rows:
        g = by_regional.setdefault(r["regional"] or "Lainnya", {"total_pdp": 0.0, "count": 0})
        g["total_pdp"] += r["total_pdp"]
        g["count"] += 1

    total_sap_uip = by_regional.get("UIP", {}).get("total_pdp", 0.0)
    return rows, totals, by_regional, total_sap_uip


def parse_rekap_settlement(wb):
    """Sheet Rekap Settlement: tabel target (kolom A-H, header baris 2) +
    5 tabel pivot 'Sum of NILAI SETTLEMENT' (kategori PDP/ATBM ke AT/ATBM)."""
    ws = wb["Rekap Settlement"]
    h = header_row(ws, row_no=2)
    c_uip = find_col(h, "uip")
    c_target26 = find_col(h, "target settlement tahun 2026")
    c_targetjun = find_col(h, "target settlement", start=c_target26 + 1)

    target = {}  # uip (Title case, mis. "Sumbagut") -> {target_2026, target_juni}
    for row in ws.iter_rows(min_row=4, max_row=30, values_only=True):
        uip = row[c_uip]
        if not uip:
            continue
        key = str(uip).strip().lower()
        if key not in BY_REKAP:
            continue  # lewati baris TOTAL/footnote
        target[str(uip).strip()] = {
            "target_2026": float(row[c_target26] or 0),
            "target_juni": float(row[c_targetjun] or 0),
        }
    # baris TOTAL (tetap di kolom yang sama)
    total_target_2026 = total_target_juni = 0.0
    for row in ws.iter_rows(min_row=4, max_row=30, values_only=True):
        uip = row[c_uip]
        if uip and str(uip).strip().upper() == "TOTAL":
            total_target_2026 = float(row[c_target26] or 0)
            total_target_juni = float(row[c_targetjun] or 0)
            break

    # --- cari 5 tabel pivot "Sum of NILAI SETTLEMENT" ---
    pivots = []  # list of dict: {kategori, settle_ke, rows: {ROWLABEL: grand_total}}
    r = 1
    max_r = ws.max_row
    while r <= max_r:
        v = ws.cell(r, 1).value
        if isinstance(v, str) and v.strip() == "KATEGORI (PDP/ATBM)":
            kategori = ws.cell(r, 2).value
            settle_ke = ws.cell(r + 1, 2).value
            # cari baris header "Row Labels" dalam 5 baris berikutnya
            hdr_r = None
            for rr in range(r, r + 6):
                if ws.cell(rr, 1).value == "Row Labels":
                    hdr_r = rr
                    break
            if hdr_r is None:
                r += 1
                continue
            headers2 = [ws.cell(hdr_r, c).value for c in range(1, 12)]
            gt_col = None
            for i, hv in enumerate(headers2):
                if hv == "Grand Total":
                    gt_col = i + 1
                    break
            data = {}
            rr = hdr_r + 1
            while rr <= max_r:
                label = ws.cell(rr, 1).value
                if label is None:
                    break
                if str(label).strip() == "Grand Total":
                    rr += 1
                    break
                gt = ws.cell(rr, gt_col).value if gt_col else None
                data[str(label).strip().upper()] = float(gt or 0)
                rr += 1
            pivots.append({"kategori": kategori, "settle_ke": settle_ke, "rows": data})
            r = rr
            continue
        r += 1

    def pivot_lookup(kategori, settle_ke):
        for p in pivots:
            if p["kategori"] == kategori and p["settle_ke"] == settle_ke:
                return p["rows"]
        return {}

    pdp_at = pivot_lookup("PDP", "AT")
    atbm_at = pivot_lookup("ATBM", "AT")

    table = []
    for u in UIPS:
        t = target.get(u["rekap"], {"target_2026": 0.0, "target_juni": 0.0})
        real_pdp = pdp_at.get(u["st"], 0.0)
        real_atbm = atbm_at.get(u["st"], 0.0)
        real_total = real_pdp + real_atbm
        pct_juni = (real_total / t["target_juni"] * 100) if t["target_juni"] else 0.0
        pct_2026 = (real_total / t["target_2026"] * 100) if t["target_2026"] else 0.0
        table.append({
            "uip": u["rekap"], "target_2026": t["target_2026"], "target_juni": t["target_juni"],
            "realisasi_pdp": real_pdp, "realisasi_atbm": real_atbm, "realisasi_total": real_total,
            "pct_vs_target_juni": round(pct_juni, 1), "pct_vs_target_2026": round(pct_2026, 1),
        })

    # --- tabel bulanan (kolom M-Y, baris TOTAL) untuk tren settlement ---
    h2 = [ws.cell(2, c).value for c in range(13, 26)]
    month_cols = {}
    for i, mv in enumerate(h2):
        if mv:
            month_cols[str(mv).strip().upper()] = 13 + i
    total_row = None
    for row in ws.iter_rows(min_row=4, max_row=30, values_only=False):
        if row[0].value and str(row[0].value).strip().upper() == "TOTAL":
            total_row = row
            break
    monthly_total = {}
    for m in ["JAN", "FEB", "MAR", "APR", "MEI", "JUN"]:
        col = month_cols.get(m)
        v = total_row[col - 1].value if (total_row and col) else 0
        monthly_total[m] = float(v or 0)
    kumulatif = {}
    running = 0.0
    for m in ["JAN", "FEB", "MAR", "APR", "MEI", "JUN"]:
        running += monthly_total[m]
        kumulatif[m] = running

    target_total = {"target_2026": total_target_2026, "target_juni": total_target_juni}
    return table, target_total, monthly_total, kumulatif


def parse_d1_d2(wb, sheet_name, kat):
    ws = wb[sheet_name]
    h = header_row(ws)
    c_unit = find_col(h, "unit")
    c_nama = find_col(h, "nama proyek")
    c_fungsi = find_col(h, "fungsi")
    c_kontrak = find_col(h, "nilai kontrak")  # sengaja tanpa "akhir": lihat catatan di README
    saldo_juni_col = find_col(h, "saldo pdp", "juni")
    c_kat = find_col(h, f"kategori {kat.lower()}")

    out = []
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, values_only=True):
        unit = row[c_unit]
        nama = row[c_nama]
        if not unit or str(unit).strip().upper() == "NIHIL" or not nama:
            continue
        out.append({
            "unit": clean(unit), "nama": text_raw(nama), "fungsi": fungsi_of(row[c_fungsi]),
            "nilai_kontrak_akhir_txt": rawtext(row[c_kontrak]),
            "saldo_juni": numf(row[saldo_juni_col]), "kategori": text_strip(row[c_kat]),
        })
    return out


def parse_d3(wb):
    ws = wb["D3"]
    h = header_row(ws)
    c_unit = find_col(h, "unit")
    c_nama = find_col(h, "nama proyek")
    c_fungsi = find_col(h, "fungsi")
    saldo_juni_col = find_col(h, "saldo pdp", "juni")
    prog_juni_col = find_col(h, "progress fisik", "juni")
    c_durasi = find_col(h, "durasi tidak berprogress")
    c_fase = find_col(h, "fase proyek")
    c_estcod = find_col(h, "estimated cod")
    c_ket = find_col(h, "keterangan", start=c_estcod)

    out = []
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, values_only=True):
        unit = row[c_unit]
        nama = row[c_nama]
        if not unit or str(unit).strip().upper() == "NIHIL" or not nama:
            continue
        out.append({
            "unit": clean(unit), "nama": text_raw(nama), "fungsi": fungsi_of(row[c_fungsi]),
            "saldo_juni": numf(row[saldo_juni_col]), "progress_fisik": frac(row[prog_juni_col]),
            "durasi_bulan": durasi_bulan(row[c_durasi]), "fase": fixcase(clean(row[c_fase])) or "—",
            "estimated_cod": text_strip(row[c_estcod]) or "—", "keterangan": cap(text_strip(row[c_ket]), 260) or "—",
        })
    return out


def parse_d4(wb):
    ws = wb["D4"]
    h = header_row(ws)
    c_unit = find_col(h, "unit")
    c_pulau = find_col(h, "pulau")
    c_rencana = find_col(h, "rencana", "pemanfaatan")
    c_nama = find_col(h, "nama material")
    c_lengkap = find_col(h, "kelengkapan material")
    saldo_juni_col = find_col(h, "saldo pdp", "juni")
    # "Keterangan" akhir (kolom Y) adalah SATU-SATUNYA header yang PERSIS
    # "keterangan" (bukan "Keterangan Mutasi ..."), harus dicari terpisah
    # supaya tidak salah kena kolom "Keterangan Mutasi <bulan>" duluan.
    c_ket = next(i for i, hh in enumerate(h) if hh == "keterangan")

    out = []
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, values_only=True):
        col0 = row[0]
        if isinstance(col0, str) and re.match(r"^[A-Za-z]\s+\S", col0):
            continue  # baris judul sub-kelompok ("A Rencana pemanfaatan ...")
        if not isinstance(col0, (int, float)):
            continue  # baris tanpa nomor urut = bukan baris data valid
        unit, nama = row[c_unit], row[c_nama]
        if not unit or not nama:
            continue
        out.append({
            "unit": clean(unit), "pulau": text_strip(row[c_pulau]),
            "rencana_pemanfaatan": text_strip(row[c_rencana]), "nama_material": text_raw(nama),
            "kelengkapan": row[c_lengkap] if row[c_lengkap] is not None else "—",
            "saldo_juni": numf(row[saldo_juni_col]), "keterangan": cap(text_strip(row[c_ket]), 200),
        })
    return out


def parse_bd(wb):
    ws = wb["Biaya Ditangguhkan"]
    h = header_row(ws)
    c_unit = find_col(h, "unit")
    c_nama = find_col(h, "nama proyek")
    c_uraian = find_col(h, "uraian pekerjaan")
    nilai_juni_col = find_col(h, "nilai", "juni")
    c_status = find_col(h, "status dan kendala")
    c_target = find_col(h, "target proyek terkontrak")

    out = []
    for row in ws.iter_rows(min_row=6, max_row=ws.max_row, values_only=True):
        if not isinstance(row[0], (int, float)):
            continue
        unit, nama = row[c_unit], row[c_nama]
        if not unit or not nama or str(unit).strip() not in BY_KEY:
            continue  # lewati unit yang tidak cocok persis dgn 11 UIP (mis. singkatan "MPA")
        out.append({
            "unit": clean(unit), "nama": text_raw(nama), "uraian": cap(text_strip(row[c_uraian]), 150),
            "nilai_juni": numf(row[nilai_juni_col]), "status": cap(text_strip(row[c_status]), 220),
            "target_cod": text_plain(row[c_target]),
        })
    return out


def parse_clusters(wb):
    ws = wb["Cluster Proyek 1-9"]
    title_re = re.compile(r"^\s*Cluster\s+(\d+)\s*:\s*(.+?)\s*$", re.I)

    # cari semua baris judul cluster (bisa di kolom mana saja)
    titles = []  # (row, full_title)
    for r in range(1, ws.max_row + 1):
        for c in range(1, 8):
            v = ws.cell(r, c).value
            if isinstance(v, str) and title_re.match(v):
                titles.append((r, v.strip()))
                break

    clusters = []
    for i, (r0, title) in enumerate(titles):
        r_end = titles[i + 1][0] - 1 if i + 1 < len(titles) else ws.max_row
        # cari baris header ("NO" di kolom B) di antara r0+1 .. r0+3
        hdr_r = None
        for rr in range(r0 + 1, min(r0 + 4, r_end + 1)):
            if str(ws.cell(rr, 2).value or "").strip().upper() == "NO":
                hdr_r = rr
                break
        if hdr_r is None:
            continue
        projects = []
        for rr in range(hdr_r + 1, r_end + 1):
            no = ws.cell(rr, 2).value
            unit = ws.cell(rr, 3).value
            nama = ws.cell(rr, 4).value
            nilai = ws.cell(rr, 5).value
            progres = ws.cell(rr, 6).value
            status = ws.cell(rr, 7).value
            if nama is not None and str(nama).strip().lower() == "total":
                continue
            if not isinstance(no, (int, float)):
                continue  # baris subjudul ("PDP Berprogress" dsb.) atau kosong
            if not unit or not nama:
                continue
            projects.append({
                "unit": clean(unit), "nama": text_raw(nama), "nilai": numf_id(nilai),
                "progres_pct": pct_num(progres),
                "status": status if isinstance(status, (int, float)) else text_strip(status),
            })
        clusters.append({
            "title": title, "count": len(projects),
            "total": sum(p["nilai"] for p in projects), "projects": projects,
        })
    return clusters


# =============================================================================
# MAIN
# =============================================================================
def main():
    wb = load_workbook()

    print("[build_data] Parsing SAP-DIVAKT ...")
    sap_rows, sap_totals, sap_by_regional, total_sap_uip = parse_sap(wb)
    print(f"    -> {len(sap_rows)} entitas")

    print("[build_data] Parsing Rekap Settlement ...")
    settle_table, settle_target_total, monthly_total, monthly_kumulatif = parse_rekap_settlement(wb)
    print(f"    -> {len(settle_table)} UIP")

    print("[build_data] Parsing D1 ...")
    d1 = parse_d1_d2(wb, "D1", "D1")
    print(f"    -> {len(d1)} proyek")

    print("[build_data] Parsing D2 ...")
    d2 = parse_d1_d2(wb, "D2", "D2")
    print(f"    -> {len(d2)} proyek")

    print("[build_data] Parsing D3 ...")
    d3 = parse_d3(wb)
    print(f"    -> {len(d3)} proyek")

    print("[build_data] Parsing D4 ...")
    d4 = parse_d4(wb)
    print(f"    -> {len(d4)} item")

    print("[build_data] Parsing Biaya Ditangguhkan ...")
    bd_items = parse_bd(wb)
    print(f"    -> {len(bd_items)} item")

    print("[build_data] Parsing Cluster Proyek 1-9 ...")
    clusters = parse_clusters(wb)
    print(f"    -> {len(clusters)} cluster")

    # ---------------- by_category ----------------
    def cat_block(items):
        return {"count": len(items), "value": sum(it["saldo_juni"] for it in items), "items": items}

    d1_block = cat_block(d1)
    d2_block = cat_block(d2)
    d4_block = cat_block(d4)

    d3_block = cat_block(d3)
    fase_dist = {}
    no_progress_buckets = {"0 bulan": 0, "1-2 bulan": 0, "3-6 bulan": 0, ">6 bulan": 0, "n/a": 0}
    trend_fisik = {"0-25%": 0, "25-50%": 0, "50-75%": 0, "75-99%": 0, "100% (Selesai)": 0}
    prog_vals = []
    for it in d3:
        fase_dist[it["fase"]] = fase_dist.get(it["fase"], 0) + 1
        d = it["durasi_bulan"]
        if d is None:
            no_progress_buckets["n/a"] += 1
        elif d == 0:
            no_progress_buckets["0 bulan"] += 1
        elif d <= 2:
            no_progress_buckets["1-2 bulan"] += 1
        elif d <= 6:
            no_progress_buckets["3-6 bulan"] += 1
        else:
            no_progress_buckets[">6 bulan"] += 1
        p = it["progress_fisik"]
        if p is not None:
            prog_vals.append(p)
            pp = p * 100
            if pp >= 100:
                trend_fisik["100% (Selesai)"] += 1
            elif pp >= 75:
                trend_fisik["75-99%"] += 1
            elif pp >= 50:
                trend_fisik["50-75%"] += 1
            elif pp >= 25:
                trend_fisik["25-50%"] += 1
            else:
                trend_fisik["0-25%"] += 1
    d3_block["fase_dist"] = fase_dist
    d3_block["no_progress_buckets"] = no_progress_buckets
    d3_block["trend_fisik"] = trend_fisik
    d3_block["progress_fisik_avg"] = round(sum(prog_vals) / len(prog_vals) * 100, 1) if prog_vals else 0.0

    by_category = {"D1": d1_block, "D2": d2_block, "D3": d3_block, "D4": d4_block}

    # ---------------- fungsi_dist (dari D3 saja) ----------------
    fungsi_dist = {}
    for it in d3:
        f = it["fungsi"]
        if f not in ("KIT", "TL", "GI"):
            continue
        g = fungsi_dist.setdefault(f, {"count": 0, "value": 0.0})
        g["count"] += 1
        g["value"] += it["saldo_juni"]

    # ---------------- unit_map ----------------
    agg = {u["key"]: {"d1": 0.0, "d2": 0.0, "d3": 0.0, "d4": 0.0, "jumlah_proyek_d3": 0} for u in UIPS}
    for it in d1:
        if it["unit"] in agg:
            agg[it["unit"]]["d1"] += it["saldo_juni"]
    for it in d2:
        if it["unit"] in agg:
            agg[it["unit"]]["d2"] += it["saldo_juni"]
    for it in d3:
        if it["unit"] in agg:
            agg[it["unit"]]["d3"] += it["saldo_juni"]
            agg[it["unit"]]["jumlah_proyek_d3"] += 1
    for it in d4:
        if it["unit"] in agg:
            agg[it["unit"]]["d4"] += it["saldo_juni"]

    map_path = ""
    if os.path.exists(MAP_PATH_FILE):
        with open(MAP_PATH_FILE, encoding="utf-8") as f:
            map_path = f.read().strip()

    unit_map = []
    for u in UIPS:
        a = agg[u["key"]]
        total = a["d1"] + a["d2"] + a["d3"] + a["d4"]
        unit_map.append({
            "unit": u["key"], "label": u["label"], "pos": u["pos"],
            "d1": a["d1"], "d2": a["d2"], "d3": a["d3"], "d4": a["d4"], "total": total,
            "jumlah_proyek_d3": a["jumlah_proyek_d3"],
        })

    # ---------------- biaya_ditangguhkan ----------------
    bd_by_unit = {}
    for it in bd_items:
        g = bd_by_unit.setdefault(it["unit"], {"count": 0, "nilai": 0.0})
        g["count"] += 1
        g["nilai"] += it["nilai_juni"]
    bd_top = sorted(bd_items, key=lambda x: -x["nilai_juni"])[:25]
    biaya_ditangguhkan = {
        "by_unit": bd_by_unit, "top": bd_top,
        "total": sum(it["nilai_juni"] for it in bd_items), "count": len(bd_items),
    }

    # ---------------- ews ----------------
    crit_all = [it for it in d3 if (it["durasi_bulan"] or 0) >= 3]
    critical_projects = sorted(crit_all, key=lambda x: -x["saldo_juni"])[:40]
    no_progress_3plus = no_progress_buckets["3-6 bulan"] + no_progress_buckets[">6 bulan"]
    ews = {
        "no_progress_buckets": no_progress_buckets, "fase_dist": fase_dist,
        "critical_projects": critical_projects,
    }

    # ---------------- kpi ----------------
    total_pmo = d1_block["value"] + d2_block["value"] + d3_block["value"] + d4_block["value"]
    total_sap = sap_totals["total_pdp"]
    selisih = abs(total_pmo - total_sap_uip)
    cluster_total_value = sum(c["total"] for c in clusters)
    cluster_total_count = sum(c["count"] for c in clusters)
    settlement_realisasi_total = sum(r["realisasi_total"] for r in settle_table)
    biaya_ditangguhkan_total = biaya_ditangguhkan["total"]

    kpi = {
        "total_pmo": total_pmo, "total_sap": total_sap, "total_sap_uip": total_sap_uip,
        "selisih_pmo_sap_uip": selisih,
        "selisih_pct": (selisih / total_sap_uip * 100) if total_sap_uip else 0.0,
        "jumlah_proyek_d1": d1_block["count"], "jumlah_proyek_d2": d2_block["count"],
        "jumlah_proyek_d3": d3_block["count"], "jumlah_item_d4": d4_block["count"],
        "avg_progress_fisik": d3_block["progress_fisik_avg"],
        "cluster_total_value": cluster_total_value, "cluster_total_count": cluster_total_count,
        "settlement_realisasi_total": settlement_realisasi_total,
        "settlement_target_2026_total": settle_target_total["target_2026"],
        "settlement_target_juni_total": settle_target_total["target_juni"],
        "biaya_ditangguhkan_total": biaya_ditangguhkan_total,
        "no_progress_3plus": no_progress_3plus,
    }

    data = {
        "map": {"viewbox": MAP_VIEWBOX, "path": map_path},
        "meta": {
            "title": f"Monitoring PDP sampai {PERIODE_LABEL} DIV SAK",
            "period": PERIODE_LABEL,
            "generated": datetime.datetime.now().strftime("%d %B %Y"),
        },
        "kpi": kpi,
        "by_category": by_category,
        "unit_map": unit_map,
        "fungsi_dist": fungsi_dist,
        "clusters": clusters,
        "sap_divakt": {"rows": sap_rows, "totals": sap_totals, "by_regional": sap_by_regional},
        "settlement": {
            "table": settle_table, "target_total": settle_target_total,
            "monthly_total": monthly_total, "monthly_kumulatif": monthly_kumulatif,
        },
        "biaya_ditangguhkan": biaya_ditangguhkan,
        "ews": ews,
    }

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))

    print(f"\n[build_data] Selesai. Ditulis ke: {OUT_PATH}")
    print(f"[build_data] Total proyek D1+D2+D3: {d1_block['count']+d2_block['count']+d3_block['count']}, D4: {d4_block['count']}")


if __name__ == "__main__":
    main()
