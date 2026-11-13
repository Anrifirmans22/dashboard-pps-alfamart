# =============================================================
#  Dashboard Eksekutif Rekapitulasi PPS Area Alfamart
#  File  : app.py
#  Jalan : streamlit run app.py
# =============================================================
import html
import io
import os
import re
import time

import pandas as pd
import streamlit as st

SEMUA = "Semua"
MODEL_DEFAULT = "gemini-2.5-flash"
DAFTAR_MODEL = ["gemini-2.5-flash", "gemini-2.5-flash-lite", "gemini-2.5-pro",
                "gemini-2.0-flash", "gemini-2.0-flash-lite"]
KOLOM_ID = {"KD TOKO", "KODE TOKO", "NAMA TOKO", "AC", "AM",
            "AREA MANAGER", "AREA COORDINATOR"}
WARNA_XLSX = {"id": "FFFF00", "orange": "FFC000", "black": "000000",
              "flop": "F1F5F9", "gray": "D9D9D9", "white": "FFFFFF"}


def _versi():
    try:
        return tuple(int(x) for x in st.__version__.split(".")[:2])
    except Exception:
        return (0, 0)


# parameter lebar penuh yang kompatibel dengan versi Streamlit lama & baru
LEBAR = {"width": "stretch"} if _versi() >= (1, 50) else {"use_container_width": True}

CSS = """
<style>
    .block-container { max-width: 1280px; margin: 0 auto; padding: 2rem 2rem 3rem 2rem; }
    .judul-utama { text-align:center; font-size:2.1rem; font-weight:800; color:#0B3C8C; margin-bottom:.2rem; }
    .sub-judul { text-align:center; font-size:1rem; color:#6b7280; margin-bottom:1rem; }
    .garis { height:4px; width:90px; margin:0 auto 1.8rem auto; border-radius:4px;
             background:linear-gradient(90deg,#E30613,#0B3C8C); }
    div[data-testid="stMetric"] { background:#fff; border:1px solid #e5e7eb; border-left:5px solid #E30613;
        border-radius:12px; padding:14px 18px; text-align:center; box-shadow:0 2px 8px rgba(0,0,0,.06); }
    div[data-testid="stMetricLabel"] { justify-content:center; }
    div[data-testid="stMetricLabel"] p { color:#4b5563 !important; }
    div[data-testid="stMetricValue"] { justify-content:center; color:#0B3C8C; font-weight:800; font-size:1.4rem; }
    .stTabs [data-baseweb="tab-list"] { justify-content:center; gap:12px; }
    .stTabs [data-baseweb="tab"] { font-weight:600; padding:10px 26px; }
    .catatan { text-align:center; color:#6b7280; font-size:.92rem; margin:.4rem 0 1rem 0; }

    .tabel-wrap { max-height:560px; overflow:auto; border:1px solid #cbd5e1; border-radius:10px; background:#fff; }
    table.pps { border-collapse:separate; border-spacing:0; font-size:12px; white-space:nowrap;
                width:max-content; min-width:100%; }
    table.pps th, table.pps td { border-right:1px solid #cbd5e1; border-bottom:1px solid #e2e8f0;
                padding:4px 9px; text-align:center; box-sizing:border-box; }
    table.pps td { background:#fff; color:#111; }
    table.pps tbody tr:nth-child(even) td { background:#f8fafc; }
    table.pps td.kiri { text-align:left; }
    table.pps td.neg { color:#E30613; font-weight:600; }
    table.pps td.tebal { font-weight:800; }
    table.pps td.b-top { background:#66FF66 !important; font-weight:800; }
    table.pps td.b-flop { background:#FFA500 !important; font-weight:800; }
    table.pps thead th { position:sticky; height:34px; font-weight:700; color:#111; z-index:2; }
    table.pps thead tr.r1 th { top:0; }
    table.pps thead tr.r2 th { top:34px; }
    .th-id { background:#FFFF00; } .th-orange { background:#FFC000; }
    .th-black { background:#000; color:#fff !important; } .th-flop { background:#f1f5f9; }
    .th-gray { background:#D9D9D9; } .th-white, .th-putih { background:#fff; }
</style>
"""


# ---------------------------------------------------------------
# FUNGSI BANTU UMUM
# ---------------------------------------------------------------
def norm(x):
    try:
        if pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return " ".join(str(x).split())


def kunci(x):
    return norm(x).upper()


def cari_kolom(meta, *nama):
    cari = {n.upper() for n in nama}
    for flat, m in meta.items():
        if not m["group"] and m["sub"].upper() in cari:
            return flat
    return None


def adalah_pct(sub_up):
    if sub_up.startswith("JT"):
        return False
    return "%" in sub_up or "ACH TO TF" in sub_up or sub_up == "TOTAL ACH"


def tentukan_warna(sub_up, group):
    if sub_up.startswith("GAP") or "ACH TO TF" in sub_up:
        return "orange"
    if sub_up == "TOTAL ACH":
        return "black"
    if sub_up == "FLOP":
        return "flop"
    if sub_up in KOLOM_ID or not group:
        return "id"
    return "white" if group.upper().startswith("APC") else "gray"


def fmt(v, m):
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if m["tipe"] == "pct":
        return f"{float(v) * 100:.{m['des']}f}%"
    if m["tipe"] == "angka":
        return f"{float(v):,.0f}" if m["bulat"] else f"{float(v):,.1f}"
    return str(v)


# ---------------------------------------------------------------
# MEMBACA EXCEL
# ---------------------------------------------------------------
def indeks_sheet_default(sheets):
    for i, s in enumerate(sheets):
        n = s.lower().replace(" ", "")
        if "alltoko" in n or "semuatoko" in n:
            return i
    for i, s in enumerate(sheets):
        if "toko" in s.lower() and "project" not in s.lower():
            return i
    return 0


@st.cache_data(show_spinner=False)
def daftar_sheet(file_id, _isi):
    return pd.ExcelFile(io.BytesIO(_isi)).sheet_names


def cari_baris_header(mentah, maks=40):
    """Baris header = baris yang punya sel 'Nama Toko' atau 'AC'."""
    for i in range(min(maks, len(mentah))):
        sel = {kunci(x) for x in mentah.iloc[i].tolist()}
        if "NAMA TOKO" in sel or "AC" in sel:
            return i
    return 0


def ambil_info(mentah, idx_h):
    periode, tf = "", []
    for i in range(idx_h):
        for x in mentah.iloc[i].tolist():
            t = norm(x)
            if t.upper().startswith("PERIODE") and not periode:
                periode = t
            elif t.upper().startswith("TF "):
                tf.append(t)
    return {"periode": periode, "tf": tf}


def bangun_tabel(mentah):
    n_kolom = mentah.shape[1]
    idx_h = cari_baris_header(mentah)
    hdr = [norm(x) for x in mentah.iloc[idx_h].tolist()]
    info = ambil_info(mentah, idx_h)

    # baris grup (APC, PSM TOTAL, SERBA GRATIS, PWP ...) tepat di atas header
    grp_raw = [""] * n_kolom
    if idx_h > 0:
        cand = [norm(x) for x in mentah.iloc[idx_h - 1].tolist()]
        valid = [c for c in cand if c and not c.upper().startswith(("PERIODE", "TF "))]
        if len(valid) >= 2:
            grp_raw = cand

    grp, aktif = [], ""
    for j in range(n_kolom):
        g = grp_raw[j]
        if g.lower().startswith("sortby"):
            aktif, g = "", ""
        elif g.upper().startswith(("PERIODE", "TF ")):
            g = ""
        elif g:
            aktif = g
        sub_up = hdr[j].upper()
        grp.append("" if (sub_up in KOLOM_ID or sub_up in ("TOTAL ACH", "FLOP")) else aktif)

    j0 = next((j for j, h in enumerate(hdr) if h.upper() in KOLOM_ID), 0)
    isi = mentah.iloc[idx_h + 1:, :]

    urut, meta, kolom_data = [], {}, {}
    for j in range(j0, n_kolom):
        sub = hdr[j]
        if not sub:
            continue
        flat = f"{grp[j]} | {sub}" if grp[j] else sub
        dasar, n = flat, 2
        while flat in meta:
            flat, n = f"{dasar} ({n})", n + 1
        sub_up = sub.upper()
        meta[flat] = {"group": grp[j], "sub": sub, "warna": tentukan_warna(sub_up, grp[j]),
                      "tipe": "teks", "des": 0, "bulat": True, "neg": sub_up.startswith("GAP")}
        kolom_data[flat] = isi.iloc[:, j].values
        urut.append(flat)

    data = pd.DataFrame(kolom_data, columns=urut)

    # buang baris kosong, baris filter, baris total
    kol_nama = cari_kolom(meta, "NAMA TOKO")
    kol_kunci = kol_nama or cari_kolom(meta, "AC") or urut[0]
    ser = data[kol_kunci].map(norm)
    mask = (ser != "") & ~ser.str.upper().isin(["NAMA TOKO", "AC", "KD TOKO"])
    if kol_nama:
        mask &= ~ser.str.upper().str.startswith(("TOTAL", "GRAND TOTAL", "JUMLAH"))
    data = data[mask.values].reset_index(drop=True)

    # tipe data per kolom
    for f in urut:
        m = meta[f]
        sub_up = m["sub"].upper()
        if sub_up in KOLOM_ID:
            data[f] = data[f].map(norm)
            continue
        try:
            num = pd.to_numeric(data[f], errors="coerce")
        except Exception:
            num = pd.Series([float("nan")] * len(data))
        ada = int(data[f].notna().sum())
        if ada and int(num.notna().sum()) >= 0.5 * ada:
            if adalah_pct(sub_up):
                m["tipe"] = "pct"
                m["des"] = 2 if (sub_up in ("ACH%", "TOTAL ACH") or "ACH TO TF" in sub_up) else 1
                v = num.dropna().abs()
                if len(v) and v.quantile(0.95) > 20:   # data sudah berupa persen, ubah ke pecahan
                    num = num / 100
            else:
                m["tipe"] = "angka"
                v = num.dropna()
                m["bulat"] = bool(((v - v.round()).abs() < 1e-6).all())
            data[f] = num
        else:
            data[f] = data[f].map(norm)

    # buang kolom yang kosong seluruhnya (mis. PSM W2-W4 yang datanya belum ada)
    if len(data):
        kosong = [f for f in urut
                  if meta[f]["sub"].upper() not in KOLOM_ID
                  and data[f].map(lambda v: norm(v) == "").all()]
        if kosong:
            data = data.drop(columns=kosong)
            meta = {k: v for k, v in meta.items() if k not in kosong}
    return data, meta, info


@st.cache_data(show_spinner="Membaca sheet Excel...")
def muat_sheet(file_id, nama_sheet, _isi):
    mentah = pd.read_excel(io.BytesIO(_isi), sheet_name=nama_sheet, header=None)
    return bangun_tabel(mentah)


@st.cache_data(show_spinner="Mencari data AM di sheet lain...")
def cari_peta_am(file_id, sheet_utama, _isi):
    """Cari sheet lain yang punya kolom AC dan AM -> peta {AC: AM}."""
    try:
        xl = pd.ExcelFile(io.BytesIO(_isi))
    except Exception:
        return {}
    for sh in xl.sheet_names:
        if sh == sheet_utama:
            continue
        try:
            cuplik = xl.parse(sh, header=None, nrows=60)
            idx = None
            for i in range(len(cuplik)):
                sel = [kunci(x) for x in cuplik.iloc[i].tolist()]
                if "AC" in sel and ("AM" in sel or "AREA MANAGER" in sel):
                    idx = i
                    break
            if idx is None:
                continue
            penuh = xl.parse(sh, header=None)
            hdr = [kunci(x) for x in penuh.iloc[idx].tolist()]
            j_ac = hdr.index("AC")
            j_am = hdr.index("AM") if "AM" in hdr else hdr.index("AREA MANAGER")
            hitung = {}
            for ac, am in zip(penuh.iloc[idx + 1:, j_ac], penuh.iloc[idx + 1:, j_am]):
                k, a = kunci(ac), norm(am)
                if k and a and k != "AC":
                    hitung.setdefault(k, {})
                    hitung[k][a] = hitung[k].get(a, 0) + 1
            peta = {k: max(v, key=v.get) for k, v in hitung.items()}
            if peta:
                return peta
        except Exception:
            continue
    return {}


def parse_manual(teks):
    """Format per baris:  NAMA AM = AC 1, AC 2"""
    peta = {}
    for baris in (teks or "").splitlines():
        if "=" not in baris:
            continue
        am, acs = baris.split("=", 1)
        for ac in acs.split(","):
            if kunci(ac) and norm(am):
                peta[kunci(ac)] = norm(am)
    return peta


def pasang_am(df, meta, peta):
    kol_ac = cari_kolom(meta, "AC", "AREA COORDINATOR")
    if not kol_ac or not peta:
        return df, meta
    import difflib

    def cari_am(x):
        k = kunci(x)
        if k in peta:
            return peta[k]
        mirip = difflib.get_close_matches(k, list(peta.keys()), n=1, cutoff=0.85)
        return peta[mirip[0]] if mirip else ""

    nilai = df[kol_ac].map(cari_am)
    if (nilai == "").all():
        return df, meta
    df = df.copy()
    df.insert(list(df.columns).index(kol_ac) + 1, "AM", nilai.values)
    meta_baru = {}
    for k in df.columns:
        meta_baru[k] = meta.get(k) or {"group": "", "sub": "AM", "warna": "id", "tipe": "teks",
                                       "des": 0, "bulat": True, "neg": False}
    return df, meta_baru


def hitung_badge(df, meta):
    """Urutkan per AC berdasarkan TOTAL ACH, lalu tandai TOP (4 teratas) & FLOP (3 terbawah)."""
    kol_total = cari_kolom(meta, "TOTAL ACH")
    if not kol_total or df.empty or not df[kol_total].notna().any():
        return df, None
    kol_ac = cari_kolom(meta, "AC", "AREA COORDINATOR")
    urut = [kol_ac, kol_total] if kol_ac else [kol_total]
    asc = [True, False] if kol_ac else [False]
    df = df.sort_values(urut, ascending=asc, na_position="last", kind="stable").reset_index(drop=True)
    badge = [""] * len(df)
    kelompok = list(df.groupby(kol_ac, sort=False).indices.values()) if kol_ac else [list(range(len(df)))]
    for idx in kelompok:
        idx = [i for i in idx if pd.notna(df.at[i, kol_total])]
        n = len(idx)
        if n < 4:
            continue
        jt, jf = min(4, n // 2), min(3, n // 2)
        for i in idx[:jt]:
            badge[i] = "TOP"
        for i in idx[-jf:]:
            badge[i] = "FLOP"
    return df, badge


# ---------------------------------------------------------------
# TAMPILAN TABEL (HTML ala laporan) & EXCEL
# ---------------------------------------------------------------
def tabel_html(df, meta, badge=None, maks_baris=600):
    cols = list(df.columns)
    ada_grup = any(meta[c]["group"] for c in cols)
    rs = ' rowspan="2"' if ada_grup else ""
    r1, r2 = [], []
    if badge is not None:
        r1.append(f'<th class="th-putih"{rs}></th>')
    i = 0
    while i < len(cols):
        m = meta[cols[i]]
        if m["group"]:
            j = i
            while (j + 1 < len(cols) and meta[cols[j + 1]]["group"] == m["group"]
                   and meta[cols[j + 1]]["warna"] == m["warna"]):
                j += 1
            r1.append(f'<th class="th-{m["warna"]}" colspan="{j - i + 1}">{html.escape(m["group"])}</th>')
            for k in range(i, j + 1):
                mk = meta[cols[k]]
                r2.append(f'<th class="th-{mk["warna"]}">{html.escape(mk["sub"])}</th>')
            i = j + 1
        else:
            r1.append(f'<th class="th-{m["warna"]}"{rs}>{html.escape(m["sub"])}</th>')
            i += 1

    h = ['<div class="tabel-wrap"><table class="pps"><thead>',
         '<tr class="r1">' + "".join(r1) + "</tr>"]
    if ada_grup:
        h.append('<tr class="r2">' + "".join(r2) + "</tr>")
    h.append("</thead><tbody>")

    nilai = df.head(maks_baris).to_numpy(dtype=object)
    for ri, baris in enumerate(nilai):
        sel = []
        if badge is not None:
            b = badge[ri]
            sel.append(f'<td class="b-{b.lower()}">{b}</td>' if b else "<td></td>")
        for c, v in zip(cols, baris):
            m = meta[c]
            kls = []
            if m["tipe"] == "teks":
                kls.append("kiri")
            elif m["neg"] and pd.notna(v) and float(v) < 0:
                kls.append("neg")
            if m["warna"] == "black":
                kls.append("tebal")
            sel.append(f'<td class="{" ".join(kls)}">{html.escape(fmt(v, m))}</td>')
        h.append("<tr>" + "".join(sel) + "</tr>")
    h.append("</tbody></table></div>")
    return "".join(h)      # satu baris, tanpa baris kosong (agar aman di st.markdown)


def ke_excel(df, meta, badge=None):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Hasil Filter"
    cols = list(df.columns)
    ada_grup = any(meta[c]["group"] for c in cols)
    off = 1 if badge is not None else 0
    r_sub = 2 if ada_grup else 1
    r_data = r_sub + 1
    tipis = Side(style="thin", color="BBBBBB")
    garis = Border(left=tipis, right=tipis, top=tipis, bottom=tipis)
    tengah = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def gaya_header(sel, warna):
        sel.fill = PatternFill("solid", fgColor=WARNA_XLSX.get(warna, "FFFFFF"))
        sel.font = Font(bold=True, color="FFFFFF" if warna == "black" else "000000")
        sel.alignment = tengah
        sel.border = garis

    if off:
        ws.cell(1, 1, "TOP/FLOP")
        gaya_header(ws.cell(1, 1), "white")
        if ada_grup:
            ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)

    i = 0
    while i < len(cols):
        m = meta[cols[i]]
        k = i + 1 + off
        if m["group"]:
            j = i
            while (j + 1 < len(cols) and meta[cols[j + 1]]["group"] == m["group"]
                   and meta[cols[j + 1]]["warna"] == m["warna"]):
                j += 1
            ws.cell(1, k, m["group"])
            for kk in range(i, j + 1):
                gaya_header(ws.cell(1, kk + 1 + off), m["warna"])
                mk = meta[cols[kk]]
                gaya_header(ws.cell(2, kk + 1 + off, mk["sub"]), mk["warna"])
            if j > i:
                ws.merge_cells(start_row=1, start_column=k, end_row=1, end_column=j + 1 + off)
            i = j + 1
        else:
            ws.cell(1, k, m["sub"])
            gaya_header(ws.cell(1, k), m["warna"])
            if ada_grup:
                ws.merge_cells(start_row=1, start_column=k, end_row=2, end_column=k)
            i += 1

    for r, idx in enumerate(range(len(df)), start=r_data):
        if off:
            b = badge[idx]
            sel = ws.cell(r, 1, b or None)
            sel.alignment, sel.border = tengah, garis
            if b:
                sel.fill = PatternFill("solid", fgColor="66FF66" if b == "TOP" else "FFA500")
                sel.font = Font(bold=True)
        for n, c in enumerate(cols):
            m = meta[c]
            v = df.iat[idx, n]
            sel = ws.cell(r, n + 1 + off)
            sel.border = garis
            if m["tipe"] == "teks":
                sel.value = v if v != "" else None
                sel.alignment = Alignment(horizontal="left")
                continue
            sel.alignment = Alignment(horizontal="center")
            if pd.isna(v):
                continue
            sel.value = float(v)
            if m["tipe"] == "pct":
                sel.number_format = "0.00%" if m["des"] == 2 else "0.0%"
            else:
                sel.number_format = "#,##0" if m["bulat"] else "#,##0.0"
            if m["neg"] and float(v) < 0:
                sel.font = Font(color="E30613")

    for n, c in enumerate(cols):
        sub = meta[c]["sub"]
        lebar = 28 if sub.upper() == "NAMA TOKO" else max(10, min(len(sub) + 4, 18))
        ws.column_dimensions[get_column_letter(n + 1 + off)].width = lebar
    if off:
        ws.column_dimensions["A"].width = 10
    ws.freeze_panes = f"A{r_data}"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ---------------------------------------------------------------
# GEMINI AI AGENT
# ---------------------------------------------------------------
INSTRUKSI_AI = (
    "Kamu adalah analis performa toko Alfamart yang teliti. Jawab dalam Bahasa Indonesia yang "
    "ringkas, jelas, dan terstruktur (boleh pakai poin-poin). Gunakan HANYA data yang diberikan. "
    "Nilai pada tabel sudah dalam format tampilan (persen sudah berupa %). Kolom bernama "
    "'GROUP | SUB' berarti SUB milik kelompok GROUP (contoh: 'APC (25%) | ACH%'). Sebut nama toko "
    "dan angka yang relevan. Jika data tidak cukup untuk menjawab, katakan dengan jujur."
)


def bangun_konteks(df, meta, info, info_filter, nama_sheet, maks=250):
    ringkas = [f"Sheet: {nama_sheet}", f"Filter aktif: {info_filter}", f"Jumlah toko: {len(df)}"]
    if info.get("periode"):
        ringkas.append(info["periode"])
    if info.get("tf"):
        ringkas.append("Target TF: " + "; ".join(info["tf"]))
    kol_nama = cari_kolom(meta, "NAMA TOKO")
    kol_total = cari_kolom(meta, "TOTAL ACH")
    if kol_nama and kol_total:
        s = df[[kol_nama, kol_total]].dropna(subset=[kol_total]).sort_values(kol_total, ascending=False)
        if len(s):
            def teks(r):
                return f"{r[kol_nama]} ({r[kol_total] * 100:.2f}%)"
            ringkas.append(f"TOTAL ACH rata-rata: {s[kol_total].mean() * 100:.2f}%")
            ringkas.append("5 TOTAL ACH tertinggi: " + "; ".join(teks(r) for _, r in s.head(5).iterrows()))
            ringkas.append("5 TOTAL ACH terendah: " + "; ".join(teks(r) for _, r in s.tail(5).iloc[::-1].iterrows()))
    tampil = df.head(maks)
    fm = pd.DataFrame({c: [fmt(v, meta[c]) for v in tampil[c]] for c in tampil.columns})
    csv = fm.to_csv(index=False)[:120000]
    if len(df) > maks:
        ringkas.append(f"(Hanya {maks} baris pertama dari {len(df)} yang dikirim.)")
    return "RINGKASAN:\n" + "\n".join(ringkas) + "\n\nDATA TOKO (CSV):\n" + csv


def ambil_kunci_tersimpan():
    """API Key permanen: dari .streamlit/secrets.toml (GEMINI_API_KEY) atau environment variable."""
    try:
        k = st.secrets["GEMINI_API_KEY"]
        if k:
            return str(k).strip()
    except Exception:
        pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def muat_daftar_model(api_key):
    """Ambil daftar model Gemini yang benar-benar tersedia untuk API Key ini."""
    from google import genai

    klien = genai.Client(api_key=api_key.strip())
    buang = ("tts", "image", "embedding", "live", "audio", "robotics", "computer-use")
    hasil = []
    for m in klien.models.list():
        nama = (m.name or "").replace("models/", "")
        aksi = getattr(m, "supported_actions", None) or []
        if not nama.startswith("gemini") or any(b in nama for b in buang):
            continue
        if aksi and "generateContent" not in aksi:
            continue
        hasil.append(nama)
    return sorted(set(hasil))


def panggil_gemini(api_key, isi, instruksi, model):
    from google import genai
    from google.genai import types

    klien = genai.Client(api_key=api_key.strip())
    for percobaan in range(2):
        try:
            resp = klien.models.generate_content(
                model=model,
                contents=isi,
                config=types.GenerateContentConfig(system_instruction=instruksi, temperature=0.3),
            )
            teks = (resp.text or "").strip()
            if teks:
                return teks
            raise RuntimeError("Gemini tidak mengembalikan teks (mungkin diblokir filter keamanan).")
        except Exception as e:
            m = str(e)
            sibuk = any(k in m for k in ("429", "500", "503", "UNAVAILABLE", "RESOURCE_EXHAUSTED", "overloaded"))
            if sibuk and percobaan < 1:
                time.sleep(3)
                continue
            raise


def jalankan_ai(api_key, isi, instruksi, model, otomatis):
    """Coba model pilihan; jika gagal dan 'otomatis' aktif, coba model cadangan satu per satu."""
    urutan = [model] + ([m for m in DAFTAR_MODEL if m != model] if otomatis else [])
    galat = None
    for m in urutan:
        try:
            return panggil_gemini(api_key, isi, instruksi, m), m
        except Exception as e:
            galat = e
            s = str(e)
            if isinstance(e, ImportError) or "API_KEY_INVALID" in s or "API key not valid" in s or "401" in s:
                raise
    raise galat


def pesan_ramah(e, model=MODEL_DEFAULT):
    m = str(e)
    if "API_KEY_INVALID" in m or "API key not valid" in m or "401" in m:
        return "API Key tidak valid. Salin ulang dari Google AI Studio (tanpa spasi di awal/akhir)."
    if "PERMISSION_DENIED" in m or "403" in m:
        return "API Key ditolak (izin/wilayah/akses API). Coba buat API Key baru di Google AI Studio."
    if "429" in m or "RESOURCE_EXHAUSTED" in m:
        return "Kuota atau batas permintaan Gemini habis. Tunggu sekitar 1 menit lalu coba lagi."
    if "503" in m or "UNAVAILABLE" in m or "overloaded" in m:
        return "Server Gemini sedang sibuk. Coba lagi beberapa saat."
    if "404" in m or "NOT_FOUND" in m:
        return f"Model {model} tidak tersedia untuk API Key ini. Pilih model lain di sidebar."
    if isinstance(e, ImportError):
        return "Pustaka google-genai belum terpasang. Jalankan: pip install google-genai"
    return "Terjadi kesalahan saat menghubungi Gemini."


# ---------------------------------------------------------------
# APLIKASI
# ---------------------------------------------------------------
def opsi(seri):
    return [SEMUA] + sorted({v for v in seri.astype(str).str.strip() if v and v.lower() != "nan"})


def main():
    st.set_page_config(page_title="Dashboard PPS Area Alfamart", page_icon="📊",
                       layout="wide", initial_sidebar_state="expanded")
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('<div class="judul-utama">📊 Dashboard Eksekutif PPS Area Alfamart</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-judul">Rekapitulasi performa toko berdasarkan Area Manager &amp; Area Coordinator</div>',
                unsafe_allow_html=True)
    st.markdown('<div class="garis"></div>', unsafe_allow_html=True)

    # ---- Sidebar: upload, sheet, API key
    with st.sidebar:
        st.header("⚙️ Panel Kontrol")
        file_excel = st.file_uploader("Upload file Excel PPS", type=["xlsx", "xls"])
        file_id = isi = pilih_sheet = None
        if file_excel is not None:
            isi = file_excel.getvalue()
            file_id = f"{file_excel.name}-{file_excel.size}"
            try:
                daftar = daftar_sheet(file_id, isi)
            except Exception as e:
                st.error(f"Gagal membuka file: {e}")
                st.stop()
            pilih_sheet = st.selectbox("Sheet data", daftar, index=indeks_sheet_default(daftar),
                                       key=f"sheet_{file_id}")
        st.divider()
        kunci_simpan = ambil_kunci_tersimpan()
        input_key = st.text_input("🔑 Gemini API Key", type="password",
                                  placeholder="Kosongkan jika sudah disimpan permanen" if kunci_simpan
                                  else "Tempel API Key di sini")
        api_key = input_key.strip() or kunci_simpan
        if kunci_simpan and not input_key.strip():
            st.caption("✅ API Key permanen terdeteksi, tidak perlu ditempel lagi.")
        else:
            st.caption("API Key yang ditempel di sini hanya berlaku selama sesi ini.")

        daftar_model = st.session_state.get("daftar_model", DAFTAR_MODEL)
        if MODEL_DEFAULT not in daftar_model:
            daftar_model = [MODEL_DEFAULT] + daftar_model
        model_pilih = st.selectbox("🧠 Model Gemini", daftar_model,
                                   index=daftar_model.index(MODEL_DEFAULT), key="model_gemini")
        otomatis = st.checkbox("Otomatis coba model lain jika gagal", value=True, key="oto_model")
        if st.button("🔄 Muat daftar model dari akun saya", **LEBAR):
            if not api_key:
                st.warning("Isi API Key dulu.")
            else:
                try:
                    st.session_state.daftar_model = muat_daftar_model(api_key)
                    st.rerun()
                except Exception as e:
                    st.error(pesan_ramah(e, model_pilih))

    if file_excel is None:
        st.info("👈 Silakan upload file Excel PPS melalui panel kontrol di sebelah kiri.")
        st.stop()

    try:
        df0, meta0, info = muat_sheet(file_id, pilih_sheet, isi)
    except Exception as e:
        st.error(f"Gagal membaca sheet '{pilih_sheet}': {e}")
        st.stop()
    if df0.empty:
        st.warning("Sheet terbaca, tetapi tidak ada data yang bisa ditampilkan.")
        st.stop()

    # ---- Sidebar: filter AM & AC
    dasar = f"{file_id}_{pilih_sheet}"
    with st.sidebar:
        st.divider()
        st.subheader("🔎 Filter Data")
        df, meta = df0, meta0
        if not cari_kolom(meta0, "AM", "AREA MANAGER"):
            peta = cari_peta_am(file_id, pilih_sheet, isi)
            if not peta:
                with st.expander("Pemetaan AM manual (AM tidak ada di file)"):
                    teks = st.text_area("Satu baris per AM: NAMA AM = AC 1, AC 2", height=120,
                                        key=f"manual_{dasar}")
                peta = parse_manual(teks)
            df, meta = pasang_am(df0, meta0, peta)

        kol_am = cari_kolom(meta, "AM", "AREA MANAGER")
        kol_ac = cari_kolom(meta, "AC", "AREA COORDINATOR")

        am_pilih = SEMUA
        if kol_am:
            am_pilih = st.selectbox("Area Manager (AM)", opsi(df[kol_am]), key=f"am_{dasar}")
        else:
            st.selectbox("Area Manager (AM)", [SEMUA], disabled=True, key=f"am_kosong_{dasar}")
            st.caption("Data AM tidak ditemukan di file ini.")
        df_am = df if (am_pilih == SEMUA or not kol_am) else df[df[kol_am] == am_pilih]

        ac_pilih = SEMUA
        if kol_ac:
            ac_pilih = st.selectbox("Area Coordinator (AC)", opsi(df_am[kol_ac]), key=f"ac_{dasar}_{am_pilih}")
        else:
            st.selectbox("Area Coordinator (AC)", [SEMUA], disabled=True, key=f"ac_kosong_{dasar}")
            st.caption("Kolom AC tidak ditemukan di sheet ini.")

    df_filter = df_am if (ac_pilih == SEMUA or not kol_ac) else df_am[df_am[kol_ac] == ac_pilih]
    info_filter = f"AM = {am_pilih}; AC = {ac_pilih}"
    df_tampil, badge = hitung_badge(df_filter, meta)

    # ---- Kartu ringkasan
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Total Toko Tersaring", f"{len(df_filter):,}".replace(",", "."))
    m2.metric("Total Seluruh Toko", f"{len(df):,}".replace(",", "."))
    m3.metric("Area Manager", am_pilih)
    m4.metric("Area Coordinator", ac_pilih)

    bagian = [f"Sheet: <b>{html.escape(pilih_sheet)}</b>"]
    if info["periode"]:
        bagian.append(html.escape(info["periode"]))
    if info["tf"]:
        bagian.append(html.escape(" · ".join(info["tf"])))
    st.markdown('<div class="catatan">' + " &nbsp;|&nbsp; ".join(bagian) + "</div>", unsafe_allow_html=True)

    tab_data, tab_ai = st.tabs(["📋 Data Toko", "🤖 AI Agent (Gemini)"])

    # ---- Tab data
    with tab_data:
        st.markdown(f'<div class="catatan">Menampilkan <b>{len(df_filter)}</b> toko | Filter: {html.escape(info_filter)}</div>',
                    unsafe_allow_html=True)
        _, tengah, _ = st.columns([1, 2, 1])
        with tengah:
            mode = st.radio("Mode tampilan", ["Format Laporan", "Tabel Interaktif"], horizontal=True,
                            label_visibility="collapsed")
        if df_tampil.empty:
            st.warning("Tidak ada data pada filter saat ini.")
        elif mode == "Format Laporan":
            if len(df_tampil) > 600:
                st.caption("Menampilkan 600 baris pertama. Unduh Excel untuk data lengkap.")
            st.markdown(tabel_html(df_tampil, meta, badge), unsafe_allow_html=True)
        else:
            lihat, konfig = df_tampil.copy(), {}
            for c in lihat.columns:
                m = meta[c]
                if m["tipe"] == "pct":
                    lihat[c] = lihat[c] * 100
                    konfig[c] = st.column_config.NumberColumn(format=f"%.{m['des']}f%%")
            st.dataframe(lihat, column_config=konfig, height=520, hide_index=True, **LEBAR)

        st.write("")
        _, kolom_tombol, _ = st.columns([1, 1, 1])
        with kolom_tombol:
            st.download_button("⬇️ Download Excel", data=ke_excel(df_tampil, meta, badge),
                               file_name="rekap_pps_filter.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               disabled=df_tampil.empty, **LEBAR)

    # ---- Tab AI Agent
    with tab_ai:
        _, tengah, _ = st.columns([1, 4, 1])
        with tengah:
            st.markdown("### 🤖 AI Agent Analisis Performa Toko")
            st.caption("Tanyakan apa saja tentang data yang sedang aktif di tabel (sesuai filter AM & AC).")
            if "riwayat_ai" not in st.session_state:
                st.session_state.riwayat_ai = []

            with st.form("form_ai"):
                pertanyaan = st.text_area(
                    "Pertanyaan atau permintaan analisis", height=110,
                    placeholder="Contoh: Toko mana yang paling perlu dibina? Buat ringkasan performa area ini.")
                kirim = st.form_submit_button("🚀 Tanya Gemini", type="primary", **LEBAR)

            c1, c2 = st.columns(2)
            tes = c1.button("🔌 Tes API Key", **LEBAR)
            hapus = c2.button("🗑️ Hapus Riwayat", **LEBAR)

            if hapus:
                st.session_state.riwayat_ai = []

            if tes or kirim:
                if not api_key.strip():
                    st.warning("Masukkan Gemini API Key di sidebar (tekan Enter setelah menempel) atau simpan permanen.")
                elif kirim and not pertanyaan.strip():
                    st.warning("Tulis pertanyaan Anda terlebih dahulu.")
                elif kirim and df_tampil.empty:
                    st.warning("Tidak ada data pada filter saat ini.")
                else:
                    try:
                        if tes:
                            with st.spinner("Menguji koneksi ke Gemini..."):
                                balasan = panggil_gemini(api_key, "Balas dengan satu kata: OK",
                                                         "Jawab singkat.", model_pilih)
                            st.success(f"Koneksi berhasil ({model_pilih}). Balasan: {balasan[:40]}")
                        else:
                            konteks = bangun_konteks(df_tampil, meta, info, info_filter, pilih_sheet)
                            isi_prompt = konteks + "\n\nPERTANYAAN PENGGUNA:\n" + pertanyaan.strip()
                            with st.spinner("Gemini sedang menganalisis data..."):
                                jawaban, dipakai = jalankan_ai(api_key, isi_prompt, INSTRUKSI_AI,
                                                               model_pilih, otomatis)
                            st.session_state.riwayat_ai.insert(0, (pertanyaan.strip(), jawaban, dipakai))
                    except Exception as e:
                        st.error(pesan_ramah(e, model_pilih))
                        with st.expander("Detail error teknis"):
                            st.code(f"{type(e).__name__}: {e}")

            for item in st.session_state.riwayat_ai:
                tanya, jawab = item[0], item[1]
                with st.container(border=True):
                    st.markdown(f"**🙋 Pertanyaan:** {tanya}")
                    st.markdown(jawab)
                    if len(item) > 2:
                        st.caption(f"Dijawab oleh model: {item[2]}")


if __name__ == "__main__":
    main()