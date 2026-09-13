import io
import sqlite3
from docx import Document
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import streamlit as st

# Konfigurasi Halaman & Tema Profesional
st.set_page_config(
    page_title="SIMANTAP - Manajemen Aset Terpadu", page_icon="💎", layout="wide"
)

st.markdown(
    """
    <style>
    .main {
        background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
        color: white;
    }
    [data-testid="stSidebar"] .stMarkdown h1, 
    [data-testid="stSidebar"] .stMarkdown h2, 
    [data-testid="stSidebar"] .stMarkdown h3, 
    [data-testid="stSidebar"] label {
        color: #ffffff !important;
        font-weight: 600 !important;
    }
    [data-testid="stSidebar"] .stRadio div[role="radiogroup"] label p {
        color: #f8fafc !important;
        font-size: 15px !important;
        font-weight: 500 !important;
    }
    [data-testid="stSidebar"] .stRadio > div {
        background-color: rgba(255, 255, 255, 0.08);
        padding: 12px;
        border-radius: 12px;
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
    [data-testid="stSidebar"] .stRadio label {
        padding: 8px 12px;
        border-radius: 8px;
        margin-bottom: 6px;
        display: block;
        transition: background 0.2s;
    }
    [data-testid="stSidebar"] .stRadio label:hover {
        background-color: rgba(56, 189, 248, 0.25);
    }
    .header-banner {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
        padding: 30px;
        border-radius: 16px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .card-kategori {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #2563eb;
        margin-bottom: 15px;
    }
    .stButton>button {
        border-radius: 8px;
        font-weight: 600;
        width: 100%;
    }
    </style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_data(table_name):
  try:
    conn = sqlite3.connect("simantap_database.db")
    df = pd.read_sql(f"SELECT * FROM {table_name}", conn)
    conn.close()
    df = df.dropna(how="all")
    if not df.empty:
      for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": None, "None": None, "": None})
      df = df.dropna(subset=[df.columns[0]])
    return df
  except Exception:
    return pd.DataFrame()


def clean_harga(df, possible_cols=["Harga", "Nilai", "Total_Harga"]):
  if df.empty:
    return pd.Series(dtype=float)
  target_col = None
  for col in possible_cols:
    if col in df.columns:
      target_col = col
      break
  if not target_col:
    for c in df.columns:
      if any(kw in c.lower() for kw in ["harga", "nilai", "cost"]):
        target_col = c
        break

  if target_col:
    cleaned = (
        df[target_col]
        .astype(str)
        .str.replace(r"[^\d.]", "", regex=True)
        .replace("", "0")
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)
  return pd.Series(0, index=df.index)


def kategorkan_kendaraan(nama_brg):
  if pd.isna(nama_brg):
    return "Mobil Dinas"
  n = str(nama_brg).upper()
  if any(
      x in n
      for x in [
          "ALAT BERAT",
          "EXCAVATOR",
          "LOADER",
          "GRADER",
          "BULLDOZER",
          "CRANE",
          "TRAKTOR",
          "FORKLIFT",
          "WHEEL",
      ]
  ):
    return "Alat Berat"
  elif any(
      x in n
      for x in [
          "SEPEDA MOTOR",
          "MOTOR",
          "TRAIL",
          "VESPA",
          "HONDA",
          "YAMAHA",
          "SUZUKI",
          "KAWASAKI",
      ]
  ):
    if not any(
        x in n
        for x in [
            "CIVIC",
            "AVANZA",
            "INNOVA",
            "SEDAN",
            "MINIBUS",
            "BUS",
            "TRUCK",
        ]
    ):
      return "Sepeda Motor"
  if any(
      x in n
      for x in [
          "PICK UP",
          "PICKUP",
          "TRUCK",
          "TRUK",
          "BOX",
          "DUMP",
          "STRADA",
          "HILUX",
      ]
  ):
    return "Pick Up / Truk"
  return "Mobil Dinas"


# --- HEADER UTAMA ---
st.markdown(
    """
    <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; margin-bottom: 20px;">
        <div>
            <h2 style="margin:0; color: #1e3a8a;">🛡️ BMD - Aset Daerah</h2>
            <p style="margin:0; color: #64748b; font-size: 14px;">Badan Pengelolaan Keuangan dan Aset Daerah</p>
        </div>
        <div>
            <span style="background: #e0f2fe; color: #0369a1; padding: 6px 14px; border-radius: 20px; font-weight: 600; font-size: 13px;">SIMANTAP v2.1 Professional</span>
        </div>
    </div>
""",
    unsafe_allow_html=True,
)

# --- SIDEBAR NAVIGASI ---
st.sidebar.markdown(
    "<h3 style='color: #38bdf8; text-align: center;'>📌 MENU NAVIGASI</h3>",
    unsafe_allow_html=True,
)
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Pilih Modul Aset:",
    [
        "🏠 Beranda & Ringkasan",
        "🚗 Kendaraan Dinas",
        "🗺️ KIB A (Tanah)",
        "🏢 KIB C (Gedung & Bangunan)",
    ],
)


# ==========================================
# 0. BERANDA & RINGKASAN EKSEKUTIF
# ==========================================
if menu == "🏠 Beranda & Ringkasan":
  st.markdown(
      """
        <div class="header-banner">
            <h1 style="margin:0; font-size: 26px;">"Kelola Aset Daerah, Untuk Pelayanan yang Lebih Baik"</h1>
            <p style="margin:8px 0 0 0; opacity: 0.95; font-size: 15px;">Dashboard rekapitulasi lengkap berdasarkan kategori spesifik aset secara real-time.</p>
        </div>
    """,
      unsafe_allow_html=True,
  )

  df_k = load_data("tabel_kendaraan")
  df_a = load_data("tabel_kib_a_tanah")
  df_c = load_data("tabel_kib_c_gedung")

  if not df_k.empty:
    df_k["Harga_Clean"] = clean_harga(df_k)
    nama_col_ref = (
        "Nama_Barang_Jenis_Barang"
        if "Nama_Barang_Jenis_Barang" in df_k.columns
        else (df_k.columns[1] if len(df_k.columns) > 1 else df_k.columns[0])
    )
    df_k["Kategori_Detail"] = df_k[nama_col_ref].apply(kategorkan_kendaraan)

    alat_berat_cnt = len(df_k[df_k["Kategori_Detail"] == "Alat Berat"])
    alat_berat_val = df_k[df_k["Kategori_Detail"] == "Alat Berat"][
        "Harga_Clean"
    ].sum()
    mobil_cnt = len(df_k[df_k["Kategori_Detail"] == "Mobil Dinas"])
    mobil_val = df_k[df_k["Kategori_Detail"] == "Mobil Dinas"][
        "Harga_Clean"
    ].sum()
    pickup_cnt = len(df_k[df_k["Kategori_Detail"] == "Pick Up / Truk"])
    pickup_val = df_k[df_k["Kategori_Detail"] == "Pick Up / Truk"][
        "Harga_Clean"
    ].sum()
    motor_cnt = len(df_k[df_k["Kategori_Detail"] == "Sepeda Motor"])
    motor_val = df_k[df_k["Kategori_Detail"] == "Sepeda Motor"][
        "Harga_Clean"
    ].sum()
  else:
    (
        alat_berat_cnt,
        alat_berat_val,
        mobil_cnt,
        mobil_val,
        pickup_cnt,
        pickup_val,
        motor_cnt,
        motor_val,
    ) = (0, 0, 0, 0, 0, 0, 0, 0)

  tanah_cnt = len(df_a) if not df_a.empty else 0
  tanah_val = 0
  if not df_a.empty:
    df_a["Harga_Clean"] = clean_harga(df_a)
    tanah_val = df_a["Harga_Clean"].sum()

  gedung_cnt = len(df_c) if not df_c.empty else 0
  gedung_val = 0
  if not df_c.empty:
    df_c["Harga_Clean"] = clean_harga(df_c)
    gedung_val = df_c["Harga_Clean"].sum()

  col1, col2, col3 = st.columns(3)
  with col1:
    st.markdown(
        f"""
        <div class="card-kategori">
            <h4 style="margin:0; color:#1e3a8a;">🚜 Alat Berat</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{alat_berat_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:14px;">Nilai: Rp {alat_berat_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="card-kategori">
            <h4 style="margin:0; color:#1e3a8a;">🚗 Mobil Dinas</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{mobil_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:14px;">Nilai: Rp {mobil_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

  with col2:
    st.markdown(
        f"""
        <div class="card-kategori" style="border-left-color: #059669;">
            <h4 style="margin:0; color:#065f46;">🚚 Pick Up / Truk</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{pickup_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:14px;">Nilai: Rp {pickup_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="card-kategori" style="border-left-color: #059669;">
            <h4 style="margin:0; color:#065f46;">🏍️ Sepeda Motor</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{motor_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:14px;">Nilai: Rp {motor_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

  with col3:
    st.markdown(
        f"""
        <div class="card-kategori" style="border-left-color: #7c3aed;">
            <h4 style="margin:0; color:#6d28d9;">🗺️ KIB A (Tanah)</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{tanah_cnt:,} Bidang</h2>
            <p style="margin:0; color:#64748b; font-size:14px;">Nilai: Rp {tanah_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="card-kategori" style="border-left-color: #7c3aed;">
            <h4 style="margin:0; color:#6d28d9;">🏢 KIB C (Gedung)</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{gedung_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:14px;">Nilai: Rp {gedung_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==========================================
# 1. MENU KENDARAAN DINAS
# ==========================================
elif menu == "🚗 Kendaraan Dinas":
  df = load_data("tabel_kendaraan")
  if df.empty:
    st.warning("Data tabel_kendaraan belum tersedia.")
  else:
    df["Harga_Clean"] = clean_harga(df)

    st.markdown(
        """
        <div class="header-banner">
            <h2 style='margin:0;'>🚗 Manajemen Aset Kendaraan Dinas</h2>
            <p style='margin:5px 0 0 0;'>Filter jumlah barang berdasarkan SKPD dan Kategori Nilai Barang.</p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    col_f1, col_f2 = st.columns(2)
    with col_f1:
      skpd_list = (
          ["Semua SKPD"] + sorted(df["SKPD"].dropna().astype(str).unique().tolist())
          if "SKPD" in df.columns
          else ["Semua SKPD"]
      )
      selected_skpd = st.selectbox("🏢 **Filter Berdasarkan SKPD:**", skpd_list)

    with col_f2:
      filter_nilai = st.selectbox(
          "💰 **Filter Berdasarkan Nilai Barang:**",
          [
              "Semua Kisaran Nilai",
              "Di bawah Rp 50 Juta",
              "Rp 50 Juta - Rp 200 Juta",
              "Di atas Rp 200 Juta",
          ],
      )

    df_filtered = df.copy()
    if selected_skpd != "Semua SKPD":
      df_filtered = df_filtered[df_filtered["SKPD"] == selected_skpd]

    if filter_nilai == "Di bawah Rp 50 Juta":
      df_filtered = df_filtered[df_filtered["Harga_Clean"] < 50000000]
    elif filter_nilai == "Rp 50 Juta - Rp 200 Juta":
      df_filtered = df_filtered[
          (df_filtered["Harga_Clean"] >= 50000000)
          & (df_filtered["Harga_Clean"] <= 200000000)
      ]
    elif filter_nilai == "Di atas Rp 200 Juta":
      df_filtered = df_filtered[df_filtered["Harga_Clean"] > 200000000]

    col_m1, col_m2 = st.columns(2)
    with col_m1:
      st.metric("Jumlah Unit Akurat", f"{len(df_filtered):,} Unit")
    with col_m2:
      st.metric(
          "Akumulasi Nilai", f"Rp {df_filtered['Harga_Clean'].sum():,.0f}"
      )

    st.markdown("---")
    keyword = st.text_input("🔎 Cari Berdasarkan No. Polisi / Merk / Pengguna:")
    df_view = df_filtered.copy()
    if keyword:
      df_view = df_view[
          df_view.astype(str)
          .apply(lambda row: row.str.contains(keyword, case=False).any(), axis=1)
      ]

    st.dataframe(
        df_view.drop(columns=["Harga_Clean"], errors="ignore"),
        use_container_width=True,
        height=400,
    )


# ==========================================
# 2. MENU KIB A (TANAH)
# ==========================================
elif menu == "🗺️ KIB A (Tanah)":
  df_a = load_data("tabel_kib_a_tanah")
  if df_a.empty:
    st.warning("Belum ada data untuk tabel KIB A (Tanah).")
  else:
    df_a["Harga_Clean"] = clean_harga(df_a)

    st.markdown(
        """
        <div class="header-banner">
            <h2 style='margin:0;'>🗺️ Manajemen Aset KIB A (Tanah)</h2>
            <p style='margin:5px 0 0 0;'>Filter jumlah bidang tanah berdasarkan SKPD dan Rentang Nilai/Harga.</p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    col_fa1, col_fa2 = st.columns(2)
    with col_fa1:
      skpd_options_a = (
          ["Semua SKPD"]
          + sorted(df_a["SKPD"].dropna().astype(str).unique().tolist())
          if "SKPD" in df_a.columns
          else ["Semua SKPD"]
      )
      selected_skpd_a = st.selectbox(
          "🏢 **Filter Berdasarkan SKPD:**", skpd_options_a, key="skpd_a"
      )
    with col_fa2:
      filter_nilai_a = st.selectbox(
          "💰 **Filter Berdasarkan Nilai Tanah:**",
          [
              "Semua Kisaran Nilai",
              "Di bawah Rp 100 Juta",
              "Rp 100 Juta - Rp 1 Miliar",
              "Di atas Rp 1 Miliar",
          ],
      )

    df_a_filtered = df_a.copy()
    if selected_skpd_a != "Semua SKPD":
      df_a_filtered = df_a_filtered[df_a_filtered["SKPD"] == selected_skpd_a]

    if filter_nilai_a == "Di bawah Rp 100 Juta":
      df_a_filtered = df_a_filtered[df_a_filtered["Harga_Clean"] < 100000000]
    elif filter_nilai_a == "Rp 100 Juta - Rp 1 Miliar":
      df_a_filtered = df_a_filtered[
          (df_a_filtered["Harga_Clean"] >= 100000000)
          & (df_a_filtered["Harga_Clean"] <= 1000000000)
      ]
    elif filter_nilai_a == "Di atas Rp 1 Miliar":
      df_a_filtered = df_a_filtered[df_a_filtered["Harga_Clean"] > 1000000000]

    col_ma1, col_ma2 = st.columns(2)
    with col_ma1:
      st.metric("Jumlah Bidang Akurat", f"{len(df_a_filtered):,} Bidang")
    with col_ma2:
      st.metric(
          "Total Nilai Tanah", f"Rp {df_a_filtered['Harga_Clean'].sum():,.0f}"
      )

    st.markdown("---")
    kw_a = st.text_input("🔎 Cari Berdasarkan Alamat / Lokasi:")
    df_a_view = df_a_filtered.copy()
    if kw_a:
      df_a_view = df_a_view[
          df_a_view.astype(str)
          .apply(lambda row: row.str.contains(kw_a, case=False).any(), axis=1)
      ]

    st.dataframe(
        df_a_view.drop(columns=["Harga_Clean"], errors="ignore"),
        use_container_width=True,
        height=400,
    )


# ==========================================
# 3. MENU KIB C (GEDUNG & BANGUNAN)
# ==========================================
elif menu == "🏢 KIB C (Gedung & Bangunan)":
  df_c = load_data("tabel_kib_c_gedung")
  if df_c.empty:
    st.warning("Belum ada data untuk tabel KIB C (Gedung & Bangunan).")
  else:
    df_c["Harga_Clean"] = clean_harga(df_c)

    st.markdown(
        """
        <div class="header-banner">
            <h2 style='margin:0;'>🏢 Manajemen Aset KIB C (Gedung & Bangunan)</h2>
            <p style='margin:5px 0 0 0;'>Filter jumlah unit gedung berdasarkan SKPD dan Rentang Nilai/Harga.</p>
        </div>
    """,
        unsafe_allow_html=True,
    )

    col_fc1, col_fc2 = st.columns(2)
    with col_fc1:
      skpd_options_c = (
          ["Semua SKPD"]
          + sorted(df_c["SKPD"].dropna().astype(str).unique().tolist())
          if "SKPD" in df_c.columns
          else ["Semua SKPD"]
      )
      selected_skpd_c = st.selectbox(
          "🏢 **Filter Berdasarkan SKPD:**", skpd_options_c, key="skpd_c"
      )
    with col_fc2:
      filter_nilai_c = st.selectbox(
          "💰 **Filter Berdasarkan Nilai Gedung:**",
          [
              "Semua Kisaran Nilai",
              "Di bawah Rp 200 Juta",
              "Rp 200 Juta - Rp 1 Miliar",
              "Di atas Rp 1 Miliar",
          ],
      )

    df_c_filtered = df_c.copy()
    if selected_skpd_c != "Semua SKPD":
      df_c_filtered = df_c_filtered[df_c_filtered["SKPD"] == selected_skpd_c]

    if filter_nilai_c == "Di bawah Rp 200 Juta":
      df_c_filtered = df_c_filtered[df_c_filtered["Harga_Clean"] < 200000000]
    elif filter_nilai_c == "Rp 200 Juta - Rp 1 Miliar":
      df_c_filtered = df_c_filtered[
          (df_c_filtered["Harga_Clean"] >= 200000000)
          & (df_c_filtered["Harga_Clean"] <= 1000000000)
      ]
    elif filter_nilai_c == "Di atas Rp 1 Miliar":
      df_c_filtered = df_c_filtered[df_c_filtered["Harga_Clean"] > 1000000000]

    col_mc1, col_mc2 = st.columns(2)
    with col_mc1:
      st.metric("Jumlah Unit Gedung Akurat", f"{len(df_c_filtered):,} Unit")
    with col_mc2:
      st.metric(
          "Total Nilai Gedung", f"Rp {df_c_filtered['Harga_Clean'].sum():,.0f}"
      )

    st.markdown("---")
    kw_c = st.text_input("🔎 Cari Berdasarkan Nama Gedung / Lokasi:")
    df_c_view = df_c_filtered.copy()
    if kw_c:
      df_c_view = df_c_view[
          df_c_view.astype(str)
          .apply(lambda row: row.str.contains(kw_c, case=False).any(), axis=1)
      ]

    st.dataframe(
        df_c_view.drop(columns=["Harga_Clean"], errors="ignore"),
        use_container_width=True,
        height=400,
    )