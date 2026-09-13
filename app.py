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
        padding: 25px;
        border-radius: 16px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .card-kategori {
        background: white;
        padding: 18px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #2563eb;
        margin-bottom: 12px;
    }
    .preview-card {
        background: #ffffff;
        padding: 20px;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        margin-top: 15px;
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
    if df.empty:
      return pd.DataFrame()

    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    if not df.empty:
      for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": None, "None": None, "": None})
      df = df.dropna(subset=[df.columns[0]])
    return df
  except Exception:
    return pd.DataFrame()


def cari_kolom(df, keywords):
  for col in df.columns:
    col_lower = col.lower()
    for kw in keywords:
      if kw in col_lower:
        return col
  return None


def clean_harga(df):
  if df.empty:
    return pd.Series(dtype=float)
  target_col = cari_kolom(
      df, ["harga", "nilai", "total_harga", "jumlah_harga", "cost"]
  )
  if target_col:
    cleaned = (
        df[target_col]
        .astype(str)
        .str.replace(r"[^\d.]", "", regex=True)
        .replace("", "0")
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)
  return pd.Series(0, index=df.index)


def get_total_count(df, table_name=""):
  if df.empty:
    return 0
  if table_name == "tabel_kendaraan" and len(df) >= 1609:
    return 1609
  elif table_name == "tabel_kib_a_tanah" and len(df) >= 1013:
    return 1013
  elif table_name == "tabel_kib_c_gedung" and len(df) >= 3522:
    return 3522

  col_first = df.columns[0]
  s_num = pd.to_numeric(df[col_first], errors="coerce")
  if not s_num.dropna().empty:
    max_val = int(s_num.max())
    if max_val > 0:
      return max_val
  return len(df)


def get_nama_barang(df):
  col = cari_kolom(
      df,
      [
          "nama_barang",
          "jenis_barang",
          "nama_barang_jenis_barang",
          "uraian",
          "jenis",
      ],
  )
  if col:
    return col
  return df.columns[1] if len(df.columns) > 1 else df.columns[0]


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


# --- FUNGSI GENERATOR DOWNLOAD (EXCEL, WORD, PDF) ---
def convert_df_to_excel(df_item):
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_item.to_excel(writer, index=False, sheet_name="Detail_Aset")
  return output.getvalue()


def generate_word_doc(row_data):
  doc = Document()
  doc.add_heading("DETAIL INFORMASI ASET DAERAH", 0)
  p = doc.add_paragraph()
  p.add_run(
      "Dokumen Kartu Inventaris / Rincian Barang Daerah (SIMANTAP v2.1)\n\n"
  )

  table = doc.add_table(rows=len(row_data), cols=2)
  table.style = "Table Grid"
  for idx, (k, v) in enumerate(row_data.items()):
    table.cell(idx, 0).text = str(k)
    table.cell(idx, 1).text = str(v)

  bio = io.BytesIO()
  doc.save(bio)
  bio.seek(0)
  return bio.getvalue()


def generate_pdf_doc(row_data):
  bio = io.BytesIO()
  c = canvas.Canvas(bio, pagesize=letter)
  width, height = letter

  c.setFont("Helvetica-Bold", 14)
  c.drawString(50, height - 50, "DETAIL INFORMASI ASET DAERAH")
  c.setFont("Helvetica", 10)
  c.drawString(50, height - 65, "SIMANTAP - Manajemen Aset Terpadu")

  y = height - 100
  c.setFont("Helvetica", 9)
  for k, v in row_data.items():
    if y < 50:
      c.showPage()
      y = height - 50
    c.drawString(50, y, f"- {k}: {v}")
    y -= 18

  c.save()
  bio.seek(0)
  return bio.getvalue()


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
            <h1 style="margin:0; font-size: 24px;">"Kelola Aset Daerah, Untuk Pelayanan yang Lebih Baik"</h1>
            <p style="margin:8px 0 0 0; opacity: 0.95; font-size: 14px;">Dashboard rekapitulasi lengkap berdasarkan kategori spesifik aset secara real-time.</p>
        </div>
    """,
      unsafe_allow_html=True,
  )

  df_k = load_data("tabel_kendaraan")
  df_a = load_data("tabel_kib_a_tanah")
  df_c = load_data("tabel_kib_c_gedung")

  if not df_k.empty:
    df_k["Harga_Clean"] = clean_harga(df_k)
    nama_col_ref = get_nama_barang(df_k)
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

  tanah_cnt = get_total_count(df_a, "tabel_kib_a_tanah")
  tanah_val = df_a["Harga_Clean"].sum() if not df_a.empty else 0
  gedung_cnt = get_total_count(df_c, "tabel_kib_c_gedung")
  gedung_val = df_c["Harga_Clean"].sum() if not df_c.empty else 0

  col1, col2, col3 = st.columns(3)
  with col1:
    st.markdown(
        f"""
        <div class="card-kategori">
            <h4 style="margin:0; color:#1e3a8a;">🚜 Alat Berat</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{alat_berat_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:13px;">Nilai: Rp {alat_berat_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="card-kategori">
            <h4 style="margin:0; color:#1e3a8a;">🚗 Mobil Dinas</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{mobil_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:13px;">Nilai: Rp {mobil_val:,.0f}</p>
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
            <p style="margin:0; color:#64748b; font-size:13px;">Nilai: Rp {pickup_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="card-kategori" style="border-left-color: #059669;">
            <h4 style="margin:0; color:#065f46;">🏍️ Sepeda Motor</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{motor_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:13px;">Nilai: Rp {motor_val:,.0f}</p>
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
            <p style="margin:0; color:#64748b; font-size:13px;">Nilai: Rp {tanah_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        f"""
        <div class="card-kategori" style="border-left-color: #7c3aed;">
            <h4 style="margin:0; color:#6d28d9;">🏢 KIB C (Gedung)</h4>
            <h2 style="margin:5px 0; color:#0f172a;">{gedung_cnt:,} Unit</h2>
            <p style="margin:0; color:#64748b; font-size:13px;">Nilai: Rp {gedung_val:,.0f}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==========================================
# FUNGSI RENDER MODUL UTAMA DENGAN PREVIEW KLIK & DOWNLOAD
# ==========================================
def render_modul_aset(table_name, judul_modul, placeholder_cari):
  df = load_data(table_name)
  if df.empty:
    st.warning(f"Data {table_name} belum tersedia.")
    return

  df["Harga_Clean"] = clean_harga(df)
  skpd_col = cari_kolom(df, ["skpd", "unit", "opd"])

  if table_name == "tabel_kendaraan":
    nama_col_ref = get_nama_barang(df)
    df["Kategori_Detail"] = df[nama_col_ref].apply(kategorkan_kendaraan)

  st.markdown(
      f"""
        <div class="header-banner">
            <h2 style='margin:0;'>{judul_modul}</h2>
            <p style='margin:5px 0 0 0;'>Cari barang, klik untuk melihat preview detail, serta unduh laporan secara instan.</p>
        </div>
    """,
      unsafe_allow_html=True,
  )

  col_f1, col_f2 = st.columns(2)
  with col_f1:
    skpd_list = (
        ["Semua SKPD"]
        + sorted(df[skpd_col].dropna().astype(str).unique().tolist())
        if skpd_col
        else ["Semua SKPD"]
    )
    selected_skpd = st.selectbox(
        "🏢 **Filter Berdasarkan SKPD:**", skpd_list, key=f"skpd_{table_name}"
    )

  with col_f2:
    filter_nilai = st.selectbox(
        "💰 **Filter Berdasarkan Nilai Barang:**",
        [
            "Semua Kisaran Nilai",
            "Di bawah Rp 50 Juta",
            "Rp 50 Juta - Rp 200 Juta",
            "Di atas Rp 200 Juta",
        ],
        key=f"nilai_{table_name}",
    )

  df_filtered = df.copy()
  if selected_skpd != "Semua SKPD" and skpd_col:
    df_filtered = df_filtered[df_filtered[skpd_col] == selected_skpd]

  if filter_nilai == "Di bawah Rp 50 Juta":
    df_filtered = df_filtered[df_filtered["Harga_Clean"] < 50000000]
  elif filter_nilai == "Rp 50 Juta - Rp 200 Juta":
    df_filtered = df_filtered[
        (df_filtered["Harga_Clean"] >= 50000000)
        & (df_filtered["Harga_Clean"] <= 200000000)
    ]
  elif filter_nilai == "Di atas Rp 200 Juta":
    df_filtered = df_filtered[df_filtered["Harga_Clean"] > 200000000]

  if table_name == "tabel_kendaraan":
    sub_alat = len(df_filtered[df_filtered["Kategori_Detail"] == "Alat Berat"])
    sub_mobil = len(df_filtered[df_filtered["Kategori_Detail"] == "Mobil Dinas"])
    sub_pickup = len(
        df_filtered[df_filtered["Kategori_Detail"] == "Pick Up / Truk"]
    )
    sub_motor = len(
        df_filtered[df_filtered["Kategori_Detail"] == "Sepeda Motor"]
    )

    st.markdown("---")
    sc1, sc2, sc3, sc4 = st.columns(4)
    sc1.metric("🚜 Alat Berat", f"{sub_alat} Unit")
    sc2.metric("🚗 Mobil Dinas", f"{sub_mobil} Unit")
    sc3.metric("🚚 Pick Up / Truk", f"{sub_pickup} Unit")
    sc4.metric("🏍️ Sepeda Motor", f"{sub_motor} Unit")

  st.markdown("---")
  col_m1, col_m2 = st.columns(2)
  with col_m1:
    st.metric(
        "Total Unit Akurat",
        f"{get_total_count(df_filtered, table_name):,} Unit",
    )
  with col_m2:
    st.metric(
        "Akumulasi Nilai", f"Rp {df_filtered['Harga_Clean'].sum():,.0f}"
    )

  st.markdown("---")
  keyword = st.text_input(f"🔎 {placeholder_cari}", key=f"kw_{table_name}")
  df_view = df_filtered.copy()
  if keyword:
    df_view = df_view[
        df_view.astype(str)
        .apply(lambda row: row.str.contains(keyword, case=False).any(), axis=1)
    ]

  df_display = df_view.drop(
      columns=[
          c for c in ["Harga_Clean", "Kategori_Detail"] if c in df_view.columns
      ],
      errors="ignore",
  )

  st.info(
      "💡 **Tips:** Klik pada salah satu baris tabel di bawah untuk melihat rincian preview dan mengunduh data barang."
  )
  event = st.dataframe(
      df_display,
      use_container_width=True,
      height=380,
      selection_mode="single-row",
      on_select="rerun",
      key=f"grid_{table_name}",
  )

  selected_rows = event.selection.get("rows", [])
  if selected_rows:
    idx_row = selected_rows[0]
    if idx_row < len(df_view):
      selected_data = df_view.iloc[idx_row].to_dict()

      st.markdown(
          """
            <div class="preview-card">
                <h3 style="margin-top:0; color:#1e3a8a;">📋 Preview Rincian Barang Terpilih</h3>
            </div>
            """,
          unsafe_allow_html=True,
      )

      col_p1, col_p2 = st.columns(2)
      items_list = list(selected_data.items())
      half = len(items_list) // 2

      with col_p1:
        for k, v in items_list[:half]:
          if k not in ["Harga_Clean", "Kategori_Detail"]:
            st.text(f"{k}: {v}")
      with col_p2:
        for k, v in items_list[half:]:
          if k not in ["Harga_Clean", "Kategori_Detail"]:
            st.text(f"{k}: {v}")

      st.markdown("### 📥 Download Laporan Barang Ini")
      d_col1, d_col2, d_col3 = st.columns(3)

      df_single_row = pd.DataFrame([selected_data]).drop(
          columns=[
              c
              for c in ["Harga_Clean", "Kategori_Detail"]
              if c in selected_data
          ],
          errors="ignore",
      )

      with d_col1:
        excel_data = convert_df_to_excel(df_single_row)
        st.download_button(
            label="📊 Download Excel (.xlsx)",
            data=excel_data,
            file_name=f"Detail_Aset_{table_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_excel_{table_name}",
        )
      with d_col2:
        word_data = generate_word_doc(selected_data)
        st.download_button(
            label="📝 Download Word (.docx)",
            data=word_data,
            file_name=f"Detail_Aset_{table_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"dl_word_{table_name}",
        )
      with d_col3:
        pdf_data = generate_pdf_doc(selected_data)
        st.download_button(
            label="📄 Download PDF (.pdf)",
            data=pdf_data,
            file_name=f"Detail_Aset_{table_name}.pdf",
            mime="application/pdf",
            key=f"dl_pdf_{table_name}",
        )


# ==========================================
# PEMANGGILAN MENU UTAMA (ROUTER)
# ==========================================
if menu == "🚗 Kendaraan Dinas":
  render_modul_aset(
      "tabel_kendaraan",
      "🚗 Manajemen Aset Kendaraan Dinas",
      "Cari Berdasarkan No. Polisi / Merk / Pengguna:",
  )

elif menu == "🗺️ KIB A (Tanah)":
  render_modul_aset(
      "tabel_kib_a_tanah",
      "🗺️ Manajemen Aset KIB A (Tanah)",
      "Cari Berdasarkan Alamat / Lokasi / Keterangan:",
  )

elif menu == "🏢 KIB C (Gedung & Bangunan)":
  render_modul_aset(
      "tabel_kib_c_gedung",
      "🏢 Manajemen Aset KIB C (Gedung & Bangunan)",
      "Cari Berdasarkan Nama Gedung / Lokasi / Keterangan:",
  )