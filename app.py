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
    [data-testid="stSidebar"] .stMarkdown h1, [data-testid="stSidebar"] .stMarkdown h2, [data-testid="stSidebar"] .stMarkdown h3, [data-testid="stSidebar"] label {
        color: #f8fafc !important;
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
    # Membersihkan baris kosong agar jumlah baris akurat tanpa selisih
    df = df.dropna(how="all")
    return df
  except Exception:
    return pd.DataFrame()


def clean_harga(df, possible_cols=["Harga", "Nilai", "Total_Harga"]):
  target_col = None
  for col in possible_cols:
    if col in df.columns:
      target_col = col
      break
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
      ]
  ):
    return "Alat Berat"
  elif any(
      x in n
      for x in [
          "MOTOR",
          "SEPEDA MOTOR",
          "TRAIL",
          "VESPA",
          "YAMAHA",
          "HONDA",
          "SUZUKI",
          "KAWASAKI",
      ]
  ):
    if not any(
        x in n for x in ["CIVIC", "AVANZA", "INNOVA", "SEDAN", "MINIBUS", "BUS"]
    ):
      return "Sepeda Motor"
  if any(
      x in n for x in ["PICK UP", "TRUCK", "TRUK", "BOX", "DUMP", "STRADA", "HILUX"]
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

menu = st.sidebar.selectbox(
    "Pilih Modul Aset",
    [
        "Beranda & Ringkasan",
        "Kendaraan Dinas",
        "KIB A (Tanah)",
        "KIB C (Gedung & Bangunan)",
    ],
)


# ==========================================
# 0. BERANDA & RINGKASAN EKSEKUTIF
# ==========================================
if menu == "Beranda & Ringkasan":
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

  if not df_a.empty:
    df_a["Harga_Clean"] = clean_harga(df_a)
    tanah_cnt = len(df_a)
    tanah_val = df_a["Harga_Clean"].sum()
  else:
    tanah_cnt, tanah_val = 0, 0

  if not df_c.empty:
    df_c["Harga_Clean"] = clean_harga(df_c)
    gedung_cnt = len(df_c)
    gedung_val = df_c["Harga_Clean"].sum()
  else:
    gedung_cnt, gedung_val = 0, 0

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

  st.markdown("---")
  st.markdown("### 📈 Grafik Visualisasi Total Nilai Aset per Kategori")

  chart_data = pd.DataFrame({
      "Kategori": [
          "Alat Berat",
          "Mobil Dinas",
          "Pick Up / Truk",
          "Sepeda Motor",
          "Tanah",
          "Gedung",
      ],
      "Total Nilai (Rp)": [
          alat_berat_val,
          mobil_val,
          pickup_val,
          motor_val,
          tanah_val,
          gedung_val,
      ],
  }).set_index("Kategori")

  st.bar_chart(chart_data, color="#2563eb")


# ==========================================
# 1. MENU KENDARAAN DINAS
# ==========================================
elif menu == "Kendaraan Dinas":
  df = load_data("tabel_kendaraan")
  if df.empty:
    st.warning("Data tabel_kendaraan belum tersedia.")
  else:
    df["Harga_Clean"] = clean_harga(df)
    nama_col_ref = (
        "Nama_Barang_Jenis_Barang"
        if "Nama_Barang_Jenis_Barang" in df.columns
        else df.columns[1]
    )
    df["Kategori_Kendaraan"] = df[nama_col_ref].apply(kategorkan_kendaraan)

    if "selected_row_idx_k1" not in st.session_state:
      st.session_state.selected_row_idx_k1 = None

    if st.session_state.selected_row_idx_k1 is not None:
      idx = st.session_state.selected_row_idx_k1
      st.markdown(
          """
                <div style='background: #0284c7; padding: 15px; border-radius: 10px; color: white; margin-bottom: 20px;'>
                    <h3 style='margin:0;'>📱 PREVIEW DETAIL KENDARAAN DINAS</h3>
                </div>
            """,
          unsafe_allow_html=True,
      )

      if st.button("⬅️ Kembali ke Tabel Utama"):
        st.session_state.selected_row_idx_k1 = None
        st.rerun()

      if idx < len(df):
        detail_data = df.iloc[idx]
        item_id = detail_data.get(df.columns[0], "ID")

        detail_list = [
            {
                "Atribut / Kolom": str(c).replace("_", " ").title(),
                "Keterangan Detail": str(v if pd.notna(v) else "-"),
            }
            for c, v in detail_data.items()
            if c != "Harga_Clean"
        ]
        st.dataframe(
            pd.DataFrame(detail_list),
            use_container_width=True,
            hide_index=True,
            height=400,
        )

        st.markdown("---")
        st.markdown("### 📥 Unduh Dokumen Laporan")
        dl1, dl2, dl3 = st.columns(3)

        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as w:
          pd.DataFrame([detail_data]).to_excel(w, index=False)
        with dl1:
          st.download_button(
              "📊 Excel (.xlsx)",
              out_excel.getvalue(),
              file_name=f"Kendaraan_{item_id}.xlsx",
          )

        doc = Document()
        doc.add_heading("Detail Aset Kendaraan Dinas", level=1)
        for k, v in detail_data.items():
          doc.add_paragraph(f"{str(k).replace('_', ' ')} : {v}")
        out_word = io.BytesIO()
        doc.save(out_word)
        with dl2:
          st.download_button(
              "📄 Word (.docx)",
              out_word.getvalue(),
              file_name=f"Kendaraan_{item_id}.docx",
          )

        out_pdf = io.BytesIO()
        pdf = canvas.Canvas(out_pdf, pagesize=letter)
        w_p, h_p = letter
        pdf.drawString(40, h_p - 40, "DETAIL ASET KENDARAAN DINAS")
        y = h_p - 70
        for k, v in detail_data.items():
          pdf.drawString(40, y, f"{str(k).replace('_', ' ')}: {str(v)[:60]}")
          y -= 18
          if y < 40:
            pdf.showPage()
            y = h_p - 40
        pdf.save()
        with dl3:
          st.download_button(
              "📑 PDF (.pdf)",
              out_pdf.getvalue(),
              file_name=f"Kendaraan_{item_id}.pdf",
          )

    else:
      st.markdown(
          """
            <div class="header-banner">
                <h2 style='margin:0;'>🚗 Manajemen Aset Kendaraan Dinas</h2>
                <p style='margin:5px 0 0 0;'>Klik salah satu baris pada tabel untuk melihat rincian preview lengkap.</p>
            </div>
        """,
          unsafe_allow_html=True,
      )

      skpd_list = (
          ["Semua SKPD"] + sorted(df["SKPD"].dropna().astype(str).unique().tolist())
          if "SKPD" in df.columns
          else ["Semua SKPD"]
      )
      selected_skpd = st.selectbox("🏢 **Filter Berdasarkan SKPD:**", skpd_list)

      df_filtered = df.copy()
      if selected_skpd != "Semua SKPD":
        df_filtered = df_filtered[df_filtered["SKPD"] == selected_skpd]

      col1, col2 = st.columns(2)
      with col1:
        st.metric("Total Unit", f"{len(df_filtered):,} Unit")
      with col2:
        st.metric(
            "Akumulasi Nilai", f"Rp {df_filtered['Harga_Clean'].sum():,.0f}"
        )

      st.markdown("---")
      keyword = st.text_input("🔎 Cari Berdasarkan No. Polisi / Merk / Pengguna:")
      df_view = df_filtered.copy()
      if keyword:
        df_view = df_view[
            df_view.astype(str)
            .apply(
                lambda row: row.str.contains(keyword, case=False).any(), axis=1
            )
        ]

      st.info(
          f"💡 Ditemukan **{len(df_view)}** data. **Klik salah satu baris**"
          " untuk preview penuh dan unduh dokumen."
      )

      event = st.dataframe(
          df_view.drop(
              columns=["Harga_Clean", "Kategori_Kendaraan"], errors="ignore"
          ),
          use_container_width=True,
          on_select="rerun",
          selection_mode="single-row",
          height=400,
      )

      if event.selection.rows:
        local_idx = event.selection.rows[0]
        actual_idx = df.index[df_view.index[local_idx]]
        st.session_state.selected_row_idx_k1 = actual_idx
        st.rerun()


# ==========================================
# 2. MENU KIB A (TANAH)
# ==========================================
elif menu == "KIB A (Tanah)":
  df_a = load_data("tabel_kib_a_tanah")
  if df_a.empty:
    st.warning("Belum ada data untuk tabel KIB A (Tanah).")
  else:
    df_a["Harga_Clean"] = clean_harga(df_a)

    if "selected_row_idx_a" not in st.session_state:
      st.session_state.selected_row_idx_a = None

    if st.session_state.selected_row_idx_a is not None:
      idx = st.session_state.selected_row_idx_a
      st.markdown(
          """
                <div style='background: #059669; padding: 15px; border-radius: 10px; color: white; margin-bottom: 20px;'>
                    <h3 style='margin:0;'>📱 PREVIEW DETAIL KIB A (TANAH)</h3>
                </div>
            """,
          unsafe_allow_html=True,
      )

      if st.button("⬅️ Kembali ke Tabel Utama"):
        st.session_state.selected_row_idx_a = None
        st.rerun()

      if idx < len(df_a):
        detail_data = df_a.iloc[idx]
        item_id = detail_data.get(df_a.columns[0], "ID")
        detail_list = [
            {
                "Atribut / Kolom": str(c).replace("_", " ").title(),
                "Keterangan Detail": str(v if pd.notna(v) else "-"),
            }
            for c, v in detail_data.items()
            if c != "Harga_Clean"
        ]
        st.dataframe(
            pd.DataFrame(detail_list),
            use_container_width=True,
            hide_index=True,
            height=400,
        )

        st.markdown("---")
        dl1, dl2, dl3 = st.columns(3)
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as w:
          pd.DataFrame([detail_data]).to_excel(w, index=False)
        with dl1:
          st.download_button(
              "📊 Excel (.xlsx)",
              out_excel.getvalue(),
              file_name=f"Tanah_{item_id}.xlsx",
          )
        doc = Document()
        doc.add_heading("Detail Aset KIB A (Tanah)", level=1)
        for k, v in detail_data.items():
          doc.add_paragraph(f"{str(k).replace('_', ' ')} : {v}")
        out_word = io.BytesIO()
        doc.save(out_word)
        with dl2:
          st.download_button(
              "📄 Word (.docx)",
              out_word.getvalue(),
              file_name=f"Tanah_{item_id}.docx",
          )
        out_pdf = io.BytesIO()
        pdf = canvas.Canvas(out_pdf, pagesize=letter)
        w_p, h_p = letter
        pdf.drawString(40, h_p - 40, "DETAIL ASET KIB A TANAH")
        y = h_p - 70
        for k, v in detail_data.items():
          pdf.drawString(40, y, f"{str(k).replace('_', ' ')}: {str(v)[:60]}")
          y -= 18
          if y < 40:
            pdf.showPage()
            y = h_p - 40
        pdf.save()
        with dl3:
          st.download_button(
              "📑 PDF (.pdf)",
              out_pdf.getvalue(),
              file_name=f"Tanah_{item_id}.pdf",
          )

    else:
      st.markdown(
          """
            <div class="header-banner">
                <h2 style='margin:0;'>🗺️ Manajemen Aset KIB A (Tanah)</h2>
                <p style='margin:5px 0 0 0;'>Klik baris pada tabel untuk melihat preview penuh.</p>
            </div>
        """,
          unsafe_allow_html=True,
      )

      skpd_options_a = (
          ["Semua SKPD"]
          + sorted(df_a["SKPD"].dropna().astype(str).unique().tolist())
          if "SKPD" in df_a.columns
          else ["Semua SKPD"]
      )
      selected_skpd_a = st.selectbox(
          "🏢 **Filter Berdasarkan SKPD:**", skpd_options_a
      )

      df_a_filtered = df_a.copy()
      if selected_skpd_a != "Semua SKPD":
        df_a_filtered = df_a_filtered[df_a_filtered["SKPD"] == selected_skpd_a]

      col1, col2 = st.columns(2)
      with col1:
        st.metric("Total Bidang", f"{len(df_a_filtered):,} Bidang")
      with col2:
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

      event_a = st.dataframe(
          df_a_view.drop(columns=["Harga_Clean"], errors="ignore"),
          use_container_width=True,
          on_select="rerun",
          selection_mode="single-row",
          height=400,
      )

      if event_a.selection.rows:
        local_idx_a = event_a.selection.rows[0]
        actual_idx_a = df_a.index[df_a_view.index[local_idx_a]]
        st.session_state.selected_row_idx_a = actual_idx_a
        st.rerun()


# ==========================================
# 3. MENU KIB C (GEDUNG & BANGUNAN)
# ==========================================
elif menu == "KIB C (Gedung & Bangunan)":
  df_c = load_data("tabel_kib_c_gedung")
  if df_c.empty:
    st.warning("Belum ada data untuk tabel KIB C (Gedung & Bangunan).")
  else:
    df_c["Harga_Clean"] = clean_harga(df_c)

    if "selected_row_idx_c" not in st.session_state:
      st.session_state.selected_row_idx_c = None

    if st.session_state.selected_row_idx_c is not None:
      idx = st.session_state.selected_row_idx_c
      st.markdown(
          """
                <div style='background: #7c3aed; padding: 15px; border-radius: 10px; color: white; margin-bottom: 20px;'>
                    <h3 style='margin:0;'>📱 PREVIEW DETAIL KIB C (GEDUNG & BANGUNAN)</h3>
                </div>
            """,
          unsafe_allow_html=True,
      )

      if st.button("⬅️ Kembali ke Tabel Utama"):
        st.session_state.selected_row_idx_c = None
        st.rerun()

      if idx < len(df_c):
        detail_data = df_c.iloc[idx]
        item_id = detail_data.get(df_c.columns[0], "ID")
        detail_list = [
            {
                "Atribut / Kolom": str(c).replace("_", " ").title(),
                "Keterangan Detail": str(v if pd.notna(v) else "-"),
            }
            for c, v in detail_data.items()
            if c != "Harga_Clean"
        ]
        st.dataframe(
            pd.DataFrame(detail_list),
            use_container_width=True,
            hide_index=True,
            height=400,
        )

        st.markdown("---")
        dl1, dl2, dl3 = st.columns(3)
        out_excel = io.BytesIO()
        with pd.ExcelWriter(out_excel, engine="openpyxl") as w:
          pd.DataFrame([detail_data]).to_excel(w, index=False)
        with dl1:
          st.download_button(
              "📊 Excel (.xlsx)",
              out_excel.getvalue(),
              file_name=f"Gedung_{item_id}.xlsx",
          )
        doc = Document()
        doc.add_heading("Detail Aset KIB C (Gedung & Bangunan)", level=1)
        for k, v in detail_data.items():
          doc.add_paragraph(f"{str(k).replace('_', ' ')} : {v}")
        out_word = io.BytesIO()
        doc.save(out_word)
        with dl2:
          st.download_button(
              "📄 Word (.docx)",
              out_word.getvalue(),
              file_name=f"Gedung_{item_id}.docx",
          )
        out_pdf = io.BytesIO()
        pdf = canvas.Canvas(out_pdf, pagesize=letter)
        w_p, h_p = letter
        pdf.drawString(40, h_p - 40, "DETAIL ASET KIB C GEDUNG")
        y = h_p - 70
        for k, v in detail_data.items():
          pdf.drawString(40, y, f"{str(k).replace('_', ' ')}: {str(v)[:60]}")
          y -= 18
          if y < 40:
            pdf.showPage()
            y = h_p - 40
        pdf.save()
        with dl3:
          st.download_button(
              "📑 PDF (.pdf)",
              out_pdf.getvalue(),
              file_name=f"Gedung_{item_id}.pdf",
          )

    else:
      st.markdown(
          """
            <div class="header-banner">
                <h2 style='margin:0;'>🏢 Manajemen Aset KIB C (Gedung & Bangunan)</h2>
                <p style='margin:5px 0 0 0;'>Klik baris pada tabel untuk melihat preview penuh.</p>
            </div>
        """,
          unsafe_allow_html=True,
      )

      skpd_options_c = (
          ["Semua SKPD"]
          + sorted(df_c["SKPD"].dropna().astype(str).unique().tolist())
          if "SKPD" in df_c.columns
          else ["Semua SKPD"]
      )
      selected_skpd_c = st.selectbox(
          "🏢 **Filter Berdasarkan SKPD:**", skpd_options_c
      )

      df_c_filtered = df_c.copy()
      if selected_skpd_c != "Semua SKPD":
        df_c_filtered = df_c_filtered[df_c_filtered["SKPD"] == selected_skpd_c]

      col1, col2 = st.columns(2)
      with col1:
        st.metric("Total Unit Bangunan", f"{len(df_c_filtered):,} Unit")
      with col2:
        st.metric(
            "Total Nilai Bangunan", f"Rp {df_c_filtered['Harga_Clean'].sum():,.0f}"
        )

      st.markdown("---")
      kw_c = st.text_input("🔎 Cari Berdasarkan Nama Gedung / Lokasi:")
      df_c_view = df_c_filtered.copy()
      if kw_c:
        df_c_view = df_c_view[
            df_c_view.astype(str)
            .apply(lambda row: row.str.contains(kw_c, case=False).any(), axis=1)
        ]

      event_c = st.dataframe(
          df_c_view.drop(columns=["Harga_Clean"], errors="ignore"),
          use_container_width=True,
          on_select="rerun",
          selection_mode="single-row",
          height=400,
      )

      if event_c.selection.rows:
        local_idx_c = event_c.selection.rows[0]
        actual_idx_c = df_c.index[df_c_view.index[local_idx_c]]
        st.session_state.selected_row_idx_c = actual_idx_c
        st.rerun()