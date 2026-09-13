import io
import pandas as pd
import streamlit as st


def render_modul_aset():
  # 1. Custom CSS untuk Tampilan Kotak, Garis, dan Kartu Rincian
  st.markdown(
      """
        <style>
        .preview-box {
            background: #ffffff;
            padding: 20px;
            border-radius: 12px;
            border: 1px solid #cbd5e1;
            box-shadow: 0 4px 6px rgba(0,0,0,0.03);
            margin-top: 15px;
            margin-bottom: 15px;
        }
        .detail-card {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 12px 15px;
            height: 100%;
        }
        .data-item {
            padding: 6px 0;
            border-bottom: 1px dashed #cbd5e1;
            font-size: 13px;
        }
        </style>
    """,
      unsafe_allow_html=True,
  )

  st.title("🚗 Kendaraan Dinas & Manajemen Aset")
  st.write(
      "Silakan pilih atau klik pada tabel di bawah untuk melihat rincian"
      " lengkap seluruh data."
  )

  # Dataset Utama Anda
  if "df_aset" not in st.session_state:
    st.session_state["df_aset"] = pd.DataFrame({
        "NOUrut": [7, 8, 9, 10],
        "Kode__Barang": [
            "02.02.01.04.001",
            "02.02.01.04.001",
            "02.02.01.04.001",
            "02.02.01.04.001",
        ],
        "Nama_Barang_Jenis_Barang": [
            "Sepeda Motor",
            "Sepeda Motor",
            "Sepeda Motor",
            "Sepeda Motor",
        ],
        "Nomor_Register": ["000003", "000004", "000005", "000006"],
        "Merk_Type": [
            "Honda NF 100 LD",
            "Honda NF 100 LD",
            "Honda NF 125 XD",
            "Honda NF 125 XD",
        ],
        "Bahan": ["Besi", "Besi", "Besi", "Besi"],
        "Tahun_Pembelian": [2005.0, 2006.0, 2010.0, 2011.0],
        "Harga": [1500000, 1600000, 2000000, 2100000],
        "SKPD": [
            "DINAS KESEHATAN",
            "DINAS KESEHATAN",
            "DINAS PERHUBUNGAN",
            "DINAS PERHUBUNGAN",
        ],
    })

  df_view = st.session_state["df_aset"]

  # 2. Tabel Interaktif Utama
  event = st.dataframe(
      df_view,
      use_container_width=True,
      selection_mode="single-row",
      on_select="rerun",
      key="tabel_aset_grid",
  )

  # 3. Tombol Download Data Keseluruhan
  col_dl1, col_dl2, _ = st.columns([1, 1, 2])
  csv_data = df_view.to_csv(index=False).encode("utf-8")
  col_dl1.download_button(
      label="📥 Download CSV",
      data=csv_data,
      file_name="data_aset.csv",
      mime="text/csv",
  )

  output_excel = io.BytesIO()
  with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    df_view.to_excel(writer, index=False, sheet_name="Aset")
  excel_data = output_excel.getvalue()

  col_dl2.download_button(
      label="📥 Download Excel",
      data=excel_data,
      file_name="data_aset.xlsx",
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

  # 4. Preview Rincian Barang (Menampilkan SEMUA LAMAN/DATA LENGKAP KETIKA DIKLIK)
  selected_rows = event.selection.get("rows", [])

  if selected_rows:
    st.markdown(
        """
        <div class="preview-box">
            <h3 style="margin-top:0; color:#1e3a8a; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px;">📋 Preview Rincian Semua Laman / Data Terpilih</h3>
        """,
        unsafe_allow_html=True,
    )

    # Menampilkan seluruh baris data dari dataframe secara penuh & rapi
    for idx, row in df_view.iterrows():
      st.markdown(
          f"<b style='color:#0f172a;'>Baris / Entri #{idx + 1} - "
          f"{row.get('Nama_Barang_Jenis_Barang', 'Aset')} (Reg:"
          f" {row.get('Nomor_Register', '-')})</b>",
          unsafe_allow_html=True,
      )

      col_p1, col_p2, col_p3 = st.columns(3)
      items_list = [
          (k, v)
          for k, v in row.items()
          if k not in ["Harga_Clean", "Kategori_Detail"]
      ]
      chunk_size = (len(items_list) + 2) // 3

      with col_p1:
        st.markdown(
            '<div class="detail-card"><b style="color:#1e3a8a;">📌 Identitas &'
            ' Kode</b><hr style="margin:5px 0 10px 0;">',
            unsafe_allow_html=True,
        )
        for k, v in items_list[:chunk_size]:
          val_str = (
              f"Rp {float(v):,.0f}" if "harga" in k.lower() and pd.notna(v) else str(v)
          )
          st.markdown(
              f"<div class='data-item'><b>{k}:</b> {val_str}</div>",
              unsafe_allow_html=True,
          )
        st.markdown("</div>", unsafe_allow_html=True)

      with col_p2:
        st.markdown(
            '<div class="detail-card"><b style="color:#1e3a8a;">🚗 Spesifikasi &'
            ' Fisik</b><hr style="margin:5px 0 10px 0;">',
            unsafe_allow_html=True,
        )
        for k, v in items_list[chunk_size : chunk_size * 2]:
          val_str = (
              f"Rp {float(v):,.0f}" if "harga" in k.lower() and pd.notna(v) else str(v)
          )
          st.markdown(
              f"<div class='data-item'><b>{k}:</b> {val_str}</div>",
              unsafe_allow_html=True,
          )
        st.markdown("</div>", unsafe_allow_html=True)

      with col_p3:
        st.markdown(
            '<div class="detail-card"><b style="color:#1e3a8a;">🏢 Pengelola &'
            ' Lainnya</b><hr style="margin:5px 0 10px 0;">',
            unsafe_allow_html=True,
        )
        for k, v in items_list[chunk_size * 2 :]:
          val_str = (
              f"Rp {float(v):,.0f}" if "harga" in k.lower() and pd.notna(v) else str(v)
          )
          st.markdown(
              f"<div class='data-item'><b>{k}:</b> {val_str}</div>",
              unsafe_allow_html=True,
          )
        st.markdown("</div>", unsafe_allow_html=True)

      st.markdown(
          "<hr style='border: 0.5px dashed #cbd5e1; margin: 15px 0;'>",
          unsafe_allow_html=True,
      )

    st.markdown("</div>", unsafe_allow_html=True)
  else:
    st.info(
        "👆 Klik baris mana saja pada tabel di atas untuk menampilkan pratinjau"
        " rincian seluruh laman data secara lengkap."
    )


if __name__ == "__main__":
  st.set_page_config(
      page_title="SIMANTAP - Manajemen Aset Terpadu", layout="wide"
  )
  render_modul_aset()