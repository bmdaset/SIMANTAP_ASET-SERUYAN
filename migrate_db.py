import sqlite3
import pandas as pd

# Koneksi ke database SQLite
conn = sqlite3.connect("simantap_database.db")

# 1. Migrasi Data Kendaraan
try:
  df_kendaraan = pd.read_excel("DATA ASET KENDARAAN.xlsx", header=0)
  df_kendaraan = df_kendaraan.dropna(how="all")
  df_kendaraan = df_kendaraan.loc[
      :, [isinstance(c, str) and not c.startswith("Unnamed") for c in df_kendaraan.columns]
  ]
  df_kendaraan.columns = [
      str(c)
      .strip()
      .replace(" ", "_")
      .replace(".", "")
      .replace("/", "_")
      for c in df_kendaraan.columns
  ]
  df_kendaraan.to_sql(
      "tabel_kendaraan", conn, if_exists="replace", index=False
  )
  print("Berhasil memperbarui: Tabel Kendaraan")
except Exception as e:
  print(f"Gagal memperbarui Kendaraan: {e}")

# 2. Migrasi Data KIB A (Tanah)
try:
  df_tanah = pd.read_excel("DATA ASET KIB A TANAH.xlsx", header=1)
  df_tanah = df_tanah.dropna(how="all")
  df_tanah = df_tanah.loc[
      :, [isinstance(c, str) and not c.startswith("Unnamed") for c in df_tanah.columns]
  ]
  df_tanah.columns = [
      str(c)
      .strip()
      .replace(" ", "_")
      .replace(".", "")
      .replace("/", "_")
      .replace("(", "")
      .replace(")", "")
      for c in df_tanah.columns
  ]
  df_tanah.to_sql("tabel_kib_a_tanah", conn, if_exists="replace", index=False)
  print("Berhasil memperbarui: Tabel KIB A (Tanah)")
except Exception as e:
  print(f"Gagal memperbarui KIB A: {e}")

# 3. Migrasi Data KIB C (Gedung & Bangunan)
try:
  df_kibc = pd.read_excel("DATA ASET KIB C GEDUNG BANGUNAN.xlsx", header=7)
  df_kibc = df_kibc.dropna(how="all")
  df_kibc = df_kibc.loc[
      :, [isinstance(c, str) and not c.startswith("Unnamed") for c in df_kibc.columns]
  ]
  df_kibc.columns = [
      str(c)
      .strip()
      .replace(" ", "_")
      .replace(".", "")
      .replace("/", "_")
      .replace("(", "")
      .replace(")", "")
      for c in df_kibc.columns
  ]
  df_kibc.to_sql("tabel_kib_c_gedung", conn, if_exists="replace", index=False)
  print("Berhasil memperbarui: Tabel KIB C (Gedung & Bangunan)")
except Exception as e:
  print(f"Gagal memperbarui KIB C: {e}")

conn.close()
print("\nSemua perubahan data database berhasil diterapkan!")