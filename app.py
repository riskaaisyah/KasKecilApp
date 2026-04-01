import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import datetime
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Aplikasi Kas Kecil", layout="centered")
st.title("💰 Aplikasi Kas Kecil")

LIMIT_KAS = 25_000_000
ID_SHEET = "1PoUMSVIEA_dDotKXjt6c--_93jLq9uV5jf9WmaL4wxk"
URL_SHEETS = f"https://docs.google.com/spreadsheets/d/{ID_SHEET}"

# ========================
# KONEKSI & AMBIL DATA
# ========================
conn = st.connection("gsheets", type=GSheetsConnection)

def fetch_data():
    try:
        # Mengambil data terbaru dari GSheets
        df = conn.read(spreadsheet=URL_SHEETS, worksheet="Sheet1", ttl=0)
        if df is None or df.empty:
            return pd.DataFrame(columns=["No","Uraian","Vendor","Tanggal","Jumlah"])
        # Pastikan format tanggal dan angka benar
        df['Tanggal'] = pd.to_datetime(df['Tanggal']).dt.date
        df['Jumlah'] = pd.to_numeric(df['Jumlah'], errors='coerce').fillna(0).astype(int)
        return df
    except:
        return pd.DataFrame(columns=["No","Uraian","Vendor","Tanggal","Jumlah"])

df_gsheets = fetch_data()

# ========================
# FORM INPUT
# ========================
with st.form("form_kas", clear_on_submit=True):
    st.subheader("Input Data Baru")
    opsi_uraian = ["Jamuan Makan Dinas", "Kebutuhan Kantor", "Karcis Parkir Kendaraan Operasional", "Isi BBM Kendaraan Operasional"]
    
    uraian_pilih = st.selectbox("Uraian", opsi_uraian)
    vendor = st.text_input("Nama Vendor")
    tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
    jumlah = st.number_input("Jumlah", min_value=0, step=1000, value=None, placeholder="Masukkan angka...")

    if jumlah:
        st.info(f"**Konfirmasi:** Rp {jumlah:,.0f}".replace(",", "."))
    
    submit = st.form_submit_button("Tambah ke Tabel")

# ========================
# LOGIKA INSERT
# ========================
if submit:
    if jumlah is None or vendor == "":
        st.error("Gagal: Nama Vendor dan Jumlah wajib diisi!")
    else:
        # Tambahkan data baru
        no_baru = len(df_gsheets) + 1
        new_row = pd.DataFrame({
            "No": [no_baru],
            "Uraian": [f"{no_baru} {uraian_pilih}"],
            "Vendor": [vendor],
            "Tanggal": [tanggal_input],
            "Jumlah": [jumlah]
        })
        
        updated_db = pd.concat([df_gsheets, new_row], ignore_index=True)
        # Update permanen ke GSheets
        conn.update(spreadsheet=URL_SHEETS, data=updated_db)
        st.success("Data Tersimpan!")
        st.rerun()

# ========================
# EDIT DATA (Bagian yang kamu mau)
# ========================
st.divider()
st.subheader("Edit Data")

if not df_gsheets.empty:
    # Mapping Nama Bulan Bahasa Indonesia
    nama_bulan = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI",
                  7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}
    
    # Buat kolom bantu untuk grouping (berdasarkan tanggal di GSheets)
    df_gsheets_dt = df_gsheets.copy()
    df_gsheets_dt['Bulan'] = pd.to_datetime(df_gsheets_dt['Tanggal']).dt.month.map(nama_bulan)
    df_gsheets_dt['Tahun'] = pd.to_datetime(df_gsheets_dt['Tanggal']).dt.year
    
    # Ambil periode yang tersedia
    periodes = df_gsheets_dt[['Bulan', 'Tahun']].drop_duplicates().sort_values(['Tahun', 'Bulan'], ascending=False)

    for _, row in periodes.iterrows():
        bln, thn = row['Bulan'], row['Tahun']
        label_tabel = f"{bln} 1" # Sesuai format digambar kamu: "JANUARI 1"
        
        # Filter data untuk periode ini
        df_periode = df_gsheets_dt[(df_gsheets_dt['Bulan'] == bln) & (df_gsheets_dt['Tahun'] == thn)].copy()
        
        st.write(f"### {label_tabel}")
        
        # Tampilkan Tabel yang Bisa Diedit (Data Editor)
        edited_df = st.data_editor(
            df_periode[["No", "Uraian", "Vendor", "Tanggal", "Jumlah"]],
            num_rows="dynamic",
            key=f"editor_{bln}_{thn}",
            use_container_width=True
        )

        # Jika ada perubahan di tabel editor
        if not edited_df.equals(df_periode[["No", "Uraian", "Vendor", "Tanggal", "Jumlah"]]):
            # Gabungkan kembali dengan data utama
            # (Ganti data lama di periode ini dengan data hasil editan)
            other_data = df_gsheets_dt[~((df_gsheets_dt['Bulan'] == bln) & (df_gsheets_dt['Tahun'] == thn))]
            final_db = pd.concat([other_data, edited_df], ignore_index=True)
            
            # Bersihkan kolom bantu sebelum upload
            final_db = final_db[["No", "Uraian", "Vendor", "Tanggal", "Jumlah"]]
            
            # Simpan Perubahan ke Google Sheets
            conn.update(spreadsheet=URL_SHEETS, data=final_db)
            st.success(f"Perubahan di {label_tabel} tersimpan!")
            st.rerun()

# ========================
# EXPORT EXCEL
# ========================
def to_excel(df):
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Rekap Kas")
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
    buf = BytesIO(); wb.save(buf)
    return buf.getvalue()

st.sidebar.divider()
if st.sidebar.button("Download Excel"):
    file_ex = to_excel(df_gsheets)
    st.sidebar.download_button("Klik Download", file_ex, "Kas_Kecil.xlsx")
