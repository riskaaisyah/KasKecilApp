import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
import datetime

# --- CONFIG ---
st.set_page_config(page_title="Aplikasi Kas Kecil", layout="centered")
st.title("💰 Aplikasi Kas Kecil (Final Stable)")

# ID Sheet kamu
ID_SHEET = "1PoUMSVIEA_dDotKXjt6c--_93jLq9uV5jf9WmaL4wxk"
# Link khusus untuk kirim data via Pandas
URL_DATA = f"https://docs.google.com/spreadsheets/d/{ID_SHEET}/gviz/tq?tqx=out:csv"

# --- FUNGSI AMBIL DATA ---
def fetch_data():
    try:
        # Membaca data langsung tanpa library connection yang ribet
        df = pd.read_csv(URL_DATA)
        if df.empty:
            return pd.DataFrame(columns=["No","Uraian","Vendor","Tanggal","Jumlah"])
        df['Tanggal'] = pd.to_datetime(df['Tanggal'], errors='coerce')
        return df
    except:
        return pd.DataFrame(columns=["No","Uraian","Vendor","Tanggal","Jumlah"])

df_gsheets = fetch_data()

# --- FORM INPUT ---
with st.form("form_kas", clear_on_submit=True):
    st.subheader("Input Data Baru")
    opsi_uraian = ["Jamuan Makan Dinas", "Kebutuhan Kantor", "Karcis Parkir Kendaraan Operasional", "Isi BBM Kendaraan Operasional"]
    uraian_pilih = st.selectbox("Uraian", opsi_uraian)
    vendor = st.text_input("Nama Vendor")
    tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
    jumlah = st.number_input("Jumlah", min_value=0, step=1000, value=None, placeholder="Masukkan angka...")
    submit = st.form_submit_button("Simpan Data 🚀")

# --- LOGIKA SIMPAN (MENGGUNAKAN LIBRARY CONNECTION) ---
if submit:
    if jumlah is None or vendor == "":
        st.error("Gagal: Nama Vendor dan Jumlah wajib diisi!")
    else:
        try:
            from streamlit_gsheets import GSheetsConnection
            conn = st.connection("gsheets", type=GSheetsConnection)
            
            # Ambil data terbaru untuk nomor urut
            db_now = fetch_data()
            no_baru = len(db_now) + 1
            
            new_row = pd.DataFrame({
                "No": [no_baru],
                "Uraian": [f"{no_baru} {uraian_pilih}"],
                "Vendor": [vendor],
                "Tanggal": [tanggal_input.strftime('%Y-%m-%d')],
                "Jumlah": [jumlah]
            })
            
            updated_db = pd.concat([db_now, new_row], ignore_index=True)
            updated_db['Tanggal'] = updated_db['Tanggal'].astype(str)
            
            # KIRIM DATA (Pastikan Secrets di Streamlit Cloud sudah benar)
            # URL harus pakai versi /edit
            clean_url = f"https://docs.google.com/spreadsheets/d/{ID_SHEET}/edit?usp=sharing"
            conn.update(spreadsheet=clean_url, data=updated_db)
            
            st.success("Data Berhasil Tersimpan Permanen!")
            st.rerun()
        except Exception as e:
            st.error(f"Error: {e}")
            st.info("Coba cek apakah link di Secrets sudah sama persis dengan link GSheets kamu.")

# --- TAMPILAN EDIT DATA (DATA EDITOR) ---
st.divider()
st.subheader("Edit Data")

if not df_gsheets.empty:
    df_display = df_gsheets.copy()
    # Pastikan tampilan tanggal cantik
    df_display['Tanggal'] = pd.to_datetime(df_display['Tanggal']).dt.strftime('%Y-%m-%d')
    
    # Biarkan user edit tabelnya
    edited_df = st.data_editor(df_display[["No", "Uraian", "Vendor", "Tanggal", "Jumlah"]], 
                               num_rows="dynamic", use_container_width=True)
    
    # Tombol Update jika ada perubahan di tabel
    if st.button("Simpan Perubahan Tabel"):
        try:
            from streamlit_gsheets import GSheetsConnection
            conn = st.connection("gsheets", type=GSheetsConnection)
            conn.update(spreadsheet=f"https://docs.google.com/spreadsheets/d/{ID_SHEET}/edit?usp=sharing", data=edited_df)
            st.success("Tabel Berhasil Diupdate!")
            st.rerun()
        except Exception as e:
            st.error(f"Gagal Update Tabel: {e}")
