import streamlit as st
import pandas as pd
from st_supabase_connection import SupabaseConnection
from io import BytesIO
from openpyxl import Workbook
import datetime

st.set_page_config(page_title="Aplikasi Kas Kecil", layout="centered")
st.title("💰 Aplikasi Kas Kecil")

# --- KONEKSI SUPABASE ---
# Pastikan URL dan Key sudah diisi di Secrets Streamlit Cloud
conn = st.connection("supabase", type=SupabaseConnection)

# --- FUNGSI AMBIL DATA ---
def fetch_data():
    try:
        # Mengambil data dari tabel 'kas_kecil'
        res = conn.table("kas_kecil").select("*").execute()
        df = pd.DataFrame(res.data)
        if df.empty:
            return pd.DataFrame(columns=["id", "uraian", "vendor", "tanggal", "jumlah"])
        # Urutkan berdasarkan ID (waktu input)
        return df.sort_values("id")
    except Exception as e:
        st.error(f"Gagal mengambil data: {e}")
        return pd.DataFrame(columns=["id", "uraian", "vendor", "tanggal", "jumlah"])

df_raw = fetch_data()

# --- FORM INPUT ---
with st.form("form_kas", clear_on_submit=True):
    st.subheader("Input Data Transaksi")
    opsi_uraian = [
        "Jamuan Makan Dinas",
        "Kebutuhan Kantor",
        "Karcis Parkir Kendaraan Operasional",
        "Isi BBM Kendaraan Operasional"
    ]
    
    uraian_pilih = st.selectbox("Uraian", opsi_uraian)
    vendor = st.text_input("Nama Vendor")
    tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
    jumlah = st.number_input("Jumlah (Nominal)", min_value=0, step=1000)
    
    submit = st.form_submit_button("Simpan ke Cloud 🚀")

if submit:
    if vendor and jumlah:
        # Simpan data asli ke Supabase
        data_to_insert = {
            "uraian": uraian_pilih,
            "vendor": vendor,
            "tanggal": str(tanggal_input),
            "jumlah": jumlah
        }
        try:
            conn.table("kas_kecil").insert(data_to_insert).execute()
            st.success("Data Berhasil Tersimpan!")
            st.rerun()
        except Exception as e:
            st.error(f"Gagal Simpan: {e}")
    else:
        st.warning("Mohon isi Vendor dan Jumlah!")

# --- REKAPITULASI (Tampilan Persis Awal) ---
st.divider()
st.subheader("📋 Rekapitulasi Kas")

if not df_raw.empty:
    # 1. Tambahkan No Urut (1, 2, 3...)
    df_raw['No'] = range(1, len(df_raw) + 1)
    
    # 2. Gabungkan No + Uraian (Hasil: "1 Jamuan Makan Dinas")
    df_raw['Uraian_Tampil'] = df_raw.apply(lambda x: f"{x['No']} {x['uraian']}", axis=1)
    
    # 3. Filter kolom untuk tampilan tabel
    df_display = df_raw[['No', 'Uraian_Tampil', 'vendor', 'tanggal', 'jumlah']].copy()
    df_display.columns = ['No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
    
    # Ambil Nama Bulan dari data terakhir untuk judul grup
    last_date = pd.to_datetime(df_raw['tanggal'].iloc[-1])
    nama_bulan = last_date.strftime('%B').upper()
    
    with st.expander(f"Data {nama_bulan} 1", expanded=True):
        # Tabel Editor (Sama persis tampilannya)
        st.data_editor(df_display, use_container_width=True, num_rows="dynamic")

# --- DOWNLOAD EXCEL ---
if st.sidebar.button("💾 Siapkan Excel"):
    wb = Workbook()
    ws = wb.active
    # Header Excel
    ws.append(["No", "Uraian", "Vendor", "Tanggal", "Jumlah"])
    # Isi Data
    for r in df_display.values:
        ws.append(list(r))
    
    buf = BytesIO()
    wb.save(buf)
    st.sidebar.download_button("⬇️ Download Sekarang", buf.getvalue(), "Rekap_Kas.xlsx")
