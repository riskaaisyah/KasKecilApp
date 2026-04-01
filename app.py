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
    vendor = st.text_input("Nama Vendor", placeholder="Contoh: Toko Buku Gramedia")
    tanggal_input = st.date_input("Tanggal", value=datetime.today().date()) # Pakai .date() agar formatnya rapi
    
    # Perubahan di sini: value=None dan placeholder
    jumlah = st.number_input(
        "Jumlah (Nominal)", 
        min_value=0, 
        step=1000, 
        value=None, 
        placeholder="Masukkan angka tanpa titik/koma..."
    )
    
    # Fitur konfirmasi agar tidak salah nol
    if jumlah:
        st.info(f"**Konfirmasi Nominal:** Rp {jumlah:,.0f}".replace(",", "."))
    
    submit = st.form_submit_button("Simpan ke Cloud 🚀")

# --- LOGIKA SIMPAN ---
if submit:
    # Pastikan jumlah tidak None dan vendor tidak kosong
    if vendor and jumlah is not None: 
        data_to_insert = {
            "uraian": uraian_pilih,
            "vendor": vendor,
            "tanggal": str(tanggal_input),
            "jumlah": int(jumlah) # Pastikan masuk sebagai angka bulat
        }
        try:
            conn.table("kas_kecil").insert(data_to_insert).execute()
            st.success("Data Berhasil Tersimpan!")
            st.rerun()
        except Exception as e:
            st.error(f"Gagal Simpan: {e}")
    else:
        st.warning("Mohon isi Nama Vendor dan Jumlah Nominal!")

# --- REKAPITULASI (OTOMATIS PISAH PER BULAN) ---
st.divider()
st.subheader("📋 Rekapitulasi Kas")

if not df_raw.empty:
    # 1. Pastikan kolom tanggal benar-benar format tanggal
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'])
    
    # 2. Buat kolom bantu Nama Bulan & Tahun (Bahasa Indonesia)
    nama_bulan_id = {
        1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI",
        7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
    }
    df_raw['Bulan_Nama'] = df_raw['tanggal_dt'].dt.month.map(nama_bulan_id)
    df_raw['Tahun'] = df_raw['tanggal_dt'].dt.year
    
    # 3. Ambil daftar Periode yang unik (misal: April 2026, Januari 2026)
    # Diurutkan dari yang terbaru (Descending)
    list_periode = df_raw[['Bulan_Nama', 'Tahun', 'tanggal_dt']].drop_duplicates(subset=['Bulan_Nama', 'Tahun'])
    list_periode = list_periode.sort_values('tanggal_dt', ascending=False)

    # 4. LOOPING: Buat kotak expander untuk SETIAP periode
    for _, row in list_periode.iterrows():
        bln = row['Bulan_Nama']
        thn = row['Tahun']
        
        # Filter data khusus bulan & tahun ini
        df_periode = df_raw[(df_raw['Bulan_Nama'] == bln) & (df_raw['Tahun'] == thn)].copy()
        
        # Hitung total per bulan (biar informatif)
        total_bulan = df_periode['jumlah'].sum()
        
        with st.expander(f"📂 Data {bln} {thn} (Total: Rp {total_bulan:,.0f})".replace(",", "."), expanded=True):
            # Penomoran ulang di dalam bulan tersebut
            df_periode['No'] = range(1, len(df_periode) + 1)
            df_periode['Uraian_Tampil'] = df_periode.apply(lambda x: f"{x['No']} {x['uraian']}", axis=1)
            
            # Rapikan tabel untuk tampilan
            df_display = df_periode[['No', 'Uraian_Tampil', 'vendor', 'tanggal', 'jumlah']].copy()
            df_display.columns = ['No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            # Format tampilan ribuan di kolom jumlah (biar enak dilihat)
            df_display_style = df_display.copy()
            df_display_style['Jumlah'] = df_display_style['Jumlah'].apply(lambda x: f"{x:,.0f}".replace(",", "."))
            
            st.data_editor(df_display_style, use_container_width=True, key=f"editor_{bln}_{thn}")

# --- DOWNLOAD EXCEL (SEMUA DATA) ---
if st.sidebar.button("💾 Siapkan Excel Semua Data"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Semua Rekap"
    ws.append(["No", "Uraian", "Vendor", "Tanggal", "Jumlah"])
    
    # Ambil data terbaru lagi untuk excel
    df_excel = fetch_data()
    for i, r in enumerate(df_excel.values, 1):
        # Format No Uraian otomatis di excel juga
        ws.append([i, f"{i} {r[1]}", r[2], r[3], r[4]])
    
    buf = BytesIO()
    wb.save(buf)
    st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "Rekap_Kas_Kecil.xlsx")
