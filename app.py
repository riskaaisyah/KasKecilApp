import streamlit as st
import pandas as pd
from st_supabase_connection import SupabaseConnection
from io import BytesIO
from openpyxl import Workbook
import datetime

st.set_page_config(page_title="Aplikasi Kas Kecil", layout="centered")
st.title("💰 Aplikasi Kas Kecil")

# --- KONEKSI SUPABASE ---
conn = st.connection("supabase", type=SupabaseConnection)

# --- FUNGSI AMBIL DATA ---
def fetch_data():
    try:
        res = conn.table("kas_kecil").select("*").execute()
        df = pd.DataFrame(res.data)
        if df.empty:
            return pd.DataFrame(columns=["id", "uraian", "vendor", "tanggal", "jumlah"])
        return df.sort_values("id")
    except Exception as e:
        return pd.DataFrame(columns=["id", "uraian", "vendor", "tanggal", "jumlah"])

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
    tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
    
    jumlah = st.number_input(
        "Jumlah (Nominal)", 
        min_value=0, 
        step=1000, 
        value=None, 
        placeholder="Masukkan angka tanpa titik/koma..."
    )
    
    if jumlah:
        st.info(f"**Konfirmasi Nominal:** Rp {jumlah:,.0f}".replace(",", "."))
    
    submit = st.form_submit_button("Simpan ke Cloud 🚀")

# --- LOGIKA SIMPAN (SUDAH DIPERBAIKI) ---
if submit:
    if vendor and jumlah is not None:
        # Menyiapkan data untuk dikirim
        data_baru = {
            "uraian": uraian_pilih,
            "vendor": vendor,
            "tanggal": str(tanggal_input),
            "jumlah": int(jumlah)
        }
        
        try:
            # PROSES KIRIM KE SUPABASE
            conn.table("kas_kecil").insert(data_baru).execute()
            st.success(f"Berhasil menyimpan: {uraian_pilih} - Rp {jumlah:,.0f}".replace(",", "."))
            st.rerun() # Refresh agar data muncul di tabel bawah
        except Exception as e:
            st.error(f"Gagal kirim ke Cloud: {e}")
    else:
        st.warning("Mohon isi Nama Vendor dan Jumlah dulu ya!")

# --- REKAPITULASI ---
st.divider()
st.subheader("📋 Rekapitulasi Kas")

df_raw = fetch_data()

if not df_raw.empty:
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'], errors='coerce')
    df_raw = df_raw.dropna(subset=['tanggal_dt'])
    
    nama_bulan_id = {
        1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI",
        7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"
    }
    df_raw['Bulan_Nama'] = df_raw['tanggal_dt'].dt.month.map(nama_bulan_id)
    df_raw['Tahun'] = df_raw['tanggal_dt'].dt.year
    
    list_periode = df_raw.sort_values('tanggal_dt', ascending=False)[['Bulan_Nama', 'Tahun']].drop_duplicates()

    for _, row in list_periode.iterrows():
        bln = row['Bulan_Nama']
        thn = row['Tahun']
        
        mask = (df_raw['Bulan_Nama'] == bln) & (df_raw['Tahun'] == thn)
        df_periode = df_raw[mask].copy()
        total_bulan = df_periode['jumlah'].sum()
        
        with st.expander(f"📂 Data {bln} {thn} (Total: Rp {total_bulan:,.0f})".replace(",", "."), expanded=True):
            df_periode = df_periode.sort_values('id')
            df_periode['No'] = range(1, len(df_periode) + 1)
            df_periode['Uraian_Tampil'] = df_periode.apply(lambda x: f"{x['No']} {x['uraian']}", axis=1)
            
            df_display = df_periode[['No', 'Uraian_Tampil', 'vendor', 'tanggal', 'jumlah']].copy()
            df_display.columns = ['No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            df_style = df_display.copy()
            df_style['Jumlah'] = df_style['Jumlah'].apply(lambda x: f"{x:,.0f}".replace(",", "."))
            
            st.data_editor(df_style, use_container_width=True, key=f"table_{bln}_{thn}")
else:
    st.info("Belum ada data yang tersimpan di Cloud Database.")

# --- DOWNLOAD EXCEL ---
if not df_raw.empty:
    if st.sidebar.button("💾 Siapkan Excel Semua Data"):
        wb = Workbook()
        ws = wb.active
        ws.title = "Semua Rekap"
        ws.append(["No", "Uraian", "Vendor", "Tanggal", "Jumlah"])
        
        for i, r in enumerate(df_raw.sort_values('tanggal_dt').values, 1):
            # r[1] = uraian, r[2] = vendor, r[3] = tanggal, r[4] = jumlah
            ws.append([i, f"{i} {r[1]}", r[2], r[3], r[4]])
        
        buf = BytesIO()
        wb.save(buf)
        st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "Rekap_Kas_Kecil.xlsx")
