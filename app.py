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

# --- LOGIKA SIMPAN ---
if submit:
    if vendor and (jumlah is not None):
        try:
            target_uraian = "Karcis Parkir Kendaraan Operasional"
            
            if uraian_pilih == target_uraian:
                # Logika Akumulasi Karcis di Bulan yang sama
                bulan_ini = datetime.date.today().month
                tahun_ini = datetime.date.today().year
                
                res = conn.table("kas_kecil").select("*").eq("uraian", target_uraian).execute()
                df_cek = pd.DataFrame(res.data)
                
                found = False
                if not df_cek.empty:
                    df_cek['tgl_temp'] = pd.to_datetime(df_cek['tanggal'], errors='coerce')
                    match = df_cek[(df_cek['tgl_temp'].dt.month == bulan_ini) & (df_cek['tgl_temp'].dt.year == tahun_ini)]
                    
                    if not match.empty:
                        id_lama = match.iloc[0]['id']
                        jumlah_baru = match.iloc[0]['jumlah'] + jumlah
                        conn.table("kas_kecil").update({"jumlah": int(jumlah_baru)}).eq("id", id_lama).execute()
                        found = True
                
                if not found:
                    data_karcis = {
                        "uraian": target_uraian,
                        "vendor": vendor,
                        "tanggal": str(datetime.date.today()), # Kasih tanggal hari ini untuk sistem batching
                        "jumlah": int(jumlah)
                    }
                    conn.table("kas_kecil").insert(data_karcis).execute()
            else:
                data_normal = {
                    "uraian": uraian_pilih,
                    "vendor": vendor,
                    "tanggal": str(tanggal_input),
                    "jumlah": int(jumlah)
                }
                conn.table("kas_kecil").insert(data_normal).execute()
            
            st.success("Data Berhasil Tersimpan!")
            st.rerun()
        except Exception as e:
            st.error(f"Gagal Simpan: {e}")
    else:
        st.warning("Mohon isi Nama Vendor dan Jumlah!")

# --- REKAPITULASI ---
st.divider()
st.subheader("📋 Rekapitulasi Kas (Limit 25jt/Sheet)")

df_raw = fetch_data()

if not df_raw.empty:
    LIMIT_KAS = 25_000_000
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'], errors='coerce')
    # Jangan gunakan dropna karena data Karcis mungkin bermasalah formatnya
    df_raw = df_raw.sort_values('id')
    
    nama_bulan_id = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
                     7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

    # --- LOGIKA BATCHING ---
    df_raw['Kelompok_Sheet'] = "" 
    # Buat temp kolom untuk grouping periode
    df_raw['temp_month'] = df_raw['tanggal_dt'].dt.month
    df_raw['temp_year'] = df_raw['tanggal_dt'].dt.year
    
    periods = df_raw[['temp_year', 'temp_month']].drop_duplicates().values

    for thn, bln_num in periods:
        if pd.isna(bln_num): continue # Lewati jika tanggal rusak
        current_batch = 1
        running_total = 0
        mask = (df_raw['temp_year'] == thn) & (df_raw['temp_month'] == bln_num)
        df_bulan = df_raw[mask].sort_values('id')
        
        for idx, row in df_bulan.iterrows():
            if running_total + row['jumlah'] > LIMIT_KAS:
                current_batch += 1
                running_total = row['jumlah']
            else:
                running_total += row['jumlah']
            label = f"{nama_bulan_id[int(bln_num)]} ({current_batch}) {int(thn)}"
            df_raw.at[idx, 'Kelompok_Sheet'] = label

    # --- TAMPILKAN EXPANDER ---
    list_kelompok = [k for k in df_raw['Kelompok_Sheet'].unique() if k != ""][::-1]

    for kelompok in list_kelompok:
        df_group = df_raw[df_raw['Kelompok_Sheet'] == kelompok].copy()
        total_group = df_group['jumlah'].sum()
        sisa_kuota = LIMIT_KAS - total_group
        
        with st.expander(f"📂 {kelompok} | Total: Rp {total_group:,.0f} | 💰 Sisa: Rp {sisa_kuota:,.0f}".replace(",", "."), expanded=True):
            df_group['No'] = range(1, len(df_group) + 1)
            
            # Sembunyikan tanggal jika Karcis
            df_group['Tanggal_Tampil'] = df_group.apply(
                lambda x: "" if x['uraian'] == "Karcis Parkir Kendaraan Operasional" else x['tanggal'], axis=1
            )
            df_group['Uraian_Tampil'] = df_group.apply(lambda x: f"{x['No']} {x['uraian']}", axis=1)
            
            df_display = df_group[['No', 'Uraian_Tampil', 'vendor', 'Tanggal_Tampil', 'jumlah']].copy()
            df_display.columns = ['No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            df_style = df_display.copy()
            df_style['Jumlah'] = df_style['Jumlah'].apply(lambda x: f"{x:,.0f}".replace(",", "."))
            st.data_editor(df_style, use_container_width=True, key=f"editor_{kelompok}")

    # --- DOWNLOAD EXCEL ---
    st.sidebar.divider()
    if st.sidebar.button("💾 Siapkan Excel Format BIOS"):
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        wb = Workbook()
        del wb['Sheet']
        
        header_fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")
        title_fill = PatternFill(start_color="CC9900", end_color="CC9900", fill_type="solid")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))

        for p in df_raw['Kelompok_Sheet'].unique():
            if p == "": continue
            ws = wb.create_sheet(title=p[:31])
            ws.merge_cells('B2:I3')
            cell_judul = ws['B2']
            cell_judul.value = "APLIKASI BIOS (BIAYA OPERASIONAL)"
            cell_judul.font = Font(bold=True, color="FFFFFF", size=14)
            cell_judul.alignment = Alignment(horizontal="center", vertical="center")
            cell_judul.fill = title_fill

            ws['A5'] = "PERTANGGUNGJAWABAN ATAS ND PENGAJUAN NOMOR KU.02.04/19/11/1/PBLU/PBLU-25"
            ws['A5'].font = Font(bold=True, size=10)

            headers = ["No", "URAIAN", "NAMA VENDOR", "POS MATA ANGGARAN", "GL ACCOUNT", "TANGGAL TRANSAKSI", "JUMLAH PENGGUNAAN", "SETELAH PPN"]
            ws.append([]) 
            ws.append(headers) 
            
            for cell in ws[7]:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            df_batch = df_raw[df_raw['Kelompok_Sheet'] == p]
            for i, row in enumerate(df_batch.itertuples(), 1):
                # row.uraian, row.vendor, row.tanggal, row.jumlah
                tgl_excel = "" if row.uraian == "Karcis Parkir Kendaraan Operasional" else row.tanggal
                ws.append([i, f"{i} {row.uraian}", row.vendor, "", "", tgl_excel, row.jumlah, row.jumlah])
                for cell in ws[ws.max_row]:
                    cell.border = thin_border
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0'

            ws.column_dimensions['B'].width = 35
            ws.column_dimensions['C'].width = 25

        buf = BytesIO()
        wb.save(buf)
        st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "Rekap_BIOS.xlsx")

else:
    st.info("Belum ada data.")

# --- CLEAR DATA ---
st.sidebar.divider()
konfirmasi_hapus = st.sidebar.checkbox("Saya yakin ingin menghapus SEMUA data")
if st.sidebar.button("🗑️ Kosongkan Data", type="primary", disabled=not konfirmasi_hapus):
    try:
        conn.table("kas_kecil").delete().neq("id", 0).execute()
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Gagal: {e}")
