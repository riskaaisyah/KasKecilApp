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
        st.warning("Mohon isi Nama Vendor dan Jumlah dulu!")

    # --- REKAPITULASI (DENGAN INFO TOTAL & SISA KUOTA) ---
st.divider()
st.subheader("📋 Rekapitulasi Kas (Limit 25jt/Sheet)")

df_raw = fetch_data()

if not df_raw.empty:
    # --- 1. PREPARASI DATA ---
    LIMIT_KAS = 25_000_000
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'], errors='coerce')
    df_raw = df_raw.dropna(subset=['tanggal_dt']).sort_values(['tanggal_dt', 'id'])
    
    nama_bulan_id = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
                     7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

    # --- 2. LOGIKA BATCHING PER BULAN ---
    df_raw['Kelompok_Sheet'] = "" 
    df_raw['temp_month'] = df_raw['tanggal_dt'].dt.month
    df_raw['temp_year'] = df_raw['tanggal_dt'].dt.year
    distinct_periods = df_raw[['temp_year', 'temp_month']].drop_duplicates().values

    for thn, bln_num in distinct_periods:
        current_batch = 1
        running_total = 0
        mask = (df_raw['temp_year'] == thn) & (df_raw['temp_month'] == bln_num)
        df_bulan = df_raw[mask].sort_values(['tanggal_dt', 'id'])
        
        for idx, row in df_bulan.iterrows():
            if running_total + row['jumlah'] > LIMIT_KAS:
                current_batch += 1
                running_total = row['jumlah']
            else:
                running_total += row['jumlah']
            
            label = f"{nama_bulan_id[bln_num]} ({current_batch}) {int(thn)}"
            df_raw.at[idx, 'Kelompok_Sheet'] = label

    df_raw = df_raw.drop(columns=['temp_month', 'temp_year'])

    # --- 3. TAMPILKAN EXPANDER DENGAN INFO SISA ---
    list_kelompok = df_raw['Kelompok_Sheet'].unique()[::-1]

    for kelompok in list_kelompok:
        df_group = df_raw[df_raw['Kelompok_Sheet'] == kelompok].copy()
        total_group = df_group['jumlah'].sum()
        
        # LOGIKA SISA KUOTA
        sisa_kuota = LIMIT_KAS - total_group
        
        # Format string untuk judul expander
        total_str = f"Rp {total_group:,.0f}".replace(",", ".")
        sisa_str = f"Rp {sisa_kuota:,.0f}".replace(",", ".")
        
        # Judul Expander sekarang lebih informatif
        with st.expander(f"📂 {kelompok} | Total: {total_str} | 💰 Sisa: {sisa_str}", expanded=True):
            df_group['No'] = range(1, len(df_group) + 1)
            df_group['Uraian_Tampil'] = df_group.apply(lambda x: f"{x['No']} {x['uraian']}", axis=1)
            
            df_display = df_group[['No', 'Uraian_Tampil', 'vendor', 'tanggal', 'jumlah']].copy()
            df_display.columns = ['No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            df_style = df_display.copy()
            df_style['Jumlah'] = df_style['Jumlah'].apply(lambda x: f"{x:,.0f}".replace(",", "."))
            
            st.data_editor(df_style, use_container_width=True, key=f"editor_{kelompok}")

else:
    st.info("Belum ada data yang tersimpan di Cloud Database.")

# --- DOWNLOAD EXCEL ---
if not df_raw.empty:
    if st.sidebar.button("💾 Siapkan Excel Format BIOS"):
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl import Workbook
        
        wb = Workbook()
        del wb['Sheet'] # Hapus sheet bawaan yang kosong

        # Warna & Style (Sesuai Gambar 2 yang kamu mau)
        header_fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid") # Tosca
        title_fill = PatternFill(start_color="CC9900", end_color="CC9900", fill_type="solid")  # Cokelat Judul
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))

        # Kita gunakan 'Kelompok_Sheet' yang sudah dibuat di bagian Rekapitulasi
        # Setiap kelompok (Januari 1, Januari 2, dst) akan jadi satu Sheet
        for kelompok in df_raw['Kelompok_Sheet'].unique():
            ws = wb.create_sheet(title=kelompok)
            
            # 1. Judul Besar (Merge Cell B2:I3)
            ws.merge_cells('B2:I3')
            cell_judul = ws['B2']
            cell_judul.value = "APLIKASI BIOS (BIAYA OPERASIONAL)"
            cell_judul.font = Font(bold=True, color="FFFFFF", size=14)
            cell_judul.alignment = Alignment(horizontal="center", vertical="center")
            cell_judul.fill = title_fill

            # 2. Sub-judul Nota Dinas (Baris 5)
            ws['A5'] = "PERTANGGUNGJAWABAN ATAS ND PENGAJUAN NOMOR KU.02.04/19/11/1/PBLU/PBLU-25"
            ws['A5'].font = Font(bold=True, size=10)

            # 3. Header Tabel (Baris 7)
            headers = ["No", "URAIAN", "NAMA VENDOR", "POS MATA ANGGARAN", "GL ACCOUNT", "TANGGAL TRANSAKSI", "JUMLAH PENGGUNAAN", "SETELAH PPN"]
            ws.append([]) # Baris kosong 6
            ws.append(headers) # Baris 7
            
            for cell in ws[7]:
                cell.fill = header_fill
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin_border

            # 4. Isi Data Khusus Batch Ini
            df_batch = df_raw[df_raw['Kelompok_Sheet'] == kelompok]
            for i, row in enumerate(df_batch.values, 1):
                # row[1]=uraian, row[2]=vendor, row[3]=tanggal, row[4]=jumlah
                baris_data = [
                    i, 
                    f"{i} {row[1]}", 
                    row[2], 
                    "", # Pos Mata Anggaran (kosong)
                    "", # GL Account (kosong)
                    row[3], 
                    row[4], 
                    row[4]
                ]
                ws.append(baris_data)
                
                # Tambah border & format ribuan
                for cell in ws[ws.max_row]:
                    cell.border = thin_border
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0'

            # 5. Atur Lebar Kolom agar rapi
            ws.column_dimensions['B'].width = 35
            ws.column_dimensions['C'].width = 25
            ws.column_dimensions['F'].width = 20
            ws.column_dimensions['G'].width = 18
            ws.column_dimensions['H'].width = 18

        # Simpan ke Buffer
        buf = BytesIO()
        wb.save(buf)
        st.sidebar.success(f"Siap! Terbagi jadi {len(df_raw['Kelompok_Sheet'].unique())} Sheet.")
        st.sidebar.download_button(
            label="⬇️ Download Excel Multi-Sheet",
            data=buf.getvalue(),
            file_name=f"2026 kas kecil _REKAPITULASI PENGGUNAAN KAS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# --- FITUR CLEAR ALL DATA (DI SIDEBAR) ---
# Checkbox konfirmasi agar tidak asal klik
konfirmasi_hapus = st.sidebar.checkbox("Saya yakin ingin menghapus SEMUA data")

if st.sidebar.button("🗑️ Kosongkan Semua Data Cloud", type="primary", disabled=not konfirmasi_hapus):
    try:
        # Menghapus semua baris di tabel 'kas_kecil'
        # Di Supabase, .neq("id", 0) adalah trik untuk memilih semua baris karena ID pasti bukan 0
        conn.table("kas_kecil").delete().neq("id", 0).execute()
        
        st.sidebar.success("Semua data telah dibersihkan!")
        st.rerun() # Refresh tampilan agar tabel jadi kosong
    except Exception as e:
        st.sidebar.error(f"Gagal menghapus: {e}")
