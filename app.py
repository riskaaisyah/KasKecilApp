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

# --- DAFTAR OPSI URAIAN ---
opsi_uraian = [
    "Jamuan Makan Dinas",
    "Kebutuhan Kantor",
    "Karcis Parkir Kendaraan Operasional",
    "Isi BBM Kendaraan Operasional"
]

# --- FORM INPUT ---
with st.form("form_kas", clear_on_submit=True):
    st.subheader("Input Data Transaksi")
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
                bulan_pilihan = tanggal_input.month
                tahun_pilihan = tanggal_input.year
                
                res = conn.table("kas_kecil").select("*").eq("uraian", target_uraian).execute()
                df_cek = pd.DataFrame(res.data)
                
                found = False
                if not df_cek.empty:
                    df_cek['tgl_temp'] = pd.to_datetime(df_cek['tanggal'], errors='coerce')
                    match = df_cek[(df_cek['tgl_temp'].dt.month == bulan_pilihan) & (df_cek['tgl_temp'].dt.year == tahun_pilihan)]
                    
                    if not match.empty:
                        id_lama = match.iloc[0]['id']
                        jumlah_baru = match.iloc[0]['jumlah'] + jumlah
                        conn.table("kas_kecil").delete().eq("id", id_lama).execute()
                        
                        data_update = {
                            "uraian": target_uraian,
                            "vendor": vendor,
                            "tanggal": str(tanggal_input),
                            "jumlah": int(jumlah_baru)
                        }
                        conn.table("kas_kecil").insert(data_update).execute()
                        found = True
                
                if not found:
                    data_baru = {"uraian": target_uraian, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}
                    conn.table("kas_kecil").insert(data_baru).execute()
            else:
                data_normal = {"uraian": uraian_pilih, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}
                conn.table("kas_kecil").insert(data_normal).execute()
            
            st.success("Data Berhasil Diperbarui!")
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
    df_raw = df_raw.sort_values('id')
    
    nama_bulan_id = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
                     7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

    df_raw['Kelompok_Sheet'] = "" 
    df_raw['temp_month'] = df_raw['tanggal_dt'].dt.month
    df_raw['temp_year'] = df_raw['tanggal_dt'].dt.year
    periods = df_raw[['temp_year', 'temp_month']].drop_duplicates().values

    for thn, bln_num in periods:
        if pd.isna(bln_num): continue
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

    # --- TAMPILKAN EXPANDER (AUTO-SAVE) ---
    list_kelompok = [k for k in df_raw['Kelompok_Sheet'].unique() if k != ""][::-1]

    for kelompok in list_kelompok:
        df_group = df_raw[df_raw['Kelompok_Sheet'] == kelompok].copy()
        total_group = df_group['jumlah'].sum()
        sisa_kuota = LIMIT_KAS - total_group
        
        with st.expander(f"📂 {kelompok} | Total: Rp {total_group:,.0f} | 💰 Sisa: Rp {sisa_kuota:,.0f}".replace(",", "."), expanded=True):
            df_edit = df_group[['id', 'uraian', 'vendor', 'tanggal_dt', 'jumlah']].copy()
            df_edit.columns = ['id', 'uraian', 'vendor', 'tanggal', 'jumlah']
            
            edited_data = st.data_editor(
                df_edit,
                key=f"editor_{kelompok}",
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "id": None,
                    "uraian": st.column_config.SelectboxColumn("Uraian", options=opsi_uraian),
                    "tanggal": st.column_config.DateColumn("Tanggal"),
                    "jumlah": st.column_config.NumberColumn("Jumlah", format="%d")
                }
            )

            # --- LOGIKA AUTO-SAVE ---
            if not edited_data.equals(df_edit):
                try:
                    ids_asli = set(df_edit['id'].tolist())
                    ids_sekarang = set(edited_data['id'].dropna().tolist())
                    ids_dihapus = ids_asli - ids_sekarang
                    
                    for d_id in ids_dihapus:
                        conn.table("kas_kecil").delete().eq("id", d_id).execute()
                    
                    for index, row_baru in edited_data.iterrows():
                        if pd.notna(row_baru['id']):
                            row_lama = df_edit[df_edit['id'] == row_baru['id']].iloc[0]
                            if not row_baru.equals(row_lama):
                                conn.table("kas_kecil").update({
                                    "uraian": row_baru['uraian'], "vendor": row_baru['vendor'],
                                    "tanggal": str(row_baru['tanggal']), "jumlah": int(row_baru['jumlah'])
                                }).eq("id", row_baru['id']).execute()
                    st.toast("Perubahan disimpan!", icon="✅")
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal simpan otomatis: {e}")
else:
    st.info("Belum ada data.")

# --- SIDEBAR ---
st.sidebar.markdown("### 💾 Simpan Rekap Kas Kecil")
if 'df_raw' in locals() and not df_raw.empty:
    if st.sidebar.button("💾 Siapkan Excel Format BIOS"):
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        wb = Workbook()
        del wb['Sheet']
        for p in df_raw['Kelompok_Sheet'].unique():
            if p == "": continue
            ws = wb.create_sheet(title=p[:31])
            ws.merge_cells('B2:I3')
            ws['B2'] = "APLIKASI BIOS (BIAYA OPERASIONAL)"
            ws['B2'].font = Font(bold=True, color="FFFFFF", size=14)
            ws['B2'].alignment = Alignment(horizontal="center", vertical="center")
            ws['B2'].fill = PatternFill(start_color="CC9900", end_color="CC9900", fill_type="solid")
            ws['A5'] = "PERTANGGUNGJAWABAN ATAS ND PENGAJUAN NOMOR KU.02.04/19/11/1/PBLU/PBLU-25"
            headers = ["No", "URAIAN", "NAMA VENDOR", "POS MATA ANGGARAN", "GL ACCOUNT", "TANGGAL TRANSAKSI", "JUMLAH PENGGUNAAN", "SETELAH PPN"]
            ws.append([])
            ws.append(headers)
            df_b = df_raw[df_raw['Kelompok_Sheet'] == p]
            for i, row in enumerate(df_b.itertuples(), 1):
                tgl = "" if row.uraian == "Karcis Parkir Kendaraan Operasional" else str(row.tanggal)
                ws.append([i, f"{i} {row.uraian}", row.vendor, "", "", tgl, row.jumlah, row.jumlah])
        buf = BytesIO()
        wb.save(buf)
        st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "Rekap_BIOS.xlsx")

st.sidebar.divider()
st.sidebar.markdown("### 🗑️ Hapus Keseluruhan Data")
konfirmasi_hapus = st.sidebar.checkbox("Saya yakin ingin menghapus SEMUA data")
if st.sidebar.button("🗑️ Kosongkan Data", type="primary", disabled=not konfirmasi_hapus):
    try:
        conn.table("kas_kecil").delete().neq("id", 0).execute()
        st.rerun()
    except Exception as e:
        st.sidebar.error(f"Gagal: {e}")
