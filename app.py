import streamlit as st
import pandas as pd
from st_supabase_connection import SupabaseConnection
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import datetime
import time

# --- 1. KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Aplikasi Kas Kecil", layout="wide")
st.title("💰 Aplikasi Kas Kecil")

# --- 2. KONEKSI SUPABASE ---
conn = st.connection("supabase", type=SupabaseConnection)

# --- 3. FUNGSI AMBIL DATA ---
def fetch_data():
    try:
        res = conn.table("kas_kecil").select("*").execute()
        df = pd.DataFrame(res.data)
        if df.empty:
            return pd.DataFrame(columns=["id", "uraian", "vendor", "tanggal", "jumlah"])
        df['id'] = df['id'].astype(int)
        return df.sort_values("id")
    except Exception:
        return pd.DataFrame(columns=["id", "uraian", "vendor", "tanggal", "jumlah"])

opsi_uraian = [
    "Jamuan Makan Dinas",
    "Kebutuhan Kantor",
    "Karcis Parkir Kendaraan Operasional",
    "Isi BBM Kendaraan Operasional"
]

# --- 4. FORM INPUT (BEBAS ANGKA 0 & DUKUNG ENTER) ---
with st.form("form_kas", clear_on_submit=True):
    st.subheader("Input Data Transaksi")
    col1, col2 = st.columns(2)
    with col1:
        uraian_pilih = st.selectbox("Uraian", opsi_uraian)
        vendor = st.text_input("Nama Vendor", placeholder="Ketik vendor lalu tekan Tab...")
    with col2:
        tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
        # FIX: Menggunakan value=None agar tidak muncul angka 0 yang harus dihapus
        jumlah = st.number_input(
            "Jumlah (Nominal)", 
            min_value=0, 
            step=1000, 
            value=None, 
            placeholder="Masukkan nominal..."
        )
    
    submit = st.form_submit_button("Simpan ke Database")

# --- 5. LOGIKA SIMPAN ---
if submit:
    if vendor and (jumlah is not None):
        try:
            target_uraian = "Karcis Parkir Kendaraan Operasional"
            if uraian_pilih == target_uraian:
                bln, thn = tanggal_input.month, tanggal_input.year
                res = conn.table("kas_kecil").select("*").eq("uraian", target_uraian).execute()
                df_cek = pd.DataFrame(res.data)
                found = False
                if not df_cek.empty:
                    df_cek['t_dt'] = pd.to_datetime(df_cek['tanggal'])
                    match = df_cek[(df_cek['t_dt'].dt.month == bln) & (df_cek['t_dt'].dt.year == thn)]
                    if not match.empty:
                        id_old = int(match.iloc[0]['id'])
                        new_amt = int(match.iloc[0]['jumlah'] + jumlah)
                        conn.table("kas_kecil").delete().eq("id", id_old).execute()
                        conn.table("kas_kecil").insert({
                            "uraian": target_uraian, "vendor": vendor, 
                            "tanggal": str(tanggal_input), "jumlah": new_amt
                        }).execute()
                        found = True
                if not found:
                    conn.table("kas_kecil").insert({"uraian": target_uraian, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}).execute()
            else:
                conn.table("kas_kecil").insert({"uraian": uraian_pilih, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}).execute()
            
            st.success("Data Berhasil Disimpan!")
            time.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Gagal: {e}")
    else:
        st.warning("Mohon isi Nama Vendor dan Jumlah!")

# --- 6. REKAPITULASI ---
st.divider()
st.subheader("📋 Rekapitulasi Kas")

df_raw = fetch_data()
nama_bulan_id = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
                 7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

if not df_raw.empty:
    LIMIT_KAS = 25_000_000
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'], errors='coerce')
    
    # LOGIKA BATCHING
    df_raw['Kelompok_Sheet'] = "" 
    periods = df_raw[['tanggal_dt']].copy()
    periods['m'], periods['y'] = periods['tanggal_dt'].dt.month, periods['tanggal_dt'].dt.year
    unique_p = periods[['y', 'm']].drop_duplicates().values

    for y, m in unique_p:
        if pd.isna(m): continue
        curr_batch, run_total = 1, 0
        mask = (df_raw['tanggal_dt'].dt.year == y) & (df_raw['tanggal_dt'].dt.month == m)
        df_m = df_raw[mask].sort_values('id')
        for idx, row in df_m.iterrows():
            if run_total + row['jumlah'] > LIMIT_KAS:
                curr_batch += 1
                run_total = row['jumlah']
            else:
                run_total += row['jumlah']
            df_raw.at[idx, 'Kelompok_Sheet'] = f"{nama_bulan_id[int(m)]} ({curr_batch}) {int(y)}"

    # --- URUTAN TAMPILAN: TERBARU DI ATAS ---
    list_kelompok = sorted(
        [k for k in df_raw['Kelompok_Sheet'].unique() if k != ""],
        key=lambda x: (int(x.split(' ')[-1]), list(nama_bulan_id.values()).index(x.split(' ')[0])),
        reverse=True
    )

    for g in list_kelompok:
        df_g = df_raw[df_raw['Kelompok_Sheet'] == g].copy()
        tot_g = df_g['jumlah'].sum()
        with st.expander(f"📂 {g} | Total: Rp {tot_g:,.0f} | 💰 Sisa: Rp {LIMIT_KAS-tot_g:,.0f}".replace(",", "."), expanded=True):
            
            df_display = df_g.copy()
            df_display['No_Urut'] = range(1, len(df_display) + 1)
            df_display['Uraian_View'] = df_display.apply(lambda x: f"{x['No_Urut']} {x['uraian']}", axis=1)
            df_display['Tgl_View'] = df_display.apply(lambda x: "" if x['uraian'] == "Karcis Parkir Kendaraan Operasional" else x['tanggal'], axis=1)
            
            df_for_edit = df_display[['id', 'No_Urut', 'Uraian_View', 'vendor', 'Tgl_View', 'jumlah']].copy()
            df_for_edit.columns = ['id', 'No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            edited_df = st.data_editor(
                df_for_edit,
                key=f"ed_{g}",
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "id": None,
                    "No": st.column_config.NumberColumn("No", width="small", disabled=True),
                    "Uraian": st.column_config.TextColumn("Uraian", width="large"),
                    "Jumlah": st.column_config.NumberColumn("Jumlah", format="Rp %d")
                }
            )

            if st.button(f"💾 Simpan Perubahan {g}", key=f"save_{g}"):
                try:
                    ids_old = set(df_for_edit['id'].tolist())
                    ids_new = set(edited_df['id'].dropna().astype(int).tolist())
                    ids_del = ids_old - ids_new
                    for d_id in ids_del:
                        conn.table("kas_kecil").delete().eq("id", d_id).execute()
                    for _, row in edited_df.iterrows():
                        if pd.notna(row['id']):
                            u_asli = row['Uraian'].split(' ', 1)[-1] if ' ' in row['Uraian'] else row['Uraian']
                            conn.table("kas_kecil").update({
                                "uraian": u_asli, "vendor": row['Vendor'],
                                "tanggal": str(row['Tanggal']), "jumlah": int(row['Jumlah'])
                            }).eq("id", int(row['id'])).execute()
                    st.success("Berhasil Update!")
                    time.sleep(1); st.rerun()
                except Exception as e:
                    st.error(f"Gagal: {e}")

# --- 7. SIDEBAR ---
st.sidebar.markdown("### Simpan Rekap Kas Kecil")
if not df_raw.empty:
    if st.sidebar.button("💾 Siapkan Excel Format BIOS"):
        wb = Workbook()
        del wb['Sheet']
        # URUTAN EXCEL: TERLAMA DI DEPAN (Maret -> April)
        list_excel = sorted(
            [k for k in df_raw['Kelompok_Sheet'].unique() if k != ""],
            key=lambda x: (int(x.split(' ')[-1]), list(nama_bulan_id.values()).index(x.split(' ')[0]))
        )
        for p in list_excel:
            ws = wb.create_sheet(title=p[:31])
            ws.merge_cells('B2:I3')
            ws['B2'] = "APLIKASI BIOS (BIAYA OPERASIONAL)"; ws['B2'].font = Font(bold=True, color="FFFFFF", size=14)
            ws['B2'].alignment = Alignment(horizontal="center", vertical="center")
            ws['B2'].fill = PatternFill(start_color="CC9900", end_color="CC9900", fill_type="solid")
            ws.append([]); ws.append(["No", "URAIAN", "NAMA VENDOR", "POS", "GL", "TANGGAL", "JUMLAH", "NET"])
            df_b = df_raw[df_raw['Kelompok_Sheet'] == p]
            for i, r in enumerate(df_b.itertuples(), 1):
                t = "" if r.uraian == "Karcis Parkir Kendaraan Operasional" else str(r.tanggal)
                ws.append([i, f"{i} {r.uraian}", r.vendor, "", "", t, r.jumlah, r.jumlah])
        buf = BytesIO(); wb.save(buf)
        st.sidebar.download_button("⬇️ Download Excel", buf.getvalue(), "Rekap_BIOS.xlsx")

st.sidebar.divider()
st.sidebar.markdown("### Hapus Keseluruhan Data")
konf = st.sidebar.checkbox("Yakin hapus SEMUA?")
if st.sidebar.button("🗑️ Kosongkan Data", type="primary", disabled=not konf):
    conn.table("kas_kecil").delete().neq("id", 0).execute()
    st.rerun()

if 'df_raw' not in locals() or df_raw.empty:
    st.info("Belum ada data di Cloud Database.")
