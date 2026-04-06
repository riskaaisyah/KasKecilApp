import streamlit as st
import pandas as pd
from st_supabase_connection import SupabaseConnection
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import datetime
import time

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Aplikasi Kas Kecil", layout="wide")
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
        # Pastikan ID adalah integer agar sinkronisasi lancar
        df['id'] = df['id'].astype(int)
        return df.sort_values("id")
    except Exception:
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
    col1, col2 = st.columns(2)
    with col1:
        uraian_pilih = st.selectbox("Uraian", opsi_uraian)
        vendor = st.text_input("Nama Vendor", placeholder="Contoh: Toko Buku Gramedia")
    with col2:
        tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
        jumlah = st.number_input("Jumlah (Nominal)", min_value=0, step=1000, value=None)
    
    submit = st.form_submit_button("Simpan ke Cloud 🚀")

# --- LOGIKA SIMPAN BARU ---
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
                        conn.table("kas_kecil").insert({"uraian": target_uraian, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": new_amt}).execute()
                        found = True
                if not found:
                    conn.table("kas_kecil").insert({"uraian": target_uraian, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}).execute()
            else:
                conn.table("kas_kecil").insert({"uraian": uraian_pilih, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}).execute()
            
            st.success("Berhasil disimpan!")
            time.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Gagal: {e}")

# --- REKAPITULASI ---
st.divider()
st.subheader("📋 Rekapitulasi Kas (Limit 25jt/Sheet)")

df_raw = fetch_data()

if not df_raw.empty:
    LIMIT_KAS = 25_000_000
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'], errors='coerce')
    
    nama_bulan_id = {1: "JANUARI", 2: "FEBRUARI", 3: "MARET", 4: "APRIL", 5: "MEI", 6: "JUNI", 
                     7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

    # BATCHING LOGIC
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

    groups = [g for g in df_raw['Kelompok_Sheet'].unique() if g != ""][::-1]

    for g in groups:
        df_g = df_raw[df_raw['Kelompok_Sheet'] == g].copy()
        tot_g = df_g['jumlah'].sum()
        with st.expander(f"📂 {g} | Total: Rp {tot_g:,.0f} | Sisa: Rp {LIMIT_KAS-tot_g:,.0f}".replace(",", "."), expanded=True):
            
            # Form untuk edit/hapus agar tidak rerun sembarangan
            df_for_edit = df_g[['id', 'uraian', 'vendor', 'tanggal_dt', 'jumlah']].copy()
            df_for_edit.columns = ['id', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            edited_df = st.data_editor(
                df_for_edit,
                key=f"ed_{g}",
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "id": None,
                    "Uraian": st.column_config.SelectboxColumn("Uraian", options=opsi_uraian),
                    "Tanggal": st.column_config.DateColumn("Tanggal"),
                    "Jumlah": st.column_config.NumberColumn("Jumlah", format="%d")
                }
            )

            if st.button(f"💾 Simpan Perubahan {g}", key=f"save_{g}"):
                with st.spinner("Sedang menyelaraskan data..."):
                    try:
                        ids_old = set(df_for_edit['id'].tolist())
                        ids_new = set(edited_df['id'].dropna().astype(int).tolist())
                        ids_del = ids_old - ids_new
                        
                        # Jalankan Hapus
                        for d_id in ids_del:
                            conn.table("kas_kecil").delete().eq("id", d_id).execute()
                        
                        # Jalankan Update
                        for _, row in edited_df.iterrows():
                            if pd.notna(row['id']):
                                old_row = df_for_edit[df_for_edit['id'] == row['id']].iloc[0]
                                if not row.equals(old_row):
                                    conn.table("kas_kecil").update({
                                        "uraian": row['Uraian'], "vendor": row['Vendor'],
                                        "tanggal": str(row['Tanggal']), "jumlah": int(row['Jumlah'])
                                    }).eq("id", int(row['id'])).execute()
                        
                        st.success("Data Sinkron!")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal Simpan: {e}")

# --- SIDEBAR ---
st.sidebar.header("⚙️ Menu")
if not df_raw.empty:
    if st.sidebar.button("💾 Export to Excel BIOS"):
        wb = Workbook()
        del wb['Sheet']
        for p in df_raw['Kelompok_Sheet'].unique():
            if not p: continue
            ws = wb.create_sheet(title=p[:31])
            ws.merge_cells('B2:I3')
            ws['B2'] = "APLIKASI BIOS (BIAYA OPERASIONAL)"
            ws['B2'].font = Font(bold=True, color="FFFFFF", size=14)
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
konfirmasi = st.sidebar.checkbox("Yakin hapus SEMUA?")
if st.sidebar.button("🗑️ Kosongkan Data Cloud", type="primary", disabled=not konfirmasi):
    conn.table("kas_kecil").delete().neq("id", 0).execute()
    st.rerun()
