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

# --- FUNGSI GENERATE EXCEL (BIAR KONSISTEN) ---
def generate_excel(df_to_export, kelompok_names, nama_bulan_id):
    wb = Workbook()
    del wb['Sheet']
    
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))
    header_fill = PatternFill(start_color="00FFFF", end_color="00FFFF", fill_type="solid")
    title_fill = PatternFill(start_color="CC9900", end_color="CC9900", fill_type="solid")
    bbm_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
    parkir_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    total_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    sorted_groups = sorted(
        kelompok_names,
        key=lambda x: (int(x.split(' ')[-1]), list(nama_bulan_id.values()).index(x.split(' ')[0]))
    )

    for p in sorted_groups:
        ws = wb.create_sheet(title=p[:31])
        ws.merge_cells('B2:I3')
        ws['B2'] = "APLIKASI BIOS (BIAYA OPERASIONAL)"
        ws['B2'].font = Font(bold=True, color="FFFFFF", size=14)
        ws['B2'].alignment = Alignment(horizontal="center", vertical="center")
        ws['B2'].fill = title_fill

        headers = ["No", "URAIAN", "NAMA VENDOR", "POS MATA ANGGARAN", "GL ACCOUNT", 
                   "TANGGAL TRANSAKSI", "JUMLAH PENGGUNAAN", "SETELAH PPN", "PPN"]
        
        header_row = 5
        for col_num, cell_val in enumerate(headers, 1):
            cell = ws.cell(row=header_row, column=col_num)
            cell.value = cell_val
            cell.fill = header_fill
            cell.font = Font(bold=True, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        df_b = df_to_export[df_to_export['Kelompok_Sheet'] == p]
        current_row = header_row + 1
        for i, r in enumerate(df_b.itertuples(), 1):
            tgl = "" if r.uraian == "Karcis Parkir Kendaraan Operasional" else pd.to_datetime(r.tanggal).strftime('%d/%m/%Y')
            row_data = [i, f"{i} {r.uraian}", r.vendor, "", "", tgl, r.jumlah, r.jumlah, ""]
            ws.append(row_data)
            
            fill = bbm_fill if "Isi BBM" in r.uraian else (parkir_fill if "Karcis Parkir" in r.uraian else None)
            for col_idx in range(1, 10):
                cell = ws.cell(row=current_row, column=col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", horizontal="left" if col_idx <= 3 else "center")
                if fill: cell.fill = fill
                if col_idx in [7, 8]: cell.number_format = '#,##0'
            current_row += 1

        ws.cell(row=current_row, column=2).value = "TOTAL"
        ws.cell(row=current_row, column=2).font = Font(bold=True)
        ws.cell(row=current_row, column=7).value = f"=SUM(G{header_row+1}:G{current_row-1})"
        ws.cell(row=current_row, column=8).value = f"=SUM(H{header_row+1}:H{current_row-1})"
        for col_idx in range(1, 10):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = Font(bold=True); cell.fill = total_fill; cell.border = thin_border
            if col_idx in [7, 8]: cell.number_format = '#,##0'

        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['F'].width = 20
    
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# --- 4. FORM INPUT ---
opsi_uraian = ["Jamuan Makan Dinas", "Kebutuhan Kantor", "Karcis Parkir Kendaraan Operasional", "Isi BBM Kendaraan Operasional"]

with st.form("form_kas", clear_on_submit=True):
    st.subheader("Input Data Transaksi")
    c1, c2 = st.columns(2)
    with c1:
        uraian_pilih = st.selectbox("Uraian", opsi_uraian)
        vendor = st.text_input("Nama Vendor", placeholder="Ketik vendor...")
    with c2:
        tanggal_input = st.date_input("Tanggal", value=datetime.date.today())
        jumlah = st.number_input("Jumlah (Nominal)", min_value=0, step=1000, value=None, placeholder="Masukkan nominal...")
    submit = st.form_submit_button("Simpan ke Cloud 🚀")

if submit and vendor and jumlah:
    try:
        if uraian_pilih == "Karcis Parkir Kendaraan Operasional":
            res = conn.table("kas_kecil").select("*").eq("uraian", uraian_pilih).execute()
            df_cek = pd.DataFrame(res.data)
            found = False
            if not df_cek.empty:
                df_cek['t_dt'] = pd.to_datetime(df_cek['tanggal'])
                m, y = tanggal_input.month, tanggal_input.year
                match = df_cek[(df_cek['t_dt'].dt.month == m) & (df_cek['t_dt'].dt.year == y)]
                if not match.empty:
                    id_old = int(match.iloc[0]['id'])
                    new_amt = int(match.iloc[0]['jumlah'] + jumlah)
                    conn.table("kas_kecil").delete().eq("id", id_old).execute()
                    conn.table("kas_kecil").insert({"uraian": uraian_pilih, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": new_amt}).execute()
                    found = True
            if not found:
                conn.table("kas_kecil").insert({"uraian": uraian_pilih, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}).execute()
        else:
            conn.table("kas_kecil").insert({"uraian": uraian_pilih, "vendor": vendor, "tanggal": str(tanggal_input), "jumlah": int(jumlah)}).execute()
        st.success("Berhasil disimpan!"); time.sleep(1); st.rerun()
    except Exception as e: st.error(f"Gagal: {e}")

# --- 5. REKAPITULASI & BATCHING ---
st.divider()
df_raw = fetch_data()
nama_bulan_id = {1:"JANUARI", 2:"FEBRUARI", 3:"MARET", 4:"APRIL", 5:"MEI", 6: "JUNI", 7: "JULI", 8: "AGUSTUS", 9: "SEPTEMBER", 10: "OKTOBER", 11: "NOVEMBER", 12: "DESEMBER"}

if not df_raw.empty:
    LIMIT_KAS = 25_000_000
    df_raw['tanggal_dt'] = pd.to_datetime(df_raw['tanggal'], errors='coerce')
    df_raw['Kelompok_Sheet'] = ""
    for (y, m), group in df_raw.groupby([df_raw['tanggal_dt'].dt.year, df_raw['tanggal_dt'].dt.month]):
        if pd.isna(m): continue
        batch, run_total = 1, 0
        for idx, row in group.sort_values('id').iterrows():
            if run_total + row['jumlah'] > LIMIT_KAS:
                batch += 1; run_total = row['jumlah']
            else: run_total += row['jumlah']
            df_raw.at[idx, 'Kelompok_Sheet'] = f"{nama_bulan_id[int(m)]} ({batch}) {int(y)}"

    list_kelompok = sorted([k for k in df_raw['Kelompok_Sheet'].unique() if k != ""], 
                           key=lambda x: (int(x.split(' ')[-1]), list(nama_bulan_id.values()).index(x.split(' ')[0])), reverse=True)

    for g in list_kelompok:
        df_g = df_raw[df_raw['Kelompok_Sheet'] == g].copy()
        tot_g = df_g['jumlah'].sum()
        with st.expander(f"📂 {g} | Total: Rp {tot_g:,.0f} | 💰 Sisa: Rp {LIMIT_KAS-tot_g:,.0f}".replace(",","."), expanded=True):
            df_disp = df_g.copy()
            df_disp['No'] = range(1, len(df_disp)+1)
            df_disp['Uraian_View'] = df_disp.apply(lambda x: f"{x['No']} {x['uraian']}", axis=1)
            df_disp['Tgl_View'] = df_disp.apply(lambda x: "" if x['uraian'] == "Karcis Parkir Kendaraan Operasional" else x['tanggal'], axis=1)
            
            df_edit = df_disp[['id', 'No', 'Uraian_View', 'vendor', 'Tgl_View', 'jumlah']].copy()
            df_edit.columns = ['id', 'No', 'Uraian', 'Vendor', 'Tanggal', 'Jumlah']
            
            # FIX: Format titik ribuan dengan format="Rp %d"
            edited = st.data_editor(df_edit, key=f"ed_{g}", num_rows="dynamic", use_container_width=True,
                                    column_config={
                                        "id": None, 
                                        "No": st.column_config.NumberColumn(disabled=True), 
                                        "Jumlah": st.column_config.NumberColumn(format="Rp %d") # Titik muncul di sini
                                    })

            if st.button(f"💾 Simpan Perubahan {g}", key=f"btn_{g}"):
                try:
                    ids_old = set(df_edit['id'].tolist()); ids_new = set(edited['id'].dropna().astype(int).tolist())
                    for d_id in (ids_old - ids_new): conn.table("kas_kecil").delete().eq("id", d_id).execute()
                    for _, row in edited.iterrows():
                        if pd.notna(row['id']):
                            u = row['Uraian'].split(' ', 1)[-1] if ' ' in row['Uraian'] else row['Uraian']
                            conn.table("kas_kecil").update({"uraian":u, "vendor":row['Vendor'], "tanggal":str(row['Tanggal']), "jumlah":int(row['Jumlah'])}).eq("id", int(row['id'])).execute()
                    st.success("Berhasil!"); time.sleep(1); st.rerun()
                except Exception as e: st.error(f"Gagal: {e}")

# --- 6. SIDEBAR ---
st.sidebar.markdown("### 💾 Simpan Rekap Kas Kecil")
if not df_raw.empty:
    all_kelompok = sorted([k for k in df_raw['Kelompok_Sheet'].unique() if k != ""], 
                          key=lambda x: (int(x.split(' ')[-1]), list(nama_bulan_id.values()).index(x.split(' ')[0])))
    
    if st.sidebar.button("💾 Siapkan Excel (Semua Tabel)"):
        data = generate_excel(df_raw, all_kelompok, nama_bulan_id)
        st.sidebar.download_button("⬇️ Download Semua", data, f"Rekap_Lengkap_{datetime.date.today()}.xlsx")
    
    st.sidebar.divider()
    st.sidebar.markdown("#### Download Per Tabel")
    pilihan_tabel = st.sidebar.selectbox("Pilih Tabel:", all_kelompok)
    if st.sidebar.button(f"💾 Siapkan Excel {pilihan_tabel}"):
        data_single = generate_excel(df_raw, [pilihan_tabel], nama_bulan_id)
        st.sidebar.download_button(f"⬇️ Download {pilihan_tabel}", data_single, f"Rekap_{pilihan_tabel.replace(' ', '_')}.xlsx")

st.sidebar.divider()
st.sidebar.markdown("### 🗑️ Hapus Keseluruhan Data")
konf = st.sidebar.checkbox("Saya Ingin Menghapus Semua Data")
if st.sidebar.button("🗑️ Kosongkan Data", type="primary", disabled=not konf):
    conn.table("kas_kecil").delete().neq("id", 0).execute()
    st.rerun()
