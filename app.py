import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="Aplikasi Kas Kecil", layout="centered")
st.title("💰 Aplikasi Kas Kecil")

LIMIT_KAS = 25_000_000

# ========================
# PILIH PERIODE
# ========================
bulan_list = ["JANUARI","FEBRUARI","MARET","APRIL","MEI","JUNI","JULI","AGUSTUS","SEPTEMBER","OKTOBER","NOVEMBER","DESEMBER"]
periode = st.selectbox("Pilih Periode", bulan_list)

# ========================
# SESSION STATE
# ========================
if "sheets" not in st.session_state:
    st.session_state.sheets = {
        f"{periode} 1": pd.DataFrame(columns=["No","Uraian","Vendor","Tanggal","Jumlah"])
    }

# ========================
# HITUNG SALDO (FITUR BARU)
# ========================
last_sheet_name = list(st.session_state.sheets.keys())[-1]
df_current = st.session_state.sheets[last_sheet_name]
total_terpakai = df_current["Jumlah"].sum() if not df_current.empty else 0
sisa_saldo = LIMIT_KAS - total_terpakai

# Tampilan Sisa Saldo di Atas
col1, col2 = st.columns(2)
with col1:
    st.metric("Total Terpakai", f"Rp {total_terpakai:,.0f}".replace(",", "."), delta=None)
with col2:
    # Warna delta hijau jika masih banyak, merah jika sisa dikit
    warna_delta = "normal" if sisa_saldo > 5_000_000 else "inverse"
    st.metric("Sisa Saldo Sheet Ini", f"Rp {sisa_saldo:,.0f}".replace(",", "."), delta=f"Limit: 25jt", delta_color=warna_delta)

st.divider()

# ========================
# FORM INPUT
# ========================
with st.form("form_kas", clear_on_submit=True):
    st.subheader(f"Input Data - {last_sheet_name}")
    opsi_uraian = [
        "Jamuan Makan Dinas",
        "Kebutuhan Kantor",
        "Karcis Parkir Kendaraan Operasional",
        "Isi BBM Kendaraan Operasional"
    ]

    uraian_pilih = st.selectbox("Uraian", opsi_uraian)
    vendor = st.text_input("Nama Vendor")
    tanggal = st.date_input("Tanggal")
    
    # Input Angka
    jumlah = st.number_input("Jumlah (Nominal)", min_value=0, step=1000, value=None, placeholder="Masukkan angka...")

    if jumlah:
        st.info(f"**Konfirmasi:** Rp {jumlah:,.0f}".replace(",", "."))
    
    submit = st.form_submit_button("Tambah ke Tabel")

# ========================
# LOGIKA INSERT
# ========================
if submit:
    if jumlah is None or vendor == "":
        st.error("Gagal: Nama Vendor dan Jumlah wajib diisi!")
    elif jumlah > LIMIT_KAS:
        st.error(f"Gagal: Jumlah input (Rp {jumlah:,.0f}) melebihi limit per sheet!")
    else:
        # Cek apakah harus pindah sheet
        if total_terpakai + jumlah <= LIMIT_KAS:
            target_sheet = last_sheet_name
            no_baru = len(df_current) + 1
        else:
            target_sheet = f"{periode} {len(st.session_state.sheets) + 1}"
            st.session_state.sheets[target_sheet] = pd.DataFrame(columns=["No","Uraian","Vendor","Tanggal","Jumlah"])
            no_baru = 1
            st.warning(f"Limit tercapai! Data otomatis dipindah ke sheet baru: {target_sheet}")

        uraian_final = f"{no_baru} {uraian_pilih}"
        new_row = pd.DataFrame({
            "No": [no_baru],
            "Uraian": [uraian_final],
            "Vendor": [vendor],
            "Tanggal": [tanggal],
            "Jumlah": [jumlah]
        })

        st.session_state.sheets[target_sheet] = pd.concat([st.session_state.sheets[target_sheet], new_row], ignore_index=True)
        st.rerun() # Refresh agar metrik saldo di atas langsung update

# ========================
# EDIT & TABEL DATA
# ========================
st.subheader("📋 Rekapitulasi Sheet")
for name, df in list(st.session_state.sheets.items()):
    with st.expander(f"Data {name} (Klik untuk Edit/Hapus)", expanded=(name == last_sheet_name)):
        df_display = df.copy()
        df_display["Jumlah"] = df_display["Jumlah"].apply(lambda x: f"{x:,.0f}".replace(",", "."))

        edited_df = st.data_editor(df_display, num_rows="dynamic", key=f"editor_{name}", use_container_width=True)

        if not edited_df.equals(df_display):
            try:
                # Balikin format titik ke angka
                edited_df["Jumlah"] = edited_df["Jumlah"].astype(str).str.replace(".", "", regex=False).astype(int)
                
                # Auto Renumber
                edited_df = edited_df.reset_index(drop=True)
                edited_df["No"] = edited_df.index + 1
                edited_df["Uraian"] = edited_df.apply(
                    lambda r: f"{r['No']} {' '.join(str(r['Uraian']).split(' ')[1:])}", axis=1
                )
                
                st.session_state.sheets[name] = edited_df
                st.rerun()
            except:
                st.error("Input jumlah tidak valid.")

# ========================
# EXPORT EXCEL
# ========================
def to_excel(sheets_dict):
    wb = Workbook()
    wb.remove(wb.active)
    thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    for name, df in sheets_dict.items():
        ws = wb.create_sheet(title=name)
        ws.append(["REKAPITULASI PENGGUNAAN KAS"])
        ws.merge_cells("A1:E1")
        ws["A1"].alignment = Alignment(horizontal="center")
        ws["A1"].font = Font(bold=True, size=12)

        ws.append([]) # Space
        ws.append(["No", "Uraian", "Nama Vendor", "Tanggal", "Jumlah"])
        
        for r in dataframe_to_rows(df, index=False, header=False):
            ws.append(r)

        # Styling
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, max_col=5):
            for cell in row:
                cell.border = thin

        # Footer Total
        total_row = ws.max_row + 1
        ws.cell(row=total_row, column=4, value="TOTAL").font = Font(bold=True)
        ws.cell(row=total_row, column=5, value=df["Jumlah"].sum()).font = Font(bold=True)
        ws.cell(row=total_row, column=5).number_format = '#,##0'

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

st.sidebar.title("Opsi")
if st.sidebar.button("💾 Siapkan File Excel"):
    file_excel = to_excel(st.session_state.sheets)
    st.sidebar.download_button(
        label="⬇️ Download Excel",
        data=file_excel,
        file_name=f"Kas_Kecil_{periode}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )