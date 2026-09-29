import streamlit as st
import pandas as pd
import numpy as np
import re
import calendar
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
from openpyxl.utils import get_column_letter

# Pengaturan Halaman Streamlit
st.set_page_config(
    page_title="Pengelolaan Data PCP | Dashboard",
    page_icon="🚀",
    layout="wide"
)

# ==========================================
# STYLING KUSTOM: TEMA FUTURISTIK / MODERN
# ==========================================
st.markdown("""
<style>
    /* Mengubah latar belakang utama agar senada dengan tema gelap modern */
    .stApp {
        background-color: #0d1117;
        color: #c9d1d9;
    }
    
    /* Styling Header Utama */
    h1 {
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        background: linear-gradient(90deg, #58a6ff, #bc8cff, #3fb950);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        margin-bottom: 0px;
    }
    
    /* Styling Kartu / Kontainer */
    .metric-card {
        background: rgba(22, 27, 34, 0.7);
        border: 1px solid #30363d;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.5);
        backdrop-filter: blur(10px);
        text-align: center;
    }

    /* Styling Sidebar */
    [data-testid="stSidebar"] {
        background-color: #161b22;
        border-right: 1px solid #30363d;
    }

    /* Efek Tombol Futuristik */
    .stButton>button {
        background: linear-gradient(135deg, #1f6feb 0%, #238636 100%);
        color: white;
        border-radius: 8px;
        border: none;
        font-weight: 600;
        transition: all 0.3s ease;
        box-shadow: 0 0 15px rgba(31, 111, 235, 0.4);
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 0 25px rgba(35, 134, 54, 0.8);
    }

    /* Styling Tab Menu */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #161b22;
        padding: 6px;
        border-radius: 10px;
        border: 1px solid #30363d;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: transparent;
        border-radius: 6px;
        color: #8b949e;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #1f6feb, #388bfd) !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Judul & Deskripsi Futuristik
st.markdown("<h1>PENGELOLAAN DATA PCP</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #8b949e; font-size: 16px;'>Sistem Otomasi Pintar & Rekapitulasi Data Produksi Berbasis Web</p>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# SIDEBAR: PENGATURAN & UNGGAH BERKAS
# ==========================================
st.sidebar.markdown("### 🎛️ Panel Kontrol Utama")
st.sidebar.markdown("---")

uploaded_master = st.sidebar.file_uploader(
    "📁 1. Master Ekspor Sistem (.xlsx)", 
    type=["xlsx"],
    help="Unggah file laporan utama dari sistem."
)

uploaded_indicator = st.sidebar.file_uploader(
    "📁 2. Master Indikator (.xlsx)", 
    type=["xlsx"],
    help="Unggah file acuan target, waktu, dan spindle."
)

st.sidebar.markdown("---")
st.sidebar.markdown("### ⚙️ Parameter Pengolahan")

bulan_pilihan = st.sidebar.selectbox(
    "🗓️ Pilih Bulan Laporan",
    ["JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
     "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"],
    index=8
)

threshold_val = st.sidebar.number_input(
    "🎯 Nilai Batas Minimum (%)",
    min_value=0.0,
    max_value=100.0,
    value=60.0,
    step=1.0,
    help="Batas minimum persentase pencapaian"
)

st.sidebar.markdown("---")
proses_btn = st.sidebar.button("⚡ EKSEKUSI PROSES DATA", use_container_width=True)

# Fungsi Normalisasi Teks
def clean_str(val):
    if pd.isna(val):
        return ""
    return re.sub(r'\s+', '', str(val)).upper()

# Fungsi Utama Pemrosesan Data
if proses_btn:
    if not uploaded_master or not uploaded_indicator:
        st.error("⚠️ Peringatan: Mohon unggah kedua file master terlebih dahulu melalui panel kontrol di sebelah kiri!")
    else:
        with st.spinner("🚀 Sistem sedang memproses algoritma data, harap tunggu..."):
            try:
                # ==========================================
                # TAHAP 1: PEMBUATAN DATA MASTER BERSIH
                # ==========================================
                df_raw = pd.read_excel(uploaded_master, header=None)

                current_tanggal = ""
                current_shift = ""
                current_nip = ""
                current_pegawai = ""
                current_group = ""

                parsed_rows = []
                temp_group_rows = []
                temp_sub_qty = 0.0
                temp_sub_total = 0.0

                def flush_group_rows(group_rows, sub_q, sub_t):
                    if len(group_rows) > 0:
                        for r in group_rows:
                            r['Total Qty Cones'] = sub_q
                            r['Total Qty KG'] = sub_t
                            parsed_rows.append(r)

                for idx, row in df_raw.iterrows():
                    row_vals = [str(row[i]).strip() for i in range(len(row)) if pd.notna(row[i]) and str(row[i]).strip() != ""]
                    if not row_vals:
                        continue

                    row_text = " ".join(row_vals)
                    if "Tanggal" in row_text and "Pegawai" in row_text and "Benang" in row_text:
                        continue

                    for v in row_vals:
                        if v.startswith("Tanggal:"):
                            current_tanggal = v.replace("Tanggal:", "").strip()
                            break

                    for v in row_vals:
                        if v.startswith("Shift:"):
                            current_shift = v.replace("Shift:", "").strip()
                            break

                    if len(row_vals) == 1 and row_vals[0].startswith("NIP:"):
                        current_nip = row_vals[0].replace("NIP:", "").strip()
                        continue

                    is_subtotal = False
                    if len(row_vals) <= 3 and any(v.replace('.', '', 1).isdigit() for v in row_vals):
                        is_subtotal = True

                    if is_subtotal:
                        numbers = []
                        for v in row_vals:
                            try:
                                numbers.append(float(v))
                            except ValueError:
                                pass

                        if len(numbers) >= 2:
                            temp_sub_qty = numbers[-2]
                            temp_sub_total = numbers[-1]
                        elif len(numbers) == 1:
                            temp_sub_total = numbers[0]

                        flush_group_rows(temp_group_rows, temp_sub_qty, temp_sub_total)
                        temp_group_rows = []
                        temp_sub_qty = 0.0
                        temp_sub_total = 0.0
                        continue

                    if len(row_vals) >= 4:
                        try:
                            potential_qty = float(row_vals[-2])
                            potential_kg = float(row_vals[-1])
                            is_transaction = True
                        except ValueError:
                            is_transaction = False

                        if is_transaction:
                            total_kg = float(row_vals[-1])
                            qty_cones = float(row_vals[-2])
                            benang = row_vals[-3]
                            mesin = row_vals[-4]
                            group_pegawai = row_vals[-5] if len(row_vals) >= 5 else current_group
                            pegawai = row_vals[-6] if len(row_vals) >= 6 else current_pegawai
                            nip_val = row_vals[-7] if len(row_vals) >= 7 and "NIP" in row_vals[-7] else current_nip

                            if nip_val:
                                current_nip = nip_val.replace("NIP:", "").strip()
                            if pegawai and not pegawai.startswith("SSM") and not pegawai.startswith("PT") and not pegawai.startswith("FADIS"):
                                current_pegawai = pegawai
                            if group_pegawai:
                                current_group = group_pegawai

                            if not mesin or "SSM" not in mesin.upper():
                                continue
                            if qty_cones == 0:
                                continue

                            row_dict = {
                                'Tanggal': current_tanggal,
                                'Shift': current_shift,
                                'NIP': current_nip,
                                'Pegawai': current_pegawai,
                                'Mesin': mesin,
                                'Group Pegawai': current_group,
                                'Benang': benang,
                                'Qty Cones': qty_cones,
                                'Total': total_kg,
                                'Total Qty Cones': 0.0,
                                'Total Qty KG': 0.0
                            }
                            temp_group_rows.append(row_dict)

                df_clean = pd.DataFrame(parsed_rows)
                final_columns = [
                    'Tanggal', 'Shift', 'NIP', 'Pegawai', 'Mesin',
                    'Group Pegawai', 'Benang', 'Qty Cones', 'Total',
                    'Total Qty Cones', 'Total Qty KG'
                ]
                if not df_clean.empty:
                    df_clean = df_clean[final_columns]

                # Simpan DATA BERSIH ke Excel Bytes
                output_clean_io = BytesIO()
                wb_clean = Workbook()
                ws_c = wb_clean.active
                ws_c.title = "Data Bersih"
                ws_c.append(final_columns)
                for _, row_data in df_clean.iterrows():
                    ws_c.append(list(row_data))
                
                for col in ws_c.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_c.column_dimensions[col_letter].width = max(max_len + 3, 12)

                wb_clean.save(output_clean_io)
                output_clean_bytes = output_clean_io.getvalue()

                # ==========================================
                # TAHAP 2 & 3: REKAP KARYAWAN & MESIN
                # ==========================================
                df_master = pd.read_excel(uploaded_indicator)
                df_clean['Mesin_Clean'] = df_clean['Mesin'].apply(clean_str)
                df_master['clean_var'] = df_master['Nama Var.'].apply(clean_str)

                spindle_col_name = None
                for col in df_master.columns:
                    col_up = str(col).upper()
                    if 'SPIND' in col_up or 'SPIN' in col_up:
                        spindle_col_name = col
                        break
                if spindle_col_name is None:
                    spindle_col_name = 'SPINDEL'

                def get_master_data(benang_val, mesin_val, col_name):
                    b_clean = clean_str(benang_val)
                    m_clean = clean_str(mesin_val)
                    matched = df_master[df_master['clean_var'] == b_clean]
                    if matched.empty:
                        return 0.0
                    for idx, row in matched.iterrows():
                        master_machine_str = str(row['MESIN'])
                        machines = [clean_str(m) for m in master_machine_str.split('/')]
                        if m_clean in machines or any(m in m_clean or m_clean in m for m in machines):
                            try:
                                val = float(row[col_name])
                                if val > 0:
                                    return val
                            except:
                                pass
                    try:
                        return float(matched.iloc[0][col_name])
                    except:
                        return 0.0

                shift_counts = df_clean.groupby(['Tanggal', 'Pegawai', 'Mesin_Clean'])['Shift'].nunique().reset_index()
                shift_counts['Status_Shift'] = shift_counts['Shift'].apply(lambda x: 'PANJANG' if x > 1 else 'PENDEK')
                df_clean = pd.merge(df_clean, shift_counts[['Tanggal', 'Pegawai', 'Mesin_Clean', 'Status_Shift']], on=['Tanggal', 'Pegawai', 'Mesin_Clean'], how='left')

                def determine_shift_final(row):
                    shift_val = str(row['Shift']).strip().upper()
                    status_val = row['Status_Shift']
                    if status_val == 'PENDEK':
                        return f"{shift_val} / PENDEK"
                    else:
                        group_rows = df_clean[
                            (df_clean['Tanggal'] == row['Tanggal']) &
                            (df_clean['Pegawai'] == row['Pegawai']) &
                            (df_clean['Mesin_Clean'] == row['Mesin_Clean'])
                        ]
                        shifts_in_group = set(group_rows['Shift'].astype(str).str.strip().str.upper())
                        if 'SIANG' in shifts_in_group and 'MALAM' in shifts_in_group:
                            return "MALAM / PANJANG"
                        elif 'PAGI' in shifts_in_group and 'SIANG' in shifts_in_group:
                            return "PAGI / PANJANG"
                        else:
                            return f"{shift_val} / PANJANG"

                df_clean['Shift_Final'] = df_clean.apply(determine_shift_final, axis=1)

                group_max_qty = df_clean.groupby(['Tanggal', 'Pegawai'])['Qty Cones'].max().reset_index()
                group_max_qty = group_max_qty.rename(columns={'Qty Cones': 'max_qty'})
                df_clean = pd.merge(df_clean, group_max_qty, on=['Tanggal', 'Pegawai'], how='left')

                real_cones_list = []
                for idx, row in df_clean.iterrows():
                    q_val = row['Qty Cones']
                    max_q = row['max_qty']
                    if q_val == max_q:
                        real_cones_list.append(round(q_val))
                    else:
                        b_val = row['Benang']
                        m_val = row['Mesin']
                        waktu_current = get_master_data(b_val, m_val, 'WAKTU (MENIT)')
                        group_rows = df_clean[(df_clean['Tanggal'] == row['Tanggal']) & (df_clean['Pegawai'] == row['Pegawai'])]
                        row_max = group_rows[group_rows['Qty Cones'] == max_q].iloc[0]
                        waktu_max_qty_item = get_master_data(row_max['Benang'], row_max['Mesin'], 'WAKTU (MENIT)')
                        if waktu_max_qty_item > 0:
                            real_cones = (q_val * waktu_current) / waktu_max_qty_item
                        else:
                            real_cones = q_val
                        real_cones_list.append(round(real_cones))
                df_clean['Qty Real Cones'] = real_cones_list

                # --- BUAT REKAP KARYAWAN (EXCEL) ---
                wb_emp = Workbook()
                wb_emp.remove(wb_emp.active)
                employees = df_clean['Pegawai'].unique()
                thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))

                for emp in employees:
                    sheet_name = str(emp)[:31].replace('/', '_').replace('\\', '_')
                    ws = wb_emp.create_sheet(title=sheet_name)
                    df_emp = df_clean[df_clean['Pegawai'] == emp].copy()

                    ws.merge_cells('F1:G1')
                    ws.cell(row=1, column=1, value="TANGGAL")
                    ws.cell(row=1, column=2, value="NAMA")
                    ws.cell(row=1, column=3, value="MESIN")
                    ws.cell(row=1, column=4, value="JUMLAH SPINDLE")
                    ws.cell(row=1, column=5, value="SHIFT")
                    ws.cell(row=1, column=6, value="HASIL SOFTWINDING")
                    ws.cell(row=1, column=8, value="HASIL TOTAL\n( CONES )")
                    ws.cell(row=1, column=9, value="TARGET 90%\n( CONES )")
                    ws.cell(row=1, column=10, value="PERSENTASE")
                    ws.cell(row=1, column=11, value="TARGET 100%\n( CONES )")
                    ws.cell(row=1, column=12, value="PERSENTASE")
                    ws.cell(row=1, column=13, value="CONES")
                    ws.cell(row=1, column=14, value="KG")

                    ws.cell(row=2, column=6, value="JENIS BENANG")
                    ws.cell(row=2, column=7, value="JUMLAH CONES")
                    for col_idx in [1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14]:
                        ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)

                    current_row = 3
                    df_emp['Tanggal_Parsed'] = pd.to_datetime(df_emp['Tanggal'], errors='coerce')
                    date_machines = df_emp.sort_values('Tanggal_Parsed')[['Tanggal', 'Mesin_Clean']].drop_duplicates().values

                    for dt, m_clean in date_machines:
                        df_group = df_emp[(df_emp['Tanggal'] == dt) & (df_emp['Mesin_Clean'] == m_clean)]
                        start_group_row = current_row
                        total_real_cones = int(df_group['Qty Real Cones'].sum()) if df_group['Qty Real Cones'].sum().is_integer() else df_group['Qty Real Cones'].sum()
                        shift_status = df_group['Status_Shift'].iloc[0]
                        max_row_in_group = df_group.loc[df_group['Qty Cones'].idxmax()]
                        ref_benang = max_row_in_group['Benang']
                        ref_mesin = max_row_in_group['Mesin']

                        spindle_val = get_master_data(ref_benang, ref_mesin, spindle_col_name)
                        if spindle_val.is_integer():
                            spindle_val = int(spindle_val)

                        col_t90 = 'TARGET 90% SHIFT PENDEK' if shift_status == 'PENDEK' else 'TARGET 90% SHIFT PANJANG'
                        col_t100 = 'TARGET 100% SHIFT PENDEK' if shift_status == 'PENDEK' else 'TARGET 100% SHIFT PANJANG'
                        target_90 = get_master_data(ref_benang, ref_mesin, col_t90)
                        if target_90.is_integer():
                            target_90 = int(target_90)
                        target_100 = get_master_data(ref_benang, ref_mesin, col_t100)
                        if target_100.is_integer():
                            target_100 = int(target_100)

                        for _, row in df_group.iterrows():
                            benang = row['Benang']
                            mesin = row['Mesin']
                            qty_cones_val = row['Qty Cones']
                            if qty_cones_val.is_integer():
                                qty_cones_val = int(qty_cones_val)
                            kg_val = row['Total'] if 'Total' in row else 0.0
                            q_real_val = row['Qty Real Cones']
                            if q_real_val.is_integer():
                                q_real_val = int(q_real_val)

                            ws.cell(row=current_row, column=1, value=str(row['Tanggal'])[:10])
                            ws.cell(row=current_row, column=2, value=row['Pegawai'])
                            ws.cell(row=current_row, column=3, value=mesin)
                            ws.cell(row=current_row, column=4, value=spindle_val).number_format = '#,##0'
                            ws.cell(row=current_row, column=5, value=row['Shift_Final'])
                            ws.cell(row=current_row, column=6, value=benang)
                            ws.cell(row=current_row, column=7, value=q_real_val).number_format = '#,##0'
                            ws.cell(row=current_row, column=8, value=total_real_cones).number_format = '#,##0'
                            ws.cell(row=current_row, column=9, value=target_90).number_format = '#,##0' if isinstance(target_90, int) else '#,##0.00'
                            
                            p90 = (total_real_cones / target_90) if target_90 > 0 else 0.0
                            ws.cell(row=current_row, column=10, value=p90).number_format = '0.0%'
                            ws.cell(row=current_row, column=11, value=target_100).number_format = '#,##0' if isinstance(target_100, int) else '#,##0.00'
                            
                            p100 = (total_real_cones / target_100) if target_100 > 0 else 0.0
                            ws.cell(row=current_row, column=12, value=p100).number_format = '0.0%'
                            ws.cell(row=current_row, column=13, value=qty_cones_val).number_format = '#,##0'
                            ws.cell(row=current_row, column=14, value=kg_val).number_format = '#,##0.00'
                            current_row += 1

                        end_group_row = current_row - 1
                        if end_group_row > start_group_row:
                            for col_idx in [1, 2, 3, 4, 5, 8, 9, 10, 11, 12]:
                                ws.merge_cells(start_row=start_group_row, start_column=col_idx, end_row=end_group_row, end_column=col_idx)

                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=14):
                        for cell in row:
                            cell.border = thin_border
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                    min_widths = {1: 14, 2: 18, 3: 12, 4: 15, 5: 18, 6: 25, 7: 14}
                    for col_idx in range(1, 15):
                        col_letter = get_column_letter(col_idx)
                        max_len = 0
                        for cell in ws[col_letter]:
                            val_str = str(cell.value or '')
                            for line in val_str.split('\n'):
                                if len(line) > max_len:
                                    max_len = len(line)
                        calculated_width = max(max_len + 3, 12)
                        if col_idx in min_widths:
                            calculated_width = max(calculated_width, min_widths[col_idx])
                        ws.column_dimensions[col_letter].width = calculated_width

                output_emp_io = BytesIO()
                wb_emp.save(output_emp_io)
                output_emp_bytes = output_emp_io.getvalue()

                # --- BUAT REKAP MESIN (EXCEL) ---
                wb_mach = Workbook()
                wb_mach.remove(wb_mach.active)
                machines_raw = df_clean['Mesin'].unique()

                def machine_sort_key(m):
                    m_str = str(m).strip().upper()
                    custom_order = {'SSM 1A': 1, 'SSM 1B': 2, 'SSM 2A': 3, 'SSM 2B': 4, 'SSM 3A': 5, 'SSM 3B': 6}
                    for key, rank in custom_order.items():
                        if key in m_str:
                            return (0, rank, m_str)
                    return (1, 99, m_str)

                machines = sorted(machines_raw, key=machine_sort_key)
                header_fill = PatternFill(start_color='9ED3DC', end_color='9ED3DC', fill_type='solid')
                color_list = ['BDB2FF', 'FFCEE3', 'FDFFD2', 'CAEDFF']

                for mach in machines:
                    safe_mach_name = str(mach).strip()
                    sheet_name = re.sub(r'[:\\/?*\[\]]', '_', safe_mach_name)[:31]
                    if not sheet_name:
                        sheet_name = "MESIN"
                    ws = wb_mach.create_sheet(title=sheet_name)
                    df_mach = df_clean[df_clean['Mesin'] == mach].copy()

                    ws.merge_cells('F1:G1')
                    ws.cell(row=1, column=1, value="TANGGAL")
                    ws.cell(row=1, column=2, value="NAMA")
                    ws.cell(row=1, column=3, value="MESIN")
                    ws.cell(row=1, column=4, value="JUMLAH SPINDLE")
                    ws.cell(row=1, column=5, value="SHIFT")
                    ws.cell(row=1, column=6, value="HASIL SOFTWINDING")
                    ws.cell(row=1, column=8, value="HASIL TOTAL\n( CONES )")
                    ws.cell(row=1, column=9, value="TARGET 90%\n( CONES )")
                    ws.cell(row=1, column=10, value="PERSENTASE")
                    ws.cell(row=1, column=11, value="TARGET 100%\n( CONES )")
                    ws.cell(row=1, column=12, value="PERSENTASE")
                    ws.cell(row=1, column=13, value="CONES")
                    ws.cell(row=1, column=14, value="KG")

                    ws.cell(row=2, column=6, value="JENIS BENANG")
                    ws.cell(row=2, column=7, value="JUMLAH CONES")
                    for col_idx in [1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14]:
                        ws.merge_cells(start_row=1, start_column=col_idx, end_row=2, end_column=col_idx)

                    for r in [1, 2]:
                        for c in range(1, 15):
                            ws.cell(row=r, column=c).fill = header_fill

                    current_row = 3
                    df_mach['Tanggal_Parsed'] = pd.to_datetime(df_mach['Tanggal'], errors='coerce')
                    group_keys = df_mach.sort_values(['Tanggal_Parsed', 'Pegawai', 'Shift'])[['Tanggal', 'Pegawai', 'Shift']].drop_duplicates().values
                    unique_dates = sorted(df_mach['Tanggal_Parsed'].dropna().unique())
                    date_to_color_idx = {dt: idx % len(color_list) for idx, dt in enumerate(unique_dates)}

                    for dt, peg, shf in group_keys:
                        df_group = df_mach[(df_mach['Tanggal'] == dt) & (df_mach['Pegawai'] == peg) & (df_mach['Shift'] == shf)]
                        if df_group.empty:
                            continue

                        start_group_row = current_row
                        total_real_cones = int(df_group['Qty Real Cones'].sum()) if df_group['Qty Real Cones'].sum().is_integer() else df_group['Qty Real Cones'].sum()
                        shift_status = df_group['Status_Shift'].iloc[0]
                        max_row_in_group = df_group.loc[df_group['Qty Cones'].idxmax()]
                        ref_benang = max_row_in_group['Benang']
                        ref_mesin = max_row_in_group['Mesin']

                        spindle_val = get_master_data(ref_benang, ref_mesin, spindle_col_name)
                        if spindle_val.is_integer():
                            spindle_val = int(spindle_val)

                        col_t90 = 'TARGET 90% SHIFT PENDEK' if shift_status == 'PENDEK' else 'TARGET 90% SHIFT PANJANG'
                        col_t100 = 'TARGET 100% SHIFT PENDEK' if shift_status == 'PENDEK' else 'TARGET 100% SHIFT PANJANG'
                        target_90 = get_master_data(ref_benang, ref_mesin, col_t90)
                        if target_90.is_integer():
                            target_90 = int(target_90)
                        target_100 = get_master_data(ref_benang, ref_mesin, col_t100)
                        if target_100.is_integer():
                            target_100 = int(target_100)

                        for _, row in df_group.iterrows():
                            benang = row['Benang']
                            mesin_item = row['Mesin']
                            qty_cones_val = row['Qty Cones']
                            if qty_cones_val.is_integer():
                                qty_cones_val = int(qty_cones_val)
                            kg_val = row['Total'] if 'Total' in row else 0.0
                            q_real_val = row['Qty Real Cones']
                            if q_real_val.is_integer():
                                q_real_val = int(q_real_val)

                            ws.cell(row=current_row, column=1, value=str(row['Tanggal'])[:10])
                            ws.cell(row=current_row, column=2, value=row['Pegawai'])
                            ws.cell(row=current_row, column=3, value=mesin_item)
                            ws.cell(row=current_row, column=4, value=spindle_val).number_format = '#,##0'
                            ws.cell(row=current_row, column=5, value=row['Shift_Final'])
                            ws.cell(row=current_row, column=6, value=benang)
                            ws.cell(row=current_row, column=7, value=q_real_val).number_format = '#,##0'
                            ws.cell(row=current_row, column=8, value=total_real_cones).number_format = '#,##0'
                            ws.cell(row=current_row, column=9, value=target_90).number_format = '#,##0' if isinstance(target_90, int) else '#,##0.00'
                            
                            p90 = (total_real_cones / target_90) if target_90 > 0 else 0.0
                            ws.cell(row=current_row, column=10, value=p90).number_format = '0.0%'
                            ws.cell(row=current_row, column=11, value=target_100).number_format = '#,##0' if isinstance(target_100, int) else '#,##0.00'
                            
                            p100 = (total_real_cones / target_100) if target_100 > 0 else 0.0
                            ws.cell(row=current_row, column=12, value=p100).number_format = '0.0%'
                            ws.cell(row=current_row, column=13, value=qty_cones_val).number_format = '#,##0'
                            ws.cell(row=current_row, column=14, value=kg_val).number_format = '#,##0.00'
                            current_row += 1

                        end_group_row = current_row - 1
                        if end_group_row > start_group_row:
                            for col_idx in [1, 2, 3, 4, 5, 8, 9, 10, 11, 12]:
                                ws.merge_cells(start_row=start_group_row, start_column=col_idx, end_row=end_group_row, end_column=col_idx)

                        dt_parsed = pd.to_datetime(dt, errors='coerce')
                        if dt_parsed in date_to_color_idx:
                            c_idx = date_to_color_idx[dt_parsed]
                            hex_color = color_list[c_idx]
                            group_fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')
                            for r in range(start_group_row, end_group_row + 1):
                                for c in range(1, 15):
                                    ws.cell(row=r, column=c).fill = group_fill

                    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=14):
                        for cell in row:
                            cell.border = thin_border
                            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

                    min_widths = {1: 14, 2: 18, 3: 12, 4: 15, 5: 18, 6: 25, 7: 14}
                    for col_idx in range(1, 15):
                        col_letter = get_column_letter(col_idx)
                        max_len = 0
                        for cell in ws[col_letter]:
                            val_str = str(cell.value or '')
                            for line in val_str.split('\n'):
                                if len(line) > max_len:
                                    max_len = len(line)
                        calculated_width = max(max_len + 3, 12)
                        if col_idx in min_widths:
                            calculated_width = max(calculated_width, min_widths[col_idx])
                        ws.column_dimensions[col_letter].width = calculated_width

                output_mach_io = BytesIO()
                wb_mach.save(output_mach_io)
                output_mach_bytes = output_mach_io.getvalue()

                output_mach_io.seek(0)
                wb_src = load_workbook(output_mach_io, data_only=True)
                sheet_names = wb_src.sheetnames

                # ==========================================
                # TAHAP 4: SUMMARY REKAP PER MINGGU
                # ==========================================
                tahun_input = 2026
                bulan_dict = {
                    'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4, 'MEI': 5, 'JUNI': 6,
                    'JULI': 7, 'AGUSTUS': 8, 'SEPTEMBER': 9, 'OKTOBER': 10, 'NOVEMBER': 11, 'DESEMBER': 12
                }
                target_month_num = bulan_dict.get(bulan_pilihan, 9)

                cal = calendar.Calendar(firstweekday=6)
                month_weeks = cal.monthdatescalendar(tahun_input, target_month_num)

                date_to_week_map = {}
                for week_idx, week in enumerate(month_weeks):
                    w_num = week_idx + 1
                    for day_date in week:
                        date_to_week_map[day_date] = w_num

                def get_calendar_week_mapping(dt):
                    if pd.isna(dt):
                        return 1
                    try:
                        dt_date = pd.to_datetime(dt).date()
                        if dt_date in date_to_week_map:
                            return date_to_week_map[dt_date]
                        first_day_grid = month_weeks[0][0]
                        last_day_grid = month_weeks[-1][-1]
                        if dt_date < first_day_grid:
                            return 1
                        elif dt_date > last_day_grid:
                            return len(month_weeks)
                        return 1
                    except:
                        return 1

                def parse_percentage_value(val_raw):
                    if val_raw is None or pd.isna(val_raw):
                        return None
                    try:
                        if isinstance(val_raw, (int, float)):
                            f_val = float(val_raw)
                            if f_val <= 1.5:
                                return f_val * 100.0
                            else:
                                return f_val
                        else:
                            p_str = str(val_raw).replace('%', '').replace(',', '.').strip()
                            return float(p_str)
                    except:
                        return None

                wb_out = Workbook()
                wb_out.remove(wb_out.active)

                days_map = {0: 'SENIN', 1: 'SELASA', 2: 'RABU', 3: 'KAMIS', 4: 'JUMAT', 5: 'SABTU', 6: 'MINGGU'}
                days_order = ['SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU']
                bold_font = Font(name='Calibri', size=11, bold=True)
                header_font = Font(name='Calibri', size=11, bold=True, color='000000')
                sub_header_font = Font(name='Calibri', size=10, bold=True, color='000000')
                highlight_fill = PatternFill(start_color='E11A45', end_color='E11A45', fill_type='solid')
                highlight_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')

                kelompok_dict = {}
                for s_name in sheet_names:
                    s_upper = s_name.upper()
                    if 'SSM 1' in s_upper or 'SSM1' in s_upper:
                        group_key = 'SSM 1'
                    elif 'SSM 2' in s_upper or 'SSM2' in s_upper:
                        group_key = 'SSM 2'
                    elif 'SSM 3' in s_upper or 'SSM3' in s_upper:
                        group_key = 'SSM 3'
                    else:
                        group_key = s_upper.split()[0] if s_upper else 'LAINNYA'
                    if group_key not in kelompok_dict:
                        kelompok_dict[group_key] = []
                    kelompok_dict[group_key].append(s_name)

                output_mach_io.seek(0)
                temp_df_dict = pd.read_excel(output_mach_io, sheet_name=None)

                for group_name, s_list in kelompok_dict.items():
                    ws_out = wb_out.create_sheet(title=group_name)
                    current_max_row = 1
                    all_combined_dfs = []

                    for sheet_idx, s_name in enumerate(s_list):
                        ws_src_sheet = wb_src[s_name]
                        if s_name in temp_df_dict:
                            df_temp = temp_df_dict[s_name]
                            if not df_temp.empty:
                                all_combined_dfs.append(df_temp)

                        if sheet_idx > 0:
                            current_max_row += 3

                        target_start_row = current_max_row
                        for row in ws_src_sheet.iter_rows(values_only=False):
                            if all(cell.value is None for cell in row):
                                current_max_row += 1
                                continue
                            for cell in row:
                                target_cell = ws_out.cell(row=current_max_row, column=cell.column, value=cell.value)
                                if cell.has_style:
                                    if cell.font:
                                        target_cell.font = Font(name=cell.font.name, size=cell.font.size, bold=cell.font.bold, italic=cell.font.italic, color=cell.font.color)
                                    if cell.fill and cell.fill.fill_type:
                                        target_cell.fill = cell.fill.__copy__()
                                    if cell.alignment:
                                        target_cell.alignment = Alignment(horizontal=cell.alignment.horizontal, vertical=cell.alignment.vertical, wrap_text=cell.alignment.wrap_text)
                                    if cell.border:
                                        target_cell.border = cell.border.__copy__()
                                    if cell.number_format:
                                        target_cell.number_format = cell.number_format
                            current_max_row += 1

                        row_shift = target_start_row - 1
                        for merged_range in ws_src_sheet.merged_cells.ranges:
                            min_col, max_col, min_row, max_row = merged_range.min_col, merged_range.max_col, merged_range.min_row, merged_range.max_row
                            new_min_row, new_max_row = min_row + row_shift, max_row + row_shift
                            if new_min_row <= new_max_row:
                                try:
                                    ws_out.merge_cells(start_row=new_min_row, start_column=min_col, end_row=new_max_row, end_column=max_col)
                                except:
                                    pass

                    if not all_combined_dfs:
                        continue

                    df_group = pd.concat(all_combined_dfs, ignore_index=True)
                    tgl_col = 'TANGGAL' if 'TANGGAL' in df_group.columns else ('Tanggal' if 'Tanggal' in df_group.columns else df_group.columns[0])
                    mesin_col = 'MESIN' if 'MESIN' in df_group.columns else ('Mesin' if 'Mesin' in df_group.columns else df_group.columns[1])
                    nama_col = 'NAMA' if 'NAMA' in df_group.columns else ('Nama' if 'Nama' in df_group.columns else (df_group.columns[2] if len(df_group.columns) > 2 else None))

                    df_group[tgl_col] = df_group[tgl_col].ffill()
                    df_group[mesin_col] = df_group[mesin_col].ffill()
                    if nama_col and nama_col in df_group.columns:
                        df_group[nama_col] = df_group[nama_col].ffill()

                    df_group['Parsed_Date'] = pd.to_datetime(df_group[tgl_col], errors='coerce')
                    df_group['Nama_Hari'] = df_group['Parsed_Date'].dt.weekday.map(days_map)
                    df_group['Minggu_Ke'] = df_group['Parsed_Date'].apply(get_calendar_week_mapping)
                    df_group['Mesin_Clean'] = df_group[mesin_col].astype(str).str.strip().str.upper()

                    df_group = df_group[
                        df_group['Mesin_Clean'].notna() &
                        (df_group['Mesin_Clean'] != '') &
                        (df_group['Mesin_Clean'] != 'NAN') &
                        (df_group['Mesin_Clean'] != 'NONE') &
                        (df_group['Mesin_Clean'] != 'NAT')
                    ]

                    max_orig_col = ws_out.max_column
                    start_right_col = max_orig_col + 3
                    mesin_in_group = sorted(df_group['Mesin_Clean'].unique())

                    col_persen_name, col_persen_idx_1based = None, None
                    for idx, col_name in enumerate(df_group.columns):
                        clean_col = str(col_name).upper().replace('\n', ' ')
                        if 'TARGET 100%' in clean_col and 'CONES' in clean_col:
                            if idx + 1 < len(df_group.columns):
                                col_persen_name = df_group.columns[idx + 1]
                                col_persen_idx_1based = idx + 2
                            break

                    if col_persen_name is None:
                        col_persen_idx_1based = 12

                    col_kg_name = None
                    for idx, col_name in enumerate(df_group.columns):
                        clean_col = str(col_name).upper().replace('\n', ' ')
                        if 'CONES' in clean_col and not 'TOTAL' in clean_col and not 'TARGET' in clean_col:
                            if idx + 1 < len(df_group.columns):
                                col_kg_name = df_group.columns[idx + 1]
                                break
                    if col_kg_name is None:
                        kg_candidates = [c for c in df_group.columns if str(c).strip().upper() == 'KG']
                        col_kg_name = kg_candidates[0] if kg_candidates else df_group.columns[-1]

                    current_check_row = 3
                    for sheet_idx, s_name in enumerate(s_list):
                        if sheet_idx > 0:
                            current_check_row += 3
                        ws_src_sheet = wb_src[s_name]
                        merged_ranges_list = ws_src_sheet.merged_cells.ranges

                        r_idx = 3
                        while r_idx <= ws_src_sheet.max_row:
                            actual_row_in_out = current_check_row + (r_idx - 3)
                            group_end_row_in_out = actual_row_in_out
                            for mr in merged_ranges_list:
                                if mr.min_row <= r_idx <= mr.max_row and mr.min_row != mr.max_row:
                                    row_shift_val = current_check_row - 3
                                    group_end_row_in_out = mr.max_row + row_shift_val
                                    break

                            cell_persen = ws_out.cell(row=actual_row_in_out, column=col_persen_idx_1based)
                            p_float = parse_percentage_value(cell_persen.value)
                            if p_float is not None:
                                if p_float < threshold_val:
                                    for target_r in range(actual_row_in_out, group_end_row_in_out + 1):
                                        for col_idx in range(1, max_orig_col + 1):
                                            row_cell = ws_out.cell(row=target_r, column=col_idx)
                                            row_cell.fill = highlight_fill
                                            row_cell.font = highlight_font

                            r_idx += (group_end_row_in_out - actual_row_in_out + 1)
                        current_check_row += ws_src_sheet.max_row

                    for m_idx, mesin in enumerate(mesin_in_group):
                        df_mesin = df_group[df_group['Mesin_Clean'] == mesin]
                        minggu_list = sorted(df_mesin['Minggu_Ke'].unique())
                        current_col = start_right_col + (m_idx * (len(minggu_list) * 3 + 2))

                        ws_out.cell(row=1, column=current_col, value="HARI").alignment = Alignment(horizontal='center', vertical='center')
                        ws_out.merge_cells(start_row=1, start_column=current_col, end_row=2, end_column=current_col)
                        ws_out.cell(row=1, column=current_col).font = header_font

                        avg_row_idx = 3 + len(days_order)
                        for r in range(1, avg_row_idx + 1):
                            ws_out.cell(row=r, column=current_col).border = thin_border

                        for d_i, day_name in enumerate(days_order):
                            c_cell = ws_out.cell(row=3 + d_i, column=current_col, value=day_name)
                            c_cell.font = bold_font
                            c_cell.alignment = Alignment(horizontal='center', vertical='center')

                        ws_out.cell(row=avg_row_idx, column=current_col, value="RATA-RATA").font = bold_font

                        col_offset = current_col + 1
                        for w in minggu_list:
                            df_minggu = df_mesin[df_mesin['Minggu_Ke'] == w]
                            header_title = f"MINGGU KE {w} BULAN {bulan_pilihan} - {mesin}"
                            ws_out.merge_cells(start_row=1, start_column=col_offset, end_row=1, end_column=col_offset + 2)
                            top_head = ws_out.cell(row=1, column=col_offset, value=header_title)
                            top_head.font = Font(name='Calibri', size=10, bold=True)
                            top_head.alignment = Alignment(horizontal='center', vertical='center')

                            sub_headers = ["TOTAL CONES", "PERSENTASE HASIL", "HASIL (KG)"]
                            for s_i, sh in enumerate(sub_headers):
                                ws_out.cell(row=2, column=col_offset + s_i, value=sh).font = sub_header_font

                            daily_cones_for_avg, daily_persen_for_avg, daily_kg_for_avg = [], [], []
                            row_data_store = []
                            col_cones_name = next((c for c in df_group.columns if 'CONES' in c.upper() and 'TOTAL' in c.upper()), df_group.columns[7])

                            for day_name in days_order:
                                df_day = df_minggu[df_minggu['Nama_Hari'] == day_name]
                                valid_cones_list, valid_persen_list, valid_kg_list = [], [], []

                                for _, row_item in df_day.iterrows():
                                    p_f = parse_percentage_value(row_item.get(col_persen_name))
                                    if p_f is not None and p_f >= threshold_val:
                                        c_val = row_item.get(col_cones_name, 0.0)
                                        if pd.notna(c_val):
                                            try: valid_cones_list.append(float(c_val))
                                            except: pass
                                        valid_persen_list.append(p_f)
                                        k_val = row_item.get(col_kg_name, 0.0)
                                        if pd.notna(k_val):
                                            try: valid_kg_list.append(float(str(k_val).replace(',', '.')))
                                            except: pass

                                val_cones = sum(valid_cones_list) if valid_cones_list else 0.0
                                val_persen = np.mean(valid_persen_list) if valid_persen_list else 0.0
                                val_kg = sum(valid_kg_list) if valid_kg_list else 0.0
                                row_data_store.append((val_cones, val_persen, val_kg))

                                if val_cones > 0: daily_cones_for_avg.append(val_cones)
                                if val_persen > 0: daily_persen_for_avg.append(val_persen)
                                if val_kg > 0: daily_kg_for_avg.append(val_kg)

                            for d_i, (val_cones, val_persen, val_kg) in enumerate(row_data_store):
                                row_target_idx = 3 + d_i
                                ws_out.cell(row=row_target_idx, column=col_offset, value=round(val_cones, 1) if val_cones > 0 else 0)
                                c2 = ws_out.cell(row=row_target_idx, column=col_offset + 1, value="-" if val_persen == 0 else round(val_persen / 100.0, 4))
                                if val_persen > 0: c2.number_format = '0.00%'
                                ws_out.cell(row=row_target_idx, column=col_offset + 2, value="-" if val_kg == 0 else round(val_kg, 2))

                            avg_cones_val = np.mean(daily_cones_for_avg) if daily_cones_for_avg else 0.0
                            avg_persen_val = np.mean(daily_persen_for_avg) if daily_persen_for_avg else 0.0
                            avg_kg_val = np.mean(daily_kg_for_avg) if daily_kg_for_avg else 0.0

                            ws_out.cell(row=avg_row_idx, column=col_offset, value=round(avg_cones_val, 1)).font = bold_font
                            ac2 = ws_out.cell(row=avg_row_idx, column=col_offset + 1, value="-" if avg_persen_val == 0 else round(avg_persen_val / 100.0, 4))
                            if avg_persen_val > 0: ac2.number_format = '0.00%'
                            ac2.font = bold_font
                            ws_out.cell(row=avg_row_idx, column=col_offset + 2, value="-" if avg_kg_val == 0 else round(avg_kg_val, 2)).font = bold_font

                            for r in range(1, avg_row_idx + 1):
                                for c_idx in range(col_offset, col_offset + 3):
                                    ws_out.cell(row=r, column=c_idx).border = thin_border
                            col_offset += 3

                    summary_start_col = start_right_col
                    summary_start_row = avg_row_idx + 4
                    ws_out.cell(row=summary_start_row, column=summary_start_col, value=f"REKAP PERSENTASE HASIL PER MINGGU & MESIN (BULAN {bulan_pilihan})").font = bold_font
                    t_header_row2 = summary_start_row + 2

                    ws_out.cell(row=t_header_row2, column=summary_start_col, value="HARI").font = header_font
                    ws_out.cell(row=t_header_row2, column=summary_start_col).border = thin_border

                    all_weeks = sorted(df_group['Minggu_Ke'].unique())
                    combo_list = []
                    for w in all_weeks:
                        for mesin in mesin_in_group:
                            if not df_group[(df_group['Mesin_Clean'] == mesin) & (df_group['Minggu_Ke'] == w)].empty:
                                combo_list.append((mesin, w))

                    col_ptr = summary_start_col + 1
                    col_data_dict = {}
                    bulan_singkatan = bulan_pilihan[:4]

                    for mesin, w in combo_list:
                        cell_head = ws_out.cell(row=t_header_row2, column=col_ptr, value=f"MINGGU {w} {bulan_singkatan}\n{mesin}")
                        cell_head.font = sub_header_font
                        cell_head.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell_head.border = thin_border

                        df_mesin_item = df_group[df_group['Mesin_Clean'] == mesin]
                        df_minggu_item = df_mesin_item[df_mesin_item['Minggu_Ke'] == w]

                        day_percentages_for_avg = []
                        row_values_to_write = []

                        for day_name in days_order:
                            df_day = df_minggu_item[df_minggu_item['Nama_Hari'] == day_name]
                            valid_p_vals = []
                            for _, row_item in df_day.iterrows():
                                p_f = parse_percentage_value(row_item.get(col_persen_name))
                                if p_f is not None and p_f >= threshold_val:
                                    valid_p_vals.append(p_f)

                            val_persen = np.mean(valid_p_vals) if valid_p_vals else 0.0
                            row_values_to_write.append(val_persen)
                            if val_persen > 0:
                                day_percentages_for_avg.append(val_persen)

                        for d_i, val_persen in enumerate(row_values_to_write):
                            row_idx = t_header_row2 + 1 + d_i
                            c_cell = ws_out.cell(row=row_idx, column=col_ptr, value="-" if val_persen == 0 else round(val_persen / 100.0, 4))
                            if val_persen > 0:
                                c_cell.number_format = '0.00%'
                            c_cell.alignment = Alignment(horizontal='center', vertical='center')
                            c_cell.border = thin_border

                        col_data_dict[col_ptr] = day_percentages_for_avg
                        col_ptr += 1

                    for d_i, day_name in enumerate(days_order):
                        row_idx = t_header_row2 + 1 + d_i
                        d_cell = ws_out.cell(row=row_idx, column=summary_start_col, value=day_name)
                        d_cell.font = bold_font
                        d_cell.border = thin_border

                    avg_row_summary = t_header_row2 + 1 + len(days_order)
                    ws_out.cell(row=avg_row_summary, column=summary_start_col, value="RATA-RATA").font = bold_font
                    ws_out.cell(row=avg_row_summary, column=summary_start_col).border = thin_border

                    col_ptr = summary_start_col + 1
                    for mesin, w in combo_list:
                        p_list = col_data_dict[col_ptr]
                        col_avg = np.mean(p_list) if p_list else 0.0
                        ac_cell = ws_out.cell(row=avg_row_summary, column=col_ptr, value="-" if col_avg == 0 else round(col_avg / 100.0, 4))
                        if col_avg > 0:
                            ac_cell.number_format = '0.00%'
                        ac_cell.font = bold_font
                        ac_cell.border = thin_border
                        col_ptr += 1

                for sheet in wb_out.worksheets:
                    for col in sheet.columns:
                        max_len = 0
                        col_letter = get_column_letter(col[0].column)
                        for cell in col:
                            if type(cell).__name__ == 'MergedCell':
                                continue
                            if cell.value is not None:
                                for line in str(cell.value).split('\n'):
                                    if len(line) > max_len:
                                        max_len = len(line)
                        sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

                output_summary_io = BytesIO()
                wb_out.save(output_summary_io)
                output_summary_bytes = output_summary_io.getvalue()

                # Simpan hasil ke session_state
                st.session_state['processed'] = True
                st.session_state['clean_bytes'] = output_clean_bytes
                st.session_state['emp_bytes'] = output_emp_bytes
                st.session_state['mach_bytes'] = output_mach_bytes
                st.session_state['summary_bytes'] = output_summary_bytes
                st.session_state['df_clean_preview'] = df_clean
                
                df_rekap_karyawan = df_clean[['Tanggal', 'Pegawai', 'Mesin', 'Shift_Final', 'Benang', 'Qty Cones', 'Qty Real Cones', 'Total']].copy()
                df_rekap_karyawan.columns = ['Tanggal', 'Nama Karyawan', 'Mesin', 'Shift', 'Jenis Benang', 'Jumlah Cones', 'Qty Real Cones', 'Total KG']
                st.session_state['df_emp_preview'] = df_rekap_karyawan

                df_rekap_mesin = df_clean[['Tanggal', 'Mesin', 'Pegawai', 'Shift_Final', 'Benang', 'Qty Cones', 'Qty Real Cones', 'Total']].copy()
                df_rekap_mesin.columns = ['Tanggal', 'Mesin', 'Nama Karyawan', 'Shift', 'Jenis Benang', 'Jumlah Cones', 'Qty Real Cones', 'Total KG']
                st.session_state['df_mach_preview'] = df_rekap_mesin

                st.success("⚡ Pemrosesan Data Sukses! Sistem Siap Digunakan.")

            except Exception as e:
                st.error(f"❌ Error Sistem: {e}")

# ==========================================
# TAMPILAN DASHBOARD UTAMA & KARTU METRIK
# ==========================================
if st.session_state.get('processed', False):
    st.markdown("---")
    
    # Kartu Metrik Ringkasan Futuristik (Glassmorphism)
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.markdown(f"""
            <div class="metric-card">
                <p style='color: #8b949e; margin-bottom: 0;'>STATUS BULAN</p>
                <h3 style='color: #58a6ff; margin-top: 5px;'>{bulan_pilihan}</h3>
            </div>
        """, unsafe_allow_html=True)
    with col_m2:
        st.markdown(f"""
            <div class="metric-card">
                <p style='color: #8b949e; margin-bottom: 0;'>THRESHOLD MINIMUM</p>
                <h3 style='color: #3fb950; margin-top: 5px;'>{threshold_val}%</h3>
            </div>
        """, unsafe_allow_html=True)
    with col_m3:
        st.markdown(f"""
            <div class="metric-card">
                <p style='color: #8b949e; margin-bottom: 0;'>TOTAL TRANSAKSI BERSIH</p>
                <h3 style='color: #bc8cff; margin-top: 5px;'>{len(st.session_state['df_clean_preview'])} Baris</h3>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("### 📊 Menu Navigasi Data & Unduhan")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "📁 Data Bersih", 
        "👥 Rekap per Karyawan", 
        "⚙️ Rekap per Mesin", 
        "📅 Summary per Minggu"
    ])

    with tab1:
        st.markdown("#### Tabel Data Bersih Sistem")
        st.download_button(
            label="📥 Unduh Berkas Data Bersih (.xlsx)",
            data=st.session_state['clean_bytes'],
            file_name="DATA BERSIH.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.dataframe(st.session_state['df_clean_preview'], use_container_width=True)

    with tab2:
        st.markdown("#### Tabel Rekap Hasil per Karyawan")
        st.download_button(
            label="📥 Unduh Berkas Rekap Karyawan (.xlsx)",
            data=st.session_state['emp_bytes'],
            file_name="REKAP HASIL PER KARYAWAN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("💡 Saring data berdasarkan nama karyawan untuk analisis performa individu:")
        
        df_emp_all = st.session_state['df_emp_preview']
        daftar_karyawan = sorted(df_emp_all['Nama Karyawan'].dropna().unique().tolist())
        pilihan_karyawan = st.selectbox("Pilih Nama Karyawan:", ["-- SEMUA KARYAWAN --"] + daftar_karyawan)

        if pilihan_karyawan != "-- SEMUA KARYAWAN --":
            df_emp_filtered = df_emp_all[df_emp_all['Nama Karyawan'] == pilihan_karyawan]
        else:
            df_emp_filtered = df_emp_all

        st.dataframe(df_emp_filtered, use_container_width=True)

    with tab3:
        st.markdown("#### Tabel Rekap Hasil per Mesin")
        st.download_button(
            label="📥 Unduh Berkas Rekap Mesin (.xlsx)",
            data=st.session_state['mach_bytes'],
            file_name="REKAP HASIL PER MESIN.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.info("💡 Saring data berdasarkan mesin produksi untuk memantau performa unit mesin:")
        
        df_mach_all = st.session_state['df_mach_preview']
        daftar_mesin = sorted(df_mach_all['Mesin'].dropna().unique().tolist())
        pilihan_mesin = st.selectbox("Pilih Mesin Produksi:", ["-- SEMUA MESIN --"] + daftar_mesin)

        if pilihan_mesin != "-- SEMUA MESIN --":
            df_mach_filtered = df_mach_all[df_mach_all['Mesin'] == pilihan_mesin]
        else:
            df_mach_filtered = df_mach_all

        st.dataframe(df_mach_filtered, use_container_width=True)

    with tab4:
        st.markdown("#### Summary Rekap per Minggu")
        st.download_button(
            label="📥 Unduh Berkas Summary Mingguan (.xlsx)",
            data=st.session_state['summary_bytes'],
            file_name="SUMMARY REKAP PER MINGGU.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.warning("⚠️ Struktur tabel Summary Mingguan dirancang secara kompleks per grup mesin. Silakan unduh langsung file Excel untuk melihat format laporan utuh.")
else:
    st.markdown("""
        <div style="text-align: center; padding: 40px; background: rgba(22, 27, 34, 0.4); border-radius: 12px; border: 1px dashed #30363d;">
            <h3>🌐 Sistem Menunggu Masukan Berkas</h3>
            <p style="color: #8b949e;">Silakan unggah berkas laporan dan tentukan parameter di panel kontrol sebelah kiri, lalu klik tombol <b>EKSEKUSI PROSES DATA</b>.</p>
        </div>
    """, unsafe_allow_html=True)
