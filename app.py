import streamlit as st
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill

# ==========================================
# UI/UX: INJEKSI TEMA DARK CYBER / FUTURISTIC
# ==========================================
st.set_page_config(page_title="Sistem Rekap Detail", layout="wide", initial_sidebar_state="collapsed")

custom_css = """
<style>
    /* Latar Belakang Utama Aplikasi (Deep Dark Navy/Black) */
    .stApp {
        background-color: #040814 !important;
        background-image: radial-gradient(circle at 50% 0%, #0a1931 0%, #040814 70%);
        color: #e0f7fa !important;
        font-family: 'Courier New', Courier, monospace;
    }

    /* Styling Judul dan Header */
    h1, h2, h3, h4, h5, h6 {
        color: #00f3ff !important;
        text-shadow: 0 0 10px rgba(0, 243, 255, 0.5);
        font-family: 'Trebuchet MS', sans-serif;
        text-transform: uppercase;
        letter-spacing: 2px;
    }

    /* Styling Kolom Upload File (Glowing Dashed Border) */
    div[data-testid="stFileUploadDropzone"] {
        background-color: rgba(4, 12, 32, 0.7) !important;
        border: 2px dashed #00f3ff !important;
        box-shadow: 0 0 15px rgba(0, 243, 255, 0.2);
        border-radius: 10px;
        transition: all 0.3s ease;
    }
    div[data-testid="stFileUploadDropzone"]:hover {
        border-color: #00ffff !important;
        box-shadow: 0 0 25px rgba(0, 255, 255, 0.6);
        background-color: rgba(0, 243, 255, 0.05) !important;
    }

    /* Warna Teks pada Uploader */
    .st-emotion-cache-1gblk4, .st-emotion-cache-1wivap2, .st-emotion-cache-1jspxgj {
        color: #00f3ff !important;
    }

    /* Tombol Utama (Hollow Neon Style) */
    div[data-testid="stButton"] > button {
        background-color: transparent !important;
        color: #00f3ff !important;
        border: 1px solid #00f3ff !important;
        box-shadow: 0 0 8px rgba(0, 243, 255, 0.4), inset 0 0 5px rgba(0, 243, 255, 0.2);
        border-radius: 4px;
        font-weight: bold;
        letter-spacing: 1px;
        text-transform: uppercase;
        transition: all 0.3s ease;
        padding: 0.5rem 1rem;
    }
    div[data-testid="stButton"] > button:hover {
        background-color: #00f3ff !important;
        color: #040814 !important;
        box-shadow: 0 0 20px rgba(0, 243, 255, 0.8);
    }

    /* Notifikasi Info / Success (Panel Transparan) */
    div[data-testid="stAlert"] {
        background-color: rgba(0, 243, 255, 0.05) !important;
        border: 1px solid #00f3ff !important;
        border-left: 5px solid #00f3ff !important;
        color: #e0f7fa !important;
        box-shadow: 0 0 10px rgba(0, 243, 255, 0.1);
    }
    
    /* Ikon dan Markdown Text */
    .stMarkdown p {
        color: #b2ebf2;
    }
    
    /* Label Input */
    label {
        color: #00f3ff !important;
        font-weight: 600 !important;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# JUDUL BARU SEBAGAI PENANDA CACHE UPDATE
st.markdown("<h1 style='text-align: center; border-bottom: 1px solid #00f3ff; padding-bottom: 15px; margin-bottom: 30px;'>⚙️ SISTEM REKAP DETAIL LAPORAN BULANAN (UPDATE V2)</h1>", unsafe_allow_html=True)

# ==========================================
# FUNGSI BACKEND
# ==========================================
def parse_number(val):
    if pd.isna(val): return 0.0
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip()
    if not s or s == '-': return 0.0
    s = re.sub(r'[^\d.,-]', '', s)
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    else:
        if s.count('.') > 1 or re.search(r'\.\d{3}$', s):
            s = s.replace('.', '')
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_single_file(file, file_idx):
    df = pd.read_excel(file, skiprows=4)
    df = df.iloc[:-4]
    df = df.dropna(how='all')
    
    ind_col = 'Kode Barang Jadi' if 'Kode Barang Jadi' in df.columns else ('Gudang' if 'Gudang' in df.columns else None)
    if ind_col:
        is_parent = df[ind_col].notna() & (df[ind_col].astype(str).str.strip() != '') & (df[ind_col].astype(str).str.lower() != 'nan')
        df['Tx_ID'] = str(file_idx) + "_" + is_parent.cumsum().astype(str)
    else:
        df['Tx_ID'] = f"{file_idx}_1"
    
    if 'Kode Item' in df.columns:
        idx = df.columns.get_loc('Kode Item')
        if isinstance(idx, int):
            kolom_ffill_dinamis = df.columns[:idx]
            df[kolom_ffill_dinamis] = df[kolom_ffill_dinamis].ffill()
    else:
        kolom_ffill = ['Gudang', 'Kode Barang Jadi', 'Nama Barang Jadi', 'Target Qty', 'Target Unit']
        cols_to_fill = [c for c in kolom_ffill if c in df.columns]
        if cols_to_fill:
            df[cols_to_fill] = df[cols_to_fill].ffill()
    
    kolom_angka = ['Target Qty', 'Qty', 'Total', 'Total Harga', 'Harga', 'Berat Netto', 'Berat Netto (Gram)'] 
    for col in kolom_angka:
        if col in df.columns:
            df[col] = df[col].apply(parse_number)
            
    return df

def get_sheet_group_and_merk(r, df_ind):
    gudang = str(r.get('Gudang', '')).upper()
    nama_bj = str(r.get('Nama Barang Jadi', ''))
    target_unit_raw = str(r.get('Target Unit', '')).upper().strip()
    
    if ("500" in gudang or "5000" in gudang) and df_ind is not None and not df_ind.empty:
        df_filter = df_ind[df_ind.iloc[:, 0].astype(str).str.upper() == gudang]
        
        nama_bj_upper = nama_bj.upper()
        nama_bj_compare = re.sub(r'(?<=\d)\.(?=\d)', '', nama_bj_upper)
        
        cols = df_ind.columns
        col_ind = next((c for c in cols if 'NAMA BARANG JADI' in c.upper() or 'INDIKATOR' in c.upper()), cols[1] if len(cols) > 1 else None)
        col_target_unit = next((c for c in cols if 'TARGET UNIT' in c.upper()), cols[2] if len(cols) > 2 else None)
        col_sheet = next((c for c in cols if 'SHEET' in c.upper()), cols[3] if len(cols) > 3 else None)
        col_merk = next((c for c in cols if 'MEREK' in c.upper() or 'MERK' in c.upper()), cols[4] if len(cols) > 4 else None)
        
        for _, ind_row in df_filter.iterrows():
            if col_ind is not None:
                if col_target_unit is not None:
                    ind_tu = str(ind_row[col_target_unit]).upper().strip()
                    if target_unit_raw != ind_tu:
                        continue 
                
                keyword_raw = str(ind_row[col_ind]).upper()
                keyword_clean = re.sub(r'(?<=\d)\.(?=\d)', '', keyword_raw)
                keywords = keyword_clean.split()
                
                match_all = True
                for kw in keywords:
                    pattern = r'(?<![A-Z0-9])' + re.escape(kw) + r'(?![A-Z0-9])'
                    if not re.search(pattern, nama_bj_compare):
                        match_all = False
                        break
                        
                if match_all:
                    nama_sheet = str(ind_row[col_sheet]) if col_sheet else nama_bj
                    merk_val = str(ind_row[col_merk]) if col_merk else nama_sheet
                    return nama_sheet, merk_val
                
    return nama_bj, nama_bj

def generate_excel_report(df_gabungan, df_indikator):
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active)

    df_gabungan['Kode Item'] = df_gabungan['Kode Item'].astype(str)
    masks = []
    
    if 'Gudang' in df_gabungan.columns:
        for gudang, group in df_gabungan.groupby('Gudang', dropna=False):
            g = str(gudang).upper()
            if "DYEING" in g or "LAB" in g or "RND" in g:
                m = group['Kode Item'].str.contains('TWP|MWP|TBM', case=True, na=False)
            elif "500" in g or "5000" in g:
                m = group['Kode Item'].str.contains('TWP', case=True, na=False)
            else:
                m = group['Kode Item'].str.contains('TWP|MWP|MBB|TBB|TBM', case=True, na=False)
            masks.append(group[m])
            
    df_filtered = pd.concat(masks, ignore_index=True) if masks else pd.DataFrame(columns=df_gabungan.columns)

    res = df_filtered.apply(lambda r: get_sheet_group_and_merk(r, df_indikator), axis=1)
    df_filtered['Sheet_Group'] = [x[0] for x in res]
    df_filtered['Merk_Group'] = [x[1] for x in res]

    if 'Target Unit' in df_filtered.columns and 'Target Qty' in df_filtered.columns:
        loc = df_filtered.columns.get_loc('Target Unit')
        idx = loc + 1 if isinstance(loc, int) else len(df_filtered.columns)
        df_filtered.insert(idx, 'Qty Target Standar (GR)', df_filtered['Target Qty'])
    else:
        df_filtered['Qty Target Standar (GR)'] = df_filtered.get('Target Qty', 0)
        
    if 'Qty' in df_filtered.columns and 'Unit' in df_filtered.columns:
        loc = df_filtered.columns.get_loc('Qty')
        idx = loc + 1 if isinstance(loc, int) else len(df_filtered.columns)
        df_filtered.insert(idx, 'Qty BB Standar (GR)', df_filtered['Qty'])
    else:
        df_filtered['Qty BB Standar (GR)'] = df_filtered.get('Qty', 0)

    mask_dyeing_lab = df_filtered['Gudang'].astype(str).str.upper().str.contains('DYEING|LAB|RND', regex=True)
    mask_kg_target = df_filtered['Target Unit'].astype(str).str.upper() == 'KG'
    mask_kg_unit = df_filtered['Unit'].astype(str).str.upper() == 'KG'

    df_filtered.loc[mask_dyeing_lab & mask_kg_target, 'Qty Target Standar (GR)'] = df_filtered['Target Qty'] * 1000
    df_filtered.loc[mask_dyeing_lab & mask_kg_unit, 'Qty BB Standar (GR)'] = df_filtered['Qty'] * 1000

    ws_summary = wb.create_sheet(title="Summary", index=0)
    master_idx = 1
    
    summary_dict = {}
    master_names_dict = {}
    kolom_total = 'Total Harga' if 'Total Harga' in df_filtered.columns else 'Total'

    for gudang, gudang_df in df_filtered.groupby('Gudang', dropna=False):
        gudang_name = str(gudang)
        g_upper = gudang_name.upper()
        
        is_500_5000 = "500" in g_upper or "5000" in g_upper
        is_dyeing_lab = "DYEING" in g_upper or "LAB" in g_upper or "RND" in g_upper

        cols_to_drop = ['Sheet_Group', 'Merk_Group', 'Tx_ID']
        if not is_dyeing_lab:
            cols_to_drop.extend(['Qty Target Standar (GR)', 'Qty BB Standar (GR)'])

        master_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', f"Master - {gudang_name}")[:31]
        ws_master = wb.create_sheet(title=master_sheet_name, index=master_idx)
        master_names_dict[gudang_name] = master_sheet_name 
        master_idx += 1
        
        btn = ws_master.cell(row=1, column=1, value="⬅️ KEMBALI KE SUMMARY")
        btn.hyperlink = "#'Summary'!A1"
        btn.font = Font(color="0563C1", underline="single", bold=True)
        
        gudang_export = gudang_df.drop(columns=cols_to_drop, errors='ignore')
        for c_idx, col_name in enumerate(gudang_export.columns, 1):
            cell = ws_master.cell(row=3, column=c_idx, value=col_name)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        for r_idx, row in enumerate(dataframe_to_rows(gudang_export, index=False, header=False), 4):
            for c_idx, value in enumerate(row, 1):
                ws_master.cell(row=r_idx, column=c_idx, value=value)
                
        summary_dict[gudang_name] = []
        grouped = gudang_df.groupby('Sheet_Group', dropna=False)
            
        for keys, group_data in grouped:
            nama_bj = str(keys) 
            
            if is_500_5000:
                sheet_name_raw = nama_bj 
                merk_display = group_data['Merk_Group'].iloc[0] if not group_data.empty else nama_bj
            else:
                # PERBAIKAN LOGIKA PEMOTONGAN SESUAI TANDA "." (TITIK)
                if "." in nama_bj:
                    penggalan = nama_bj.split(".", 1)[1].strip()
                else:
                    penggalan = nama_bj.strip()
                    
                sheet_name_raw = penggalan
                merk_display = penggalan
                
            # MENGUBAH KARAKTER ILEGAL MENJADI STRIP "-"
            sheet_name_clean = re.sub(r'[\\/*?:\[\]]', '-', sheet_name_raw)[:31]
            base_sheet_name = sheet_name_clean
            counter = 1
            while base_sheet_name in wb.sheetnames:
                base_sheet_name = f"{sheet_name_clean[:28]}_{counter}"
                counter += 1
            final_sheet_name = base_sheet_name
            
            ws = wb.create_sheet(title=final_sheet_name)
            
            btn = ws.cell(row=1, column=1, value="⬅️ KEMBALI KE SUMMARY")
            btn.hyperlink = "#'Summary'!A1"
            btn.font = Font(color="0563C1", underline="single", bold=True)
            
            group_export = group_data.drop(columns=cols_to_drop, errors='ignore')
            for r_idx, r in enumerate(dataframe_to_rows(group_export, index=False, header=True), 3):
                for c_idx, val in enumerate(r, 1):
                    ws.cell(row=r_idx, column=c_idx, value=val)
            for cell in ws[3]:
                cell.font = Font(bold=True)
                
            max_col_master = ws.max_column
            start_col_rekap = max_col_master + 3
            
            rekap_headers = [
                "Kode Barang Jadi", "Nama Barang Jadi", "Target Qty", "Target Unit",
                "Kode Item", "Nama Item", "Unit", "Qty", "Total Harga"
            ]
            for i, header in enumerate(rekap_headers):
                cell = ws.cell(row=3, column=start_col_rekap + i, value=header)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

            current_row = 4
            sub_grouped = group_data.groupby('Kode Barang Jadi', dropna=False)
            
            for kode_bj, sub_df in sub_grouped:
                nama_bj_val = sub_df['Nama Barang Jadi'].iloc[0]
                target_unit_val = sub_df['Target Unit'].iloc[0] if 'Target Unit' in sub_df.columns else ""
                
                display_target_unit = target_unit_val if is_500_5000 else "GR"
                
                if is_dyeing_lab:
                    target_qty_val = sub_df.drop_duplicates(subset=['Tx_ID'])['Qty Target Standar (GR)'].sum()
                    item_grouped = sub_df.groupby(['Kode Item', 'Nama Item'], dropna=False).agg({
                        'Unit': 'first',
                        'Qty BB Standar (GR)': 'sum',
                        kolom_total: 'sum'
                    }).reset_index()
                    item_grouped.rename(columns={'Qty BB Standar (GR)': 'Qty'}, inplace=True)
                else:
                    target_qty_val = sub_df.drop_duplicates(subset=['Tx_ID'])['Target Qty'].sum()
                    item_grouped = sub_df.groupby(['Kode Item', 'Nama Item'], dropna=False).agg({
                        'Unit': 'first',
                        'Qty': 'sum',
                        kolom_total: 'sum'
                    }).reset_index()

                jumlah_item = len(item_grouped)
                start_merge_row = current_row
                
                for _, item_row in item_grouped.iterrows():
                    display_item_unit = item_row['Unit'] if is_500_5000 else "GR"
                    
                    ws.cell(row=current_row, column=start_col_rekap, value=kode_bj)
                    ws.cell(row=current_row, column=start_col_rekap + 1, value=nama_bj_val)
                    ws.cell(row=current_row, column=start_col_rekap + 2, value=target_qty_val)
                    ws.cell(row=current_row, column=start_col_rekap + 3, value=display_target_unit)
                    ws.cell(row=current_row, column=start_col_rekap + 4, value=item_row['Kode Item'])
                    ws.cell(row=current_row, column=start_col_rekap + 5, value=item_row['Nama Item'])
                    ws.cell(row=current_row, column=start_col_rekap + 6, value=display_item_unit)
                    ws.cell(row=current_row, column=start_col_rekap + 7, value=item_row['Qty'])
                    ws.cell(row=current_row, column=start_col_rekap + 8, value=item_row[kolom_total]) 
                    current_row += 1
                    
                if jumlah_item > 1:
                    end_merge_row = current_row - 1
                    for col_offset in range(4): 
                        ws.merge_cells(start_row=start_merge_row, start_column=start_col_rekap + col_offset,
                                       end_row=end_merge_row, end_column=start_col_rekap + col_offset)
                        ws.cell(row=start_merge_row, column=start_col_rekap + col_offset).alignment = Alignment(vertical='center', horizontal='center')

            target_u = group_data['Target Unit'].iloc[0] if not group_data.empty and 'Target Unit' in group_data.columns else ""
            
            if is_dyeing_lab:
                t_qty_sum = group_data.drop_duplicates(subset=['Tx_ID'])['Qty Target Standar (GR)'].sum()
                q_gram_sum = t_qty_sum
            else:
                t_qty_sum = group_data.drop_duplicates(subset=['Tx_ID'])['Target Qty'].sum()
                q_gram_sum = t_qty_sum
            
            if is_500_5000:
                q_gram_sum = group_data['Qty'].sum()
                satuan_target = target_u
                satuan = group_data['Unit'].iloc[0] if not group_data.empty and 'Unit' in group_data.columns else ""
            else:
                satuan_target = "GR"
                satuan = "GR"

            summary_dict[gudang_name].append({
                'DIVISI': gudang_name,
                'MERK': merk_display,
                'SATUAN TARGET': satuan_target,
                'QTY TARGET': t_qty_sum,
                'SATUAN': satuan,
                'QTY GRAM': q_gram_sum,
                'Sheet_Link': final_sheet_name 
            })
            
    # 4. PENULISAN SHEET SUMMARY MENYAMPING
    summary_headers = ['DIVISI', 'MERK', 'SATUAN TARGET', 'QTY TARGET', 'SATUAN', 'QTY GRAM']
    current_start_col = 1
    
    for gudang, rows in summary_dict.items():
        cell_title = ws_summary.cell(row=1, column=current_start_col, value=f"REKAP - {gudang}")
        master_link = master_names_dict.get(gudang, "")
        if master_link:
            cell_title.hyperlink = f"#'{master_link}'!A1"
        cell_title.font = Font(bold=True, color="0563C1", underline="single")
        
        for i, h in enumerate(summary_headers):
            c = ws_summary.cell(row=2, column=current_start_col + i, value=h)
            c.font = Font(bold=True)
            c.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            
        for r_idx, row_dict in enumerate(rows, 3):
            ws_summary.cell(row=r_idx, column=current_start_col, value=row_dict['DIVISI'])
            
            c_merk = ws_summary.cell(row=r_idx, column=current_start_col + 1, value=row_dict['MERK'])
            c_merk.hyperlink = f"#'{row_dict['Sheet_Link']}'!A1"
            c_merk.font = Font(color="0563C1", underline="single")
            
            ws_summary.cell(row=r_idx, column=current_start_col + 2, value=row_dict['SATUAN TARGET'])
            ws_summary.cell(row=r_idx, column=current_start_col + 3, value=row_dict['QTY TARGET'])
            ws_summary.cell(row=r_idx, column=current_start_col + 4, value=row_dict['SATUAN'])
            ws_summary.cell(row=r_idx, column=current_start_col + 5, value=row_dict['QTY GRAM'])
            
        current_start_col += len(summary_headers) + 2 
        
    wb.save(output)
    return output.getvalue()

# ==========================================
# UI STREAMLIT: KOLOM UPLOAD
# ==========================================
st.markdown("<p style='text-align: center; color: #b2ebf2; margin-bottom: 30px;'>⚙️ SISTEM PENGELOLAAN DETAIL PRODUK DATA HASIL PRODUKSI</p>", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    uploaded_files = st.file_uploader("📂 MASTER DATA HASIL PRODUKSI (.xlsx)", type=['xlsx'], accept_multiple_files=True)
with col2:
    indicator_file = st.file_uploader("🎯 INDIKATOR PENGELOMPOKKAN 500 & 5000 (Opsional)", type=['xlsx'])

if uploaded_files:
    st.info(f"🔄 Memproses {len(uploaded_files)} file data... Mohon tunggu sebentar.")
    
    try:
        df_indikator = None
        if indicator_file is not None:
            df_indikator = pd.read_excel(indicator_file)
            
        list_df_bersih = []
        all_cols_ordered = []
        
        for idx_file, file in enumerate(uploaded_files):
            df_bersih = clean_single_file(file, idx_file)
            list_df_bersih.append(df_bersih)
            for c in df_bersih.columns:
                if c not in all_cols_ordered:
                    all_cols_ordered.append(c)
            
        df_gabungan = pd.concat(list_df_bersih, ignore_index=True)
        df_gabungan = df_gabungan.reindex(columns=all_cols_ordered) 
        
        processed_excel = generate_excel_report(df_gabungan, df_indikator)
        
        st.success("✅ Proses selesai! Data berhasil dinormalisasi, digabungkan, dan siap diunduh.")
        
        st.download_button(
            label="⬇️ UNDUH LAPORAN HASIL REKAP",
            data=processed_excel,
            file_name="Laporan_Rekap_Detail_Bulanan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")
