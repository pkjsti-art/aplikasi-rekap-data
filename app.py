import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl

def process_and_generate_outputs(input_filename='REKAP HASIL PER MESIN.xlsx'):
    print(f"1. Membaca file sumber: {input_filename}...")
    try:
        xls = pd.ExcelFile(input_filename)
        sheet_names = xls.sheet_names
    except Exception as e:
        print(f"Error membaca file Excel: {e}")
        return

    all_combined_dfs = []
    for s_name in sheet_names:
        try:
            df_temp = pd.read_excel(input_filename, sheet_name=s_name)
            if not df_temp.empty:
                df_temp['Source_Sheet'] = s_name
                all_combined_dfs.append(df_temp)
        except Exception as e:
            print(f"Catatan: Sheet {s_name} dilewati: {e}")

    if not all_combined_dfs:
        print("Tidak ada data yang valid.")
        return

    df_group = pd.concat(all_combined_dfs, ignore_index=True)

    # 2. Identifikasi & Perbaikan Parsing Tanggal (Mendukung Serial Angka Excel & String)
    tgl_col = None
    for col in df_group.columns:
        c_upper = str(col).upper()
        if 'TANGGAL' in c_upper or 'DATE' in c_upper or 'TGL' in c_upper:
            tgl_col = col
            break
    if tgl_col is None:
        tgl_col = df_group.columns[0]

    df_group[tgl_col] = df_group[tgl_col].ffill()

    def parse_excel_date(val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (pd.Timestamp, datetime)):
            return pd.to_datetime(val)
        if isinstance(val, (int, float)):
            try:
                # Konversi angka serial Excel (misal: 45800) ke datetime
                return pd.to_datetime(val, unit='D', origin='1899-12-30')
            except:
                return pd.NaT
        try:
            return pd.to_datetime(val, errors='coerce')
        except:
            return pd.NaT

    df_group['Parsed_Date'] = df_group[tgl_col].apply(parse_excel_date)
    days_map = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}
    df_group['Nama_Hari'] = df_group['Parsed_Date'].dt.weekday.map(days_map)
    df_group['Minggu_Ke'] = df_group['Parsed_Date'].apply(lambda dt: f"Minggu ke-{dt.isocalendar()[1]}" if pd.notna(dt) else "Tidak Valid")

    # 3. Identifikasi Kolom Mesin
    mesin_col = None
    for col in df_group.columns:
        c_upper = str(col).upper()
        if 'MESIN' in c_upper or 'MACHINE' in c_upper:
            mesin_col = col
            break
    if mesin_col is None:
        mesin_col = df_group.columns[1] if len(df_group.columns) > 1 else df_group.columns[0]

    df_group['Mesin_Clean'] = df_group[mesin_col].astype(str).str.strip().str.upper()

    # 4. Identifikasi & Perbaikan Parsing Kolom Persentase
    col_persen_name = None
    for col in df_group.columns:
        c_upper = str(col).upper()
        if 'PERSENTASE' in c_upper or '%' in c_upper:
            col_persen_name = col
            break
    if col_persen_name is None:
        col_persen_name = df_group.columns[-1]

    def parse_percentage_value(val):
        if pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            return val * 100.0 if val <= 1.0 else float(val)
        try:
            return float(str(val).replace('%', '').strip().replace(',', '.'))
        except:
            return None

    df_group['Parsed_Persen'] = df_group[col_persen_name].apply(parse_percentage_value)

    print("3. Memproses pembuatan 4 File Output...")

    # Output 1: Rekap Performa Mesin per Minggu (Excel)
    file_1 = "Output_1_Rekap_Mesin_Mingguan.xlsx"
    summary_df = df_group.groupby(['Mesin_Clean', 'Minggu_Ke']).agg(
        Rata_Rata_Persentase=('Parsed_Persen', 'mean'),
        Total_Data=('Parsed_Persen', 'count')
    ).reset_index()
    summary_df.to_excel(file_1, index=False)

    # Output 2: Laporan Lengkap Data Terbersih (Excel)
    file_2 = "Output_2_Data_Cleaned_Full.xlsx"
    df_group.to_excel(file_2, index=False)

    # Output 3: Ringkasan Harian per Mesin (CSV)
    file_3 = "Output_3_Ringkasan_Harian.csv"
    daily_df = df_group.groupby(['Mesin_Clean', 'Parsed_Date']).agg(
        Rata_Persentase=('Parsed_Persen', 'mean')
    ).reset_index()
    daily_df.to_csv(file_3, index=False)

    # Output 4: Laporan Eksekutif (TXT / Ringkasan Dokumen)
    file_4 = "Output_4_Laporan_Eksekutif.txt"
    with open(file_4, 'w', encoding='utf-8') as f:
        f.write("=== LAPORAN EKSEKUTIF KINERJA MESIN ===\n")
        f.write(f"Tanggal Generate: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        f.write(f"Total Baris Data Diproses: {len(df_group)}\n\n")
        f.write(summary_df.to_string(index=False))

    print("Selesai! Keempat file berhasil di-generate.")

if __name__ == '__main__':
    process_and_generate_outputs()
