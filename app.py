import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def process_machine_recap(input_filename, output_filename="Hasil_Rekap_Per_Mesin_Fixed.xlsx"):
    print(f"Membaca file: {input_filename}...")
    
    # 1. Baca semua sheet dari file Excel
    try:
        xls = pd.ExcelFile(input_filename)
        sheet_names = xls.sheet_names
    except Exception as e:
        print(f"Gagal membaca file Excel: {e}")
        return

    all_combined_dfs = []
    
    for s_name in sheet_names:
        try:
            df_temp = pd.read_excel(input_filename, sheet_name=s_name)
            if not df_temp.empty:
                # Tambahkan informasi asal sheet jika diperlukan
                df_temp['Source_Sheet'] = s_name
                all_combined_dfs.append(df_temp)
        except Exception as e:
            print(f"Catatan: Sheet {s_name} dilewati karena error: {e}")

    if not all_combined_dfs:
        print("Tidak ada data yang berhasil dibaca dari sheet manapun.")
        return

    # Gabungkan semua data
    df_group = pd.concat(all_combined_dfs, ignore_index=True)

    # 2. Identifikasi Kolom Tanggal secara Dinamis
    tgl_col = None
    for col in df_group.columns:
        c_upper = str(col).upper()
        if 'TANGGAL' in c_upper or 'DATE' in c_upper or 'TGL' in c_upper:
            tgl_col = col
            break
    
    if tgl_col is None:
        # Fallback ambil kolom pertama jika tidak ditemukan
        tgl_col = df_group.columns[0]
    
    # Forward fill tanggal jika ada cell kosong di bawahnya (merged cell style di Excel)
    df_group[tgl_col] = df_group[tgl_col].ffill()

    # 3. Fungsi Konversi Tanggal yang Aman (Mendukung Serial Excel & String)
    def parse_excel_date(val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (pd.Timestamp, datetime)):
            return pd.to_datetime(val)
        if isinstance(val, (int, float)):
            try:
                # Konversi angka serial Excel ke datetime
                return pd.to_datetime(val, unit='D', origin='1899-12-30')
            except:
                return pd.NaT
        try:
            return pd.to_datetime(val, errors='coerce')
        except:
            return pd.NaT

    df_group['Parsed_Date'] = df_group[tgl_col].apply(parse_excel_date)

    # Pemetaan Hari
    days_map = {0: 'Senin', 1: 'Selasa', 2: 'Rabu', 3: 'Kamis', 4: 'Jumat', 5: 'Sabtu', 6: 'Minggu'}
    df_group['Nama_Hari'] = df_group['Parsed_Date'].dt.weekday.map(days_map)

    # Fungsi Pemetaan Minggu Kalender
    def get_calendar_week_mapping(dt):
        if pd.isna(dt):
            return "Minggu Tidak Valid"
        # Contoh pengelompokan berdasarkan tanggal dalam bulan atau minggu ke-X
        # Anda dapat menyesuaikan logika mingguan ini sesuai kebutuhan bisnis Anda
        week_num = dt.isocalendar()[1]
        return f"Minggu ke-{week_num}"

    df_group['Minggu_Ke'] = df_group['Parsed_Date'].apply(get_calendar_week_mapping)

    # 4. Identifikasi Kolom Mesin secara Dinamis
    mesin_col = None
    for col in df_group.columns:
        c_upper = str(col).upper()
        if 'MESIN' in c_upper or 'MACHINE' in c_upper or 'NO MESIN' in c_upper:
            mesin_col = col
            break
    
    if mesin_col is None:
        mesin_col = df_group.columns[1] if len(df_group.columns) > 1 else df_group.columns[0]

    df_group['Mesin_Clean'] = df_group[mesin_col].astype(str).str.strip().str.upper()

    # 5. Identifikasi Kolom Persentase & Total Cones secara Dinamis
    col_persen_name = None
    for col in df_group.columns:
        c_upper = str(col).upper()
        if 'PERSENTASE' in c_upper or '%' in c_upper:
            col_persen_name = col
            break
            
    if col_persen_name is None:
        col_persen_name = df_group.columns[-1] # Default ambil kolom terakhir jika tidak ketemu

    # Helper parsing persentase
    def parse_percentage_value(val):
        if pd.isna(val):
            return None
        if isinstance(val, (int, float)):
            # Jika nilainya desimal (misal 0.85) ubah ke skala 100 (85)
            if val <= 1.0:
                return val * 100.0
            return float(val)
        try:
            clean_str = str(val).replace('%', '').strip().replace(',', '.')
            return float(clean_str)
        except:
            return None

    # Filter data valid (misal persentase tidak kosong dan di atas threshold minimum tertentu jika ada)
    df_group['Parsed_Persen'] = df_group[col_persen_name].apply(parse_percentage_value)
    
    # Contoh rekapitulasi per Mesin dan per Minggu
    print("Memproses rekapitulasi data...")
    summary_df = df_group.groupby(['Mesin_Clean', 'Minggu_Ke']).agg(
        Rata_Rata_Persentase=('Parsed_Persen', 'mean'),
        Total_Data_Masuk=('Parsed_Persen', 'count')
    ).reset_index()

    # Simpan ke file Excel baru dengan format rapi
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df_group.to_excel(writer, sheet_name='Data_Cleaned', index=False)
        summary_df.to_excel(writer, sheet_name='Rekap_Mingguan', index=False)

    print(f"Selesai! File rekap berhasil disimpan sebagai: {output_filename}")

# Contoh cara menjalankan fungsi di atas:
# process_machine_recap('REKAP HASIL PER MESIN.xlsx')
