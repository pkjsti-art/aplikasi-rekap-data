import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

st.set_page_config(page_title="Aplikasi Rekap Data Mesin", layout="wide")

st.title("📊 Aplikasi Rekapitulasi Hasil Per Mesin")
st.write("Upload file Excel `REKAP HASIL PER MESIN.xlsx` untuk memproses rekapitulasi mingguan secara otomatis.")

# Widget Upload File
uploaded_file = st.file_uploader("Pilih file Excel sumber", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        with st.spinner("Sedang memproses data..."):
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names

            all_combined_dfs = []
            for s_name in sheet_names:
                df_temp = pd.read_excel(uploaded_file, sheet_name=s_name)
                if not df_temp.empty:
                    df_temp['Source_Sheet'] = s_name
                    all_combined_dfs.append(df_temp)

            if not all_combined_dfs:
                st.error("File Excel kosong atau tidak memiliki data yang valid.")
            else:
                df_group = pd.concat(all_combined_dfs, ignore_index=True)

                # Identifikasi Tanggal
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

                # Identifikasi Mesin
                mesin_col = None
                for col in df_group.columns:
                    c_upper = str(col).upper()
                    if 'MESIN' in c_upper or 'MACHINE' in c_upper:
                        mesin_col = col
                        break
                if mesin_col is None:
                    mesin_col = df_group.columns[1] if len(df_group.columns) > 1 else df_group.columns[0]

                df_group['Mesin_Clean'] = df_group[mesin_col].astype(str).str.strip().str.upper()

                # Identifikasi Persentase
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

                # Buat Data Rekap
                summary_df = df_group.groupby(['Mesin_Clean', 'Minggu_Ke']).agg(
                    Rata_Rata_Persentase=('Parsed_Persen', 'mean'),
                    Total_Data=('Parsed_Persen', 'count')
                ).reset_index()

                daily_df = df_group.groupby(['Mesin_Clean', 'Parsed_Date']).agg(
                    Rata_Persentase=('Parsed_Persen', 'mean')
                ).reset_index()

                st.success("Data berhasil diproses dengan sukses!")

                # Tampilkan Preview Data
                st.subheader("Preview Rekap Performa Mesin per Minggu")
                st.dataframe(summary_df)

                # Tombol-tombol Download untuk 4 File Output
                st.markdown("---")
                st.subheader("📥 Download File Hasil Rekap")

                col1, col2 = st.columns(2)

                with col1:
                    # Output 1: Excel Rekap Mingguan
                    output_1 = io.BytesIO()
                    with pd.ExcelWriter(output_1, engine='openpyxl') as writer:
                        summary_df.to_excel(writer, index=False)
                    st.download_button(
                        label="⬇️ Download Rekap Mingguan (Excel)",
                        data=output_1.getvalue(),
                        file_name="Output_1_Rekap_Mesin_Mingguan.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # Output 2: Excel Data Cleaned Full
                    output_2 = io.BytesIO()
                    with pd.ExcelWriter(output_2, engine='openpyxl') as writer:
                        df_group.to_excel(writer, index=False)
                    st.download_button(
                        label="⬇️ Download Data Cleaned Full (Excel)",
                        data=output_2.getvalue(),
                        file_name="Output_2_Data_Cleaned_Full.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                with col2:
                    # Output 3: CSV Ringkasan Harian
                    csv_data = daily_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="⬇️ Download Ringkasan Harian (CSV)",
                        data=csv_data,
                        file_name="Output_3_Ringkasan_Harian.csv",
                        mime="text/csv"
                    )

                    # Output 4: TXT Laporan Eksekutif
                    txt_content = f"=== LAPORAN EKSEKUTIF KINERJA MESIN ===\nTanggal: {datetime.now()}\n\n" + summary_df.to_string(index=False)
                    st.download_button(
                        label="⬇️ Download Laporan Eksekutif (TXT)",
                        data=txt_content.encode('utf-8'),
                        file_name="Output_4_Laporan_Eksekutif.txt",
                        mime="text/plain"
                    )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
    st.info("Silakan upload file Excel terlebih dahulu melalui tombol di atas untuk memulai.")
