import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# Konfigurasi halaman web
st.set_page_config(
    page_title="Aplikasi Rekapitulasi Data Pabrik",
    page_icon="📊",
    layout="wide"
)

# Styling CSS tambahan agar tampilan web lebih estetik dan rapi
st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        color: #1f77b4;
        font-weight: 700;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #555555;
        margin-bottom: 25px;
    }
    .step-box {
        background-color: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #1f77b4;
        margin-bottom: 20px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">📊 Aplikasi Rekapitulasi & Analisis Data Pabrik</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Sistem Pipeline 4 Tahap: Data Bersih ➔ Karyawan ➔ Mesin ➔ Summary Mingguan</p>', unsafe_allow_html=True)

# ==========================================
# SIDEBAR: INPUT FILE TAHAP 1
# ==========================================
st.sidebar.header("📁 Upload File Sumber (Tahap 1)")
st.sidebar.write("Silakan upload kedua file master Anda untuk memulai proses.")

uploaded_master = st.sidebar.file_uploader("1. Upload File Master Utama", type=["xlsx", "xls"])
uploaded_indikator = st.sidebar.file_uploader("2. Upload File Master Indikator", type=["xlsx", "xls"])

# Tombol Eksekusi Pipeline
run_pipeline = st.sidebar.button("🚀 Jalankan Proses Pipeline", type="primary", use_container_width=True)

# ==========================================
# UTAMA: LOGIKA 4 TAHAP PIPELINE
# ==========================================
if run_pipeline:
    if uploaded_master is not None and uploaded_indikator is not None:
        with st.spinner("Sedang memproses seluruh tahapan pipeline data..."):
            try:
                # --- TAHAP 1: PEMBUATAN DATA BERSIH ---
                xls_master = pd.ExcelFile(uploaded_master)
                sheet_names_master = xls_master.sheet_names
                
                all_combined_dfs = []
                for s_name in sheet_names_master:
                    df_temp = pd.read_excel(uploaded_master, sheet_name=s_name)
                    if not df_temp.empty:
                        df_temp['Source_Sheet'] = s_name
                        all_combined_dfs.append(df_temp)
                
                df_master = pd.concat(all_combined_dfs, ignore_index=True)
                df_indikator = pd.read_excel(uploaded_indikator)

                # Pembersihan dan standarisasi Tanggal & Kolom
                tgl_col = next((col for col in df_master.columns if 'TANGGAL' in str(col).upper() or 'DATE' in str(col).upper()), df_master.columns[0])
                df_master[tgl_col] = df_master[tgl_col].ffill()

                def parse_excel_date(val):
                    if pd.isna(val): return pd.NaT
                    if isinstance(val, (pd.Timestamp, datetime)): return pd.to_datetime(val)
                    if isinstance(val, (int, float)):
                        try: return pd.to_datetime(val, unit='D', origin='1899-12-30')
                        except: return pd.NaT
                    try: return pd.to_datetime(val, errors='coerce')
                    except: return pd.NaT

                df_master['Parsed_Date'] = df_master[tgl_col].apply(parse_excel_date)
                df_master['Minggu_Ke'] = df_master['Parsed_Date'].apply(lambda dt: f"Minggu ke-{dt.isocalendar()[1]}" if pd.notna(dt) else "Tidak Valid")

                # Standarisasi Mesin & Karyawan
                mesin_col = next((col for col in df_master.columns if 'MESIN' in str(col).upper()), df_master.columns[1] if len(df_master.columns)>1 else df_master.columns[0])
                df_master['Mesin_Clean'] = df_master[mesin_col].astype(str).str.strip().str.upper()

                karyawan_col = next((col for col in df_master.columns if 'KARYAWAN' in str(col).upper() or 'NAMA' in str(col).upper() or 'OPERATOR' in str(col).upper()), df_master.columns[2] if len(df_master.columns)>2 else df_master.columns[0])
                df_master['Karyawan_Clean'] = df_master[karyawan_col].astype(str).str.strip().str.title()

                # Persentase parsing
                col_persen_name = next((col for col in df_master.columns if 'PERSENTASE' in str(col).upper() or '%' in str(col).upper()), df_master.columns[-1])
                def parse_percentage_value(val):
                    if pd.isna(val): return None
                    if isinstance(val, (int, float)): return val * 100.0 if val <= 1.0 else float(val)
                    try: return float(str(val).replace('%', '').strip().replace(',', '.'))
                    except: return None

                df_master['Parsed_Persen'] = df_master[col_persen_name].apply(parse_percentage_value)

                # OUTPUT 1: Data Bersih
                df_data_bersih = df_master.copy()

                # --- TAHAP 2: REKAP HASIL PER KARYAWAN (Mengambil dari Output 1) ---
                df_rekap_karyawan = df_data_bersih.groupby('Karyawan_Clean').agg(
                    Total_Pekerjaan=('Parsed_Persen', 'count'),
                    Rata_Rata_Persentase=('Parsed_Persen', 'mean')
                ).reset_index()

                # --- TAHAP 3: REKAP HASIL PER MESIN (Mengambil dari Output 1) ---
                df_rekap_mesin = df_data_bersih.groupby(['Mesin_Clean', 'Minggu_Ke']).agg(
                    Total_Operasi=('Parsed_Persen', 'count'),
                    Rata_Rata_Persentase=('Parsed_Persen', 'mean')
                ).reset_index()

                # --- TAHAP 4: SUMMARY REKAP PER MINGGU (Mengambil dari Output 3) ---
                df_summary_mingguan = df_rekap_mesin.groupby('Minggu_Ke').agg(
                    Total_Mesin_Beroperasi=('Mesin_Clean', 'nunique'),
                    Rata_Rata_Kinerja_Mingguan=('Rata_Rata_Persentase', 'mean')
                ).reset_index()

                st.success("✅ Seluruh 4 tahapan pipeline berhasil dieksekusi dengan sempurna!")

                # ==========================================
                # TAMPILAN TAB HASIL (USER FRIENDLY)
                # ==========================================
                tab1, tab2, tab3, tab4 = st.tabs([
                    "📁 1. Data Bersih", 
                    "👤 2. Rekap Karyawan", 
                    "⚙️ 3. Rekap Mesin", 
                    "📅 4. Summary Mingguan"
                ])

                with tab1:
                    st.subheader("Hasil Tahap 1: Data Bersih (Gabungan Master & Indikator)")
                    st.dataframe(df_data_bersih, use_container_width=True)
                    
                    # Tombol Download
                    out1 = io.BytesIO()
                    with pd.ExcelWriter(out1, engine='openpyxl') as writer:
                        df_data_bersih.to_excel(writer, index=False)
                    st.download_button("⬇️ Download Excel Data Bersih", out1.getvalue(), "1_Data_Bersih.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                with tab2:
                    st.subheader("Hasil Tahap 2: Rekap Hasil per Karyawan")
                    st.dataframe(df_rekap_karyawan, use_container_width=True)
                    
                    out2 = io.BytesIO()
                    with pd.ExcelWriter(out2, engine='openpyxl') as writer:
                        df_rekap_karyawan.to_excel(writer, index=False)
                    st.download_button("⬇️ Download Excel Rekap Karyawan", out2.getvalue(), "2_Rekap_Karyawan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                with tab3:
                    st.subheader("Hasil Tahap 3: Rekap Hasil per Mesin")
                    st.dataframe(df_rekap_mesin, use_container_width=True)
                    
                    out3 = io.BytesIO()
                    with pd.ExcelWriter(out3, engine='openpyxl') as writer:
                        df_rekap_mesin.to_excel(writer, index=False)
                    st.download_button("⬇️ Download Excel Rekap Mesin", out3.getvalue(), "3_Rekap_Mesin.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                with tab4:
                    st.subheader("Hasil Tahap 4: Summary Rekap per Minggu")
                    st.dataframe(df_summary_mingguan, use_container_width=True)
                    
                    out4 = io.BytesIO()
                    with pd.ExcelWriter(out4, engine='openpyxl') as writer:
                        df_summary_mingguan.to_excel(writer, index=False)
                    st.download_button("⬇️ Download Excel Summary Mingguan", out4.getvalue(), "4_Summary_Rekap_Mingguan.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            except Exception as e:
                st.error(f"Terjadi kesalahan saat memproses data di pipeline: {e}")
    else:
        st.warning("⚠️ Harap upload **File Master Utama** dan **File Master Indikator** terlebih dahulu melalui panel di sebelah kiri.")
else:
    st.info("👈 Silakan upload file di panel sidebar kiri, lalu klik tombol **'Jalankan Proses Pipeline'** untuk melihat hasil dari 4 tahapan.")
