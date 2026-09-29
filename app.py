import calendar
import io
import re
from datetime import datetime
import numpy as np
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Konfigurasi Halaman Streamlit
st.set_page_config(
    page_title="Pengelola Data PCP", page_icon="📊", layout="wide"
)

st.title("📊 Pengelola Data PCP")
st.markdown(
    "Aplikasi otomatisasi pengolahan data produksi, rekap per mesin, hingga summary rekap per minggu."
)

# Sidebar untuk Input Pengguna
st.sidebar.header("⚙️ Konfigurasi Parameter")
bulan_input = (
    st.sidebar.selectbox(
        "Pilih Bulan Acuan",
        [
            "JANUARI",
            "FEBRUARI",
            "MARET",
            "APRIL",
            "MEI",
            "JUNI",
            "JULI",
            "AGUSTUS",
            "SEPTEMBER",
            "OKTOBER",
            "NOVEMBER",
            "DESEMBER",
        ],
        index=8,  # Default September
    )
    .strip()
    .upper()
)

min_threshold_input = st.sidebar.number_input(
    "Batas Minimum Persentase (%) untuk Highlight",
    min_value=0.0,
    max_value=100.0,
    value=60.0,
    step=1.0,
)

st.sidebar.markdown("---")
st.sidebar.subheader("📂 Unggah Berkas Sumber")
uploaded_clean = st.sidebar.file_uploader(
    "Upload File Data Mentah/Bersih (.xlsx)", type=["xlsx"]
)
uploaded_master = st.sidebar.file_uploader(
    "Upload File Master Indikator (.xlsx)", type=["xlsx"]
)

# Inisialisasi Session State agar halaman tidak reset saat tombol download diklik
if "processed" not in st.session_state:
  st.session_state.processed = False
if "output_rekap_bytes" not in st.session_state:
  st.session_state.output_rekap_bytes = None
if "output_summary_bytes" not in st.session_state:
  st.session_state.output_summary_bytes = None

if st.sidebar.button("Proses Data", type="primary"):
  if uploaded_clean is not None and uploaded_master is not None:
    with st.spinner(
        "Sedang memproses data, menghitung shift, dan membuat rekap..."
    ):
      try:
        # --- BAGIAN 1 & 2: MEMBACA & MEMBERSIHKAN DATA ---
        df_clean = pd.read_excel(uploaded_clean)
        df_master = pd.read_excel(uploaded_master)

        def clean_str(val):
          if pd.isna(val):
            return ""
          text = str(val)
          text = re.sub(r"\s+", "", text).upper()
          return text

        df_clean["Mesin_Clean"] = df_clean["Mesin"].apply(clean_str)
        df_master["clean_var"] = df_master["Nama Var."].apply(clean_str)

        spindle_col_name = None
        for col in df_master.columns:
          col_up = str(col).upper()
          if "SPIND" in col_up or "SPIN" in col_up:
            spindle_col_name = col
            break
        if spindle_col_name is None:
          spindle_col_name = "SPINDEL"

        def get_master_data(benang_val, mesin_val, col_name):
          b_clean = clean_str(benang_val)
          m_clean = clean_str(mesin_val)
          matched = df_master[df_master["clean_var"] == b_clean]
          if matched.empty:
            return 0.0
          for idx, row in matched.iterrows():
            master_machine_str = str(row["MESIN"])
            machines = [
                clean_str(m) for m in master_machine_str.split("/")
            ]
            if m_clean in machines or any(
                m in m_clean or m_clean in m for m in machines
            ):
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

        # Deteksi Shift
        shift_counts = (
            df_clean.groupby(["Tanggal", "Pegawai", "Mesin_Clean"])["Shift"]
            .nunique()
            .reset_index()
        )
        shift_counts["Status_Shift"] = shift_counts["Shift"].apply(
            lambda x: "PANJANG" if x > 1 else "PENDEK"
        )
        df_clean = pd.merge(
            df_clean,
            shift_counts[
                ["Tanggal", "Pegawai", "Mesin_Clean", "Status_Shift"]
            ],
            on=["Tanggal", "Pegawai", "Mesin_Clean"],
            how="left",
        )

        def determine_shift_final(row):
          shift_val = str(row["Shift"]).strip().upper()
          status_val = row["Status_Shift"]
          if status_val == "PENDEK":
            return f"{shift_val} / PENDEK"
          else:
            group_rows = df_clean[
                (df_clean["Tanggal"] == row["Tanggal"])
                & (df_clean["Pegawai"] == row["Pegawai"])
                & (df_clean["Mesin_Clean"] == row["Mesin_Clean"])
            ]
            shifts_in_group = set(
                group_rows["Shift"].astype(str).str.strip().str.upper()
            )
            if "SIANG" in shifts_in_group and "MALAM" in shifts_in_group:
              return "MALAM / PANJANG"
            elif "PAGI" in shifts_in_group and "SIANG" in shifts_in_group:
              return "PAGI / PANJANG"
            else:
              return f"{shift_val} / PANJANG"

        df_clean["Shift_Final"] = df_clean.apply(
            determine_shift_final, axis=1
        )

        # Perhitungan Qty Real Cones dengan pengaman .empty
        group_max_qty = (
            df_clean.groupby(["Tanggal", "Pegawai"])["Qty Cones"]
            .max()
            .reset_index()
        )
        group_max_qty = group_max_qty.rename(columns={"Qty Cones": "max_qty"})
        df_clean = pd.merge(
            df_clean, group_max_qty, on=["Tanggal", "Pegawai"], how="left"
        )

        real_cones_list = []
        for idx, row in df_clean.iterrows():
          q_val = row["Qty Cones"]
          max_q = row["max_qty"]
          if q_val == max_q:
            real_cones_list.append(round(q_val))
          else:
            b_val = row["Benang"]
            m_val = row["Mesin"]
            waktu_current = get_master_data(b_val, m_val, "WAKTU (MENIT)")
            group_rows = df_clean[
                (df_clean["Tanggal"] == row["Tanggal"])
                & (df_clean["Pegawai"] == row["Pegawai"])
            ]
            sub_max = group_rows[group_rows["Qty Cones"] == max_q]
            if not sub_max.empty:
              row_max = sub_max.iloc[0]
              waktu_max_qty_item = get_master_data(
                  row_max["Benang"], row_max["Mesin"], "WAKTU (MENIT)"
              )
            else:
              waktu_max_qty_item = 0

            if waktu_max_qty_item > 0:
              real_cones = (q_val * waktu_current) / waktu_max_qty_item
            else:
              real_cones = q_val
            real_cones_list.append(round(real_cones))
        df_clean["Qty Real Cones"] = real_cones_list

        # --- BAGIAN 3: PEMBUATAN REKAP HASIL PER MESIN ---
        wb = Workbook()
        wb.remove(wb.active)
        machines_raw = df_clean["Mesin"].unique()

        def machine_sort_key(m):
          m_str = str(m).strip().upper()
          custom_order = {
              "SSM 1A": 1,
              "SSM 1B": 2,
              "SSM 2A": 3,
              "SSM 2B": 4,
              "SSM 3A": 5,
              "SSM 3B": 6,
          }
          for key, rank in custom_order.items():
            if key in m_str:
              return (0, rank, m_str)
          return (1, 99, m_str)

        machines = sorted(machines_raw, key=machine_sort_key)
        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )
        header_fill = PatternFill(
            start_color="9ED3DC", end_color="9ED3DC", fill_type="solid"
        )
        color_list = ["BDB2FF", "FFCEE3", "FDFFD2", "CAEDFF"]

        for mach in machines:
          safe_mach_name = str(mach).strip()
          sheet_name = re.sub(r"[:\\/?*\[\]]", "_", safe_mach_name)[:31]
          if not sheet_name:
            sheet_name = "MESIN"
          ws = wb.create_sheet(title=sheet_name)
          df_mach = df_clean[df_clean["Mesin"] == mach].copy()

          ws.merge_cells("F1:G1")
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
            ws.merge_cells(
                start_row=1, start_column=col_idx, end_row=2, end_column=col_idx
            )
          for r in [1, 2]:
            for c in range(1, 15):
              ws.cell(row=r, column=c).fill = header_fill

          current_row = 3
          df_mach["Tanggal_Parsed"] = pd.to_datetime(
              df_mach["Tanggal"], errors="coerce"
          )
          group_keys = (
              df_mach.sort_values(["Tanggal_Parsed", "Pegawai", "Shift"])[
                  ["Tanggal", "Pegawai", "Shift"]
              ]
              .drop_duplicates()
              .values
          )
          unique_dates = sorted(df_mach["Tanggal_Parsed"].dropna().unique())
          date_to_color_idx = {
              dt: idx % len(color_list) for idx, dt in enumerate(unique_dates)
          }

          for dt, peg, shf in group_keys:
            df_group = df_mach[
                (df_mach["Tanggal"] == dt)
                & (df_mach["Pegawai"] == peg)
                & (df_mach["Shift"] == shf)
            ]
            if df_group.empty:
              continue
            start_group_row = current_row
            total_real_cones = (
                int(df_group["Qty Real Cones"].sum())
                if df_group["Qty Real Cones"].sum().is_integer()
                else df_group["Qty Real Cones"].sum()
            )
            shift_status = df_group["Status_Shift"].iloc[0]
            max_row_in_group = df_group.loc[df_group["Qty Cones"].idxmax()]
            ref_benang = max_row_in_group["Benang"]
            ref_mesin = max_row_in_group["Mesin"]

            spindle_val = get_master_data(
                ref_benang, ref_mesin, spindle_col_name
            )
            if spindle_val.is_integer():
              spindle_val = int(spindle_val)

            col_t90 = (
                "TARGET 90% SHIFT PENDEK"
                if shift_status == "PENDEK"
                else "TARGET 90% SHIFT PANJANG"
            )
            col_t100 = (
                "TARGET 100% SHIFT PENDEK"
                if shift_status == "PENDEK"
                else "TARGET 100% SHIFT PANJANG"
            )
            target_90 = get_master_data(ref_benang, ref_mesin, col_t90)
            if target_90.is_integer():
              target_90 = int(target_90)
            target_100 = get_master_data(ref_benang, ref_mesin, col_t100)
            if target_100.is_integer():
              target_100 = int(target_100)

            for idx, row in df_group.iterrows():
              benang = row["Benang"]
              mesin = row["Mesin"]
              qty_cones_val = row["Qty Cones"]
              if qty_cones_val.is_integer():
                qty_cones_val = int(qty_cones_val)
              kg_val = row["Total"] if "Total" in row else 0.0
              q_real_val = row["Qty Real Cones"]
              if q_real_val.is_integer():
                q_real_val = int(q_real_val)

              ws.cell(row=current_row, column=1, value=str(row["Tanggal"])[:10])
              ws.cell(row=current_row, column=2, value=row["Pegawai"])
              ws.cell(row=current_row, column=3, value=mesin)
              cell_spindle = ws.cell(
                  row=current_row, column=4, value=spindle_val
              )
              cell_spindle.number_format = "#,##0"
              ws.cell(row=current_row, column=5, value=row["Shift_Final"])
              ws.cell(row=current_row, column=6, value=benang)
              cell_qreal = ws.cell(row=current_row, column=7, value=q_real_val)
              cell_qreal.number_format = "#,##0"
              cell_tot = ws.cell(row=current_row, column=8, value=total_real_cones)
              cell_tot.number_format = "#,##0"
              cell_t90 = ws.cell(row=current_row, column=9, value=target_90)
              cell_t90.number_format = (
                  "#,##0" if isinstance(target_90, int) else "#,##0.00"
              )
              p90 = (
                  (total_real_cones / target_90) if target_90 > 0 else 0.0
              )
              cell_p90 = ws.cell(row=current_row, column=10, value=p90)
              cell_p90.number_format = "0.0%"
              cell_t100 = ws.cell(row=current_row, column=11, value=target_100)
              cell_t100.number_format = (
                  "#,##0" if isinstance(target_100, int) else "#,##0.00"
              )
              p100 = (
                  (total_real_cones / target_100) if target_100 > 0 else 0.0
              )
              cell_p100 = ws.cell(row=current_row, column=12, value=p100)
              cell_p100.number_format = "0.0%"
              cell_cones = ws.cell(
                  row=current_row, column=13, value=qty_cones_val
              )
              cell_cones.number_format = "#,##0"
              cell_kg = ws.cell(row=current_row, column=14, value=kg_val)
              cell_kg.number_format = "#,##0.00"
              current_row += 1

            end_group_row = current_row - 1
            if end_group_row > start_group_row:
              for col_idx in [1, 2, 3, 4, 5, 8, 9, 10, 11, 12]:
                ws.merge_cells(
                    start_row=start_group_row,
                    start_column=col_idx,
                    end_row=end_group_row,
                    end_column=col_idx,
                )

            dt_parsed = pd.to_datetime(dt, errors="coerce")
            if dt_parsed in date_to_color_idx:
              c_idx = date_to_color_idx[dt_parsed]
              hex_color = color_list[c_idx]
              group_fill = PatternFill(
                  start_color=hex_color, end_color=hex_color, fill_type="solid"
              )
              for r in range(start_group_row, end_group_row + 1):
                for c in range(1, 15):
                  ws.cell(row=r, column=c).fill = group_fill

          for row in ws.iter_rows(
              min_row=1, max_row=ws.max_row, min_col=1, max_col=14
          ):
            for cell in row:
              cell.border = thin_border
              cell.alignment = Alignment(
                  horizontal="center", vertical="center", wrap_text=True
              )

          min_widths = {1: 14, 2: 18, 3: 12, 4: 15, 5: 18, 6: 25, 7: 14}
          for col_idx in range(1, 15):
            col_letter = get_column_letter(col_idx)
            max_len = 0
            for cell in ws[col_letter]:
              val_str = str(cell.value or "")
              for line in val_str.split("\n"):
                if len(line) > max_len:
                  max_len = len(line)
            calculated_width = max(max_len + 3, 12)
            if col_idx in min_widths:
              calculated_width = max(
                  calculated_width, min_widths[col_idx]
              )
            ws.column_dimensions[col_letter].width = calculated_width

        # Simpan rekap per mesin ke bytes buffer
        rekap_io = io.BytesIO()
        wb.save(rekap_io)
        rekap_io.seek(0)

        # --- BAGIAN 4: SUMMARY REKAP PER MINGGU ---
        wb_src = load_workbook(rekap_io, data_only=True)
        sheet_names = wb_src.sheetnames

        tahun_input = 2026
        bulan_dict = {
            "JANUARI": 1,
            "FEBRUARI": 2,
            "MARET": 3,
            "APRIL": 4,
            "MEI": 5,
            "JUNI": 6,
            "JULI": 7,
            "AGUSTUS": 8,
            "SEPTEMBER": 9,
            "OKTOBER": 10,
            "NOVEMBER": 11,
            "DESEMBER": 12,
        }
        target_month_num = bulan_dict.get(bulan_input, 9)

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
            if dt_date < month_weeks[0][0]:
              return 1
            elif dt_date > month_weeks[-1][-1]:
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
              p_str = (
                  str(val_raw).replace("%", "").replace(",", ".").strip()
              )
              return float(p_str)
          except:
            return None

        wb_out = Workbook()
        wb_out.remove(wb_out.active)
        days_order = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU"]
        days_map = {
            0: "SENIN",
            1: "SELASA",
            2: "RABU",
            3: "KAMIS",
            4: "JUMAT",
            5: "SABTU",
            6: "MINGGU",
        }

        thin_side = Side(style="thin", color="000000")
        thin_border = Border(
            left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
        )
        bold_font = Font(name="Calibri", size=11, bold=True)
        header_font = Font(name="Calibri", size=11, bold=True, color="000000")
        sub_header_font = Font(name="Calibri", size=10, bold=True, color="000000")
        highlight_fill = PatternFill(
            start_color="E11A45", end_color="E11A45", fill_type="solid"
        )
        highlight_font = Font(
            name="Calibri", size=11, bold=True, color="FFFFFF"
        )

        kelompok_dict = {}
        for s_name in sheet_names:
          s_upper = s_name.upper()
          if "SSM 1" in s_upper or "SSM1" in s_upper:
            group_key = "SSM 1"
          elif "SSM 2" in s_upper or "SSM2" in s_upper:
            group_key = "SSM 2"
          elif "SSM 3" in s_upper or "SSM3" in s_upper:
            group_key = "SSM 3"
          else:
            group_key = s_upper.split()[0] if s_upper else "LAINNYA"
          if group_key not in kelompok_dict:
            kelompok_dict[group_key] = []
          kelompok_dict[group_key].append(s_name)

        rekap_io.seek(0)
        for group_name, s_list in kelompok_dict.items():
          ws_out = wb_out.create_sheet(title=group_name)
          current_max_row = 1
          all_combined_dfs = []

          for sheet_idx, s_name in enumerate(s_list):
            ws_src_sheet = wb_src[s_name]
            df_temp = pd.read_excel(rekap_io, sheet_name=s_name)
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
                target_cell = ws_out.cell(
                    row=current_max_row,
                    column=cell.column,
                    value=cell.value,
                )
                if cell.has_style:
                  if cell.font:
                    target_cell.font = Font(
                        name=cell.font.name,
                        size=cell.font.size,
                        bold=cell.font.bold,
                        italic=cell.font.italic,
                        color=cell.font.color,
                    )
                  if cell.fill and cell.fill.fill_type:
                    target_cell.fill = cell.fill.__copy__()
                  if cell.alignment:
                    target_cell.alignment = Alignment(
                        horizontal=cell.alignment.horizontal,
                        vertical=cell.alignment.vertical,
                        wrap_text=cell.alignment.wrap_text,
                    )
                  if cell.border:
                    target_cell.border = cell.border.__copy__()
                  if cell.number_format:
                    target_cell.number_format = cell.number_format
              current_max_row += 1

            row_shift = target_start_row - 1
            for merged_range in ws_src_sheet.merged_cells.ranges:
              try:
                ws_out.merge_cells(
                    start_row=merged_range.min_row + row_shift,
                    start_column=merged_range.min_col,
                    end_row=merged_range.max_row + row_shift,
                    end_column=merged_range.max_col,
                )
              except:
                pass

          if not all_combined_dfs:
            continue
          df_group = pd.concat(all_combined_dfs, ignore_index=True)
          tgl_col = (
              "TANGGAL"
              if "TANGGAL" in df_group.columns
              else ("Tanggal" if "Tanggal" in df_group.columns else df_group.columns[0])
          )
          mesin_col = (
              "MESIN"
              if "MESIN" in df_group.columns
              else (
                  "Mesin" if "Mesin" in df_group.columns else df_group.columns[1]
              )
          )
          nama_col = (
              "NAMA"
              if "NAMA" in df_group.columns
              else (
                  "Nama"
                  if "Nama" in df_group.columns
                  else (df_group.columns[2] if len(df_group.columns) > 2 else None)
              )
          )

          df_group[tgl_col] = df_group[tgl_col].ffill()
          df_group[mesin_col] = df_group[mesin_col].ffill()
          if nama_col and nama_col in df_group.columns:
            df_group[nama_col] = df_group[nama_col].ffill()

          df_group["Parsed_Date"] = pd.to_datetime(
              df_group[tgl_col], errors="coerce"
          )
          df_group["Nama_Hari"] = df_group["Parsed_Date"].dt.weekday.map(
              days_map
          )
          df_group["Minggu_Ke"] = df_group["Parsed_Date"].apply(
              get_calendar_week_mapping
          )
          df_group["Mesin_Clean"] = (
              df_group[mesin_col].astype(str).str.strip().str.upper()
          )
          df_group = df_group[
              df_group["Mesin_Clean"].notna()
              & (df_group["Mesin_Clean"] != "")
              & (df_group["Mesin_Clean"] != "NAN")
          ]

          max_orig_col = ws_out.max_column
          start_right_col = max_orig_col + 3
          mesin_in_group = sorted(df_group["Mesin_Clean"].unique())

          col_persen_idx_1based = 12
          for idx, col_name in enumerate(df_group.columns):
            clean_col = str(col_name).upper().replace("\n", " ")
            if "TARGET 100%" in clean_col and "CONES" in clean_col:
              if idx + 1 < len(df_group.columns):
                col_persen_idx_1based = idx + 2
              break

          col_kg_name = None
          for idx, col_name in enumerate(df_group.columns):
            clean_col = str(col_name).upper().replace("\n", " ")
            if (
                "CONES" in clean_col
                and not "TOTAL" in clean_col
                and not "TARGET" in clean_col
            ):
              if idx + 1 < len(df_group.columns):
                col_kg_name = df_group.columns[idx + 1]
                break
          if col_kg_name is None:
            kg_candidates = [
                c for c in df_group.columns if str(c).strip().upper() == "KG"
            ]
            col_kg_name = (
                kg_candidates[0] if kg_candidates else df_group.columns[-1]
            )

          # Highlight Baris
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
                if (
                    mr.min_row <= r_idx <= mr.max_row
                    and mr.min_row != mr.max_row
                ):
                  group_end_row_in_out = mr.max_row + (current_check_row - 3)
                  break
              cell_persen = ws_out.cell(
                  row=actual_row_in_out, column=col_persen_idx_1based
              )
              p_float = parse_percentage_value(cell_persen.value)
              if p_float is not None and p_float < min_threshold_input:
                for target_r in range(
                    actual_row_in_out, group_end_row_in_out + 1
                ):
                  for col_idx in range(1, max_orig_col + 1):
                    row_cell = ws_out.cell(row=target_r, column=col_idx)
                    row_cell.fill = highlight_fill
                    row_cell.font = highlight_font
              r_idx += group_end_row_in_out - actual_row_in_out + 1
            current_check_row += ws_src_sheet.max_row

          # 1. Tabel Rekap Mingguan per Mesin
          last_rekap_end_col = start_right_col
          for m_idx, mesin in enumerate(mesin_in_group):
            df_mesin = df_group[df_group["Mesin_Clean"] == mesin]
            minggu_list = sorted(df_mesin["Minggu_Ke"].unique())
            current_col = start_right_col + (
                m_idx * (len(minggu_list) * 3 + 2)
            )

            ws_out.cell(row=1, column=current_col, value="HARI").alignment = (
                Alignment(horizontal="center", vertical="center")
            )
            ws_out.merge_cells(
                start_row=1,
                start_column=current_col,
                end_row=2,
                end_column=current_col,
            )
            ws_out.cell(row=1, column=current_col).font = header_font
            ws_out.column_dimensions[get_column_letter(current_col)].width = (
                14
            )

            avg_row_idx = 3 + len(days_order)
            for r in range(1, avg_row_idx + 1):
              ws_out.cell(row=r, column=current_col).border = thin_border

            for d_i, day_name in enumerate(days_order):
              c_cell = ws_out.cell(
                  row=3 + d_i, column=current_col, value=day_name
              )
              c_cell.font = bold_font
              c_cell.alignment = Alignment(
                  horizontal="center", vertical="center"
              )

            avg_cell = ws_out.cell(
                row=avg_row_idx, column=current_col, value="RATA-RATA"
            )
            avg_cell.font = bold_font
            avg_cell.alignment = Alignment(horizontal="center", vertical="center")

            col_offset = current_col + 1
            for w in minggu_list:
              df_minggu = df_mesin[df_mesin["Minggu_Ke"] == w]
              header_title = f"MINGGU KE {w} BULAN {bulan_input} - {mesin}"
              ws_out.merge_cells(
                  start_row=1,
                  start_column=col_offset,
                  end_row=1,
                  end_column=col_offset + 2,
              )
              top_head = ws_out.cell(
                  row=1, column=col_offset, value=header_title
              )
              top_head.font = Font(
                  name="Calibri", size=10, bold=True, color="000000"
              )
              top_head.alignment = Alignment(
                  horizontal="center", vertical="center"
              )

              ws_out.column_dimensions[get_column_letter(col_offset)].width = (
                  16
              )
              ws_out.column_dimensions[
                  get_column_letter(col_offset + 1)
              ].width = 18
              ws_out.column_dimensions[
                  get_column_letter(col_offset + 2)
              ].width = 16

              sub_headers = ["TOTAL CONES", "PERSENTASE HASIL", "HASIL (KG)"]
              for s_i, sh in enumerate(sub_headers):
                scell = ws_out.cell(row=2, column=col_offset + s_i, value=sh)
                scell.font = sub_header_font
                scell.alignment = Alignment(
                    horizontal="center", vertical="center", wrap_text=True
                )

              daily_cones_for_avg, daily_persen_for_avg, daily_kg_for_avg = (
                  [],
                  [],
                  [],
              )
              row_data_store = []
              col_cones_name = next(
                  (
                      c
                      for c in df_group.columns
                      if "CONES" in c.upper() and "TOTAL" in c.upper()
                  ),
                  df_group.columns[7],
              )

              for d_i, day_name in enumerate(days_order):
                df_day = df_minggu[df_minggu["Nama_Hari"] == day_name]
                valid_cones_list, valid_persen_list, valid_kg_list = [], [], []

                for _, row_item in df_day.iterrows():
                  p_val_raw = (
                      row_item[col_persen_name]
                      if col_persen_name in df_day.columns
                      else None
                  )
                  p_f = parse_percentage_value(p_val_raw)
                  if p_f is not None and p_f >= min_threshold_input:
                    c_val = (
                        row_item[col_cones_name]
                        if col_cones_name in df_day.columns
                        else 0.0
                    )
                    if pd.notna(c_val):
                      try:
                        valid_cones_list.append(float(c_val))
                      except:
                        pass
                    valid_persen_list.append(p_f)
                    k_val = (
                        row_item[col_kg_name]
                        if col_kg_name in df_day.columns
                        else 0.0
                    )
                    if pd.notna(k_val):
                      try:
                        valid_kg_list.append(
                            float(str(k_val).replace(",", "."))
                        )
                      except:
                        pass

                val_cones = sum(valid_cones_list) if valid_cones_list else 0.0
                val_persen = (
                    np.mean(valid_persen_list) if valid_persen_list else 0.0
                )
                val_kg = sum(valid_kg_list) if valid_kg_list else 0.0
                row_data_store.append((val_cones, val_persen, val_kg))

                if val_cones > 0:
                  daily_cones_for_avg.append(val_cones)
                if val_persen > 0:
                  daily_persen_for_avg.append(val_persen)
                if val_kg > 0:
                  daily_kg_for_avg.append(val_kg)

              for d_i, (val_cones, val_persen, val_kg) in enumerate(
                  row_data_store
              ):
                row_target_idx = 3 + d_i
                c1 = ws_out.cell(
                    row=row_target_idx,
                    column=col_offset,
                    value=round(val_cones, 1) if val_cones > 0 else 0,
                )
                c2 = ws_out.cell(
                    row=row_target_idx,
                    column=col_offset + 1,
                    value="-"
                    if val_persen == 0
                    else round(val_persen / 100.0, 4),
                )
                if val_persen > 0:
                  c2.number_format = "0.00%"
                c3 = ws_out.cell(
                    row=row_target_idx,
                    column=col_offset + 2,
                    value="-" if val_kg == 0 else round(val_kg, 2),
                )
                for c in [c1, c2, c3]:
                  c.alignment = Alignment(
                      horizontal="center", vertical="center"
                  )

              avg_cones_val = (
                  np.mean(daily_cones_for_avg) if daily_cones_for_avg else 0.0
              )
              avg_persen_val = (
                  np.mean(daily_persen_for_avg) if daily_persen_for_avg else 0.0
              )
              avg_kg_val = np.mean(daily_kg_for_avg) if daily_kg_for_avg else 0.0

              ac1 = ws_out.cell(
                  row=avg_row_idx,
                  column=col_offset,
                  value=round(avg_cones_val, 1),
              )
              ac2 = ws_out.cell(
                  row=avg_row_idx,
                  column=col_offset + 1,
                  value="-"
                  if avg_persen_val == 0
                  else round(avg_persen_val / 100.0, 4),
              )
              if avg_persen_val > 0:
                ac2.number_format = "0.00%"
              ac3 = ws_out.cell(
                  row=avg_row_idx,
                  column=col_offset + 2,
                  value="-" if avg_kg_val == 0 else round(avg_kg_val, 2),
              )

              for c in [ac1, ac2, ac3]:
                c.font = bold_font
                c.alignment = Alignment(horizontal="center", vertical="center")

              for r in range(1, avg_row_idx + 1):
                for c_idx in range(col_offset, col_offset + 3):
                  ws_out.cell(row=r, column=c_idx).border = thin_border
              col_offset += 3
              last_rekap_end_col = col_offset

          # 2. Tabel Rekap Persentase Semua Minggu
          summary_start_col = start_right_col
          summary_start_row = avg_row_idx + 4
          ws_out.cell(
              row=summary_start_row,
              column=summary_start_col,
              value=f"REKAP PERSENTASE HASIL PER MINGGU & MESIN (BULAN {bulan_input})",
          ).font = bold_font

          t_header_row2 = summary_start_row + 2
          ws_out.cell(
              row=t_header_row2, column=summary_start_col, value="HARI"
          ).alignment = Alignment(horizontal="center", vertical="center")
          h_cell = ws_out.cell(row=t_header_row2, column=summary_start_col)
          h_cell.font = header_font
          h_cell.border = thin_border

          all_weeks = sorted(df_group["Minggu_Ke"].unique())
          combo_list = []
          for w in all_weeks:
            for mesin in mesin_in_group:
              df_check = df_group[
                  (df_group["Mesin_Clean"] == mesin)
                  & (df_group["Minggu_Ke"] == w)
              ]
              if not df_check.empty:
                combo_list.append((mesin, w))

          col_ptr = summary_start_col + 1
          col_data_dict = {}
          bulan_singkatan = bulan_input[:4]

          for mesin, w in combo_list:
            header_text = f"MINGGU {w} {bulan_singkatan}\n{mesin}"
            cell_head = ws_out.cell(
                row=t_header_row2, column=col_ptr, value=header_text
            )
            cell_head.font = sub_header_font
            cell_head.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )
            cell_head.border = thin_border
            ws_out.column_dimensions[get_column_letter(col_ptr)].width = 18

            df_mesin_item = df_group[df_group["Mesin_Clean"] == mesin]
            df_minggu_item = df_mesin_item[df_mesin_item["Minggu_Ke"] == w]
            day_percentages_for_avg, row_values_to_write = [], []

            for d_i, day_name in enumerate(days_order):
              df_day = df_minggu_item[df_minggu_item["Nama_Hari"] == day_name]
              valid_p_vals = []
              for _, row_item in df_day.iterrows():
                p_val_raw = (
                    row_item[col_persen_name]
                    if col_persen_name in df_day.columns
                    else None
                )
                p_f = parse_percentage_value(p_val_raw)
                if p_f is not None and p_f >= min_threshold_input:
                  valid_p_vals.append(p_f)
              val_persen = np.mean(valid_p_vals) if valid_p_vals else 0.0
              row_values_to_write.append(val_persen)
              if val_persen > 0:
                day_percentages_for_avg.append(val_persen)

            for d_i, val_persen in enumerate(row_values_to_write):
              row_idx = t_header_row2 + 1 + d_i
              c_cell = ws_out.cell(
                  row=row_idx,
                  column=col_ptr,
                  value="-"
                  if val_persen == 0
                  else round(val_persen / 100.0, 4),
              )
              if val_persen > 0:
                c_cell.number_format = "0.00%"
              c_cell.font = Font(name="Calibri", size=11)
              c_cell.alignment = Alignment(
                  horizontal="center", vertical="center"
              )
              c_cell.border = thin_border

            col_data_dict[col_ptr] = day_percentages_for_avg
            col_ptr += 1

          for d_i, day_name in enumerate(days_order):
            row_idx = t_header_row2 + 1 + d_i
            d_cell = ws_out.cell(
                row=row_idx, column=summary_start_col, value=day_name
            )
            d_cell.font = bold_font
            d_cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            d_cell.border = thin_border

          avg_row_summary = t_header_row2 + 1 + len(days_order)
          avg_label_cell = ws_out.cell(
              row=avg_row_summary, column=summary_start_col, value="RATA-RATA"
          )
          avg_label_cell.font = bold_font
          avg_label_cell.alignment = Alignment(
              horizontal="center", vertical="center"
          )
          avg_label_cell.border = thin_border

          col_ptr = summary_start_col + 1
          for mesin, w in combo_list:
            p_list = col_data_dict[col_ptr]
            col_avg = np.mean(p_list) if p_list else 0.0
            ac_cell = ws_out.cell(
                row=avg_row_summary,
                column=col_ptr,
                value="-" if col_avg == 0 else round(col_avg / 100.0, 4),
            )
            if col_avg > 0:
              ac_cell.number_format = "0.00%"
            ac_cell.font = bold_font
            ac_cell.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            ac_cell.border = thin_border
            col_ptr += 1

        # Auto-fit columns untuk sheet hasil akhir
        for sheet in wb_out.worksheets:
          for col in sheet.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
              if type(cell).__name__ == "MergedCell":
                continue
              if cell.value is not None:
                for line in str(cell.value).split("\n"):
                  if len(line) > max_len:
                    max_len = len(line)
            sheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

        summary_io = io.BytesIO()
        wb_out.save(summary_io)
        summary_io.seek(0)

        # Simpan ke session state
        st.session_state.output_rekap_bytes = rekap_io.getvalue()
        st.session_state.output_summary_bytes = summary_io.getvalue()
        st.session_state.processed = True
        st.success("✅ Seluruh data berhasil diproses dengan sukses!")

      except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
  else:
    st.warning("Mohon unggah kedua file Excel (.xlsx) terlebih dahulu pada sidebar.")

# Bagian Tampilan Hasil dan Tombol Unduh (Memanfaatkan Session State)
if st.session_state.processed:
  st.markdown("---")
  st.subheader("📥 Unduh Hasil Berkas Excel")

  col1, col2 = st.columns(2)
  with col1:
    st.download_button(
        label="Download REKAP HASIL PER MESIN.xlsx",
        data=st.session_state.output_rekap_bytes,
        file_name="REKAP HASIL PER MESIN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
  with col2:
    st.download_button(
        label="Download SUMMARY REKAP PER MINGGU.xlsx",
        data=st.session_state.output_summary_bytes,
        file_name="SUMMARY REKAP PER MINGGU.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
