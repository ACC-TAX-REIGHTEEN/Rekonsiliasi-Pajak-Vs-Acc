import configparser
import glob
import os
import warnings
import pandas as pd


def read_excel_adaptive(file_path, preferred_engine=None, **kwargs):
    if preferred_engine:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return pd.read_excel(file_path, engine=preferred_engine, **kwargs), preferred_engine

    try:
        df = pd.read_excel(file_path, engine="calamine", **kwargs)
        return df, "calamine"
    except ImportError:
        pass
    except Exception as e:
        print(f"--> Calamine gagal ({e}), beralih ke openpyxl...")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        df = pd.read_excel(file_path, engine="openpyxl", **kwargs)
        return df, "openpyxl"


def safe_fillna_empty(df):
    return df.astype(object).where(pd.notna(df), "")


def auto_adjust_columns(worksheet, df):
    for i, col in enumerate(df.columns):
        max_len_data = df[col].astype(str).map(len).max()
        if pd.isna(max_len_data):
            max_len_data = 0
        worksheet.set_column(i, i, max(max_len_data, len(str(col))) + 2)


def main():
    print("--> Memulai proses...")

    config_file = "config.conf"
    calc_gap = 0.0

    if os.path.exists(config_file):
        try:
            config = configparser.ConfigParser()
            config.read(config_file)

            if "CALC" in config and "calc_gap" in config["CALC"]:
                calc_gap = float(config["CALC"]["calc_gap"])
                print(f"--> Konfigurasi ditemukan: Toleransi selisih (calc_gap) = {calc_gap}")
            else:
                print("--> Seksi [CALC] atau kunci 'calc_gap' tidak ditemukan. Menggunakan default gap = 0.0")
        except Exception as e:
            print(f"--> Gagal membaca {config_file}: {e}. Menggunakan default gap = 0.0")
    else:
        print(f"--> File {config_file} tidak ditemukan. Menggunakan default gap = 0.0")

    file_acc = "Acc_temp.xlsx"
    files_export = glob.glob("data_export*.xlsx")

    if not os.path.exists(file_acc):
        print(f"--> File {file_acc} tidak ditemukan.")
        return

    if not files_export:
        print("--> File data_export tidak ditemukan.")
        return

    file_export_path = files_export[0]
    print(f"--> Membaca file: {file_acc} dan {file_export_path}")

    try:
        df_acc, active_engine = read_excel_adaptive(file_acc)
        print(f"--> Engine aktif: {active_engine}")
        df_exp, _ = read_excel_adaptive(file_export_path, preferred_engine=active_engine)
    except Exception as e:
        print(f"--> Gagal membaca file: {e}")
        return

    print("--> Membersihkan format Referensi dan Tanggal...")

    df_acc["No. Referensi"] = df_acc["No. Referensi"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    df_exp["Referensi"] = df_exp["Referensi"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)

    df_acc["Tanggal"] = pd.to_datetime(df_acc["Tanggal"], dayfirst=True, errors="coerce")
    df_acc["Tgl. Pajak"] = pd.to_datetime(df_acc["Tgl. Pajak"], dayfirst=True, errors="coerce")
    df_exp["Tanggal Faktur Pajak"] = pd.to_datetime(df_exp["Tanggal Faktur Pajak"], errors="coerce")

    df_acc["Jumlah Pajak"] = pd.to_numeric(df_acc["Jumlah Pajak"], errors="coerce").fillna(0)
    df_exp["PPN"] = pd.to_numeric(df_exp["PPN"], errors="coerce").fillna(0)

    total_pajak_acc = df_acc["Jumlah Pajak"].sum()
    total_ppn_exp = df_exp["PPN"].sum()
    selisih_total = total_pajak_acc - total_ppn_exp

    print("--> Menggabungkan data...")

    df_merge = pd.merge(
        df_acc,
        df_exp[["Tanggal Faktur Pajak", "Referensi", "PPN", "Status Faktur"]],
        left_on="No. Referensi",
        right_on="Referensi",
        how="outer",
        indicator=True
    )

    print("--> Menganalisis selisih...")

    def analisa_baris(row):
        if row["_merge"] == "left_only":
            return "Hanya ada di Acc_temp"
        if row["_merge"] == "right_only":
            return "Hanya ada di Data Export"

        keterangan = []
        val_pajak = row["Jumlah Pajak"] if pd.notnull(row["Jumlah Pajak"]) else 0
        val_ppn = row["PPN"] if pd.notnull(row["PPN"]) else 0

        if abs(val_pajak - val_ppn) > calc_gap:
            keterangan.append("Selisih Nominal")

        tgl_acc = row["Tgl. Pajak"]
        tgl_exp = row["Tanggal Faktur Pajak"]

        if pd.notnull(tgl_acc) and pd.notnull(tgl_exp):
            if tgl_acc.date() != tgl_exp.date():
                keterangan.append("Beda Tanggal")

        return ", ".join(keterangan) if keterangan else "Cocok"

    df_merge["Keterangan Analisis"] = df_merge.apply(analisa_baris, axis=1)

    df_merge["Tanggal (Str)"] = df_merge["Tanggal"].dt.strftime("%d/%m/%Y").fillna("")
    df_merge["Tgl. Pajak (Str)"] = df_merge["Tgl. Pajak"].dt.strftime("%d/%m/%Y").fillna("")
    df_merge["Tanggal Faktur Pajak (Str)"] = df_merge["Tanggal Faktur Pajak"].dt.strftime("%d/%m/%Y").fillna("")

    cols_final = [
        "No. Referensi", "Referensi",
        "Jumlah Pajak", "PPN",
        "Tanggal (Str)", "Tgl. Pajak (Str)", "Tanggal Faktur Pajak (Str)",
        "Status Faktur",
        "Keterangan Analisis"
    ]

    df_final = df_merge[cols_final].copy()

    numeric_cols = ["Jumlah Pajak", "PPN"]
    for col in numeric_cols:
        df_final[col] = df_final[col].fillna(0)

    for col in df_final.columns:
        if col not in numeric_cols:
            df_final[col] = df_final[col].astype(object).where(pd.notna(df_final[col]), "")

    df_final["Sort_Key"] = df_final["Keterangan Analisis"].apply(lambda x: 1 if x == "Cocok" else 0)
    df_final = df_final.sort_values(by=["Sort_Key", "No. Referensi"]).drop(columns=["Sort_Key"])

    print("--> Membuat file output dengan format Excel...")

    output_filename = "Komparasi Acc vs Ctx.xlsx"
    df_acc_clean = safe_fillna_empty(df_acc)
    df_exp_clean = safe_fillna_empty(df_exp)

    with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_header_green = workbook.add_format({"bold": True, "bg_color": "#C6EFCE", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_header_orange = workbook.add_format({"bold": True, "bg_color": "#FFEB9C", "border": 1, "align": "center", "valign": "vcenter"})
        fmt_currency = workbook.add_format({"num_format": "#,##0", "border": 1})
        fmt_border = workbook.add_format({"border": 1})
        fmt_red_text = workbook.add_format({"font_color": "#9C0006", "bg_color": "#FFC7CE", "border": 1})

        worksheet = workbook.add_worksheet("Hasil Analisis")

        worksheet.merge_range("A1:B1", "A. Ringkasan Total", fmt_header_orange)
        worksheet.write("A2", "Total Jumlah Pajak (Acc_temp)", fmt_border)
        worksheet.write("B2", total_pajak_acc, fmt_currency)
        worksheet.write("A3", "Total PPN (Data Export)", fmt_border)
        worksheet.write("B3", total_ppn_exp, fmt_currency)
        worksheet.write("A4", "Selisih Total", fmt_border)
        worksheet.write("B4", selisih_total, fmt_currency)

        start_row = 6
        worksheet.merge_range(f"A{start_row}:I{start_row}", "B. Detail Data Komparasi", fmt_header_green)

        for col_num, value in enumerate(df_final.columns):
            worksheet.write(start_row, col_num, value, fmt_header_green)

        for row_num, row_data in enumerate(df_final.values):
            current_row = start_row + 1 + row_num
            status = row_data[-1]

            for col_num, cell_value in enumerate(row_data):
                cell_fmt = fmt_border
                if col_num in (2, 3):
                    cell_fmt = fmt_currency
                if status != "Cocok" and col_num == 8:
                    cell_fmt = fmt_red_text
                worksheet.write(current_row, col_num, cell_value, cell_fmt)

        auto_adjust_columns(worksheet, df_final)

        df_acc_clean.to_excel(writer, sheet_name="Data Asli Acc_temp", index=False)
        auto_adjust_columns(writer.sheets["Data Asli Acc_temp"], df_acc_clean)

        df_exp_clean.to_excel(writer, sheet_name="Data Asli Export", index=False)
        auto_adjust_columns(writer.sheets["Data Asli Export"], df_exp_clean)

    print(f"--> Selesai. Hasil tersimpan di: {output_filename}")


if __name__ == "__main__":
    main()