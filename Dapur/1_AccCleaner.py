import sys
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
        print(f"--> Calamine gagal ({e}), beralih ke fallback engine...")

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            df = pd.read_excel(file_path, engine="openpyxl", **kwargs)
            return df, "openpyxl"
        except Exception:
            df = pd.read_excel(file_path, engine="xlrd", **kwargs)
            return df, "xlrd"


def find_header_row(df_raw, target_columns):
    for idx, row in df_raw.iterrows():
        row_values = [str(x).strip() for x in row.values if pd.notna(x)]
        matches = sum(1 for col in target_columns if col in row_values)
        if matches >= 4:
            return idx
    return -1


def clean_id_number(value):
    if pd.isna(value):
        return ""

    if isinstance(value, (int, float)):
        return str(int(value))

    val_str = str(value).strip()
    if val_str.endswith(",00") or val_str.endswith(".00"):
        val_str = val_str[:-3]
    elif val_str.endswith(".0"):
        val_str = val_str[:-2]

    return val_str.replace(".", "").replace(",", "")


def clean_currency(value):
    if pd.isna(value):
        return 0

    if isinstance(value, (int, float)):
        return float(value)

    val_str = str(value).strip().replace(".", "").replace(",", ".")

    try:
        return float(val_str)
    except ValueError:
        return 0


def clean_date_indo(value):
    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value

    val_str = str(value).strip()
    month_map = {
        "Jan": "Jan", "Feb": "Feb", "Mar": "Mar", "Apr": "Apr",
        "Mei": "May", "Jun": "Jun", "Jul": "Jul", "Agu": "Aug",
        "Sep": "Sep", "Okt": "Oct", "Nop": "Nov", "Nov": "Nov", "Des": "Dec"
    }

    for indo, eng in month_map.items():
        if indo in val_str:
            val_str = val_str.replace(indo, eng)
            break

    try:
        return pd.to_datetime(val_str, format="%d %b %Y")
    except Exception:
        return pd.to_datetime(val_str, errors="coerce")


def main():
    try:
        file_path = "Acc.xls"
        print(f"--> Membaca file {file_path}")

        target_columns = [
            "Tanggal", "Tgl. Pajak", "No. Referensi", "No. Faktur Pajak",
            "Nama Pelanggan", "Negara Pelanggan", "Jumlah Pajak", "Nomor Pajak Pelanggan"
        ]

        df_preview, active_engine = read_excel_adaptive(file_path, header=None, nrows=50)
        print(f"--> Engine aktif: {active_engine}")

        header_row_index = find_header_row(df_preview, target_columns)

        if header_row_index == -1 and len(df_preview) == 50:
            df_raw, _ = read_excel_adaptive(file_path, preferred_engine=active_engine, header=None)
            header_row_index = find_header_row(df_raw, target_columns)

        if header_row_index == -1:
            print("--> Header kolom tidak ditemukan.")
            sys.exit(1)

        print(f"--> Header ditemukan pada baris {header_row_index + 1}")

        df, _ = read_excel_adaptive(
            file_path,
            preferred_engine=active_engine,
            header=header_row_index
        )

        df = df[target_columns].copy()
        df = df.dropna(subset=["No. Referensi", "Nama Pelanggan"], how="all")

        print("--> Memproses data dan format kolom")

        for col in ["No. Referensi", "No. Faktur Pajak", "Nomor Pajak Pelanggan"]:
            df[col] = df[col].apply(clean_id_number)

        df["Jumlah Pajak"] = df["Jumlah Pajak"].apply(clean_currency)

        for col in ["Tanggal", "Tgl. Pajak"]:
            df[col] = df[col].apply(clean_date_indo)

        output_file = "Acc_temp.xlsx"
        print(f"--> Menyimpan data ke {output_file}")

        with pd.ExcelWriter(
            output_file,
            engine="xlsxwriter",
            datetime_format="dd/mm/yyyy",
            date_format="dd/mm/yyyy"
        ) as writer:
            df.to_excel(writer, index=False, sheet_name="Data")

            workbook = writer.book
            worksheet = writer.sheets["Data"]

            currency_format = workbook.add_format({"num_format": "#,##0"})
            date_format_custom = workbook.add_format({"num_format": "dd/mm/yyyy"})
            text_format = workbook.add_format({"num_format": "@"})

            col_formats = [
                ("Jumlah Pajak", 20, currency_format),
                ("Tanggal", 15, date_format_custom),
                ("Tgl. Pajak", 15, date_format_custom),
                ("Nomor Pajak Pelanggan", 25, text_format),
                ("No. Faktur Pajak", 25, text_format),
                ("No. Referensi", 20, text_format),
            ]

            for col_name, width, cell_format in col_formats:
                col_idx = df.columns.get_loc(col_name)
                worksheet.set_column(col_idx, col_idx, width, cell_format)

        print("--> Selesai. File berhasil dibuat.")

    except Exception as e:
        print(f"--> Terjadi kesalahan: {e}")


if __name__ == "__main__":
    main()