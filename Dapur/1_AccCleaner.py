import pandas as pd
import io

def clean_id_number(value):
    if pd.isna(value):
        return ""
    
    if isinstance(value, (int, float)):
        return str(int(value))
        
    val_str = str(value).strip()
    if val_str.endswith(",00"):
        val_str = val_str[:-3]
    if val_str.endswith(".00"):
        val_str = val_str[:-3]
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
        
    val_str = val_str.replace(".", "")
    val_str = val_str.replace(",", "")
    return val_str

def clean_currency(value):
    if pd.isna(value):
        return 0
    
    if isinstance(value, (int, float)):
        return float(value)
        
    val_str = str(value).strip()
    val_str = val_str.replace(".", "")
    val_str = val_str.replace(",", ".")
    
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
        "Sep": "Sep", "Okt": "Oct", "Nov": "Nop", "Des": "Dec"
    }
    
    for indo, eng in month_map.items():
        if indo in val_str:
            val_str = val_str.replace(indo, eng)
            break
            
    try:
        return pd.to_datetime(val_str, format="%d %b %Y")
    except:
        return pd.to_datetime(val_str, errors='coerce')

try:
    file_path = 'Acc.xls'
    print(f"--> Membaca file {file_path}")
    
    df_raw = pd.read_excel(file_path, header=None)

    target_columns = [
        "Tanggal", "Tgl. Pajak", "No. Referensi", "No. Faktur Pajak", 
        "Nama Pelanggan", "Negara Pelanggan", "Jumlah Pajak", "Nomor Pajak Pelanggan"
    ]

    header_row_index = -1
    for idx, row in df_raw.iterrows():
        row_values = [str(x).strip() for x in row.values if pd.notna(x)]
        matches = sum(1 for col in target_columns if col in row_values)
        if matches >= 4:
            header_row_index = idx
            break

    if header_row_index == -1:
        print("--> Header kolom tidak ditemukan.")
        exit()

    print(f"--> Header ditemukan pada baris {header_row_index + 1}")

    df = pd.read_excel(file_path, header=header_row_index)
    
    df = df[target_columns].copy()

    df = df.dropna(subset=["No. Referensi", "Nama Pelanggan"], how='all')

    print("--> Memproses data dan format kolom")

    df["No. Referensi"] = df["No. Referensi"].apply(clean_id_number)
    df["No. Faktur Pajak"] = df["No. Faktur Pajak"].apply(clean_id_number)
    df["Nomor Pajak Pelanggan"] = df["Nomor Pajak Pelanggan"].apply(clean_id_number)

    df["Jumlah Pajak"] = df["Jumlah Pajak"].apply(clean_currency)

    df["Tanggal"] = df["Tanggal"].apply(clean_date_indo)
    df["Tgl. Pajak"] = df["Tgl. Pajak"].apply(clean_date_indo)

    output_file = 'Acc_temp.xlsx'
    print(f"--> Menyimpan data ke {output_file}")

    writer = pd.ExcelWriter(output_file, engine='xlsxwriter', datetime_format='dd/mm/yyyy', date_format='dd/mm/yyyy')
    df.to_excel(writer, index=False, sheet_name='Data')

    workbook = writer.book
    worksheet = writer.sheets['Data']

    currency_format = workbook.add_format({'num_format': '#,##0'})
    date_format_custom = workbook.add_format({'num_format': 'dd/mm/yyyy'})
    text_format = workbook.add_format({'num_format': '@'})

    col_idx_jumlah = df.columns.get_loc("Jumlah Pajak")
    worksheet.set_column(col_idx_jumlah, col_idx_jumlah, 20, currency_format)

    col_idx_tanggal = df.columns.get_loc("Tanggal")
    col_idx_tgl_pajak = df.columns.get_loc("Tgl. Pajak")
    
    worksheet.set_column(col_idx_tanggal, col_idx_tanggal, 15, date_format_custom)
    worksheet.set_column(col_idx_tgl_pajak, col_idx_tgl_pajak, 15, date_format_custom)

    col_idx_nopajak = df.columns.get_loc("Nomor Pajak Pelanggan")
    col_idx_faktur = df.columns.get_loc("No. Faktur Pajak")
    col_idx_ref = df.columns.get_loc("No. Referensi")

    worksheet.set_column(col_idx_nopajak, col_idx_nopajak, 25, text_format)
    worksheet.set_column(col_idx_faktur, col_idx_faktur, 25, text_format)
    worksheet.set_column(col_idx_ref, col_idx_ref, 20, text_format)

    writer.close()
    print("--> Selesai. File berhasil dibuat.")

except Exception as e:
    print(f"--> Terjadi kesalahan: {e}")