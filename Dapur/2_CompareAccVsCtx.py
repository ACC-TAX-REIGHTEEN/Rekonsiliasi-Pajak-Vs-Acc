import pandas as pd
import glob
import os
import xlsxwriter

def main():
    print("--> Memulai proses...")

    file_acc = 'Acc_temp.xlsx'
    files_export = glob.glob('data_export*.xlsx')

    if not os.path.exists(file_acc):
        print(f"--> File {file_acc} tidak ditemukan.")
        return
    
    if not files_export:
        print("--> File data_export tidak ditemukan.")
        return
    
    file_export_path = files_export[0]
    print(f"--> Membaca file: {file_acc} dan {file_export_path}")

    try:
        df_acc = pd.read_excel(file_acc)
        df_exp = pd.read_excel(file_export_path)
    except Exception as e:
        print(f"--> Gagal membaca file: {e}")
        return

    print("--> Membersihkan format Referensi dan Tanggal...")

    df_acc['No. Referensi'] = df_acc['No. Referensi'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    df_exp['Referensi'] = df_exp['Referensi'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

    df_acc['Tanggal'] = pd.to_datetime(df_acc['Tanggal'], dayfirst=True, errors='coerce')
    df_acc['Tgl. Pajak'] = pd.to_datetime(df_acc['Tgl. Pajak'], dayfirst=True, errors='coerce')
    df_exp['Tanggal Faktur Pajak'] = pd.to_datetime(df_exp['Tanggal Faktur Pajak'], errors='coerce')

    df_acc['Jumlah Pajak'] = pd.to_numeric(df_acc['Jumlah Pajak'], errors='coerce').fillna(0)
    df_exp['PPN'] = pd.to_numeric(df_exp['PPN'], errors='coerce').fillna(0)

    total_pajak_acc = df_acc['Jumlah Pajak'].sum()
    total_ppn_exp = df_exp['PPN'].sum()
    selisih_total = total_pajak_acc - total_ppn_exp

    print("--> Menggabungkan data...")

    df_merge = pd.merge(
        df_acc, 
        df_exp[['Tanggal Faktur Pajak', 'Referensi', 'PPN', 'Status Faktur']], 
        left_on='No. Referensi', 
        right_on='Referensi', 
        how='outer',
        indicator=True
    )

    print("--> Menganalisis selisih...")

    def analisa_baris(row):
        keterangan = []
        
        if row['_merge'] == 'left_only':
            return 'Hanya ada di Acc_temp'
        elif row['_merge'] == 'right_only':
            return 'Hanya ada di Data Export'
        
        val_pajak = row['Jumlah Pajak'] if pd.notnull(row['Jumlah Pajak']) else 0
        val_ppn = row['PPN'] if pd.notnull(row['PPN']) else 0
        
        selisih_nilai = abs(val_pajak - val_ppn)
        if selisih_nilai > 1.0: 
            keterangan.append('Selisih Nominal')
            
        tgl_acc = row['Tgl. Pajak']
        tgl_exp = row['Tanggal Faktur Pajak']
        
        if pd.notnull(tgl_acc) and pd.notnull(tgl_exp):
            if tgl_acc.date() != tgl_exp.date():
                keterangan.append('Beda Tanggal')
        
        if not keterangan:
            return 'Cocok'
        return ', '.join(keterangan)

    df_merge['Keterangan Analisis'] = df_merge.apply(analisa_baris, axis=1)

    df_merge['Tanggal (Str)'] = df_merge['Tanggal'].dt.strftime('%d/%m/%Y').fillna('')
    df_merge['Tgl. Pajak (Str)'] = df_merge['Tgl. Pajak'].dt.strftime('%d/%m/%Y').fillna('')
    df_merge['Tanggal Faktur Pajak (Str)'] = df_merge['Tanggal Faktur Pajak'].dt.strftime('%d/%m/%Y').fillna('')

    cols_final = [
        'No. Referensi', 'Referensi',
        'Jumlah Pajak', 'PPN',
        'Tanggal (Str)', 'Tgl. Pajak (Str)', 'Tanggal Faktur Pajak (Str)',
        'Status Faktur',
        'Keterangan Analisis'
    ]
    
    df_final = df_merge[cols_final].copy()
    
    df_final['Jumlah Pajak'] = df_final['Jumlah Pajak'].fillna(0)
    df_final['PPN'] = df_final['PPN'].fillna(0)
    df_final = df_final.fillna('')
    
    df_final['Sort_Key'] = df_final['Keterangan Analisis'].apply(lambda x: 1 if x == 'Cocok' else 0)
    df_final = df_final.sort_values(by=['Sort_Key', 'No. Referensi'])
    df_final = df_final.drop(columns=['Sort_Key'])

    print("--> Membuat file output dengan format Excel...")

    output_filename = 'Komparasi Acc vs Ctx.xlsx'
    
    df_acc_clean = df_acc.fillna('')
    df_exp_clean = df_exp.fillna('')
    
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    workbook = writer.book

    fmt_header_green = workbook.add_format({'bold': True, 'bg_color': '#C6EFCE', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_header_orange = workbook.add_format({'bold': True, 'bg_color': '#FFEB9C', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    fmt_currency = workbook.add_format({'num_format': '#,##0', 'border': 1})
    fmt_border = workbook.add_format({'border': 1})
    fmt_red_text = workbook.add_format({'font_color': '#9C0006', 'bg_color': '#FFC7CE', 'border': 1})
    
    sheet_name_analisis = 'Hasil Analisis'
    worksheet = workbook.add_worksheet(sheet_name_analisis)

    worksheet.merge_range('A1:B1', 'A. Ringkasan Total', fmt_header_orange)
    worksheet.write('A2', 'Total Jumlah Pajak (Acc_temp)', fmt_border)
    worksheet.write('B2', total_pajak_acc, fmt_currency)
    worksheet.write('A3', 'Total PPN (Data Export)', fmt_border)
    worksheet.write('B3', total_ppn_exp, fmt_currency)
    worksheet.write('A4', 'Selisih Total', fmt_border)
    worksheet.write('B4', selisih_total, fmt_currency)

    start_row = 6
    worksheet.merge_range(f'A{start_row}:I{start_row}', 'B. Detail Data Komparasi', fmt_header_green)
    
    headers = df_final.columns.tolist()
    for col_num, value in enumerate(headers):
        worksheet.write(start_row, col_num, value, fmt_header_green)

    for row_num, row_data in enumerate(df_final.values):
        current_row = start_row + 1 + row_num
        status = row_data[-1] 
        
        for col_num, cell_value in enumerate(row_data):
            cell_fmt = fmt_border
            
            if col_num in [2, 3]: 
                cell_fmt = fmt_currency
            
            if status != 'Cocok':
                if col_num == 8: 
                    cell_fmt = fmt_red_text
            
            worksheet.write(current_row, col_num, cell_value, cell_fmt)

    for i, col in enumerate(df_final.columns):
        series_str = df_final[col].astype(str)
        max_len_data = series_str.map(len).max()
        if pd.isna(max_len_data): max_len_data = 0
        max_len = max(max_len_data, len(str(col))) + 2
        worksheet.set_column(i, i, max_len)

    df_acc_clean.to_excel(writer, sheet_name='Data Asli Acc_temp', index=False)
    worksheet_acc = writer.sheets['Data Asli Acc_temp']
    
    for i, col in enumerate(df_acc_clean.columns):
        series_str = df_acc_clean[col].astype(str)
        max_len_data = series_str.map(len).max()
        if pd.isna(max_len_data): max_len_data = 0
        max_len = max(max_len_data, len(str(col))) + 2
        worksheet_acc.set_column(i, i, max_len)
    
    df_exp_clean.to_excel(writer, sheet_name='Data Asli Export', index=False)
    worksheet_exp = writer.sheets['Data Asli Export']
    
    for i, col in enumerate(df_exp_clean.columns):
        series_str = df_exp_clean[col].astype(str)
        max_len_data = series_str.map(len).max()
        if pd.isna(max_len_data): max_len_data = 0
        max_len = max(max_len_data, len(str(col))) + 2
        worksheet_exp.set_column(i, i, max_len)

    writer.close()
    print(f"--> Selesai. Hasil tersimpan di: {output_filename}")

if __name__ == "__main__":
    main()