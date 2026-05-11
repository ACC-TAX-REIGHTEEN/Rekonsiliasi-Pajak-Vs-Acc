import os
import shutil
import glob
import subprocess
import sys

def clean_dapur_folder(folder_path):
    extensions = ['*.xls', '*.xlsx']
    for ext in extensions:
        files = glob.glob(os.path.join(folder_path, ext))
        for file in files:
            try:
                os.remove(file)
            except OSError:
                pass

root_dir = os.getcwd()
dapur_dir = os.path.join(root_dir, "Dapur")
acc_file = "Acc.xls"
export_pattern = "data_export*.xlsx"
dapur_required_files = ["__init__.py", "1_AccCleaner.py", "2_CompareAccVsCtx.py"]

missing_items = []

if not os.path.exists(os.path.join(root_dir, acc_file)):
    missing_items.append(acc_file)

export_files = glob.glob(os.path.join(root_dir, export_pattern))
if not export_files:
    missing_items.append(export_pattern)

if not os.path.isdir(dapur_dir):
    missing_items.append("Folder: Dapur")
else:
    for f in dapur_required_files:
        if not os.path.exists(os.path.join(dapur_dir, f)):
            missing_items.append(os.path.join("Dapur", f))

if missing_items:
    print("--> Proses Gagal. Item berikut tidak ditemukan:")
    for item in missing_items:
        print(f"--> {item}")
    sys.exit()

try:
    clean_dapur_folder(dapur_dir)

    print(f"--> Menyalin {acc_file} ke Dapur")
    shutil.copy(os.path.join(root_dir, acc_file), dapur_dir)

    target_export_file = export_files[0]
    print(f"--> Menyalin {os.path.basename(target_export_file)} ke Dapur")
    shutil.copy(target_export_file, dapur_dir)

    print("--> Menjalankan 1_AccCleaner.py")
    subprocess.run([sys.executable, "1_AccCleaner.py"], cwd=dapur_dir, check=True)

    print("--> Menjalankan 2_CompareAccVsCtx.py")
    subprocess.run([sys.executable, "2_CompareAccVsCtx.py"], cwd=dapur_dir, check=True)

    result_filename = "Komparasi Acc vs Ctx.xlsx"
    result_path_dapur = os.path.join(dapur_dir, result_filename)
    result_path_root = os.path.join(root_dir, result_filename)

    if os.path.exists(result_path_dapur):
        print(f"--> Memindahkan hasil: {result_filename}")
        if os.path.exists(result_path_root):
            os.remove(result_path_root)
        shutil.move(result_path_dapur, result_path_root)
    else:
        print(f"--> {result_filename} tidak ditemukan di Dapur setelah proses")

except subprocess.CalledProcessError as e:
    print(f"--> Error saat menjalankan script di Dapur: {e}")
except Exception as e:
    print(f"--> Terjadi kesalahan: {e}")

finally:
    print("--> Membersihkan file Excel di Dapur")
    clean_dapur_folder(dapur_dir)
    print("--> Selesai")