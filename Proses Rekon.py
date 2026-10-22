import subprocess
import sys
from pathlib import Path
import shutil


def clean_dapur_folder(dapur_path: Path) -> None:
    for pattern in ("*.xls", "*.xlsx"):
        for file in dapur_path.glob(pattern):
            try:
                file.unlink()
            except OSError:
                pass


def validate_requirements(
    root_dir: Path,
    dapur_dir: Path,
    acc_file: Path,
    export_files: list[Path],
    required_scripts: list[str],
) -> list[str]:
    missing = []

    if not acc_file.exists():
        missing.append(acc_file.name)

    if not export_files:
        missing.append("data*.xlsx")

    if not dapur_dir.is_dir():
        missing.append(f"Folder: {dapur_dir.name}")
    else:
        for script_name in required_scripts:
            script_path = dapur_dir / script_name
            if not script_path.exists():
                missing.append(f"Dapur/{script_name}")

    return missing


def main() -> None:
    root_dir = Path.cwd()
    dapur_dir = root_dir / "Dapur"
    acc_file = root_dir / "Acc.xls"
    export_files = list(root_dir.glob("data*.xlsx"))

    dapur_required_scripts = [
        "__init__.py",
        "0_TransformData.py",
        "1_AccCleaner.py",
        "2_CompareAccVsCtx.py",
    ]
    pipeline_scripts = [
        "0_TransformData.py",
        "1_AccCleaner.py",
        "2_CompareAccVsCtx.py",
    ]
    result_filename = "Komparasi Acc vs Ctx.xlsx"

    missing_items = validate_requirements(
        root_dir,
        dapur_dir,
        acc_file,
        export_files,
        dapur_required_scripts,
    )
    if missing_items:
        print("--> Proses Gagal. Item berikut tidak ditemukan:")
        for item in missing_items:
            print(f"--> {item}")
        sys.exit(1)

    try:
        clean_dapur_folder(dapur_dir)

        print(f"--> Menyalin {acc_file.name} ke Dapur")
        shutil.copy(acc_file, dapur_dir)

        for file in export_files:
            print(f"--> Menyalin {file.name} ke Dapur")
            shutil.copy(file, dapur_dir)

        for script in pipeline_scripts:
            print(f"--> Menjalankan {script}")
            subprocess.run(
                [sys.executable, script],
                cwd=dapur_dir,
                check=True,
            )

        result_dapur = dapur_dir / result_filename
        result_root = root_dir / result_filename

        if result_dapur.exists():
            print(f"--> Memindahkan hasil: {result_filename}")
            if result_root.exists():
                result_root.unlink()
            shutil.move(str(result_dapur), str(result_root))
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


if __name__ == "__main__":
    main()
