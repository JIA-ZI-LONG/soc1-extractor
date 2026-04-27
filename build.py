# PyInstaller Build Script for Windows
# Usage: python build.py

import PyInstaller.__main__
import os
import shutil
import sys

APP_NAME = "SOC1Extractor"
MAIN_SCRIPT = "soc1_extractor_dify.py"

# Set UTF-8 encoding for Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')


def clean_build():
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"Cleaned: {dir_name}")

    spec_file = f"{APP_NAME}.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"Cleaned: {spec_file}")


def build_exe():
    print("=" * 60)
    print("SOC1 Extractor Build Script")
    print("=" * 60)
    print(f"Platform: {sys.platform}")

    clean_build()

    pyinstaller_args = [
        MAIN_SCRIPT,
        '--name=' + APP_NAME,
        '--onefile',
        '--console',
        '--clean',
        '--noconfirm',
        '--hidden-import=pdfplumber',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=requests',
        '--hidden-import=PIL',
        '--hidden-import=json',
        '--hidden-import=base64',
        '--hidden-import=uuid',
    ]

    print("\nStarting PyInstaller build...")
    print("=" * 60)

    PyInstaller.__main__.run(pyinstaller_args)

    print("\n" + "=" * 60)
    print("Build Complete!")
    print("=" * 60)

    exe_name = f"{APP_NAME}.exe" if sys.platform == 'win32' else APP_NAME
    exe_path = os.path.join('dist', exe_name)
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / 1024 / 1024
        print(f"\nExecutable: {exe_path}")
        print(f"Size: {size_mb:.2f} MB")


if __name__ == "__main__":
    build_exe()