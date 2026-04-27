# PyInstaller打包脚本
# 使用方法: python build.py
# Windows平台需要先安装Tesseract OCR引擎

import PyInstaller.__main__
import os
import shutil
import sys

# 打包参数
APP_NAME = "SOC1Extractor"
MAIN_SCRIPT = "soc1_extractor_dify.py"

# Windows下Tesseract OCR默认安装路径
TESSERACT_WIN_PATHS = [
    r"C:\Program Files\Tesseract-OCR",
    r"C:\Program Files (x86)\Tesseract-OCR",
]


def find_tesseract_path():
    """查找Windows下Tesseract OCR安装路径"""
    if sys.platform != 'win32':
        return None

    for path in TESSERACT_WIN_PATHS:
        if os.path.exists(path):
            print(f"检测到Tesseract安装路径: {path}")
            return path

    print("[警告] 未检测到Tesseract OCR安装路径")
    print("请确保已安装Tesseract OCR: https://github.com/UB-Mannheim/tesseract/wiki")
    return None


# 清理之前的构建文件
def clean_build():
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"已清理: {dir_name}")

    # 清理spec文件
    spec_file = f"{APP_NAME}.spec"
    if os.path.exists(spec_file):
        os.remove(spec_file)
        print(f"已清理: {spec_file}")


# 执行打包
def build_exe():
    print("=" * 50)
    print("开始打包 SOC1 Extractor...")
    print(f"平台: {sys.platform}")
    print("=" * 50)

    clean_build()

    # 基础打包参数
    pyinstaller_args = [
        MAIN_SCRIPT,
        '--name=' + APP_NAME,
        '--onefile',              # 打包成单个exe
        '--console',              # 保持命令行窗口（交互式输入）
        '--clean',                # 清理临时文件
        '--noconfirm',            # 不询问确认
        # 包含依赖的数据文件
        '--hidden-import=pdfplumber',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=requests',
        '--hidden-import=PIL',
        '--hidden-import=json',
        '--hidden-import=base64',
        '--hidden-import=uuid',
        '--hidden-import=pytesseract',
    ]

    # Windows平台特殊处理：包含Tesseract OCR
    if sys.platform == 'win32':
        tesseract_path = find_tesseract_path()
        if tesseract_path:
            # 包含Tesseract可执行文件
            tesseract_exe = os.path.join(tesseract_path, "tesseract.exe")
            if os.path.exists(tesseract_exe):
                pyinstaller_args.extend([
                    '--add-binary=' + tesseract_exe + ';tesseract',
                ])

            # 包含Tesseract数据文件（语言包）
            tessdata_path = os.path.join(tesseract_path, "tessdata")
            if os.path.exists(tessdata_path):
                pyinstaller_args.extend([
                    '--add-data=' + tessdata_path + ';tessdata',
                ])

            print(f"已添加Tesseract OCR文件到打包")

    PyInstaller.__main__.run(pyinstaller_args)

    print("\n" + "=" * 50)
    print("打包完成!")
    print("=" * 50)

    # 输出文件位置
    exe_name = f"{APP_NAME}.exe" if sys.platform == 'win32' else APP_NAME
    exe_path = os.path.join('dist', exe_name)
    if os.path.exists(exe_path):
        print(f"\n可执行文件位置: {exe_path}")
        print(f"文件大小: {os.path.getsize(exe_path) / 1024 / 1024:.2f} MB")

    # Windows平台使用提示
    if sys.platform == 'win32':
        print("\n[使用提示]")
        print("运行时需要确保Tesseract OCR已安装在系统上")
        print("或使用打包时包含的Tesseract文件")


if __name__ == "__main__":
    build_exe()