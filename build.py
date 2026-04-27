# PyInstaller打包脚本 - Windows开箱即用版
# 使用方法: python build.py
#
# 打包时会自动包含Tesseract OCR，无需用户额外安装
# Windows下需要先安装Tesseract OCR用于打包（仅打包时需要）

import PyInstaller.__main__
import os
import shutil
import sys
import glob

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
    print("请从以下地址下载安装（用于打包）:")
    print("https://github.com/UB-Mannheim/tesseract/wiki")
    print("\n打包后的exe将包含Tesseract，用户无需安装")
    return None


def get_tesseract_files(tesseract_path):
    """
    获取Tesseract OCR所有需要打包的文件
    包括：可执行文件、DLL依赖、语言数据包
    """
    files_to_include = []

    # 1. Tesseract可执行文件
    tesseract_exe = os.path.join(tesseract_path, "tesseract.exe")
    if os.path.exists(tesseract_exe):
        files_to_include.append(('binary', tesseract_exe, 'tesseract'))

    # 2. 核心DLL文件（libtesseract, liblept等）
    dll_patterns = ['*.dll']
    for pattern in dll_patterns:
        dll_files = glob.glob(os.path.join(tesseract_path, pattern))
        for dll in dll_files:
            dll_name = os.path.basename(dll)
            files_to_include.append(('binary', dll, 'tesseract'))

    # 3. tessdata目录（语言包）
    tessdata_path = os.path.join(tesseract_path, "tessdata")
    if os.path.exists(tessdata_path):
        files_to_include.append(('data', tessdata_path, 'tessdata'))

    return files_to_include


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
    print("=" * 60)
    print("SOC1 Extractor 打包脚本")
    print("=" * 60)
    print(f"平台: {sys.platform}")
    print(f"目标: 生成开箱即用的可执行文件")

    clean_build()

    # 基础打包参数
    pyinstaller_args = [
        MAIN_SCRIPT,
        '--name=' + APP_NAME,
        '--onefile',              # 打包成单个exe
        '--console',              # 保持命令行窗口
        '--clean',
        '--noconfirm',
        # Python依赖
        '--hidden-import=pdfplumber',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=requests',
        '--hidden-import=PIL',
        '--hidden-import=json',
        '--hidden-import=base64',
        '--hidden-import=uuid',
        '--hidden-import=pytesseract',
        '--hidden-import=sys',
        '--hidden-import=os',
    ]

    # Windows平台：包含完整的Tesseract OCR
    if sys.platform == 'win32':
        tesseract_path = find_tesseract_path()
        if tesseract_path:
            tesseract_files = get_tesseract_files(tesseract_path)

            print(f"\n包含Tesseract OCR文件:")
            for file_type, src_path, dest_dir in tesseract_files:
                if file_type == 'binary':
                    pyinstaller_args.append(f'--add-binary={src_path};{dest_dir}')
                    print(f"  [DLL/EXE] {os.path.basename(src_path)}")
                elif file_type == 'data':
                    pyinstaller_args.append(f'--add-data={src_path};{dest_dir}')
                    # 计算语言包数量
                    lang_files = glob.glob(os.path.join(src_path, '*.traineddata'))
                    print(f"  [tessdata] {len(lang_files)} 个语言包")

            print(f"\n总计包含 {len(tesseract_files)} 个Tesseract相关项")

    print("\n开始PyInstaller打包...")
    print("=" * 60)

    PyInstaller.__main__.run(pyinstaller_args)

    print("\n" + "=" * 60)
    print("打包完成!")
    print("=" * 60)

    # 输出文件信息
    exe_name = f"{APP_NAME}.exe" if sys.platform == 'win32' else APP_NAME
    exe_path = os.path.join('dist', exe_name)
    if os.path.exists(exe_path):
        size_mb = os.path.getsize(exe_path) / 1024 / 1024
        print(f"\n可执行文件: {exe_path}")
        print(f"文件大小: {size_mb:.2f} MB")

        if sys.platform == 'win32':
            print("\n[特性] 开箱即用 - 无需额外安装Tesseract OCR")
            print("[使用] 双击运行，输入PDF文件夹路径即可")


if __name__ == "__main__":
    build_exe()