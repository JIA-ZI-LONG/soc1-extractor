# PyInstaller打包脚本
# 使用方法: python build.py

import PyInstaller.__main__
import os
import shutil

# 打包参数
APP_NAME = "SOC1Extractor"
MAIN_SCRIPT = "soc1_extractor_dify.py"

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
    print("=" * 50)

    clean_build()

    PyInstaller.__main__.run([
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
    ])

    print("\n" + "=" * 50)
    print("打包完成!")
    print("=" * 50)

    # 输出文件位置
    exe_path = os.path.join('dist', f"{APP_NAME}.exe" if os.name == 'nt' else APP_NAME)
    if os.path.exists(exe_path):
        print(f"\n可执行文件位置: {exe_path}")
        print(f"文件大小: {os.path.getsize(exe_path) / 1024 / 1024:.2f} MB")

if __name__ == "__main__":
    build_exe()