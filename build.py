import PyInstaller.__main__
import os

# 切换到当前脚本所在目录
os.chdir(os.path.dirname(os.path.abspath(__file__)))

PyInstaller.__main__.run([
    'main.py',
    '--onefile',
    '--noconsole',
    '--name=Word工具集',
    '--hidden-import=openpyxl',
    '--collect-all=docx',
    '--clean',
])
