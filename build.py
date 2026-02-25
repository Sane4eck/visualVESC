#build.py
import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    '--onefile',
    '--windowed',
    '--icon=ico/VESC_ico.ico',
    '--clean',
    '--noconfirm',
    '--name=VESC_Cyclogram',
    '--collect-all=pyvesc',
    '--hidden-import=openpyxl',
    '--add-data=ico/VESC_ico.ico;ico',
])
