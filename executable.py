import shutil
import os
import PyInstaller.__main__

for folder in ['build', 'dist']:
    if os.path.exists(folder):
        shutil.rmtree(folder)

PyInstaller.__main__.run([
    'app.py',
    '--onefile',
    '--windowed',
    '--name', 'MyTimesheet_v1.0.1',

    '--icon', 'schrodinger_icon.ico',
])