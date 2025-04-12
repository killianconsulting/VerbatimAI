import PyInstaller.__main__
import os
import sys
import subprocess
import shutil

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Define paths
main_script = os.path.join(current_dir, 'main.py')
output_dir = os.path.join(current_dir, 'dist')
icon_path = os.path.join(current_dir, 'verbatim.ico')
logo_path = os.path.join(current_dir, 'smbteam-logo.png')

# Install tkinterdnd2 if not already installed
try:
    import tkinterdnd2
except ImportError:
    subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'tkinterdnd2'])

# Get the site-packages directory
site_packages = None
for path in sys.path:
    if 'site-packages' in path:
        site_packages = path
        break

if site_packages:
    # Copy tkinterdnd2 to the project directory
    tkdnd_src = os.path.join(site_packages, 'tkinterdnd2')
    tkdnd_dst = os.path.join(current_dir, 'tkinterdnd2')
    
    if os.path.exists(tkdnd_dst):
        shutil.rmtree(tkdnd_dst)
    
    if os.path.exists(tkdnd_src):
        shutil.copytree(tkdnd_src, tkdnd_dst)
        
        # Copy tkdnd DLL to the project directory
        tkdnd_dll_src = os.path.join(tkdnd_src, 'tkdnd', 'win64', 'libtkdnd2.9.2.dll')
        tkdnd_dll_dst = os.path.join(current_dir, 'libtkdnd2.9.2.dll')
        if os.path.exists(tkdnd_dll_src):
            shutil.copy2(tkdnd_dll_src, tkdnd_dll_dst)

# Run PyInstaller
PyInstaller.__main__.run([
    main_script,
    '--name=VerbatimAI',
    '--onefile',
    '--windowed',
    f'--icon={icon_path}',
    '--add-data=verbatim.ico;.',
    '--add-data=smbteam-logo.png;.',
    '--add-data=tkinterdnd2;tkinterdnd2',
    '--add-data=libtkdnd2.9.2.dll;.',
    '--hidden-import=tkinterdnd2',
    f'--distpath={output_dir}'
]) 