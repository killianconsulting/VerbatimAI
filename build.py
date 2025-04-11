import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Define the main script path
main_script = os.path.join(current_dir, 'main.py')

# Define the output directory
output_dir = os.path.join(current_dir, 'dist')

# Define icon path
icon_path = os.path.join(current_dir, 'verbatim.ico')

# Run PyInstaller
PyInstaller.__main__.run([
    '--name=VerbatimAI',
    '--onefile',
    '--windowed',
    '--add-data=requirements.txt;.',
    '--add-data=smbteam-logo.png;.',
    '--add-data=verbatim.ico;.',
    f'--icon={icon_path}',
    main_script
]) 