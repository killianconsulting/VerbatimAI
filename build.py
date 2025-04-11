import PyInstaller.__main__
import os

# Get the current directory
current_dir = os.path.dirname(os.path.abspath(__file__))

# Define the main script path
main_script = os.path.join(current_dir, 'main.py')

# Define the output directory
output_dir = os.path.join(current_dir, 'dist')

# Run PyInstaller
PyInstaller.__main__.run([
    '--name=DocumentWebpageComparer',
    '--onefile',
    '--windowed',
    '--add-data=requirements.txt;.',
    '--icon=NONE',
    main_script
]) 