# Verbatim AI

A powerful tool for comparing draft content with live website content.

## Features

- Drag-and-drop interface for easy file handling
- Batch processing of multiple DOCX files
- Detailed HTML comparison reports
- Markdown summary reports
- Progress tracking and error handling
- Modern, professional UI

## Documentation

- [HTML Report Guide](docs/html_report_guide.md) - Detailed documentation about the HTML report structure and styling

## Usage

1. Launch Verbatim AI
2. Either:
   - Drag and drop DOCX files or folders onto the application window
   - Click "Start AutoCompare" and select a folder containing DOCX files
3. Enter the corresponding URLs for each DOCX file
4. Wait for the comparison to complete
5. Review the generated reports in the selected folder

## Report Types

### HTML Reports
- Individual HTML reports for each DOCX file
- Side-by-side comparison of draft and live content
- Color-coded differences
- Similarity scores and visual indicators
- Professional, modern design

### Markdown Report
- Single markdown file summarizing all comparisons
- Quick overview of similarity scores
- Error reporting
- Easy to read and share

## Requirements

- Windows 10 or later
- Python 3.11 or later (if running from source)
- Internet connection for live website comparison

## Building from Source

1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Run the build script: `python build.py`
4. Find the executable in the `dist` folder

## License

Copyright © 2024 SMB Team. All rights reserved.