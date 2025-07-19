# PDF Bookmark Splitter

This script splits a PDF file based on its bookmarks, creating separate PDF files and organizing them into directories.

## Features

1. **Top-level bookmark splitting**: Creates separate PDFs for each top-level bookmark with filename prefix "_"
2. **Directory creation**: Creates folders with the same name as top-level bookmarks with prefix "_"
3. **Sub-level bookmark splitting**: Creates separate PDFs for each sub-level bookmark with filename prefix "_<page_number>"
4. **Sub-level directory creation**: Creates directories for each sub-level bookmark within their top-level directories
5. **File organization**: Moves sub-level PDFs to their respective sub-level directories
6. **Page splitting**: Splits each sub-level PDF into individual pages and stores them in "_pages" directories

## Requirements

- Python 3.6+
- PyMuPDF (fitz)
- tkinter (usually included with Python)

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Method 1: GUI File Selection (Recommended)
```bash
python pdf_bookmark_splitter.py
```
- A file dialog will open - select your PDF file
- The script will process the PDF and create directories in the same location as the PDF

### Method 2: Command Line with File Path
```bash
python pdf_bookmark_splitter.py "path/to/your/file.pdf"
```
- The script will process the PDF and create directories in the same location as the PDF

### Method 3: Help Command
```bash
python pdf_bookmark_splitter.py --help
# or
python pdf_bookmark_splitter.py -h
```
- Shows detailed usage information and examples

## Output Structure

The script will create the following structure:

```
├── _TopLevelBookmark1/
│   ├── _TopLevelBookmark1.pdf
│   ├── _SubBookmark1/
│   │   ├── _1_SubBookmark1.pdf
│   │   └── _pages/
│   │       ├── page_001.pdf
│   │       ├── page_002.pdf
│   │       └── ...
│   ├── _SubBookmark2/
│   │   ├── _4_SubBookmark2.pdf
│   │   └── _pages/
│   │       ├── page_001.pdf
│   │       ├── page_002.pdf
│   │       └── ...
│   └── ...
├── _TopLevelBookmark2/
│   ├── _TopLevelBookmark2.pdf
│   ├── _SubBookmark1/
│   │   ├── _43_SubBookmark1.pdf
│   │   └── _pages/
│   │       ├── page_001.pdf
│   │       ├── page_002.pdf
│   │       └── ...
│   └── ...
└── ...
```

## How it works

1. **Bookmark Analysis**: The script reads the PDF's table of contents and separates top-level (level 1) and sub-level (level 2) bookmarks
2. **Top-level Splitting**: Creates separate PDFs for each top-level bookmark, spanning from the bookmark's page to the next top-level bookmark
3. **Directory Creation**: Creates directories named after each top-level bookmark
4. **Sub-level Splitting**: Creates separate PDFs for each sub-level bookmark with page number prefix
5. **File Organization**: Moves sub-level PDFs to the appropriate top-level directory based on page ranges

## Notes

- The script automatically cleans filenames by removing invalid characters
- All created files and directories have the "_" prefix as requested
- Sub-level PDFs include the page number in their filename for easy identification
- The script handles edge cases like the last bookmark extending to the end of the document
- **Important**: All output files and directories are created in the same directory as the input PDF file
- The script includes error handling and user-friendly message boxes for feedback 