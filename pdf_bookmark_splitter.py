import fitz
import os
import shutil
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import sys
import argparse

def clean_filename(filename):
    """Clean filename by removing/replacing invalid characters"""
    # Remove or replace invalid characters for filenames
    cleaned = re.sub(r'[<>:"/\\|?*]', '_', filename)
    # Remove leading/trailing spaces and dots
    cleaned = cleaned.strip('. ')
    return cleaned

def get_bookmark_hierarchy(toc):
    """Organize bookmarks into hierarchy with top-level and sub-level bookmarks"""
    top_level_bookmarks = []
    sub_level_bookmarks = []
    
    for entry in toc:
        level, title, page = entry
        if level == 1:
            top_level_bookmarks.append(entry)
        elif level == 2:
            sub_level_bookmarks.append(entry)
    
    return top_level_bookmarks, sub_level_bookmarks

def split_pdf_by_top_level_bookmarks(doc, top_level_bookmarks):
    """Split PDF into separate files based on top-level bookmarks"""
    created_files = []
    
    for i, (level, title, page) in enumerate(top_level_bookmarks):
        page_start = page
        
        # Determine end page
        if i + 1 < len(top_level_bookmarks):
            page_end = top_level_bookmarks[i + 1][2] - 1
        else:
            page_end = doc.page_count
        
        # Ensure page range is valid
        if page_start > page_end:
            print(f"Warning: Invalid page range for '{title}' (start: {page_start}, end: {page_end}), skipping...")
            continue
        
        # Ensure page numbers are within document bounds
        page_start = max(1, min(page_start, doc.page_count))
        page_end = max(1, min(page_end, doc.page_count))
        
        # Create new document for this section
        new_doc = fitz.open()
        
        # Insert pages from start to end
        for page_id in range(page_start, page_end + 1):
            new_doc.insert_pdf(doc, from_page=page_id, to_page=page_id)
        
        # Only save if document has pages
        if new_doc.page_count > 0:
            # Clean filename and save
            clean_title = clean_filename(title)
            filename = f"_{clean_title}.pdf"
            new_doc.save(filename)
            created_files.append((filename, clean_title))
            print(f"Created top-level PDF: {filename} (pages {page_start}-{page_end})")
        else:
            print(f"Warning: No pages found for '{title}', skipping...")
        
        new_doc.close()
    
    return created_files

def split_pdf_by_sub_level_bookmarks(doc, sub_level_bookmarks):
    """Split PDF into separate files based on sub-level bookmarks"""
    created_files = []
    
    for i, (level, title, page) in enumerate(sub_level_bookmarks):
        page_start = page
        
        # Determine end page
        if i + 1 < len(sub_level_bookmarks):
            page_end = sub_level_bookmarks[i + 1][2] - 1
        else:
            page_end = doc.page_count
        
        # Ensure page range is valid
        if page_start > page_end:
            print(f"Warning: Invalid page range for '{title}' (start: {page_start}, end: {page_end}), skipping...")
            continue
        
        # Ensure page numbers are within document bounds
        page_start = max(1, min(page_start, doc.page_count))
        page_end = max(1, min(page_end, doc.page_count))
        
        # Create new document for this section
        new_doc = fitz.open()
        
        # Insert pages from start to end
        for page_id in range(page_start, page_end + 1):
            new_doc.insert_pdf(doc, from_page=page_id, to_page=page_id)
        
        # Only save if document has pages
        if new_doc.page_count > 0:
            # Clean filename and save with page number prefix
            clean_title = clean_filename(title)
            filename = f"_{page}_{clean_title}.pdf"
            new_doc.save(filename)
            created_files.append((filename, clean_title, page))
            print(f"Created sub-level PDF: {filename} (pages {page_start}-{page_end})")
        else:
            print(f"Warning: No pages found for '{title}', skipping...")
        
        new_doc.close()
    
    return created_files

def create_top_level_directories(top_level_files):
    """Create directories for top-level bookmarks"""
    created_dirs = []
    
    for filename, title in top_level_files:
        dir_name = f"_{title}"
        if not os.path.exists(dir_name):
            os.makedirs(dir_name)
            created_dirs.append(dir_name)
            print(f"Created directory: {dir_name}")
    
    return created_dirs

def organize_sub_level_files(sub_level_files, top_level_bookmarks):
    """Move sub-level PDFs to appropriate top-level directories"""
    # Create a mapping of page ranges to top-level bookmarks
    top_level_ranges = []
    
    for i, (level, title, page) in enumerate(top_level_bookmarks):
        if i + 1 < len(top_level_bookmarks):
            end_page = top_level_bookmarks[i + 1][2] - 1
        else:
            end_page = float('inf')  # Use infinity for the last section
        
        top_level_ranges.append((page, end_page, title))
    
    # Move each sub-level file to appropriate directory
    for filename, title, page in sub_level_files:
        # Find which top-level section this sub-level belongs to
        target_dir = None
        for start_page, end_page, top_title in top_level_ranges:
            if start_page <= page <= end_page:
                target_dir = f"_{clean_filename(top_title)}"
                break
        
        if target_dir and os.path.exists(target_dir):
            source_path = filename
            target_path = os.path.join(target_dir, filename)
            
            try:
                shutil.move(source_path, target_path)
                print(f"Moved {filename} to {target_dir}/")
            except Exception as e:
                print(f"Error moving {filename}: {e}")
        else:
            print(f"Could not determine target directory for {filename}")

def move_top_level_files_to_directories(top_level_files):
    """Move top-level PDFs to their respective directories"""
    for filename, title in top_level_files:
        dir_name = f"_{title}"
        if os.path.exists(dir_name):
            source_path = filename
            target_path = os.path.join(dir_name, filename)
            
            try:
                shutil.move(source_path, target_path)
                print(f"Moved {filename} to {dir_name}/")
            except Exception as e:
                print(f"Error moving {filename}: {e}")
        else:
            print(f"Directory {dir_name} not found for {filename}")

def select_pdf_file():
    """Open file dialog to select PDF file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select PDF file to split",
        filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
    )
    
    root.destroy()
    return file_path

def show_message(title, message):
    """Show a message box"""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(title, message)
    root.destroy()

def process_pdf(pdf_path):
    """Process the selected PDF file"""
    # Check if PDF exists
    if not os.path.exists(pdf_path):
        show_message("Error", f"File not found: {pdf_path}")
        return False
    
    print(f"Processing PDF: {pdf_path}")
    
    # Open PDF and get table of contents
    try:
        doc = fitz.open(pdf_path)
        toc = doc.get_toc(simple=True)
    except Exception as e:
        show_message("Error", f"Could not open PDF: {str(e)}")
        return False
    
    if not toc:
        show_message("Warning", "No bookmarks found in the PDF!")
        doc.close()
        return False
    
    print(f"Found {len(toc)} bookmarks")
    
    # Organize bookmarks
    top_level_bookmarks, sub_level_bookmarks = get_bookmark_hierarchy(toc)
    
    print(f"Top-level bookmarks: {len(top_level_bookmarks)}")
    print(f"Sub-level bookmarks: {len(sub_level_bookmarks)}")
    
    # Debug: Print bookmark details
    print("\nTop-level bookmarks:")
    for level, title, page in top_level_bookmarks:
        print(f"  Level {level}: '{title}' at page {page}")
    
    print("\nSub-level bookmarks:")
    for level, title, page in sub_level_bookmarks:
        print(f"  Level {level}: '{title}' at page {page}")
    
    # Change to the directory where the PDF is located
    original_dir = os.getcwd()
    pdf_dir = os.path.dirname(os.path.abspath(pdf_path))
    os.chdir(pdf_dir)
    
    try:
        # Step 1: Split PDF by top-level bookmarks
        print("\n=== Step 1: Splitting by top-level bookmarks ===")
        top_level_files = split_pdf_by_top_level_bookmarks(doc, top_level_bookmarks)
        
        # Step 2: Create directories for top-level bookmarks
        print("\n=== Step 2: Creating top-level directories ===")
        created_dirs = create_top_level_directories(top_level_files)
        
        # Step 3: Split PDF by sub-level bookmarks
        print("\n=== Step 3: Splitting by sub-level bookmarks ===")
        sub_level_files = split_pdf_by_sub_level_bookmarks(doc, sub_level_bookmarks)
        
        # Step 4: Move sub-level files to appropriate directories
        print("\n=== Step 4: Organizing sub-level files ===")
        organize_sub_level_files(sub_level_files, top_level_bookmarks)
        
        # Step 5: Move top-level files to their directories
        print("\n=== Step 5: Moving top-level files to directories ===")
        move_top_level_files_to_directories(top_level_files)
        
        # Close the document
        doc.close()
        
        print("\n=== Process completed successfully! ===")
        print(f"Created {len(top_level_files)} top-level PDFs")
        print(f"Created {len(sub_level_files)} sub-level PDFs")
        print(f"Created {len(created_dirs)} directories")
        
        show_message("Success", f"PDF processing completed!\n\nCreated {len(top_level_files)} top-level PDFs\nCreated {len(sub_level_files)} sub-level PDFs\nCreated {len(created_dirs)} directories\n\nFiles saved in: {pdf_dir}")
        
        return True
        
    except Exception as e:
        doc.close()
        show_message("Error", f"An error occurred during processing: {str(e)}")
        return False
    finally:
        # Return to original directory
        os.chdir(original_dir)

def show_help():
    """Display help information"""
    help_text = """
PDF Bookmark Splitter
====================

This script splits a PDF file based on its bookmarks, creating separate PDF files 
and organizing them into directories.

USAGE:
    python pdf_bookmark_splitter.py                    # Open GUI file selector
    python pdf_bookmark_splitter.py <pdf_file>         # Process specific PDF
    python pdf_bookmark_splitter.py -h, --help         # Show this help

EXAMPLES:
    python pdf_bookmark_splitter.py
        Opens a file dialog to select a PDF file
    
    python pdf_bookmark_splitter.py "C:\\Documents\\book.pdf"
        Processes the specified PDF file
    
    python pdf_bookmark_splitter.py --help
        Shows this help message

FEATURES:
    1. Splits PDF by top-level bookmarks (creates separate PDFs)
    2. Creates directories for each top-level bookmark
    3. Splits PDF by sub-level bookmarks (creates separate PDFs)
    4. Organizes all files into appropriate directories
    5. All output is created in the same directory as the input PDF

OUTPUT STRUCTURE:
    _TopLevelBookmark1/
        _TopLevelBookmark1.pdf
        _1_SubBookmark1.pdf
        _4_SubBookmark2.pdf
        ...
    _TopLevelBookmark2/
        _TopLevelBookmark2.pdf
        _43_SubBookmark1.pdf
        ...

REQUIREMENTS:
    - Python 3.6+
    - PyMuPDF (fitz)
    - tkinter (for GUI)

For more information, see the README.md file.
"""
    print(help_text)

def main():
    """Main function to orchestrate the PDF splitting process"""
    # Set up command line argument parser
    parser = argparse.ArgumentParser(
        description="Split PDF files based on bookmarks",
        add_help=False  # We'll handle help manually
    )
    parser.add_argument('pdf_file', nargs='?', help='Path to the PDF file to process')
    parser.add_argument('-h', '--help', action='store_true', help='Show help message')
    
    # Parse arguments
    args = parser.parse_args()
    
    # Handle help command
    if args.help or len(sys.argv) == 1:
        show_help()
        if len(sys.argv) == 1:
            print("\nNo PDF file specified. Opening file selection dialog...")
            pdf_path = select_pdf_file()
            if pdf_path:
                process_pdf(pdf_path)
            else:
                print("No file selected.")
        return
    
    # Process PDF file
    if args.pdf_file:
        print("PDF Bookmark Splitter")
        print("=" * 50)
        print(f"Processing PDF: {args.pdf_file}")
        process_pdf(args.pdf_file)
    else:
        print("Error: No PDF file specified.")
        print("Use 'python pdf_bookmark_splitter.py --help' for usage information.")

if __name__ == "__main__":
    main() 