# Directory Identity Sanitizer

This Python script takes a set of names, and a directory, then recursively walks through the directory to remove every instance of the name, replacing it with John Smith.

- Renames files and directories if they contain one of the provided names.
- Replaces the contents of ASCII files (such as .txt or .html)
- Replaces the contents of PDF files by using the PyPDF2 library (and pypdftk for uncompression/compression).
- Replaces the contents of DOCX word documents using the python-docx library.
- Replaces the contents of XLSX excel documents using the openpyxl library.
- Replaces the contents of PPTX excel documents using the python-pptx library.
- Extracts ZIP and TAR-based archives into a temporary directory, recusrively iterates through the contents, then re-archives the directory and replaces the original file.

## Usage
For Windows users, double click on identity_eraser.bat, or type `python3 directory_identity_eraser.py` into Command Prompt.
