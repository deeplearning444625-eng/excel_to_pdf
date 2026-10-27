# excel2pdf

Converts Excel (.xlsx/.xlsm) to PDF, 100% local (no internet, no upload).
Rule: "Khmer OS Moul Light" -> "Times New Roman" before conversion; Khmer letters
get a real Khmer font so they stay readable.

## For users
1. Download `excel2pdf.exe` from the Releases page.
2. Use it one of three ways:
   - Double-click -> choose Excel files -> PDFs are saved next to them
   - Drag & drop Excel files (or a folder) onto the exe
   - Command line: `excel2pdf.exe report.xlsx -o report.pdf --khmer-font "Khmer OS System"`

Requires Windows + Microsoft Excel (best Khmer rendering) or LibreOffice, and a Khmer
font installed (Khmer OS System, Khmer UI, ...).
Windows SmartScreen may warn because the exe is not code-signed: More info -> Run anyway.
Verify the download with the .sha256 file if needed.

## Speed
- One hidden Excel instance is reused for all files (batch is much faster)
- Files that do not use the Khmer OS Moul Light font are converted directly, untouched

## For developers
    pip install -r requirements.txt pyinstaller
    pyinstaller --onefile --name excel2pdf --hidden-import win32com.client excel_to_pdf.py

## Release
    git tag v1.1.0
    git push origin v1.1.0
GitHub Actions builds `excel2pdf.exe` and attaches it to the release.
