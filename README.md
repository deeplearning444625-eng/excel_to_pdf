# excel2pdf  (v1.3.0)

Converts Excel (.xlsx/.xlsm) to PDF. 100% local: no internet, nothing is uploaded.
Rule: "Khmer OS Moul Light" -> "Times New Roman" before conversion. Khmer letters in those
cells get a real Khmer font so they stay readable. The original Excel file is never changed.

## For users
1. Download `excel2pdf.exe` from the Releases page and double-click it.
   Your browser opens the app (address `http://127.0.0.1:8765`). It runs on your PC only.
2. Add Excel files: **Add files...**, **Add folder...**, drag them onto the page, or type a
   file/folder path on your PC. Then click **Convert to PDF**.
   Each file shows Waiting / Converting / Done / Failed, with a plain-language reason.
   Files that are skipped (wrong type, missing, `~$` lock files, errors) are listed in a message box.
3. Click **PDF** on a row, **Download all (ZIP)**, or **Open output folder**.
   PDFs are saved to the folder shown in Options (default `Documents\excel2pdf`).

Closing the browser tab closes the app a few seconds later (or click **Quit app**).
You can also drag Excel files onto the exe icon: they are added to the list.
Starting the exe again while it is running just opens the same page.
Needs Windows with Microsoft Excel (best Khmer rendering) or LibreOffice, and a Khmer font
(Khmer OS System, Khmer UI, ...). If none is found the page shows a warning.

Security: the server listens on 127.0.0.1 only and rejects requests from other websites.
If the browser does not open, start with `excel2pdf.exe --no-browser` and open the address yourself.
A crash log is written to `%APPDATA%\excel2pdf\error.log`.

SmartScreen may warn because the exe is not code-signed: More info -> Run anyway.
`excel2pdf-fast-start.zip` is the same program as a folder: it opens faster (unzip, run excel2pdf-fast.exe).

## Page fit
Options -> **Page fit** (command line: `--fit`):
- **Auto** (default): a sheet that is 1 page in the original Excel file stays 1 page in the PDF
  (it is shrunk slightly to fit). Sheets with several pages keep their normal page setup.
  Needs the Microsoft Excel engine (LibreOffice cannot count pages, so Auto keeps the original setup).
- **Fit every sheet to 1 page**: every sheet is shrunk to one page. Works with Excel and LibreOffice.
  Very long sheets become small, so use it only when you want that.
- **Keep the Excel page setup**: no change.

## Command line
    excel2pdf.exe report.xlsx --cli
    excel2pdf.exe D:\excel_folder -o D:\pdf --jobs 3 --khmer-font "Khmer OS System"
    excel2pdf.exe D:\excel_folder --skip-existing --timeout 300
    excel2pdf.exe report.xlsx --cli --fit all

## What makes it fast
- Fonts are edited inside the .xlsx (no re-save): pictures, shapes and charts are kept
- Several files at once, each worker keeps one hidden Excel open; all sheets export in one pass
- Excel runs with manual calculation and no screen updates; biggest files start first
- Stuck files are stopped after the timeout and retried once; broken files do not stop the batch
- Optional: skip files whose PDF is already newer than the Excel file
- Note: with manual calculation the PDF shows the values saved in the file

## Developers
    pip install -r requirements.txt pyinstaller
    python -m unittest discover -s tests
    pyinstaller --onefile --windowed --name excel2pdf --hidden-import win32com.client --exclude-module tkinter excel_to_pdf.py

    python excel_to_pdf.py            # run from source: starts the server and opens the browser

Files: `excel_to_pdf.py` (entry point), `web.py` (server + API), `web_ui.py` (the page),
`core.py` (engines, workers), `xlsx_fonts.py` (font replacement). `gui.py` (old tkinter window) is no longer used.

## Release
    git tag v1.3.0
    git push origin v1.3.0
GitHub Actions runs the tests, builds the exe and attaches it to the release.
