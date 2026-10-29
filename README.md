# Invoice Converter

Drop a scanned MLL timesheet PDF on the window; it makes one Excel invoice per
timesheet, using `template.xlsx` for the layout. Works on Mac and Windows.

## Windows: download the app

1. On GitHub, open **Actions** → the latest green **Build Windows app** run.
2. At the bottom, download **Invoice-Converter-Windows** and unzip it.
3. Open the folder and double-click **Invoice Converter.exe**
   (first time: "Windows protected your PC" → **More info** → **Run anyway**).

Everything is inside the folder, including the text reader (Tesseract). Keep the
folder together; you can make a desktop shortcut to the .exe.

Each push to this repository builds the Windows app automatically, runs a
self-test, converts the PDFs in `tests/` and checks the totals before publishing it.

## Mac

Run the installer once (it installs Homebrew/Tesseract if needed, builds the app,
puts it in Applications and checks it starts):

    bash "Install Invoice Converter.command"

## Command line

    python pdf_to_invoice.py "Document.pdf" --invoice-no UN-B-84
