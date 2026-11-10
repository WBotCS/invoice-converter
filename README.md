# Invoice Converter

Drop a scanned MLL timesheet PDF on the window; it makes one Excel invoice per
timesheet, using `template.xlsx` for the layout, plus the employee label sheet
(`MLL_LABELS_<invoice no>.docx`, laid out like `labels_template.docx` for the
sticker paper). Works on Mac and Windows.

## Checks

Some employees get part of their pay as a check. Before saving, the app shows
everyone on the check list who is on this timesheet, with their total and check
amount, so the amounts can be confirmed or changed (or someone added/removed).
Their label then reads `Total Pay = 335 + 300`: the cash part (total rounded to
the dollar, minus the check) in bold, then the check. The confirmed list is saved
in `checks.json` (Mac: `~/Library/Application Support/Invoice Converter`,
Windows: `%APPDATA%\Invoice Converter`) and starts from the UW entries of the
payroll workbook's ChecksTable.

## Windows: install on the PC itself

Unzip the project, open the folder and double-click
**Install Invoice Converter (Windows).bat**. It installs Python 3.12 and Tesseract
if needed, builds the app, and adds it to the Desktop and Start menu (about 10
minutes the first time). Run it again to update.

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
