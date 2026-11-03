==============================================================
 PDF TOOL  -  version 1.0.0  (released 2026-10-27)
==============================================================

WHAT IT DOES
  Converts Word, Excel, PowerPoint files and images to PDF.
  You can make one separate PDF per file, or combine many files
  into ONE PDF (in an order you control).

FILES IN THIS PACKAGE
  run_pdf_tool.bat   Launcher. Checks for Python (offers to install it),
                     then starts the tool.
  pdf_tool.py        The tool itself.
  README.txt         This file.

--------------------------------------------------------------
 QUICK START
--------------------------------------------------------------
  1. Copy run_pdf_tool.bat and pdf_tool.py into the folder that
     contains the files you want to convert.
  2. Click the address bar of that folder, type  cmd  and press Enter.
  3. Type  run_pdf_tool.bat  and press Enter.
     (If Python is already installed you can also use:
      python pdf_tool.py)
  4. Press a number from the menu and follow the questions.

  The tool works on the files in the folder where you run it.
  It does not look inside sub-folders.

--------------------------------------------------------------
 THE MENU
--------------------------------------------------------------
  1 - Word        -> PDF
  2 - Images      -> PDF
  3 - Excel       -> PDF
  4 - PowerPoint  -> PDF
  5 - EVERYTHING  -> ONE PDF (Word + Excel + PowerPoint + images)
  8 - About / version history
  9 - System check / install missing components
  0 - Exit

  Next to each option the tool shows how many matching files it found.

  For options 1 to 4 you then choose:
    1 - Combine all into ONE PDF
    2 - Separate PDF for each file
    0 - Back
  Option 5 always makes one combined PDF.

--------------------------------------------------------------
 HOW COMBINING AND order.txt WORK
--------------------------------------------------------------
  When you combine, the tool creates a file called order.txt in the
  folder and opens it in your text editor. It lists the files, one
  name per line, sorted by name (1, 2, 10 in natural order).

    - Move a line up or down to change the order.
    - Delete a line to leave that file out.
    - You can add the name of another supported file from the same
      folder to include it.
    - Lines starting with # are ignored.
    - Save the file, go back to the tool window and press Enter.

  Files not listed in order.txt are NOT converted.
  order.txt is created fresh every time you combine.
  Names that do not exist in the folder are skipped with a warning.

--------------------------------------------------------------
 WHERE THE RESULTS GO
--------------------------------------------------------------
  Everything goes into a folder called PDF_Output (created next to
  your files). Your original files are never changed.

    Combined PDF:  Combined_<Type>_<date_time>.pdf
                   Example: Combined_All_20261027_153012.pdf
                   It has a bookmark for each source file.
    Separate PDFs: same name as the original, ending in .pdf
                   If two files would get the same name (report.doc
                   and report.docx) the second becomes report_docx.pdf.
    Log file:      pdf_tool_log.txt (what worked, what failed).

  If one file fails (damaged, password-protected) the tool tells you
  why, continues with the others, and asks whether to combine the rest.

--------------------------------------------------------------
 SUPPORTED FILE TYPES
--------------------------------------------------------------
  Word        .doc .docx .docm .dot .dotx .dotm .rtf .odt .txt
  Excel       .xls .xlsx .xlsm .xlsb .xlt .xltx .xltm .csv .ods
  PowerPoint  .ppt .pptx .pptm .pps .ppsx .ppsm .pot .potx .potm .odp
  Images      .jpg .jpeg .jpe .jfif .png .gif .bmp .dib .tif .tiff
              .webp .ico .heic .heif .ppm .pgm .pbm .tga .jp2
              (.heic/.heif need the optional pillow-heif package;
               option 9 can install it)

  Files are recognised by their real extension, so a folder mixing
  jpg, gif, png and so on works fine.
  Temporary files starting with ~$ are ignored.

  Images:
    - Each image becomes one A4 page, scaled to fit, centred, with a
      small white border. Wide images get a landscape page.
    - Transparent PNG/GIF areas become white.
    - Animated GIF and WebP: the first frame is used.
    - Multi-page TIFF: every page is used.
    - Photos taken in portrait are rotated correctly.

--------------------------------------------------------------
 HOW OFFICE FILES ARE CONVERTED
--------------------------------------------------------------
  The tool picks an engine for each file type:
    1. Microsoft Office (if installed) - best quality. Excel sheets
       are fitted to page width so columns do not split.
    2. LibreOffice (free) - used when Office is missing, or if
       Office fails on a file.
  Images need neither.

  At start-up the tool shows which engine it found for Word, Excel
  and PowerPoint. If you want LibreOffice first on every PC (same
  output everywhere), open pdf_tool.py in Notepad and change
      PREFER_OFFICE = True   to   PREFER_OFFICE = False

--------------------------------------------------------------
 SYSTEM CHECK AND AUTOMATIC INSTALL (option 9)
--------------------------------------------------------------
  The tool checks itself every time it starts. If something is
  missing it asks "Install now? (Y/N)".

    Python (launcher)  Offers to install Python 3.12 (winget, or a
                       download from python.org for the current user).
    Python packages    Pillow, pypdf, pywin32 (only when Office is
                       present), pillow-heif (optional). Installed
                       with pip. The tool restarts afterwards.
    LibreOffice        Offered when Office is not available
                       (about 350 MB; Windows may ask for permission).
    Microsoft Office   Cannot be installed by the tool.

  An internet connection is needed the first time on each PC.
  If a download is blocked (company firewall) the tool prints the
  manual download link instead.
  Images and combining work even before LibreOffice is installed.

--------------------------------------------------------------
 SETTINGS (top of pdf_tool.py, safe to edit with Notepad)
--------------------------------------------------------------
  PREFER_OFFICE      True = Office first, False = LibreOffice first
  OUTPUT_DIR_NAME    Name of the output folder
  IMAGE_DPI          Image page resolution (default 200)
  IMAGE_QUALITY      JPEG quality inside PDFs (default 90)
  IMAGE_MARGIN_INCH  White border around images (default 0.3)

--------------------------------------------------------------
 KNOWN LIMITS
--------------------------------------------------------------
  - Windows is needed for the Microsoft Office engine.
  - With LibreOffice, very wide Excel sheets can split across pages.
    Microsoft Office does not have this problem.
  - Missing fonts on a PC can change the layout slightly when
    LibreOffice is used.
  - Password-protected files are skipped. An old protected .ppt may
    make PowerPoint wait for a password - close it if that happens.
  - Arabic text in .txt and .csv files depends on the file encoding.
    UTF-8 with BOM is the safest.
  - Macros are never run.
  - Do not keep a file open in Word/Excel/PowerPoint while
    converting it.
  - v1.0.0 was tested with LibreOffice and images on a test set.
    The Microsoft Office route and the installers could not be
    tested on a real Windows PC before release.

--------------------------------------------------------------
 IF SOMETHING GOES WRONG
--------------------------------------------------------------
  - Run option 9 (System check) first.
  - Open PDF_Output\pdf_tool_log.txt - it lists every file and the
    reason for any failure.
  - To get help or request changes, send the message shown on screen
    together with the version number (top of the window).

--------------------------------------------------------------
 VERSION HISTORY
--------------------------------------------------------------
  v1.0.0  (2026-10-27)  First release.
          Word, Excel, PowerPoint and images to PDF (separate or
          combined). Option 5 combines everything. order.txt.
          Microsoft Office engine with LibreOffice fallback.
          System check with automatic install of missing parts.

  ASKING FOR A NEW VERSION
  Bring pdf_tool.py to a new chat, tell Claude the version shown
  on screen and what you want added. The new file will have a higher
  version number. Keep the old file renamed (for example
  pdf_tool_v1.0.0.py) as a backup.
==============================================================
