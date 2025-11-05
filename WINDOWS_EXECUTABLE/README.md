RosterToMatrix (Windows Executable)
===================================

This folder holds the packaged Windows build of the Excel roster-to-matrix converter. The binary was created with PyInstaller and bundles drag-and-drop support via TkinterDnD2 so non-technical teammates can use the GUI without installing Python.

Contents
--------

* `RosterToMatrix.exe` – the standalone GUI application.
* `courses.txt` – sample course list. Copy your real catalog into this folder before running the app.

Running the App
---------------

1. Place your roster workbook (`.xlsx` or `.xlsm`) and the appropriate `courses.txt` in the same directory as the executable.
2. Double-click `RosterToMatrix.exe`. Windows SmartScreen may warn you about an unrecognized app; choose **More info** → **Run anyway** the first time.
3. In the GUI:
   - Click **Browse…** or drag a workbook into the **Input workbook** field.
   - If you already have a `courses.txt`, click **Browse…** or drag it into the **Courses list** field. If you do not have one yet, see the next section.
   - Review the auto-generated output path, then press **Convert**.
4. The tool writes a timestamped `.xlsx` next to the input file (e.g., `251103-2342-Original.xlsx`).

Getting or Replacing `courses.txt`
----------------------------------

The converter expects a plain-text file containing one course name per line. If you do not have `courses.txt` yet:

1. Open a text editor (Notepad works fine).
2. Enter the exact course names you want to appear in the output matrix, one per line.
3. Save the file as `courses.txt` inside this `WINDOWS_EXECUTABLE` folder.

You can maintain different catalogs by saving alternate `courses_*.txt` files and selecting them through the GUI when needed.

Drag-and-Drop Requirements
---------------------------

Drag-and-drop relies on the bundled TkinterDnD2 assets. If you rebuild the executable on another machine, ensure you install TkinterDnD2 before running PyInstaller:

```bash
python -m pip install tkinterdnd2 pyinstaller
python -m PyInstaller --name RosterToMatrix --windowed --onefile ^
    --add-data \"courses.txt;.\" ^
    --collect-data tkinterdnd2 ^
    excel_converter.py
```

Resolving “File is not a zip file” Errors
-----------------------------------------

If the converter reports `zipfile.BadZipFile: File is not a zip file`, the source workbook is not in the modern Excel format even if it has a `.xlsx` or `.xlsm` extension. Fix it by:

1. Opening the roster in Excel.
2. Choosing **File → Save As**.
3. Selecting **Excel Workbook (*.xlsx)** (or **Excel Macro-Enabled Workbook (*.xlsm)** if you truly need macros).
4. Running the converter again with the newly saved file.

Need More Help?
---------------

If you run into other issues:

* Confirm you are opening a workbook that contains the required columns (`LocName`, `Course Name`, `JobStatus`, `Start Date`, and `End Date`).
* Ensure `courses.txt` is present and readable; the app will prompt if it cannot find the file.
* Reach out to the project maintainer with the exact error message and a copy of the problematic workbook if the problem persists.
