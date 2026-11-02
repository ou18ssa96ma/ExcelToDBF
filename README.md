# Excel → DBF Converter

A focused, clean desktop tool to convert multi-sheet Excel files into individual dBase (`.dbf`) files. **Desktop GUI only** — no CLI mode, no sidecar files.

## Features

- **Modular Architecture** — `converter.py` (engine), `gui.py` (UI), `main.py` (launcher)
- **Worksheet Selection Table** — preview sheet names, row/column counts, and choose which sheets to convert
- **Broad Format Support** — `.xlsx`, `.xlsm`, and legacy `.xls`
- **Robust Type Inference** — auto-detects `C`, `N`, `D`, `L` field types
- **UTF-8, CP1256 (Arabic), CP1252** encoding options
- **dBase III / IV / FoxPro** format options
- **Numeric Precision Safety** — numbers with more than 15 digits are stored as text
- **Drag & Drop** — drop an Excel file directly onto the window
- **Settings Memory** — remembers last folders and preferences

## Requirements

- Python 3.10+
- Windows 10/11 (for the Windows build)
- See `requirements.txt`

## Install

Clone the repository:

```bash
git clone https://github.com/ou18ssa96ma/ExcelToDBF.git
cd ExcelToDBF
```

Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

## Run

Double-click `run.bat`, or run:

```bash
python main.py
```

Optionally, open the application with a workbook pre-selected:

```bash
python main.py "path\to\workbook.xlsx"
```

## Build for Windows

This project uses **PyInstaller** to package the application into a standalone Windows executable.

### 1. Prepare the environment

Make sure Python 3.10+ is installed on Windows.

Create and activate the virtual environment if you haven't already:

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 2. Run the build script

Double-click `build.bat`, or execute it from CMD:

```bat
build.bat
```

The script automatically:

- Installs the project dependencies.
- Installs PyInstaller.
- Cleans previous build files.
- Packages the application in Windows GUI mode.
- Creates the executable using PyInstaller's `--onedir` mode.

### 3. Find the executable

After a successful build, the output will be:

```text
dist/
└── ExcelToDBF/
    └── ExcelToDBF.exe
```

Run:

```text
dist\ExcelToDBF\ExcelToDBF.exe
```

### 4. Distribute the application

To share the application with other Windows users:

1. Open the `dist` folder.
2. Compress the entire `ExcelToDBF` folder into a ZIP archive.
3. Share the ZIP file or upload it to GitHub Releases.
4. Users can extract the archive and run `ExcelToDBF.exe`.

**Important:** Do not distribute only the `.exe` file when using `--onedir`. The `_internal` folder is required.

## Project Structure

```text
ExcelToDBF/
├── .venv/
├── tests/
├── converter.py
├── gui.py
├── main.py
├── requirements.txt
├── run.bat
├── build.bat
└── README.md
```
