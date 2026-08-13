# Advanced Data Forge

A lightweight desktop GUI for quickly loading, cleaning, and reshaping tabular
data — built with **Python**, **Tkinter**, and **pandas**. No spreadsheet
formulas or scripting needed for common data-cleaning tasks.

## Features

- 📂 **Load Data** — import CSV, TXT (tab-delimited), or Excel (`.xlsx` / `.xls`) files
- 🧩 **Select Columns** — pick exactly which columns you want to keep
- 🔍 **Filter Builder** — stack multiple conditions (exact match, contains, starts with, ends with) and apply them together
- 🔄 **Pivot Tool** — reshape data by choosing index, pivot, and value columns
- 📊 **Quick Analysis** — pick any X/Y columns and instantly plot a bar, line, or scatter chart of the current (filtered) data
- 🧹 **Remove Duplicates** — clean up repeated rows in one click
- ↩️ **Back / Reset** — undo the last change, or reset to the originally loaded file
- 💾 **Export** — save the current view to a new Excel file

## Screenshot

*(Add a screenshot of the app here before publishing — it makes a big difference on GitHub/LinkedIn!)*

## Getting Started

### Prerequisites
- Python 3.9+

### Installation

```bash
git clone https://github.com/<your-username>/advanced-data-forge.git
cd advanced-data-forge
pip install -r requirements.txt
```

### Run

```bash
python data_forge.py
```

## Notes

- The **Export** feature automatically opens the saved file after export; this
  currently relies on `os.startfile()`, which is Windows-only. On macOS/Linux
  the file will still save correctly to `output.xlsx`, but it won't auto-open.

## Tech Stack

- Python
- Tkinter (GUI)
- pandas (data processing)
- openpyxl (Excel read/write)

## Author

Developed by **BaHaa**
