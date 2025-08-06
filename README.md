# CSV Compare Tool (Side-by-Side + Unique Values)

This Python script compares two CSV files by a chosen column and generates an Excel file with:

1. **Side_by_Side Sheet**  
   - Two columns: values from the selected column in **File 1** and **File 2**  
   - Conditional formatting:  
     - **Green** → Value exists in both files  
     - **Red** → Value exists in only one file

2. **Unique_Values Sheet**  
   - A single column containing all values that appear in **only one** of the two files (no duplicates between them)

---

## Features
- Works regardless of row order — matches purely by chosen column.
- Conditional formatting in Excel for instant visual checks.
- Automatically saves to your **Downloads** folder.
- Appends a number to the file name if a file with the same name already exists.

---

## Requirements
- Python 3.8+
- [pandas](https://pandas.pydata.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

Install dependencies:
```bash
pip install pandas openpyxl
