# New Hire PDF → Excel Importer

Reads all New Employee Information Sheet PDFs in a folder and outputs a `new_hires.xlsx` with:
- FirstName
- LastName
- SO-EmployeeID (Badge #)
- EmployeeID (WD ID)

Names are parsed from the PDF filename. Badge # and WD ID are extracted from the PDF content.

---

## One-Time Setup

Install Python from **https://python.org/downloads**
> ⚠️ During install, check **"Add Python to PATH"** before clicking Install.

Open PowerShell as your **regular user (not admin)** and run:

```
python -m pip install pymupdf openpyxl
```

You only need to do this once.

---

## How to Use

1. Copy `pdf_to_excel.py` into the folder that contains all the PDF files
2. Open PowerShell as your **regular user (not admin)**
3. `cd` into that folder, for example:
   ```
   cd "C:\Users\jylee\OneDrive - Jefferson County, CO\Documents\New Hires\New Employees 07_06_26"
   ```
4. Run:
   ```
   python pdf_to_excel.py --pdf_dir .
   ```
5. Done — `new_hires.xlsx` will be created in the same folder

---

## Notes

- PDF filenames must follow the format:
  `Redacted [FirstName] [LastName] New Employee Information Sheet Signed.pdf`
- For each new batch, copy `pdf_to_excel.py` into that batch's folder and repeat steps 2–4
- Each run creates a fresh `new_hires.xlsx` in the same folder as the PDFs
- The script does not modify any existing Excel files
