"""
New Hire PDF -> Excel Importer
Parses FirstName + LastName from filename, Badge# and WD ID from PDF content.
Output: new_hires.xlsx with columns FirstName, LastName, SO-EmployeeID, EmployeeID

Requirements: pip install pymupdf openpyxl
Usage: python pdf_to_excel.py --pdf_dir .
"""

import argparse, re, sys
from pathlib import Path
import fitz
import openpyxl

def parse_name(filename):
    """Extract name from: 'Redacted [Name] New Employee Information Sheet Signed.pdf'"""
    stem = Path(filename).stem  # drop .pdf
    m = re.match(r"Redacted\s+(.+?)\s+New Employee Information Sheet Signed", stem)
    if not m:
        return "", ""
    name_parts = m.group(1).strip().split()
    first = name_parts[0] if name_parts else ""
    last  = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
    return first, last

def extract_ids(pdf_path):
    text = ""
    for page in fitz.open(str(pdf_path)):
        text += page.get_text() + "\n"
    wd_id = re.search(r"(\d{6,})", text)
    badge = re.search(r"(\d{2,}-\d{2,})", text)
    return {
        "SO-EmployeeID": badge.group(1) if badge else "",
        "EmployeeID":    wd_id.group(1) if wd_id else "",
    }

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf_dir", required=True)
    args = parser.parse_args()

    pdfs = sorted(Path(args.pdf_dir).glob("*.pdf"))
    if not pdfs:
        print("No PDFs found.", file=sys.stderr); sys.exit(1)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["FirstName", "LastName", "SO-EmployeeID", "EmployeeID"])

    for pdf_path in pdfs:
        first, last = parse_name(pdf_path.name)
        ids = extract_ids(pdf_path)
        print(f"  {first} {last} -> badge={ids['SO-EmployeeID']}  WD={ids['EmployeeID']}")
        ws.append([first, last, ids["SO-EmployeeID"], ids["EmployeeID"]])

    out = Path(args.pdf_dir) / "new_hires.xlsx"
    wb.save(out)
    print(f"\nSaved -> {out}")

if __name__ == "__main__":
    main()
