# PDF Contact Extractor

A Python script that scans a folder of PDFs and pulls out names, emails, and phone numbers — then dumps everything into a neat Excel file and CSV.

Basically: if you've ever had to open 80 PDFs one by one to copy contact info into a spreadsheet, this script is for you. One command, done in seconds.

---

## What it does

- Scans every `.pdf` in a folder you point it at
- Extracts **name**, **email**, and **phone number** from each one
- Saves results to `contacts_output.xlsx` (formatted, blue headers, alternating rows) and `contacts_output.csv`
- Shows live progress while it runs
- Flags files where something's missing so you know which ones to check manually

---

## Requirements

- Python 3.7+
- Works on Windows, macOS, Linux

Install the two dependencies:

```bash
pip install pdfplumber openpyxl
```

---

## Usage

**Step 1.** Put all your PDFs in one folder.

**Step 2.** Open `pdf_contact_extractor.py` and set your folder path near the top:

```python
PDF_FOLDER = "C:/Users/YourName/Desktop/my_pdfs"   # Windows
PDF_FOLDER = "/home/yourname/documents/my_pdfs"    # macOS / Linux
```

> **Windows tip:** use forward slashes `/` not backslashes `\`. Python's fine with it, and it saves you a headache.

**Step 3.** Run it:

```bash
python pdf_contact_extractor.py
```

When it finishes, `contacts_output.xlsx` and `contacts_output.csv` appear in the same folder as the script.

---

## Output

### Excel (`contacts_output.xlsx`)

| File | Name | Email | Phone |
|------|------|-------|-------|
| resume_john.pdf | John Smith | john@email.com | 01711234567 |
| form_sara.pdf | Sara Ahmed | sara@gmail.com | +8801812345678 |
| doc_unknown.pdf | | info@company.com | |

Blue header row, alternating shading, auto-width columns.

### CSV (`contacts_output.csv`)

Same data, plain text. Easy to open in Google Sheets or import into a database.

### Terminal output

```
Found 3 PDF(s). Processing...

  [1/3] resume_john.pdf
  [2/3] form_sara.pdf
  [3/3] doc_unknown.pdf
         ⚠ name missing, phone missing

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  Done! Processed 3 PDF(s)
  ✔ Full data (name+email+phone): 2
  ⚠ Partial data:                 1
  ✗ Nothing found:                0
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```

---

## How it extracts data

### Email
Regex. Reliable for basically any email format you'll encounter.

### Phone numbers
Handles most formats people actually use:
- `01711234567`
- `+8801711234567`
- `(017) 123-4567`
- `017-123-4567`

Country codes work too.

### Names
Two strategies, tried in order:

1. **Labelled fields** — scans for patterns like `Name: John Doe`, `Full Name:`, `Candidate:`, etc. Most reliable when the PDF has a form-style layout.
2. **Top-of-document scan** — if no label is found, checks the first 15 lines for a standalone "First Last" pattern. Works well for resumes.

If neither strategy finds a name, the cell is left blank and the file gets flagged in the terminal output.

---

## Limitations

**Text-based PDFs only.** If your PDFs are scanned images (i.e. you can't highlight the text), this won't work. You'd need OCR first — something like Tesseract.

**One contact per PDF.** It pulls the first name, email, and phone number it finds. If a PDF has multiple contacts, only the first is captured.

**Names are heuristic.** Unusual layouts can confuse it. If you're getting a lot of blank name fields, the `extract_name()` function is easy to tweak for your specific document format.

---

## Folder structure

```
pdf_extractor/
├── pdf_contact_extractor.py    ← the script
├── resume_1.pdf
├── resume_2.pdf
├── ...
├── contacts_output.xlsx        ← created after running
└── contacts_output.csv         ← created after running
```

---

## Common errors

**`SyntaxError: unterminated string literal`** (Windows)

Path ends with a backslash. Switch to forward slashes:
```python
# Causes the error
PDF_FOLDER = "C:\Users\FHP\Desktop\pdf_extractor\"

# Works fine
PDF_FOLDER = "C:/Users/FHP/Desktop/pdf_extractor"
```

**`'python' is not recognized`**

Python isn't on your PATH. Reinstall from [python.org](https://python.org) — on the first screen of the installer, check **"Add Python to PATH"** before clicking Install.

**`ModuleNotFoundError: No module named 'pdfplumber'`**

```bash
pip install pdfplumber openpyxl
```

---

## License

MIT. Use it, modify it, do whatever.
