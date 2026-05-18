# Rebar Schedule PDF Extractor

A Python tool that reads structural rebar schedule PDFs, extracts bar details, calculates weights using the standard engineering formula, and exports clean results to Excel.

## What it does

- Reads native vector PDFs using `pdfplumber` (text-selectable drawings)
- Falls back to `Tesseract OCR + OpenCV` for scanned image PDFs (fully offline, no API needed)
- Extracts: member name, bar mark, diameter, length, quantity, spacing
- Calculates unit weight and total weight using the standard BBS formula: **W = D² / 162 (kg/m)**
- Exports a formatted Excel file with a summary row

## Standard weight formula

```
Unit Weight (kg/m) = D² / 162
Total Weight (kg)  = Unit Weight × Length (m) × Quantity
```

This follows IS 456 and is the universally accepted formula for rebar weight calculation in BBS.

## Installation

```bash
pip install pdfplumber pymupdf pandas openpyxl fpdf2 pillow pytesseract opencv-python-headless
```

For scanned PDFs, also install the Tesseract OCR engine:

```bash
# Mac
brew install tesseract

# Linux
sudo apt install tesseract-ocr

# Windows
# Download installer from https://github.com/UB-Mannheim/tesseract/wiki
```

## Usage

### Run the built-in demo (no PDF needed)

Generates a realistic sample rebar schedule PDF and runs the full extraction pipeline on it.

```bash
python rebar_pdf_extractor.py --demo
```

### Run on your own PDF

```bash
python rebar_pdf_extractor.py --pdf your_rebar_schedule.pdf
```

### Specify output file name

```bash
python rebar_pdf_extractor.py --pdf drawing.pdf --output results.xlsx
```

## Output

The Excel file contains one row per rebar entry with these columns:

| Member | Bar Mark | Diameter (mm) | Length (m) | Quantity | Spacing (mm) | Unit Weight (kg/m) | Total Length (m) | Total Weight (kg) |
|--------|----------|--------------|------------|----------|--------------|--------------------|-----------------|------------------|
| Footing F1 | F1-01 | 16 | 3.20 | 12 | 150 | 1.580 | 38.400 | 60.672 |
| ... | | | | | | | | |
| **TOTAL** | | | | | | | **X.XXX** | **X.XXX** |

## Supported rebar diameters

6mm, 8mm, 10mm, 12mm, 16mm, 20mm, 25mm, 28mm, 32mm, 40mm

## Tech stack

- `pdfplumber` — vector PDF table and text extraction
- `pytesseract` + `Pillow` — OCR fallback for scanned PDFs
- `opencv-python-headless` — image preprocessing for better OCR accuracy
- `pandas` + `openpyxl` — data processing and Excel export
- `fpdf2` — sample PDF generation

## Notes

- For best results on scanned drawings, ensure scan resolution is at least 300 DPI
- The parser targets standard BBS table layouts; custom formats may need minor adjustments
- No cloud API or internet connection required for the default local pipeline

## Author

**Divyanshu Raghuwanshi**  
[GitHub](https://github.com/divyanshu964) | [LinkedIn](https://www.linkedin.com/in/divyanshu-raghuwanshi-85037b160/) | [Portfolio](https://www.datascienceportfol.io/divyanshu964)
