# 📊 Annual Report Financial Data Extractor

A **browser-based web portal** that extracts Balance Sheet and P&L data from company annual report PDFs and lets you download the results as a formatted Excel file — with **no backend, no server, no data upload**.

## 🌐 Live Web Portal

**👉 [Launch the Web Portal](https://pranaymahendrakar.github.io/annual-report-extractor/)**

Upload your annual report PDF → Extract financial data → Download Excel in seconds.

---

## ✨ Features

- **PDF Upload** — Drag & drop or browse to upload any annual report PDF
- **Auto Text Extraction** — Uses PDF.js to extract text directly in the browser
- **Balance Sheet Parsing** — Extracts Equity, Current Liabilities, Non-Current Liabilities, Current Assets, Non-Current Assets
- **P&L Parsing** — Extracts Revenue, COGS, Gross Profit, Employee Expenses, Interest, Depreciation, Other Expenses, Tax, Net Profit
- **Current & Prior Year** — Side-by-side comparison of two financial years
- **Excel Download** — One-click download of a formatted .xlsx file via SheetJS
- **Raw Text Preview** — View the extracted PDF text to verify parsing
- **Privacy First** — Everything runs in your browser; no data leaves your device

---

## 🖥️ How to Use the Web Portal

1. Go to **[https://pranaymahendrakar.github.io/annual-report-extractor/](https://pranaymahendrakar.github.io/annual-report-extractor/)**
2. Drag & drop your annual report PDF, or click **Browse PDF File**
3. Enter/confirm the company name
4. Click **⚡ Extract Financial Data**
5. Review the Balance Sheet and P&L tables
6. Click **📥 Download Excel Report** to get the .xlsx file

---

## 🐍 Python CLI (Offline / Batch Use)

For processing PDFs offline or in batch, use the Python script:

### Installation

```bash
pip install pdfplumber openpyxl
```

### Usage

```bash
python extractor.py <annual_report.pdf> [output.xlsx] [Company Name]
```

**Example:**
```bash
python extractor.py Madhav_Infra_FY2223.pdf Madhav_Output.xlsx "Madhav Infra Projects Ltd"
```

### Output

The script prints a summary to the console and saves a formatted Excel file with:
- Balance Sheet (both years)
- P&L Statement (both years)
- Difference check (Assets − Liabilities − Equity)

---

## 📁 Project Structure

```
annual-report-extractor/
├── index.html        # Web portal (single HTML file, runs in browser)
├── extractor.py      # Python CLI script for offline/batch processing
└── README.md         # This file
```

---

## 🔧 Tech Stack

| Component | Technology |
|-----------|-----------|
| Web Portal | Pure HTML5, CSS3, Vanilla JavaScript |
| PDF Parsing | [PDF.js](https://mozilla.github.io/pdf.js/) (Mozilla) |
| Excel Export | [SheetJS (xlsx)](https://sheetjs.com/) |
| Python PDF | [pdfplumber](https://github.com/jsvine/pdfplumber) |
| Python Excel | [openpyxl](https://openpyxl.readthedocs.io/) |
| Hosting | GitHub Pages |

---

## ⚠️ Limitations

- Parsing accuracy depends on the PDF's text structure. Scanned/image PDFs won't work.
- The parser is tuned for annual reports with standard Indian financial statement formats.
- For unusual layouts, you may need to adjust the regex patterns in the code.

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

*Built with ❤️ — Runs entirely in your browser. No server. No data upload. No backend.*
