# 📊 Excel Data Cleaner & Analyzer

A Python tool that turns **messy spreadsheets into clean, analysis-ready data in one click**.
It removes duplicates, fixes formatting, flags missing data, and generates summary charts automatically.

> Cut a client's monthly report cleaning from **2 hours → 30 seconds**.

---

## ✨ Features

- ✅ Removes duplicate rows
- ✅ Cleans column names (lowercase, no spaces)
- ✅ Fixes text & date formatting
- ✅ Auto-detects numbers and dates
- ✅ Flags missing / invalid values
- ✅ Generates summary charts (PNG)
- ✅ Exports a clean Excel file with a **Quality Report** sheet

---

## 🛠️ Tech Stack

- **Python 3**
- **pandas** – data cleaning
- **openpyxl** – Excel read/write
- **matplotlib** – charts

---

## 🚀 Installation

```bash
# 1. Clone the repo
git clone https://github.com/waseemqadir345/Excel-Data-Cleaner-Analyzer.git
cd Excel-Data-Cleaner-Analyzer

# 2. Install dependencies
pip install -r requirements.txt
```

📖 Usage



python excel_cleaner.py sample_data/messy_sample.csv

Or with your own file:



python excel_cleaner.py your_file.xlsx

Output

your_file_cleaned.xlsx → clean data + quality report
your_file_chart_<column>.png → summary charts


📂 Project Structure



excel-data-cleaner/
├── README.md
├── requirements.txt
├── excel_cleaner.py
├── .gitignore
└── sample_data/
    └── messy_sample.csv


👤 Author
M Waseem Q — Python Developer & Web Automation Specialist
📩 Available for freelance work on Upwork.

📝 License
MIT License — free to use and modify.
