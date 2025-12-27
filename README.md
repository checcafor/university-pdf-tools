# 🎓 University PDF Tools
Developed to save time during exam sessions! ☕

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.x](https://img.shields.io/badge/python-3.x-blue.svg)](https://www.python.org/)

A simple yet powerful Python toolkit designed to automate the management of university lecture notes. It handles batch conversion from Word to PDF and merges multiple lecture files into a single document using **smart sorting**.

## 🚀 Features

* **Batch Conversion:** Converts an entire folder of `.docx` or `.doc` files to PDF automatically.
* **Format Preservation:** Uses the native Microsoft Word engine to ensure perfect formatting in the output PDF.
* **Smart Sorting:** Unlike standard sorting, the merge tool understands that `Lecture 2` comes before `Lecture 10` (Natural Sorting).
* **Drag & Drop Friendly:** Scripts are optimized to handle file paths dragged directly into the terminal (supports PowerShell and CMD formatting).
* **Organized Output:** Automatically creates subfolders for converted files to keep your workspace clean.

---

## 🛠️ Scripts Included

### 1. `word2pdf.py` (Batch Converter)
Takes a folder containing Word documents and converts them all to PDF.
* **Input:** Folder path.
* **Output:** Creates a `PDF_Convertiti` subfolder containing the PDFs.
* **Safety:** Skips temporary Word files (start with `~$`) to avoid errors.

### 2. `unisci_pdf.py` (Smart Merger)
Merges all PDF files in a specific folder into one single master file (e.g., `RIEPILOGO_LEZIONI_UNITO.pdf`).
* **Smart Sort Logic:** parses filenames to find numbers (e.g., "Lezione_1", "Lezione_2") and sorts them numerically rather than alphabetically.

---

## ⚙️ Prerequisites

Before running the scripts, ensure you have the following:

1.  **Python 3.x** installed.
2.  **Microsoft Word** installed on your machine (Required for `docx2pdf` conversion).
3.  **OS:** Windows (Recommended) or macOS.

---

## 📦 Installation

1.  **Clone the repository:**
    ```bash
    git clone [https://github.com/checcafor/university-pdf-tools.git](https://github.com/checcafor/university-pdf-tools.git)
    cd university-pdf-tools
    ```

2.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

---

## 📖 Usage

### Convert Word files to PDF
Run the script and paste/drag the folder path when prompted:
```bash
python word2pdf.py
```
### Merge PDFs into one file
Run the script and paste/drag the folder containing the PDFs:
```bash
python unisci_pdf.py
```

## 📂 Project Structure
university-pdf-tools/
├── word2pdf.py          # Script for batch conversion
├── unisci_pdf.py        # Script for merging PDFs
├── requirements.txt     # List of dependencies
├── .gitignore           # Files to ignore (docs, envs)
└── README.md            # Project documentation

## 📝 License
This project is licensed under the MIT License - see the LICENSE file for details.
