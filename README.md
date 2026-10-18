# Word Number Comparison Tool

A professional Streamlit web application to compare numerical and text data across Word document tables. Designed for auditing, accounting, and data verification tasks.

## 🚀 Key Features
- **Table Structure Audit**: Automatically inspects table quantity and row/column dimensions between documents to detect structural variations.
- **Smart Row Alignment**: Sequence-based alignment algorithm that prevents false positive mismatches when extra rows or notes are inserted or removed in one document.
- **Context-Aware Comparison**: Extracts row labels/item descriptions so every comparison entry is clearly identified.
- **Automated Word Comments & Cell Shading**: Automatically highlights mismatched cells in red and extra rows in yellow, adding native Word comments with detailed discrepancy breakdowns directly onto both `.docx` files.
- **Multi-Format Support**: Handles both US (`1,234.56`) and Vietnam (`1.234,56`) number formats.
- **Comprehensive Export Options**:
  - Download annotated `.docx` for Document 1.
  - Download annotated `.docx` for Document 2.
  - Download structured multi-sheet Excel report (`.xlsx`).
  - Download complete package as a single `.zip` archive.

## 🛠️ Tech Stack
- **Frontend**: Streamlit
- **Processing**: Pandas, python-docx
- **Exporting**: Openpyxl, Zipfile

## ⚙️ Installation & Local Setup

1. **Clone the repository**:
   ```bash
   git clone <your-repository-url>
   cd Word_NumCompare_Antigravity_V2
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   streamlit run app.py
   ```

## ☁️ Deployment on Streamlit Cloud

1. Push this repository to GitHub.
2. Log in to [Streamlit Cloud](https://share.streamlit.io/).
3. Click "New App", select your repository, branch, and `app.py` as the main file.
4. Deploy!

---
Developed for professional auditing and data comparison tasks.
