# Word Number Comparison Tool

A professional Streamlit web application to compare numerical and text data across Word document tables. Designed for auditing and data verification tasks.

## 🚀 Features
- **Smart Extraction**: Extract numbers or text labels from Word tables automatically.
- **Multi-Format Support**: Handles both US (1,234.56) and Vietnam (1.234,56) number formats.
- **Visual Insights**: Highlights discrepancies in real-time with a clean, modern UI.
- **Exportable Reports**: Download comparison results directly to Excel for further analysis.
- **Navigation Guidance**: Built-in instructions to find specific tables in Word using Ctrl+G.

## 🛠️ Tech Stack
- **Frontend**: Streamlit
- **Processing**: Pandas, python-docx
- **Exporting**: Openpyxl

## ⚙️ Installation & Local Setup

1. **Clone the repository**:
   ```bash
   git clone <your-repository-url>
   cd Word_NumCompare_Antigravity
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
