# 🏥 Georgian Medical Document Translation Pipeline

Automated translation and data digitization pipeline for converting complex Georgian clinical documents (Form #100 certificates, medical invoices, diagnostic panels, and laboratory blood tests) into professionally formatted English PDFs and structured Excel workbooks.

---

## 🚀 Live Streamlit Deployment

### Option A: Free 1-Click Deployment to Streamlit Community Cloud
1. Push this directory to your GitHub account:
   ```bash
   git init
   git add .
   git commit -m "Initial commit for Medical Document Translator"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo-name>.git
   git push -u origin main
   ```
2. Go to [share.streamlit.io](https://share.streamlit.io/) and log in with your GitHub account.
3. Click **"New app"** > select your repository, branch (`main`), and set the main file path to:
   ```text
   app.py
   ```
4. Click **"Advanced settings..."** > **"Secrets"**, and paste your API key:
   ```toml
   GEMINI_API_KEY = "your-gemini-api-key-here"
   ```
5. Click **"Deploy!"**
   * Streamlit Cloud will install system libraries from `packages.txt` and Python packages from `requirements.txt`. Your live public HTTPS URL will be ready in ~2 minutes!

---

### Option B: Run Locally
```bash
# In the streamlit_pdf folder:
streamlit run app.py
```
Open [http://localhost:8501](http://localhost:8501) in your browser.

---

## 📁 Repository Structure
```text
streamlit_pdf/
├── app.py                   # Streamlit main application & translation engine
├── pdf_html_pdf.py          # Dual CLI / module entry point
├── requirements.txt         # Python dependencies
├── packages.txt             # Linux system dependencies (Pango, Cairo, GObject for WeasyPrint)
├── .streamlit/
│   └── config.toml          # Streamlit theme & file upload configuration
├── sample_documents/        # Ready-to-test Georgian medical documents
│   ├── hostile_real_invoice_GEO.pdf
│   ├── hostile_real_forma_100_GEO.pdf
│   ├── fake_medical_invoice_GEO.pdf
│   ├── fake_forma_100_GEO.pdf
│   └── bolgova_F100_GEO.pdf
└── README.md
```
