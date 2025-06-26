
# QC Sheet Generator (Streamlit)

This Streamlit app generates QC (Quality Control) sheets for fashion production based on a size spec Excel file, a QC sheet template, and a signature/logo image.

## How to run locally
```bash
pip install -r requirements.txt
streamlit run qc_sheet_app.py
```

## Deploy to Streamlit Community Cloud
1. Push this folder to a **public GitHub repository**.
2. Go to https://share.streamlit.io → **Create app**.
3. Select your repository, branch (e.g. `main`), and `qc_sheet_app.py` as the main file.
4. Click **Deploy** – your app will be available at  
   `https://your-username-your-repo-name.streamlit.app`

## Folder structure
```
qc_sheet_app/
├── qc_sheet_app.py
├── requirements.txt
├── README.md
└── uploaded/
    ├── spec/            # Example spec Excel file
    ├── template/        # QC sheet template
    └── image/           # Signature / logo image
```
