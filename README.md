
# Quotation PDF Generator

FastAPI web application for the supplied Zyflox quotation template.

## What it does

1. Accepts quotation/customer details.
2. Allows any number of:
   - Shared Storage Service items
   - Handling Service items
   - Add On Service items
3. Calculates:
   - Shared Storage: `amount = rate × qty`
   - Handling: `amount = rate`
   - Add On: `amount = rate`
   - GST: `amount × 9%`
4. Populates `template.docx`.
5. Converts it to PDF using LibreOffice.
6. Returns only the PDF to the browser.

The source DOCX is never sent to the user.

## Run on Windows

Install Python 3.11+ and LibreOffice.

From this folder:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open:

http://127.0.0.1:8000

If LibreOffice is installed but not in PATH:

```powershell
$env:LIBREOFFICE_PATH="C:\Program Files\LibreOffice\program\soffice.exe"
uvicorn app.main:app --reload
```

## Linux server

Install LibreOffice, then:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

## Important template note

The uploaded template has some placeholder inconsistencies, including a space-split
`{{last_day_of_ date _month}}` and a few incorrect/mismatched sample-row
placeholders. The application does not depend on those item placeholders for
the service rows; it expands the existing service-table prototype rows and
writes the calculated values directly into the corresponding columns.

The existing template also contains `{{storage_date}}` and `{{Produt_cat}}`.
`storage_date` is populated. `Produt_cat` is not part of the form requested,
so it remains unchanged unless you add that field.

## Change the template

Replace `template.docx` with the revised DOCX while keeping the same three
service tables. The app identifies:
- table 1 = Shared Storage
- table 2 = Handling
- table 3 = Add On

The first data row in each table is used as the formatting prototype.
