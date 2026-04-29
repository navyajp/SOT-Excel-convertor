# NSDL Statement Converter

Convert ICICI NSDL transaction statement PDFs into clean, downloadable Excel files.

## Features

- 📄 Upload any ICICI NSDL transaction statement PDF
- ⚡ Parses all transactions automatically (tested on 3962-page statements)
- 📊 Outputs Excel with 8 columns + a Summary sheet
- 🔍 In-browser preview of first 50 rows before download
- 🖥️ Beautiful, responsive frontend UI

## Excel Output Columns

| Column | Description |
|---|---|
| Client Name | Account holder name |
| Security Name | ISIN security name |
| Transaction Date | Booking date (DD-Mon-YYYY) |
| Transaction Number | NSDL transaction reference |
| Opening Balance | Quantity before transaction |
| Credit | Units credited (By transactions) |
| Debit | Units debited (To transactions) |
| Closing Balance | Quantity after transaction |

## Quick Start

### 1. Clone the repository

```bash
git clone https://github.com/YOUR_USERNAME/nsdl-converter.git
cd nsdl-converter
```

### 2. Set up the backend

```bash
cd backend
pip install -r requirements.txt
python app.py
```

The backend runs on `http://localhost:5000`.

### 3. Open the frontend

Open `frontend/index.html` in your browser (no build step needed).

Or serve it with Python:

```bash
cd frontend
python -m http.server 8080
```

Then go to `http://localhost:8080`.

## Usage

1. Open the frontend in your browser
2. Make sure the API Server URL points to your backend (`http://localhost:5000`)
3. Drag & drop or click to upload your NSDL PDF
4. Click **Convert to Excel**
5. Preview the data, then click **Download Excel File**

## Project Structure

```
nsdl-converter/
├── backend/
│   ├── app.py           # Flask API server
│   ├── parser.py        # PDF parsing & Excel generation logic
│   └── requirements.txt
├── frontend/
│   └── index.html       # Single-file frontend (no dependencies)
└── README.md
```

## API Endpoints

### `POST /convert`
Upload a PDF and receive the Excel file as a download.

**Request:** `multipart/form-data` with `file` field (PDF)  
**Response:** `.xlsx` file download

### `POST /preview`
Upload a PDF and get a JSON preview of the first 50 rows.

**Response:**
```json
{
  "client_name": "TARA TULSI SHAH KELTON",
  "total_records": 4821,
  "unique_securities": 87,
  "preview": [...]
}
```

### `GET /health`
Health check — returns `{"status": "ok"}`.

## Requirements

- Python 3.8+
- pdfplumber
- openpyxl
- flask
- flask-cors

## Notes

- Works with NSDL transaction statements from ICICI Bank
- Handles multi-beneficiary entries (Beneficiary / Beneficiary-Blocked)
- Supports statements with thousands of pages
- Credit = "By" transactions, Debit = "To" transactions

## License

MIT
