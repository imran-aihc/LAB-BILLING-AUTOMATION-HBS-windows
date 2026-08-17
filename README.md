# HBS Lab Bill Automation System

Automates the complete 9-step monthly lab billing workflow on hbs.chealth.pk

---

## Exact Steps Automated

| Step | Action |
|------|--------|
| 1 | Read Patient Name (Col D) + Admin No (Col C) from your Excel |
| 2 | Auto-login → hbs.chealth.pk (user: 7786) |
| 3 | Click FRONTDESK in sidebar |
| 4 | Click Final Settlement under Sales |
| 5 | Click Patients button (top right) |
| 6 | Search patient by name → if multiple found → open each → match by Admission No |
| 7 | Check Payable → if ZERO → mark as Paid (added to Excel, no PDF download) |
| 8 | Click Print button |
| 9 | Click "In Patient Invoice" → save PDF as AdmissionNo.pdf |
| 10 | Claude AI extracts LabInvestigation tests → build Lab Bill Excel |

---

## Project Structure

```
hims-automation/
├── backend/
│   ├── automation.py     ← Full 9-step workflow + AI extraction + Excel builder
│   ├── server.py         ← Flask API (connects React frontend to Python)
│   ├── requirements.txt
│   └── .env.example      ← Add your Claude API key here
├── frontend/
│   ├── src/App.jsx       ← React UI (4 screens: upload → config → progress → download)
│   ├── src/index.js
│   ├── public/index.html
│   └── package.json
├── START.bat             ← Double-click to start on Windows
└── README.md
```

---

## One-Time Setup

### Backend
```bash
cd hims-automation/backend

# Create virtual environment
python -m venv venv

# Activate (Windows)
venv\Scripts\activate

# Install packages
pip install -r requirements.txt

# Install browser for automation
playwright install chromium

# Add your Claude API key
copy .env.example .env
# Open .env and paste your key from console.anthropic.com
```

### Frontend
```bash
cd hims-automation/frontend
npm install
```

---

## Every Month — How to Use

1. Double-click **START.bat**
2. Open **http://localhost:3000**
3. Upload your **May26.xlsx** (or this month's file)
4. Check the month label → click **Start Automation**
5. Watch live progress — system works automatically
6. Click **Download Lab Bill Excel** when done

---

## Excel Output Format (per patient block)

```
HBS GENERAL HOSPITAL
Laboratory Investigation Month of May 26
              Invoice
Patient Name: SHAREEN BIBI        Admission No. 44478
         Laboratory Investigation
Sr.No. | Investigation Name  | Rate  | Q | Amount
  1    | BLOOD C.P           |  400  | 1 |    400
  2    | HBsAG SCREENING     |  250  | 1 |    250
  3    | HCV AB SCREENING    |  250  | 1 |    250
  4    | LFT                 |  600  | 1 |    600
  5    | HISTHO LARGE        | 1,800 | 1 |  1,800
       | Sub Total Of Laboratory        |  3,300
                    CHECKED BY
       _______________________________
```

One block per patient, all in a single sheet. Plus a Summary sheet.

---

## Zero Payable Patients

If a patient's Payable amount is 0 (already paid / SSP / package):
- They are **included** in the Excel
- Marked with ⚠️ PAYABLE = 0 (PAID / SKIP) in red
- No PDF downloaded for them
- Sub Total shows 0

---

## Get Claude API Key (Free)

1. Go to https://console.anthropic.com
2. Sign up / login
3. Go to API Keys → Create Key
4. Paste it in `backend/.env` as: `ANTHROPIC_API_KEY=sk-ant-...`
