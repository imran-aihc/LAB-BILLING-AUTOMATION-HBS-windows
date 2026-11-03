# HBS Lab Bill Automation — Windows Setup

This project has TWO parts that must both run at the same time:
- `backend`  → Python (Flask + Playwright automation)
- `frontend` → React (the web page you see in the browser)

---

## First-time setup (only do this once)

### 1. Install Python (if not already installed)
Download from python.org → check "Add Python to PATH" during install.

### 2. Install Node.js (if not already installed)
Download from nodejs.org → install the LTS version.

### 3. Set up the backend
Open **Command Prompt** (or PowerShell) in the `backend` folder and run:

```
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
playwright install
```

### 4. Set up the frontend
Open a **second** Command Prompt window in the `frontend` folder and run:

```
npm install
```

---

## Running the project (every time after setup)

### Easiest way — double-click `START.bat`
This automatically opens both the backend and frontend for you.

### Manual way — two separate terminals

**Terminal 1 — Backend:**
```
cd backend
venv\Scripts\activate
python server.py
```

**Terminal 2 — Frontend:**
```
cd frontend
npm start
```

Then open your browser at:
```
http://localhost:3000
```

---

## What changed in `automation.py` (Windows update)

1. **Clicks "Unpaid"** instead of "In Patient Invoice" after pressing Print.
2. **Much more accurate lab test extraction** — scrolls fully top-to-bottom
   before reading the table, and has a flexible row parser as backup so
   tests are no longer skipped.
3. **Duplicate lab tests are merged** — if the same test appears twice on
   an invoice, it now shows as one row with Qty = 2 (or more) instead of
   two separate rows.
4. **Correct patient matching when names repeat** — if two patients share
   the same name, the script now checks the Admission Number on screen
   before accepting a match, retrying up to 3 times if needed.
5. **Windows event loop fix** — added the correct asyncio event loop
   policy so Playwright's browser can launch properly on Windows
   (this is required only on Windows, Mac/Linux are unaffected).
6. **Windows-safe filenames** — patient names with special characters
   (like `\ / : * ? " < > |`) are automatically cleaned before being
   used in a PDF filename, since Windows does not allow those characters
   in file names.

---

## If something goes wrong

**"python is not recognized"**
→ Python was not added to PATH during install. Reinstall Python and
   check the "Add Python to PATH" box.

**"Failed building wheel for greenlet" or similar build errors**
→ This usually only happens with very new Python versions (3.13+).
   If you hit this, install Python 3.11 or 3.12 instead and recreate
   the virtual environment.

**Browser does not open / Playwright error**
→ Run `playwright install` again inside the activated venv.

**Port already in use**
→ Close any other terminal windows running `python server.py` or
   `npm start`, then try again.
