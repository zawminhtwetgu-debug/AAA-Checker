# AAA Checker

A desktop application for checking AAA (Authentication, Authorization, Accounting) user statuses. Built with Python, Tkinter, and Playwright.

![AAA Checker Screenshot](screenshot.png)

## Features

- **Bulk Username Checking** — Enter multiple usernames and check their status in parallel (5 concurrent workers).
- **Login Status Detection** — Identifies whether users are Online or Offline.
- **Network Port Information** — Retrieves Port ID, Site ID, and RX Power data.
- **Account Status** — Shows Active Date, Suspend Date, and current status.
- **Search & Filter** — Real-time filtering of results.
- **Sortable Columns** — Click any column header to sort.
- **Excel Export** — Export results to a formatted `.xlsx` file with auto-sized columns and borders.
- **Progress Tracking** — Real-time progress bar showing checked/total count.

---

## 📥 Quick Download (Recommended)

**No Python installation needed!** Just download and run:

1. Go to the [**Releases**](../../releases) page
2. Download **`AAA_Active.exe`**
3. Double-click to run — that's it!

> **Note:** Windows SmartScreen may show a warning since the app isn't code-signed. Click **"More info"** → **"Run anyway"** to proceed.

---

## 🛠️ Run from Source (For Developers)

If you prefer to run from source code:

### Requirements

- Python 3.8+
- Playwright (Chromium)

### Installation

```bash
# Clone the repository
git clone https://github.com/zawminhtwetgu-debug/AAA-Checker.git
cd AAA-Checker

# Install dependencies
pip install -r requirements.txt

# Install Chromium browser for Playwright
playwright install chromium
```

### Run

```bash
python AAA_Active.py
```

---

## 📖 How to Use

1. **Launch** the application (either the `.exe` or via `python AAA_Active.py`)
2. **Enter usernames** in the text area (one per line)
3. Click **Start Check** to begin checking
4. View results in the table — you can:
   - **Sort** by clicking column headers
   - **Search/Filter** results in real-time using the search box
5. Click **Export to Excel** to save results as a formatted `.xlsx` file
6. Click **Stop Check** to cancel in-progress checks
7. Click **Clear Result** to reset everything

### Result Columns

| Column | Description |
|--------|-------------|
| Username | The AAA username checked |
| Login Status | Online / Offline |
| Login Fail | Error message if login check failed |
| Status | Account status code |
| Active Date | When the account was activated |
| Suspend Date | When the account was suspended (if applicable) |
| Login Date | Last login date and time |
| Port RX Power | OLT port receive power reading |
| Port ID | OLT port identifier |
| Site ID | First 7 characters of Port ID |
| RX Power Status | Whether RX power is good or not |

---

## Author

Created by **Zaw Min Htwe**
