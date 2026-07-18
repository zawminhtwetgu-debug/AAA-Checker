# AAA Checker

A desktop application for checking AAA (Authentication, Authorization, Accounting) user statuses. Built with Python, Tkinter, and Playwright.

## Features

- **Bulk Username Checking** — Enter multiple usernames and check their status in parallel.
- **Login Status Detection** — Identifies whether users are Online or Offline.
- **Network Port Information** — Retrieves Port ID, Site ID, and RX Power data.
- **Account Status** — Shows Active Date, Suspend Date, and current status.
- **Search & Filter** — Real-time filtering of results.
- **Sortable Columns** — Click any column header to sort.
- **Excel Export** — Export results to a formatted `.xlsx` file.

## Requirements

- Python 3.8+
- Playwright (Chromium)
- pandas
- openpyxl

## Installation

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

## Usage

```bash
python AAA_Active.py
```

Enter usernames (one per line) in the text area, then click **Start Check**.

## Author

Created by **Zaw Min Htwe**
