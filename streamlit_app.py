"""
AAA Checker Active - Streamlit Web App
Converted from Tkinter desktop app by Zaw Min Htwe
Run with: streamlit run streamlit_app.py
"""
import os
import streamlit as st
import pandas as pd
import threading
import time
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# Load variables from a local .env file if present (for local/dev runs).
# In production, set these as real environment variables instead.
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ─── Configuration ──────────────────────────────────────────────────────
LOGIN_URL = os.environ.get("AAA_LOGIN_URL", "http://10.201.1.160/metro/aaacheck.asp")
USERNAME = os.environ.get("AAA_USERNAME")
PASSWORD = os.environ.get("AAA_PASSWORD")

if not USERNAME or not PASSWORD:
    st.error(
        "⚠️ AAA_USERNAME and AAA_PASSWORD environment variables are not set. "
        "Set them before running this app (see README for instructions)."
    )
    st.stop()

COLUMNS = [
    "Username", "Login Status", "Login Fail", "Status",
    "Active Date", "Suspend Date", "Login Date", "Port RX Power",
    "Port ID", "Site ID", "RX Power Status"
]

# ─── Module-level state (thread-safe, NOT st.session_state) ─────────────
_results = []          # shared list of result dicts
_lock = threading.Lock()
_progress = {"checked": 0, "total": 0, "running": False, "stop": False}

# ─── Scraping functions (same as original) ──────────────────────────────
def scrape_user(username):
    data = {col: "" for col in COLUMNS}
    data["Username"] = username
    data["Login Fail"] = "Unknown error"
    data["Login Status"] = "Unknown"

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            context.route("**/*", lambda route, request: route.abort()
                          if request.resource_type in ["image", "stylesheet", "font"]
                          else route.continue_())
            page = context.new_page()
            page.goto(LOGIN_URL, timeout=15000)
            page.fill('input[name="username"]', USERNAME)
            page.fill('input[name="password"]', PASSWORD)
            page.click('input[type="submit"]')
            page.wait_for_selector('input[name="username"]', timeout=5000)
            page.fill('input[name="username"]', username)
            page.click('input[name="Check"]')

            try:
                page.wait_for_selector(
                    'img[src="images/online_1.png"], img[src="images/offline_1.png"]',
                    timeout=5000
                )
                if page.is_visible('img[src="images/online_1.png"]'):
                    data["Login Status"] = "Online"
                    data["Login Fail"] = ""
                elif page.is_visible('img[src="images/offline_1.png"]'):
                    data["Login Status"] = "Offline"
                    data["Login Fail"] = ""
            except PlaywrightTimeoutError:
                pass

            if page.is_visible("span.style15"):
                data["Login Fail"] = page.inner_text("span.style15").strip()

            try:
                data["Status"] = page.locator(
                    "xpath=/html/body/form/table/tbody/tr[13]/td[2]"
                ).inner_text().strip()
                data["Active Date"] = page.locator(
                    "xpath=/html/body/form/table/tbody/tr[14]/td[2]/strong"
                ).inner_text().strip()
                data["Suspend Date"] = page.locator(
                    "xpath=/html/body/form/table/tbody/tr[16]/td[2]"
                ).inner_text().strip()
            except Exception:
                pass

            for tag in page.query_selector_all("strong"):
                text = tag.inner_text().strip()
                if "OLT" in text and ":" in text:
                    data["Port ID"] = text
                    data["Site ID"] = text[:7]
                elif "/" in text and ":" in text:
                    data["Login Date"] = text

            if page.is_visible("td:has-text('OLT:')"):
                data["Port RX Power"] = page.inner_text("td:has-text('OLT:')")

            html = page.content()
            if "Rx power is not good" in html:
                data["RX Power Status"] = "Rx power is not good"
            elif "Rx power is good" in html:
                data["RX Power Status"] = "Rx power is good"

            browser.close()
            return data
    except Exception as e:
        print(f"Error for {username}: {e}")
        return data


def _run_batch(usernames):
    """Background worker — does NOT touch st.session_state."""
    global _results, _progress
    _progress["stop"] = False

    with ThreadPoolExecutor(max_workers=5) as executor:
        fut_map = {executor.submit(scrape_user, u): u for u in usernames}
        for f in as_completed(fut_map):
            if _progress["stop"]:
                break
            data = f.result()
            username = fut_map[f]
            with _lock:
                _results = [r for r in _results if r["Username"] != username]
                _results.append(data)
                _progress["checked"] += 1

    _progress["running"] = False


def generate_excel(results):
    df = pd.DataFrame(results)
    buffer = io.BytesIO()
    df.to_excel(buffer, index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb.active

    border = Border(
        left=Side(border_style="thin"), right=Side(border_style="thin"),
        top=Side(border_style="thin"), bottom=Side(border_style="thin"),
    )
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value and len(str(cell.value)) > max_len:
                    max_len = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max_len + 2

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ─── Page config ────────────────────────────────────────────────────────
st.set_page_config(page_title="AAA Checker Active", layout="wide")
st.markdown("""
    <style>
        .stApp { background-color: #f5f5f5; }
        .main-header {
            background: linear-gradient(90deg, #4287f5, #2b6edb);
            color: white; padding: 1.2rem 2rem; border-radius: 10px;
            margin-bottom: 1.5rem;
        }
        .main-header h1 { margin: 0; font-size: 1.8rem; }
        .main-header p { margin: 0; opacity: 0.9; font-size: 0.9rem; }
        div[data-testid="stButton"] button[kind="primary"] {
            background: #4CAF50; color: white; font-weight: bold;
        }
        .stProgress > div > div > div > div { background: #4287f5; }
        .status-msg { font-size: 1rem; padding: 0.5rem 0; }
        .block-container { padding-top: 1.5rem; }
    </style>
""", unsafe_allow_html=True)

# ─── UI ─────────────────────────────────────────────────────────────────
st.markdown('<div class="main-header"><h1>🔍 AAA Checker</h1>'
            '<p>Created by Zaw Min Htwe</p></div>', unsafe_allow_html=True)

# Status bar
status_placeholder = st.empty()
progress_placeholder = st.empty()

# Input section
with st.container(border=True):
    st.markdown("### 📋 Enter Usernames")
    usernames_text = st.text_area(
        "One username per line",
        height=120,
        label_visibility="collapsed",
        placeholder="User1\nUser2\nUser3\n..."
    )

# Buttons
col_btns = st.columns(5)
with col_btns[0]:
    start_clicked = st.button(
        "▶ Start Check", type="primary", use_container_width=True,
        disabled=_progress["running"]
    )
with col_btns[1]:
    stop_clicked = st.button(
        "⏹ Stop", use_container_width=True,
        disabled=not _progress["running"]
    )
with col_btns[2]:
    clear_clicked = st.button(
        "🗑 Clear", use_container_width=True,
        disabled=_progress["running"]
    )
with col_btns[3]:
    has_results = len(_results) > 0
    export_clicked = st.button(
        "📥 Export Excel", use_container_width=True,
        disabled=not has_results or _progress["running"]
    )

# ─── Handle actions ─────────────────────────────────────────────────────
if start_clicked:
    usernames = [u.strip() for u in usernames_text.splitlines() if u.strip()]
    if not usernames:
        st.warning("Please enter at least one username.")
    else:
        _progress["running"] = True
        _progress["checked"] = 0
        _progress["total"] = len(usernames)
        _results.clear()
        threading.Thread(target=_run_batch, args=(usernames,), daemon=True).start()
        st.rerun()

if stop_clicked:
    _progress["stop"] = True

if clear_clicked:
    _results.clear()
    _progress["checked"] = 0
    _progress["total"] = 0

if export_clicked and _results:
    excel_bytes = generate_excel(_results)
    st.download_button(
        label="💾 Download Excel File",
        data=excel_bytes,
        file_name="AAA_Checker_Results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ─── Progress & status ──────────────────────────────────────────────────
if _progress["running"]:
    pct = int((_progress["checked"] / _progress["total"]) * 100) if _progress["total"] else 0
    progress_placeholder.progress(
        pct / 100,
        text=f"⏳ Checking... {_progress['checked']}/{_progress['total']} ({pct}%)"
    )
    status_placeholder.info(
        f"🔄 Processing: {_progress['checked']}/{_progress['total']} users checked"
    )
    time.sleep(0.5)
    st.rerun()
elif _progress["checked"] > 0 and not _progress["running"]:
    progress_placeholder.progress(1.0, text=f"✅ Complete — {_progress['checked']} users checked")
    status_placeholder.success(f"✅ Completed! {_progress['checked']} users checked.")
elif _progress["stop"]:
    progress_placeholder.info("⏹ Stopped")
    status_placeholder.warning("⏹ Stopped by user")
else:
    progress_placeholder.empty()
    status_placeholder.info("Ready. Enter usernames and click **Start Check**.")

# ─── Results table ──────────────────────────────────────────────────────
if _results:
    st.markdown("### 📊 Results")
    df = pd.DataFrame(_results)
    if all(c in df.columns for c in COLUMNS):
        df = df[COLUMNS]

    st.dataframe(
        df,
        use_container_width=True,
        height=400,
        hide_index=True,
    )
    st.caption(f"Total: {len(_results)} users checked")
