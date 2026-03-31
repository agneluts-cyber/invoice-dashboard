#!/usr/bin/env python3
"""
Scrapes invoice data from Bolt Market Store Hub and updates the CSV + dashboard.
Runs locally (LaunchAgent) or in the cloud (GitHub Actions).

Usage:
    python3 scrape_invoices.py              # scrape and update CSV + dashboard
    python3 scrape_invoices.py --preview    # scrape, screenshot, show columns (no CSV write)
    python3 scrape_invoices.py --login      # manual login to save session (local only)
"""

import csv
import os
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

IS_CI = bool(os.environ.get('GITHUB_ACTIONS'))

SCRIPT_DIR = Path(__file__).parent.resolve()
ENV_PATH = SCRIPT_DIR / '.env'
CSV_PATH = SCRIPT_DIR / 'invoices overview - Sheet1.csv'
DASHBOARD_PY = SCRIPT_DIR / 'dashboard.py'
SCREENSHOT_DIR = SCRIPT_DIR / 'screenshots'
LOG_PATH = SCRIPT_DIR / 'scrape.log'
SESSION_PATH = SCRIPT_DIR / 'bolt_session.json'

BASE_URL = 'https://wms.bolt.eu'
LOGIN_URL = BASE_URL
INVOICES_FIRST_PAGE = BASE_URL + '/store/24/invoices'
INVOICES_NEXT_PAGE = BASE_URL + '/store/24/invoices?page={page}'

CSV_HEADER = [
    'Received', 'Inv. Date', 'Due', 'Invoice #', 'Supplier/PO #',
    'DN#/Delivery', 'Invoice Total+VAT', 'Delivery Total+VAT',
    'Credit Total+VAT', 'Del.-Inv. Diff', 'Post-Cred. Diff',
    'Remark', 'Payment status'
]

# Maps Bolt table header text (lowercase) to CSV column name.
# Adjust these if the Bolt page uses different header names.
COLUMN_MAP = {
    'received': 'Received',
    'received date': 'Received',
    'invoice date': 'Inv. Date',
    'inv. date': 'Inv. Date',
    'due': 'Due',
    'due date': 'Due',
    'invoice': 'Invoice #',
    'invoice #': 'Invoice #',
    'invoice number': 'Invoice #',
    'supplier': 'Supplier/PO #',
    'supplier/po': 'Supplier/PO #',
    'supplier/po #': 'Supplier/PO #',
    'dn': 'DN#/Delivery',
    'dn#/delivery': 'DN#/Delivery',
    'dn#': 'DN#/Delivery',
    'delivery': 'DN#/Delivery',
    'invoice total': 'Invoice Total+VAT',
    'invoice total+vat': 'Invoice Total+VAT',
    'delivery total': 'Delivery Total+VAT',
    'delivery total+vat': 'Delivery Total+VAT',
    'credit total': 'Credit Total+VAT',
    'credit total+vat': 'Credit Total+VAT',
    'del.-inv. diff': 'Del.-Inv. Diff',
    'difference': 'Del.-Inv. Diff',
    'diff': 'Del.-Inv. Diff',
    'post-cred. diff': 'Post-Cred. Diff',
    'remark': 'Remark',
    'remarks': 'Remark',
    'payment status': 'Payment status',
    'status': 'Payment status',
}


def log(msg):
    ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{ts}] {msg}'
    print(line)
    with open(LOG_PATH, 'a') as f:
        f.write(line + '\n')


def login(page, email, password):
    log('Navigating to login page...')
    page.goto(LOGIN_URL, wait_until='domcontentloaded', timeout=45000)
    page.wait_for_timeout(3000)

    log('Waiting for login form...')
    page.wait_for_selector('input[type="password"]', timeout=15000)

    log('Filling credentials...')
    username_field = page.locator('input[type="text"], input[name="username"], input[name="email"], input[placeholder*="mail"], input[placeholder*="user"]').first
    password_field = page.locator('input[type="password"]').first

    username_field.click()
    username_field.fill(email)
    page.wait_for_timeout(500)
    password_field.click()
    password_field.fill(password)
    page.wait_for_timeout(500)

    log('Clicking Log In...')
    login_btn = page.locator('button:has-text("Log In"), button:has-text("Login"), button[type="submit"]').first
    login_btn.click()

    page.wait_for_timeout(5000)
    try:
        page.wait_for_load_state('domcontentloaded', timeout=15000)
    except Exception:
        pass
    log('Login submitted, waiting for redirect...')


def detect_table_columns(page):
    """Read the table headers from the current page and return ordered list."""
    headers = page.locator('table thead th, table thead td, [role="columnheader"]')
    count = headers.count()
    if count == 0:
        headers = page.locator('table tr:first-child th, table tr:first-child td')
        count = headers.count()
    cols = []
    for i in range(count):
        text = headers.nth(i).inner_text().strip()
        cols.append(text)
    return cols


def map_columns(bolt_headers):
    """Map detected Bolt column names to CSV column names. Returns index mapping."""
    mapping = {}
    for i, h in enumerate(bolt_headers):
        key = h.lower().strip()
        if key in COLUMN_MAP:
            mapping[COLUMN_MAP[key]] = i
        else:
            for map_key, csv_col in COLUMN_MAP.items():
                if map_key in key or key in map_key:
                    if csv_col not in mapping:
                        mapping[csv_col] = i
                        break
    return mapping


def clean_cell(text):
    """Join multiline cell values into a single line (e.g. '15:09\\n16/03/2026' -> '15:0916/03/2026')."""
    return text.replace('\n', '').replace('\r', '').strip()


def scrape_table_rows(page):
    """Extract all data rows from the table on the current page."""
    rows = page.locator('table tbody tr')
    count = rows.count()
    data = []
    for i in range(count):
        cells = rows.nth(i).locator('td')
        cell_count = cells.count()
        row = []
        for j in range(cell_count):
            row.append(clean_cell(cells.nth(j).inner_text()))
        if any(row):
            data.append(row)
    return data


def scrape_all_pages(page):
    """Scrape all pages of invoices. Returns (bolt_headers, all_rows)."""
    bolt_headers = None
    all_rows = []
    page_index = 0
    prev_first_row = None

    while True:
        if page_index == 0:
            url = INVOICES_FIRST_PAGE
        else:
            url = INVOICES_NEXT_PAGE.format(page=page_index)

        log(f'Loading page {page_index + 1}: {url}')
        page.goto(url, wait_until='domcontentloaded', timeout=45000)
        page.wait_for_timeout(3000)

        if bolt_headers is None:
            bolt_headers = detect_table_columns(page)
            log(f'Detected columns: {bolt_headers}')

        rows = scrape_table_rows(page)
        log(f'Page {page_index + 1}: {len(rows)} rows')

        if not rows:
            log(f'Page {page_index + 1} is empty, stopping.')
            break

        first_row = rows[0] if rows else None
        if prev_first_row and first_row == prev_first_row:
            log(f'Page {page_index + 1} repeats previous data, stopping.')
            break
        prev_first_row = first_row

        all_rows.extend(rows)
        page_index += 1

        if page_index > 50:
            log('Safety limit: stopped at 50 pages.')
            break

    log(f'Total rows scraped across {page_index} page(s): {len(all_rows)}')
    return bolt_headers, all_rows


def build_csv_rows(bolt_headers, raw_rows, col_mapping):
    """Convert scraped rows into CSV format using column mapping."""
    csv_rows = []
    for raw in raw_rows:
        row = {}
        for csv_col in CSV_HEADER:
            if csv_col in col_mapping and col_mapping[csv_col] < len(raw):
                row[csv_col] = raw[col_mapping[csv_col]]
            else:
                row[csv_col] = '-'
        csv_rows.append(row)
    return csv_rows


def write_csv(csv_rows):
    with open(CSV_PATH, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADER)
        writer.writeheader()
        writer.writerows(csv_rows)
    log(f'CSV written: {CSV_PATH} ({len(csv_rows)} rows)')


def fetch_tracker_sheet():
    """Download latest discrepancy tracker from Google Sheets."""
    import urllib.request
    url = (
        'https://docs.google.com/spreadsheets/d/'
        '17-4ObBdkyw9dRi_a4psZJ-p3rm2YnvIZTY922y7mPsY'
        '/export?format=csv&gid=1114317169'
    )
    dest = SCRIPT_DIR / '🇵🇹 Invoice Discrepancy Tracker [PT-2026] - 🔎 Tracker - Discrepancies.csv'
    try:
        urllib.request.urlretrieve(url, str(dest))
        log(f'Tracker sheet downloaded ({dest.name})')
    except Exception as e:
        log(f'Tracker sheet download failed (will use cached version): {e}')


def regenerate_dashboard():
    log('Regenerating dashboard...')
    import subprocess
    result = subprocess.run(
        ['python3', str(DASHBOARD_PY)],
        capture_output=True, text=True, cwd=str(SCRIPT_DIR)
    )
    if result.returncode == 0:
        log('Dashboard regenerated successfully.')
        log(result.stdout.strip())
    else:
        log(f'Dashboard generation failed: {result.stderr}')


RCLONE_BIN = Path.home() / 'bin' / 'rclone'
GDRIVE_FOLDER_ID = '14a3bEP26ixPLmqQ16XtZZ3SbYlpazZ5G'
NETLIFY_SITE_ID_FILE = SCRIPT_DIR / '.netlify_site_id'
HISTORY_PATH = SCRIPT_DIR / 'invoice_history.json'


def upload_to_google_drive():
    """Upload dashboard.html to Google Drive using rclone (local only)."""
    if IS_CI:
        return
    import subprocess

    src = SCRIPT_DIR / 'dashboard.html'
    if not src.exists():
        log('Dashboard file not found — skipping Drive upload.')
        return

    if not RCLONE_BIN.exists():
        log(f'rclone not found at {RCLONE_BIN} — skipping Drive upload.')
        return

    today = datetime.now().strftime('%Y-%m-%d')
    dated_name = f'Daily_Invoice_Overview_{today}.html'
    dated_path = SCRIPT_DIR / dated_name

    import shutil
    shutil.copy2(str(src), str(dated_path))

    try:
        result = subprocess.run([
            str(RCLONE_BIN), 'copy', str(dated_path),
            f'gdrive:',
            f'--drive-root-folder-id={GDRIVE_FOLDER_ID}',
        ], capture_output=True, text=True, timeout=60)

        if result.returncode == 0:
            log(f'Dashboard uploaded to Google Drive as {dated_name}')
        else:
            log(f'Drive upload failed: {result.stderr}')
    except Exception as e:
        log(f'Drive upload error: {e}')
    finally:
        dated_path.unlink(missing_ok=True)


def save_daily_history(total_count):
    """Save today's invoice count to a history file for day-over-day comparison."""
    import json
    history = {}
    if HISTORY_PATH.exists():
        try:
            history = json.loads(HISTORY_PATH.read_text())
        except Exception:
            history = {}
    today_key = datetime.now().strftime('%Y-%m-%d')
    history[today_key] = total_count
    HISTORY_PATH.write_text(json.dumps(history, indent=2))
    log(f'History updated: {today_key} = {total_count} invoices')


def deploy_to_netlify():
    """Deploy dashboard.html to Netlify for instant web viewing."""
    import json, zipfile

    src = SCRIPT_DIR / 'dashboard.html'
    if not src.exists():
        log('Dashboard file not found — skipping Netlify deploy.')
        return

    site_id = os.environ.get('NETLIFY_SITE_ID', '').strip()
    token = os.environ.get('NETLIFY_TOKEN', '').strip()

    if not site_id and NETLIFY_SITE_ID_FILE.exists():
        site_id = NETLIFY_SITE_ID_FILE.read_text().strip()

    if not token:
        netlify_cfg = Path.home() / 'Library' / 'Preferences' / 'netlify' / 'config.json'
        if netlify_cfg.exists():
            try:
                with open(netlify_cfg) as f:
                    cfg = json.load(f)
                for uid, udata in cfg.get('users', {}).items():
                    token = udata.get('auth', {}).get('token', '')
                    if token:
                        break
            except Exception:
                pass

    if not site_id:
        log('Netlify site ID not found — skipping web deploy.')
        return
    if not token:
        log('Netlify token not found — skipping web deploy.')
        return

    try:
        zip_path = SCRIPT_DIR / '_deploy.zip'
        with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as z:
            z.write(str(src), 'index.html')
            z.writestr('_headers', '/\n  Content-Type: text/html; charset=UTF-8\n/*\n  Content-Type: text/html; charset=UTF-8\n')

        import urllib.request
        req = urllib.request.Request(
            f'https://api.netlify.com/api/v1/sites/{site_id}/deploys',
            data=zip_path.read_bytes(),
            headers={
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/zip',
            },
            method='POST'
        )
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read())
            url = result.get('ssl_url') or result.get('url', '?')
            log(f'Dashboard deployed to web: {url}')

        zip_path.unlink(missing_ok=True)
    except Exception as e:
        log(f'Netlify deploy error: {e}')


def is_logged_in(page):
    """Check if we're on the invoices page (logged in) or on the login page."""
    page.goto(INVOICES_FIRST_PAGE, wait_until='domcontentloaded', timeout=45000)
    page.wait_for_timeout(4000)
    url = page.url.lower()
    if 'invoices' in url and 'login' not in url:
        has_table = page.locator('table').count() > 0
        if has_table:
            return True
    return False


def manual_login(page, context):
    """Open login page for the user to log in manually, then save the session."""
    log('Session expired or missing. Opening browser for manual login...')
    log('Please log in to Bolt in the browser window. DO NOT close the browser.')
    try:
        page.goto(LOGIN_URL, wait_until='domcontentloaded', timeout=45000)
    except Exception:
        pass
    page.wait_for_timeout(3000)

    for attempt in range(120):
        try:
            page.wait_for_timeout(5000)
            url = page.url.lower()
            if 'invoices' in url or ('store' in url and 'login' not in url):
                has_table = page.locator('table').count() > 0
                if has_table:
                    log('Login detected! Saving session...')
                    context.storage_state(path=str(SESSION_PATH))
                    log(f'Session saved to {SESSION_PATH.name}')
                    return True
        except Exception as e:
            if 'closed' in str(e).lower():
                log('ERROR: Browser was closed before login completed. Please try again and keep the browser open.')
                return False
    log('ERROR: Timed out waiting for manual login (10 minutes).')
    return False


def main():
    preview_mode = '--preview' in sys.argv
    login_mode = '--login' in sys.argv

    if not IS_CI:
        try:
            from dotenv import load_dotenv
            load_dotenv(ENV_PATH)
        except ImportError:
            pass

    email = os.environ.get('BOLT_EMAIL', '')
    password = os.environ.get('BOLT_PASSWORD', '')

    SCREENSHOT_DIR.mkdir(exist_ok=True)

    with sync_playwright() as p:
        launch_headless = IS_CI or (not login_mode and not preview_mode and email and password)
        browser = p.chromium.launch(
            headless=launch_headless,
            args=['--disable-blink-features=AutomationControlled',
                  '--no-sandbox', '--disable-dev-shm-usage']
        )

        use_session = (not IS_CI) and SESSION_PATH.exists()
        storage = str(SESSION_PATH) if use_session else None
        context = browser.new_context(
            viewport={'width': 1400, 'height': 900},
            user_agent='Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            locale='pt-PT',
            timezone_id='Europe/Tallinn',
            storage_state=storage,
        )
        page = context.new_page()

        try:
            if login_mode:
                if not manual_login(page, context):
                    sys.exit(1)
                log('Login saved! You can now run the script without --login.')
                browser.close()
                return

            if IS_CI:
                if not email or not password:
                    log('ERROR: BOLT_EMAIL and BOLT_PASSWORD must be set in GitHub Secrets.')
                    sys.exit(1)
                log('Cloud mode: logging in with credentials...')
                login(page, email, password)
            else:
                logged_in = storage and is_logged_in(page)
                if not logged_in:
                    if email and password and 'example' not in email:
                        log('Session missing/expired. Logging in with credentials...')
                        login(page, email, password)
                        context.storage_state(path=str(SESSION_PATH))
                        log('Session saved for future runs.')
                    else:
                        if not manual_login(page, context):
                            sys.exit(1)

            page.screenshot(path=str(SCREENSHOT_DIR / 'after_login.png'))
            log(f'Screenshot saved: {SCREENSHOT_DIR}/after_login.png')

            bolt_headers, all_rows = scrape_all_pages(page)

            page.screenshot(path=str(SCREENSHOT_DIR / 'invoices_table.png'))

            if not bolt_headers:
                log('ERROR: Could not detect table columns. Check screenshots/ folder.')
                page.screenshot(path=str(SCREENSHOT_DIR / 'error_no_table.png'))
                sys.exit(1)

            col_mapping = map_columns(bolt_headers)
            log(f'Column mapping: { {k: bolt_headers[v] for k, v in col_mapping.items()} }')

            unmapped = [h for h in bolt_headers if h.lower().strip() not in COLUMN_MAP
                        and not any(h.lower().strip() in k or k in h.lower().strip() for k in COLUMN_MAP)]
            if unmapped:
                log(f'WARNING: Unmapped Bolt columns: {unmapped}')

            missing = [c for c in CSV_HEADER if c not in col_mapping]
            if missing:
                log(f'NOTE: CSV columns with no Bolt match (will be "-"): {missing}')

            if preview_mode:
                log('\n=== PREVIEW MODE ===')
                log(f'Bolt headers: {bolt_headers}')
                log(f'Mapped to CSV: { {k: bolt_headers[v] for k, v in col_mapping.items()} }')
                log(f'Total rows: {len(all_rows)}')
                if all_rows:
                    log(f'Sample row: {all_rows[0]}')
                log('Screenshots saved to screenshots/ folder.')
                log('Review the mapping above. If correct, run without --preview to update the CSV.')
            else:
                csv_rows = build_csv_rows(bolt_headers, all_rows, col_mapping)
                write_csv(csv_rows)
                save_daily_history(len(csv_rows))
                fetch_tracker_sheet()
                regenerate_dashboard()
                upload_to_google_drive()

                total = len(csv_rows)
                from dashboard import load_invoices, compute_metrics
                dash_rows = load_invoices(str(CSV_PATH))
                dm = compute_metrics(dash_rows)
                log(f'Done! {total} invoices. Overdue: {dm["overdue_count"]}, >7d: {dm["over7d_count"]}')

                if not IS_CI:
                    context.storage_state(path=str(SESSION_PATH))
                log('Dashboard updated and deployed.')

        except PwTimeout as e:
            log(f'TIMEOUT ERROR: {e}')
            try:
                page.screenshot(path=str(SCREENSHOT_DIR / 'error_timeout.png'))
            except Exception:
                pass
            sys.exit(1)
        except Exception as e:
            log(f'ERROR: {e}')
            try:
                page.screenshot(path=str(SCREENSHOT_DIR / 'error.png'))
            except Exception:
                pass
            raise
        finally:
            browser.close()


if __name__ == '__main__':
    main()
