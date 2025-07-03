import sqlite3
import os
from datetime import datetime
import pandas as pd

try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
except ImportError:
    gspread = None

DB_FILE = 'elisa.db'
EXCEL_FILE = 'elisa.xlsx'
GOOGLE_CREDENTIALS = 'credentials.json'
GOOGLE_SHEET_NAME = 'ElisaData'


def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS plates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS wells (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            plate_id INTEGER,
            well TEXT,
            sample TEXT,
            value REAL,
            category TEXT,
            FOREIGN KEY(plate_id) REFERENCES plates(id)
        )
    ''')
    # Ensure "category" column exists when upgrading from older versions
    cursor.execute("PRAGMA table_info(wells)")
    cols = [c[1] for c in cursor.fetchall()]
    if 'category' not in cols:
        cursor.execute('ALTER TABLE wells ADD COLUMN category TEXT')
    conn.commit()
    conn.close()


def get_gsheet_client():
    if gspread is None:
        raise RuntimeError('gspread is not installed')
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS, scope)
    client = gspread.authorize(creds)
    return client


def add_plate():
    name = input('Plate name: ').strip()
    rows = []
    print('Enter well data. Leave well blank to finish.')
    while True:
        well = input('Well: ').strip()
        if not well:
            break
        sample = input('Sample label: ').strip()
        value = input('Value: ').strip()
        try:
            value = float(value)
        except ValueError:
            print('Invalid value, skipping well')
            continue
        rows.append({'well': well, 'sample': sample, 'value': value})

    if not rows:
        print('No data entered.')
        return

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('INSERT INTO plates (name) VALUES (?)', (name,))
    plate_id = cursor.lastrowid
    cursor.executemany(
        'INSERT INTO wells (plate_id, well, sample, value) VALUES (?, ?, ?, ?)',
        [(plate_id, r['well'], r['sample'], r['value']) for r in rows]
    )
    conn.commit()
    conn.close()

    df = pd.DataFrame(rows)
    df.insert(0, 'plate', name)
    if os.path.exists(EXCEL_FILE):
        with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='new', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=name, index=False)
    else:
        df.to_excel(EXCEL_FILE, sheet_name=name, index=False)
    print(f'Plate {name} saved to {EXCEL_FILE} and {DB_FILE}.')

    if gspread:
        try:
            client = get_gsheet_client()
            sheet = client.open(GOOGLE_SHEET_NAME)
            ws = sheet.add_worksheet(title=name, rows=str(len(df)+1), cols=str(len(df.columns)))
            ws.update([df.columns.tolist()] + df.values.tolist())
            print(f'Plate {name} uploaded to Google Sheets ({GOOGLE_SHEET_NAME}).')
        except Exception as e:
            print('Google Sheets upload failed:', e)
    else:
        print('gspread not installed; skipping Google Sheets upload.')


def fetch_local():
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query('SELECT plates.name as plate, wells.well, wells.sample, wells.value, wells.category '
                           'FROM wells JOIN plates ON wells.plate_id = plates.id', conn)
    conn.close()
    print(df)


def fetch_online():
    if not gspread:
        print('gspread not installed.')
        return
    try:
        client = get_gsheet_client()
        sheet = client.open(GOOGLE_SHEET_NAME)
    except Exception as e:
        print('Unable to open Google Sheet:', e)
        return
    for ws in sheet.worksheets():
        data = ws.get_all_records()
        df = pd.DataFrame(data)
        print(f'Worksheet {ws.title}:')
        print(df)


if __name__ == '__main__':
    import argparse
    parser = argparse.ArgumentParser(description='Manage ELISA results.')
    parser.add_argument('--add', action='store_true', help='Add a new plate')
    parser.add_argument('--fetch-local', action='store_true', help='Show local results')
    parser.add_argument('--fetch-online', action='store_true', help='Show results from Google Sheets')
    args = parser.parse_args()

    init_db()

    if args.add:
        add_plate()
    elif args.fetch_local:
        fetch_local()
    elif args.fetch_online:
        fetch_online()
    else:
        parser.print_help()
