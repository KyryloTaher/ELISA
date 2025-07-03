import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import os
import re
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


# Database initialization with category support
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute('''
        CREATE TABLE IF NOT EXISTS plates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    cur.execute('''
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
    cur.execute("PRAGMA table_info(wells)")
    cols = [c[1] for c in cur.fetchall()]
    if 'category' not in cols:
        cur.execute('ALTER TABLE wells ADD COLUMN category TEXT')
    conn.commit()
    conn.close()


def get_gsheet_client():
    if gspread is None:
        raise RuntimeError('gspread is not installed')
    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS, scope)
    client = gspread.authorize(creds)
    return client


def parse_table(text):
    lines = [l.strip() for l in text.strip().splitlines() if l.strip()]
    table = [re.split(r'\t|,|\s+', l) for l in lines]
    return table


def parse_wells(text):
    return set(w.strip().upper() for w in re.split(r'[\s,]+', text) if w.strip())


def save_plate(names_text, values_text, plate_name, kpos, kneg_sap, kneg_buf, blank,
               to_excel=False, to_google=False):
    name_table = parse_table(names_text)
    value_table = parse_table(values_text)
    if len(name_table) != len(value_table):
        messagebox.showerror('Error', 'Names and values table size mismatch')
        return

    wells = []
    for i, (n_row, v_row) in enumerate(zip(name_table, value_table)):
        for j, (n, v) in enumerate(zip(n_row, v_row)):
            well = f"{chr(65+i)}{j+1}"
            try:
                v = float(v)
            except ValueError:
                v = None
            wells.append({'well': well, 'sample': n, 'value': v})

    # assign categories
    cat_map = {}
    for w in parse_wells(kpos):
        cat_map[w] = 'K+'
    for w in parse_wells(kneg_sap):
        cat_map[w] = 'K- healthy'
    for w in parse_wells(kneg_buf):
        cat_map[w] = 'K- buffer'
    for w in parse_wells(blank):
        cat_map[w] = 'substrate blank'
    for w in wells:
        w['category'] = cat_map.get(w['well'], '')

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute('INSERT INTO plates (name) VALUES (?)', (plate_name,))
    pid = cur.lastrowid
    cur.executemany('INSERT INTO wells (plate_id, well, sample, value, category) '
                    'VALUES (?, ?, ?, ?, ?)',
                    [(pid, w['well'], w['sample'], w['value'], w['category']) for w in wells])
    conn.commit()
    conn.close()

    df = pd.DataFrame(wells)
    df.insert(0, 'plate', plate_name)

    if to_excel:
        if os.path.exists(EXCEL_FILE):
            with pd.ExcelWriter(EXCEL_FILE, mode='a', if_sheet_exists='new', engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=plate_name, index=False)
        else:
            df.to_excel(EXCEL_FILE, sheet_name=plate_name, index=False)

    if to_google and gspread:
        try:
            client = get_gsheet_client()
            sheet = client.open(GOOGLE_SHEET_NAME)
            ws = sheet.add_worksheet(title=plate_name, rows=str(len(df)+1), cols=str(len(df.columns)))
            ws.update([df.columns.tolist()] + df.values.tolist())
        except Exception as e:
            messagebox.showwarning('Google Sheets', f'Upload failed: {e}')

    messagebox.showinfo('Saved', f'Plate {plate_name} saved to database.')


def fetch_plate(plate_name):
    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query(
        'SELECT wells.well, wells.sample, wells.value, wells.category '
        'FROM wells JOIN plates ON wells.plate_id = plates.id WHERE plates.name=?',
        conn, params=(plate_name,))
    conn.close()
    return df


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ELISA Manager')
        self.geometry('800x600')
        self.build_ui()
        init_db()

    def build_ui(self):
        frm_top = ttk.Frame(self)
        frm_top.pack(fill='x', pady=5)
        ttk.Label(frm_top, text='Plate name:').pack(side='left')
        self.entry_plate = ttk.Entry(frm_top)
        self.entry_plate.pack(side='left', fill='x', expand=True, padx=5)

        frm_tables = ttk.Frame(self)
        frm_tables.pack(fill='both', expand=True)

        self.text_names = tk.Text(frm_tables, width=40, height=15)
        self.text_names.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        self.text_values = tk.Text(frm_tables, width=40, height=15)
        self.text_values.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        frm_cat = ttk.Frame(self)
        frm_cat.pack(fill='x', pady=5)
        ttk.Label(frm_cat, text='K+ wells:').grid(row=0, column=0, sticky='e')
        ttk.Label(frm_cat, text='K- healthy:').grid(row=1, column=0, sticky='e')
        ttk.Label(frm_cat, text='K- buffer:').grid(row=2, column=0, sticky='e')
        ttk.Label(frm_cat, text='Substrate blank:').grid(row=3, column=0, sticky='e')
        self.entry_kpos = ttk.Entry(frm_cat)
        self.entry_kneg_sap = ttk.Entry(frm_cat)
        self.entry_kneg_buf = ttk.Entry(frm_cat)
        self.entry_blank = ttk.Entry(frm_cat)
        self.entry_kpos.grid(row=0, column=1, sticky='ew', padx=2)
        self.entry_kneg_sap.grid(row=1, column=1, sticky='ew', padx=2)
        self.entry_kneg_buf.grid(row=2, column=1, sticky='ew', padx=2)
        self.entry_blank.grid(row=3, column=1, sticky='ew', padx=2)
        frm_cat.columnconfigure(1, weight=1)

        frm_opts = ttk.Frame(self)
        frm_opts.pack(fill='x', pady=5)
        self.var_excel = tk.BooleanVar(value=False)
        self.var_google = tk.BooleanVar(value=False)
        ttk.Checkbutton(frm_opts, text='Save to Excel', variable=self.var_excel).pack(side='left', padx=5)
        ttk.Checkbutton(frm_opts, text='Save to Google Sheets', variable=self.var_google).pack(side='left', padx=5)

        frm_btn = ttk.Frame(self)
        frm_btn.pack(fill='x', pady=5)
        ttk.Button(frm_btn, text='Save Plate', command=self.save).pack(side='left', padx=5)
        ttk.Button(frm_btn, text='Fetch Plate', command=self.fetch).pack(side='left', padx=5)

        self.text_output = tk.Text(self, height=10)
        self.text_output.pack(fill='both', expand=True, padx=5, pady=5)

    def save(self):
        save_plate(
            self.text_names.get('1.0', 'end'),
            self.text_values.get('1.0', 'end'),
            self.entry_plate.get().strip(),
            self.entry_kpos.get(),
            self.entry_kneg_sap.get(),
            self.entry_kneg_buf.get(),
            self.entry_blank.get(),
            self.var_excel.get(),
            self.var_google.get()
        )

    def fetch(self):
        plate = self.entry_plate.get().strip()
        if not plate:
            messagebox.showwarning('Plate', 'Enter plate name to fetch')
            return
        df = fetch_plate(plate)
        self.text_output.delete('1.0', 'end')
        if df.empty:
            self.text_output.insert('end', 'No data found\n')
        else:
            self.text_output.insert('end', df.to_string(index=False))


if __name__ == '__main__':
    App().mainloop()
