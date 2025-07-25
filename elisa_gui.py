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


def ensure_working_directory():
    """Ensure default working directory exists and switch to it."""
    target = os.path.join(os.path.expanduser('~'), 'projects', 'ELISA')
    os.makedirs(target, exist_ok=True)
    os.chdir(target)


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
            serum TEXT,
            normalized REAL,
            result TEXT,
            FOREIGN KEY(plate_id) REFERENCES plates(id)
        )
    ''')
    cur.execute("PRAGMA table_info(wells)")
    cols = [c[1] for c in cur.fetchall()]
    if 'category' not in cols:
        cur.execute('ALTER TABLE wells ADD COLUMN category TEXT')
    if 'serum' not in cols:
        cur.execute('ALTER TABLE wells ADD COLUMN serum TEXT')
    if 'normalized' not in cols:
        cur.execute('ALTER TABLE wells ADD COLUMN normalized REAL')
    if 'result' not in cols:
        cur.execute('ALTER TABLE wells ADD COLUMN result TEXT')
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


def save_plate_data(wells, plate_name, to_excel=False, to_google=False):
    if not wells:
        messagebox.showwarning('No data', 'Plate is empty')
        return
    if not plate_name:
        messagebox.showwarning('Plate', 'Plate name required')
        return

    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute('INSERT INTO plates (name) VALUES (?)', (plate_name,))
    pid = cur.lastrowid
    cur.executemany(
        'INSERT INTO wells (plate_id, well, sample, value, category, serum) '
        'VALUES (?, ?, ?, ?, ?, ?)',
        [(
            pid,
            w['well'],
            w['sample'],
            w['value'],
            w.get('category', ''),
            w.get('serum', '')
        ) for w in wells]
    )
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
        'SELECT wells.well, wells.sample, wells.value, wells.category, '
        'wells.serum, wells.normalized, wells.result '
        'FROM wells JOIN plates ON wells.plate_id = plates.id WHERE plates.name=?',
        conn, params=(plate_name,))
    conn.close()
    return df


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ELISA Manager')
        self.geometry('800x600')
        self.name_cells = {}
        self.value_cells = {}
        self.categories = {}
        self.serums = {}
        self.selected = set()
        self.cat_colors = {
            'K+': '#b6fcb6',
            'K- healthy': '#ffd79f',
            'K- buffer': '#ffb6c1',
            'substrate blank': '#e0e0e0',
        }
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

        lf_names = ttk.LabelFrame(frm_tables, text='Sample names')
        lf_names.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        lf_values = ttk.LabelFrame(frm_tables, text='Values')
        lf_values.pack(side='left', fill='both', expand=True, padx=5, pady=5)

        rows = [chr(65+i) for i in range(8)]
        cols = [str(i+1) for i in range(12)]

        for c, col in enumerate(cols):
            ttk.Label(lf_names, text=col).grid(row=0, column=c+1)
            ttk.Label(lf_values, text=col).grid(row=0, column=c+1)

        for r, row in enumerate(rows):
            ttk.Label(lf_names, text=row).grid(row=r+1, column=0)
            ttk.Label(lf_values, text=row).grid(row=r+1, column=0)
            for c in range(12):
                e_name = tk.Entry(lf_names, width=8)
                e_name.grid(row=r+1, column=c+1, padx=1, pady=1)
                e_value = tk.Entry(lf_values, width=8)
                e_value.grid(row=r+1, column=c+1, padx=1, pady=1)
                self.name_cells[(r, c)] = e_name
                self.value_cells[(r, c)] = e_value
                e_name.bind('<Button-1>', lambda e, rc=(r, c): self.toggle_select(rc))
                e_value.bind('<Button-1>', lambda e, rc=(r, c): self.toggle_select(rc))

        ttk.Button(lf_names, text='Paste', command=lambda: self.paste_clipboard('names')).grid(row=9, column=1, columnspan=12, sticky='ew')
        ttk.Button(lf_values, text='Paste', command=lambda: self.paste_clipboard('values')).grid(row=9, column=1, columnspan=12, sticky='ew')

        frm_cat = ttk.Frame(self)
        frm_cat.pack(fill='x', pady=5)

        labels = [
            ('K+', 'K+ wells:'),
            ('K- healthy', 'K- healthy:'),
            ('K- buffer', 'K- buffer:'),
            ('substrate blank', 'Substrate blank:')
        ]
        self.cat_entries = {}
        for i, (cat, text) in enumerate(labels):
            ttk.Label(frm_cat, text=text).grid(row=i, column=0, sticky='e')
            ent = ttk.Entry(frm_cat)
            ent.grid(row=i, column=1, sticky='ew', padx=2)
            ttk.Button(frm_cat, text='Set selected', command=lambda c=cat: self.assign_selected(c)).grid(row=i, column=2, padx=2)
            self.cat_entries[cat] = ent
        ttk.Label(frm_cat, text='Serum name:').grid(row=len(labels), column=0, sticky='e')
        self.entry_serum = ttk.Entry(frm_cat)
        self.entry_serum.grid(row=len(labels), column=1, sticky='ew', padx=2)
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
        ttk.Button(frm_btn, text='Clear Selection', command=self.clear_selection).pack(side='left', padx=5)

        frm_calc = ttk.Frame(self)
        frm_calc.pack(fill='x', pady=5)
        ttk.Label(frm_calc, text='Method:').pack(side='left')
        self.var_method = tk.StringVar(value='SD')
        ttk.Combobox(frm_calc, textvariable=self.var_method, values=['SD', 'Multiple'], width=10, state='readonly').pack(side='left', padx=2)
        ttk.Label(frm_calc, text='Multiplier:').pack(side='left')
        self.entry_multiplier = ttk.Entry(frm_calc, width=5)
        self.entry_multiplier.insert(0, '3')
        self.entry_multiplier.pack(side='left', padx=2)
        ttk.Button(frm_calc, text='Calculate Results', command=self.calculate_results).pack(side='left', padx=5)

        self.text_output = tk.Text(self, height=10)
        self.text_output.pack(fill='both', expand=True, padx=5, pady=5)

    def well_to_rc(self, well):
        m = re.match(r'^([A-H])(\d{1,2})$', well.upper())
        if not m:
            return None
        r = ord(m.group(1)) - 65
        c = int(m.group(2)) - 1
        if 0 <= r < 8 and 0 <= c < 12:
            return r, c
        return None

    def _update_cell_color(self, rc):
        color = self.cat_colors.get(self.categories.get(rc, ''), 'white')
        self.name_cells[rc].config(bg=color)
        self.value_cells[rc].config(bg=color)

    def toggle_select(self, rc):
        if rc in self.selected:
            self.selected.remove(rc)
            self._update_cell_color(rc)
        else:
            self.selected.add(rc)
            self.name_cells[rc].config(bg='cyan')
            self.value_cells[rc].config(bg='cyan')

    def assign_selected(self, cat):
        serum = self.entry_serum.get().strip()
        for rc in list(self.selected):
            self.categories[rc] = cat
            if serum:
                self.serums[rc] = serum
            self.selected.remove(rc)
            self._update_cell_color(rc)
        for w in parse_wells(self.cat_entries[cat].get()):
            rc = self.well_to_rc(w)
            if rc:
                self.categories[rc] = cat
                if serum:
                    self.serums[rc] = serum
                self._update_cell_color(rc)

    def clear_selection(self):
        for rc in list(self.selected):
            self.selected.remove(rc)
            self.serums.pop(rc, None)
            self._update_cell_color(rc)

    def assign_from_entries(self):
        serum = self.entry_serum.get().strip()
        for cat, ent in self.cat_entries.items():
            for w in parse_wells(ent.get()):
                rc = self.well_to_rc(w)
                if rc:
                    self.categories[rc] = cat
                    if serum:
                        self.serums[rc] = serum
                    self._update_cell_color(rc)

    def paste_clipboard(self, target):
        try:
            text = self.clipboard_get()
        except tk.TclError:
            return
        rows = text.strip().splitlines()
        for r, line in enumerate(rows):
            if r >= 8:
                break
            cells = re.split(r'\t', line)
            for c, cell in enumerate(cells):
                if c >= 12:
                    break
                widget = self.name_cells[(r, c)] if target == 'names' else self.value_cells[(r, c)]
                widget.delete(0, 'end')
                widget.insert(0, cell.strip())

    def collect_data(self):
        wells = []
        for r in range(8):
            for c in range(12):
                well = f"{chr(65+r)}{c+1}"
                name = self.name_cells[(r, c)].get().strip()
                val_text = self.value_cells[(r, c)].get().strip()
                try:
                    value = float(val_text) if val_text else None
                except ValueError:
                    value = None
                cat = self.categories.get((r, c), '')
                serum = self.serums.get((r, c), '')
                wells.append({'well': well, 'sample': name, 'value': value, 'category': cat, 'serum': serum})
        return wells

    def save(self):
        self.assign_from_entries()
        wells = self.collect_data()
        save_plate_data(
            wells,
            self.entry_plate.get().strip(),
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
        for rc in list(self.categories):
            self.categories.pop(rc, None)
            self._update_cell_color(rc)
        for r in range(8):
            for c in range(12):
                self.name_cells[(r, c)].delete(0, 'end')
                self.value_cells[(r, c)].delete(0, 'end')
        if df.empty:
            self.text_output.insert('end', 'No data found\n')
            return
        for _, row in df.iterrows():
            rc = self.well_to_rc(row['well'])
            if not rc:
                continue
            r, c = rc
            self.name_cells[(r, c)].insert(0, str(row['sample']))
            if row['value'] is not None:
                self.value_cells[(r, c)].insert(0, str(row['value']))
            if row['category']:
                self.categories[(r, c)] = row['category']
                self._update_cell_color((r, c))
            if row['serum']:
                self.serums[(r, c)] = row['serum']
        self.text_output.insert('end', df.to_string(index=False))
        if df.empty:
            self.text_output.insert('end', 'No data found\n')
        else:
            self.text_output.insert('end', df.to_string(index=False))

    def calculate_results(self):
        self.assign_from_entries()
        wells = self.collect_data()
        df = pd.DataFrame(wells)
        if df['value'].dropna().empty:
            messagebox.showwarning('Data', 'No numeric values to analyze')
            return

        blank_vals = df[df['category'] == 'substrate blank']['value'].dropna()
        blank = blank_vals.mean() if not blank_vals.empty else 0.0
        df['normalized'] = df['value'] - blank

        healthy = df[df['category'] == 'K- healthy']['normalized'].dropna()
        if healthy.empty:
            messagebox.showwarning('Controls', 'No healthy sap values provided')
            return
        mean_h = healthy.mean()
        try:
            mult = float(self.entry_multiplier.get())
        except ValueError:
            mult = 1.0
        if self.var_method.get() == 'SD':
            threshold = mean_h + mult * healthy.std()
        else:
            threshold = mean_h * mult
        df['result'] = df['normalized'].apply(lambda x: 'positive' if x > threshold else 'negative')

        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute('SELECT id FROM plates WHERE name=?', (self.entry_plate.get().strip(),))
        row = cur.fetchone()
        if row:
            pid = row[0]
            for _, r in df.iterrows():
                cur.execute(
                    'UPDATE wells SET serum=?, category=?, normalized=?, result=? '
                    'WHERE plate_id=? AND well=?',
                    (r.get('serum', ''), r['category'], r['normalized'], r['result'], pid, r['well'])
                )
            conn.commit()
        conn.close()

        self.text_output.delete('1.0', 'end')
        self.text_output.insert('end', df[["well", "sample", "serum", "normalized", "result"]].to_string(index=False))


if __name__ == '__main__':
    ensure_working_directory()
    App().mainloop()
