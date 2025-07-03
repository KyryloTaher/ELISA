# ELISA Manager

This project provides both a command line interface and a small graphical
application to store ELISA plate results. Data entered by the user are saved
locally in a SQLite database (`elisa.db`). They can also be exported to an Excel
workbook (`elisa.xlsx`) and optionally uploaded to a Google Sheets document for
online access.

## Requirements

* Python 3.12+
* Packages listed in `requirements.txt`

Install dependencies with:

```bash
pip install -r requirements.txt
```

To enable Google Sheets integration create a service account, download its credentials
JSON file and place it as `credentials.json` in the project directory. Create a Google
Sheet named `ElisaData` and share it with the service account email.

## Usage

Add a new plate:

```bash
python elisa_app.py --add
```

You will be asked for the plate name and data for each well. Press enter with an empty
well to finish.

Fetch data stored locally:

```bash
python elisa_app.py --fetch-local
```

Fetch data from the online Google Sheet:

```bash
python elisa_app.py --fetch-online
```

Both local and online results are printed to the console.

### Graphical interface

Launch the simple GUI with:

```bash
python elisa_gui.py
```

Paste two tables copied from Excel (sample names and corresponding values),
enter the control well positions and click **Save Plate**. Use the check boxes to
also store the plate to Excel or Google Sheets.
