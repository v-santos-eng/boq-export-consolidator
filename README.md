# BoQ Export Consolidator

Turns a folder of bill-of-quantities CSV exports, one file per sales order or quotation, into a single Excel workbook that Power BI can read directly.

## The problem

A configurator writes one CSV per sales order or quotation. Each file lists the modules sold and, underneath each module, its characteristics as key/value pairs. That layout suits the ERP import it was designed for, and it is useless for analysis: answering "how many 24 V modules did we quote last quarter" meant opening files one at a time.

The goal was the same data as a flat table, one row per module, refreshed without re-reading thousands of files each time.

## What it does

- Reads every CSV in one or more source folders.
- Classifies each file as a sales order or a quotation from the prefix of its document number.
- Flattens each module and its characteristics into one row, using a column mapping kept in `config.json`.
- Appends only documents that are not already in the workbook, so a daily run processes just the new files.
- Writes two sheets, `Sales Orders` and `Quotations`, formatted as Excel tables so Power BI picks up new rows on refresh.

Files that are not exports are skipped; unreadable files are reported and skipped.

## Stack

Python 3.10+ · openpyxl · pytest

## Running it

```bash
pip install -r requirements.txt
python make_sample_data.py              # writes 20 synthetic exports to sample_data/
python consolidate.py sample_data       # creates output/database.xlsx
python consolidate.py sample_data       # second run adds 0 rows
pytest
```

Options: `python consolidate.py FOLDER [FOLDER ...] -o path/to/workbook.xlsx -c path/to/config.json`

## Input layout

Semicolon-separated, Latin-1. The customer name is the part of the file name after the last underscore.

```
5000100;BoQ export
;;MOD-R100;4;;;;;12.03.2025        <- module line: code, quantity, delivery date
;;;;LEN_MM;1200                    <- characteristics of the module above
;;;;DRIVE_VOLTAGE;24
;;MOD-C90;1;;;;;12.03.2025
;;;;LEN_MM;600
```

To add a column, add a `"Column name": "CHARACTERISTIC_KEY"` pair to `config.json`. Document prefixes, the module prefix, delimiter and encoding are set there too.

## About the data

This is a generic rewrite of a tool I built at work. The sample files, product codes, characteristic keys, document numbers and customer names here are all invented; nothing from an employer or customer is included.

## License

MIT. See [LICENSE](LICENSE).
