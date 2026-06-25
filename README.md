# Tok Tenging

Tok Tenging is a Windows desktop tool for preparing Icelandic accounting input files and automating entry into Tok.

## What It Does

- Formats Tok input spreadsheets into the required column layout.
- Detects and converts bank exports from Arion, Landsbankinn, and Islandsbanki.
- Applies auto-coding suggestions from local training data.
- Reviews low-confidence coding clusters before export.
- Compresses matching transaction groups into summarized entries.

## Project Layout

- `tok/TokTenging.py` - application shell, navigation, settings, and Tok automation flow.
- `tok/bank_formatter.py` - bank formatter workflow, review dialogs, and transaction compression.
- `tok/bank_detection.py` - bank format definitions and Excel format detection.
- `tok/bank_utils.py` - date, amount, Excel, and identifier formatting helpers.
- `tok/ui_components.py` - shared Tkinter theme and reusable UI widgets.
- `tok/auto_coder.py` - transaction auto-coding engine.

## Run Locally

```powershell
pip install -r requirements.txt
python tok\TokTenging.py
```

The app expects private training files to live outside Git history or in ignored local folders.
