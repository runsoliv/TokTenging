# Tok Tenging

Tok Tenging is a Windows desktop application for preparing Icelandic accounting
files, reviewing bank transactions, and automating entry into Tok.

## Features

- Detects and converts exports from Arion, Landsbankinn, and Íslandsbanki.
- Produces the column layout required by Tok while preserving accounting signs,
  dates, identifiers, and source references.
- Keeps the established rule/history-based autocoder as the default.
- Offers an optional ML-assisted mode that can only supplement unresolved rows
  after a model passes grouped validation and production safety gates.
- Reviews uncertain coding clusters before export.
- Supports automatic and manual transaction compression with repeat-safe
  reference history.
- Combines compatible coded workbooks without modifying the sources.
- Downloads supported email attachments and prints folder contents.
- Provides a searchable Hluthafaskrá with source-row traceability.

## ML safety status

The ML pipeline is implemented, but the current evaluated model did not meet
the required 95% precision gate and was not activated. Selecting ML-assisted
mode therefore returns the legacy autocoder result unchanged. Missing, corrupt,
or incompatible model artifacts also fail closed to the legacy result.

See [ML_EVALUATION.md](ML_EVALUATION.md) for the current measurements and the
reproducible evaluation commands. Training does not run automatically.

## Project layout

- `tok/TokTenging.py` — application shell, navigation, settings, and Tok input.
- `tok/bank_formatter.py` — bank conversion, coding review, and compression.
- `tok/bank_detection.py` — bank-format definitions and detection.
- `tok/auto_coder.py` — legacy rule/history-based autocoder.
- `tok/ml_*.py` — dataset validation, features, models, evaluation, artifacts,
  training, and guarded runtime inference.
- `tok/coded_file_tools.py` — coded-workbook combination and compression tools.
- `tok/shareholder_registry*.py` — registry loading, linking, and interface.
- `tok/email_downloader.py` and `tok/folder_printer.py` — desktop utilities.
- `tests/` — unit and integration regression tests.

## Local setup

Python 3.13 is used for the current Windows build.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python tok\TokTenging.py
```

Training workbooks, generated artifacts, settings, logs, and build outputs are
intentionally ignored. Keep them in the existing `tok/trainingCoded`,
`tok/trainingRaw`, and `tok/trainingData` folders.

## Tests

```powershell
.venv\Scripts\python.exe -m unittest discover -s tests -v
```

## ML evaluation and gated rebuild

Generate a report without changing the active model:

```powershell
.venv\Scripts\python.exe -m tok.ml_training evaluate --training-data-dir tok\trainingData --legacy-training-dir tok\trainingCoded --folds 3 --output artifacts\ml\evaluation_report.json
```

Evaluate and activate only if every production gate passes:

```powershell
.venv\Scripts\python.exe -m tok.ml_training rebuild --training-data-dir tok\trainingData --legacy-training-dir tok\trainingCoded --folds 3 --output artifacts\ml\evaluation_report.json --artifact-dir tok\trainingData\ml_artifacts
```

## Windows EXE

Install the development dependency and build with the main specification:

```powershell
python -m pip install -r requirements-dev.txt
python -m PyInstaller --noconfirm --clean "Tok Tenging.spec"
```

The main specification includes local legacy training data, the registry
workbook, and an activated ML artifact only when those ignored paths exist.
`Tok Tenging External Training.spec` omits private data.
