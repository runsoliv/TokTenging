"""Bank format definitions and Excel format detection."""

import pandas as pd


BANK_CONFIGS = {
    'islandsbanki': {
        'name': 'Íslandsbanki',
        'columns': ['Dags.', 'Mótaðili', 'Kennitala móttakanda', 'Upphæð'],
        'rename': {'Dags.': 'DATE', 'Mótaðili': 'TEXT', 'Kennitala móttakanda': 'ID', 'Upphæð': 'AMOUNT'},
        'sheet': 'Yfirlit',
        'detect_cols': {'Dags.', 'Kennitala móttakanda'}  # unique identifiers
    },
    'islandsbanki_kort': {
        'name': 'Íslandsbanki Kreditkort',
        'columns': ['Dags.', 'Skýring', 'Fjárhæð'],
        'rename': {'Dags.': 'DATE', 'Skýring': 'TEXT', 'Fjárhæð': 'AMOUNT'},
        'sheet': 'Yfirlit',
        'detect_cols': {'Dags.', 'Skýring', 'Upprunamynt', 'Fjárhæð'},
        'custom_processor': True  # no ID column in source file
    },
    'islandsbanki_reikningsyfirlit': {
        'name': 'Íslandsbanki Reikningsyfirlit',
        'columns': ['Dagsetning', 'Mótaðili', 'Tilvísun', 'Upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Mótaðili': 'TEXT', 'Tilvísun': 'ID', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dagsetning', 'Mótaðili', 'Tilvísun', 'Texti', 'Upphæð', 'Staða'},
    },
    'islandsbanki_special': {
        'name': 'Íslandsbanki Special',
        'columns': ['Dagsetning', 'Mótaðili', 'Upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Mótaðili': 'TEXT', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dagsetning', 'Mótaðili', 'Upphæð'},
        'custom_processor': True,  # no ID column in source file
        'skip_auto_coding': True,
        'auto_coding_note': 'Skip merchant/purpose auto-coding; this special format is calculation/subtraction based.'
    },
    'landsbanki': {
        'name': 'Landsbankinn',
        'columns': ['Dags', 'Texti', 'Tilvísun', 'Upphæð'],
        'rename': {'Dags': 'DATE', 'Texti': 'TEXT', 'Tilvísun': 'ID', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dags', 'Texti'}  # unique identifiers
    },
    'landsbankinn_kort': {
        'name': 'Landsbankinn Kort',
        'columns': ['Dagsetning', 'Söluaðili eða skýring', 'Upphæð(ISK)'],
        'rename': {'Dagsetning': 'DATE', 'Söluaðili eða skýring': 'TEXT', 'Upphæð(ISK)': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dagsetning', 'Söluaðili eða skýring', 'Upphæð(ISK)'},
        'custom_processor': True  # no ID column in source file
    },
    'arion': {
        'name': 'Arion Banki',
        'columns': ['Dagsetning', 'Skýring', 'Tilvísun', 'Upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Skýring': 'TEXT', 'Tilvísun': 'ID', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dagsetning', 'Skýring'}  # unique identifiers
    },
    'arion_kort': {
        'name': 'Arion Kort',
        'columns': ['Dagsetning', 'Lýsing', 'Innlend upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Lýsing': 'TEXT', 'Innlend upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dagsetning', 'Lýsing', 'Innlend upphæð'},
        'custom_processor': True  # no ID column in source file
    },
    'islandsbanki_innheimta': {
        'name': 'Íslandsbanki Innheimta',
        'columns': ['Kennitala', 'Greiðandi', 'Eindagi', 'Kröfunúmer', 'Upphæð', 'Fjármagnstekjuskattur', 'Dráttarvextir', 'Greidd upphæð', 'Greiðsludagur', 'Rst. Upphæð'],
        'sheet': None,
        'detect_cols': {'Kennitala', 'Greiðsludagur', 'Fjármagnstekjuskattur', 'Dráttarvextir', 'Rst. Upphæð'},
        'custom_processor': True,  # uses special processing logic
        'skip_auto_coding': True,
        'auto_coding_note': 'Skip merchant/purpose auto-coding; this format already uses fixed innheimta calculation codes.'
    },
    'sala_yfirlit': {
        'name': 'Sala Yfirlit',
        'columns': ['Nafn', 'Kennitala', 'Upphæð með vsk', 'Upphæð án vsk', 'Reikningur nr', 'Dagsetning'],
        'sheet': None,
        'detect_cols': {'Nafn', 'Kennitala', 'Upphæð með vsk', 'Reikningur nr', 'Dagsetning'},
        'custom_processor': True,  # uses special processing logic
        'skip_auto_coding': True,
        'auto_coding_note': 'Skip merchant/purpose auto-coding; Reikningur nr is incremental/internal, not a purchase category.'
    }
}

def detect_bank_type(file_path):
    """
    Reads the Excel file and detects which bank format it is based on column names.
    Returns (bank_key, df) or (None, None) if unknown.
    """
    def _find_header_row(path, required_cols, max_rows=30):
        try:
            preview = pd.read_excel(path, header=None, nrows=max_rows)
        except Exception:
            return None
        for idx in range(len(preview)):
            row_values = [v for v in preview.iloc[idx].tolist() if pd.notna(v)]
            if not row_values:
                continue
            row_cols = set(str(v).strip() for v in row_values)
            if required_cols.issubset(row_cols):
                return idx
        return None

    # Try Islandsbanki first (has specific sheet)
    try:
        df = pd.read_excel(file_path, sheet_name='Yfirlit')
        cols = set(df.columns)
        if BANK_CONFIGS['islandsbanki_kort']['detect_cols'].issubset(cols):
            return 'islandsbanki_kort', df
        if BANK_CONFIGS['islandsbanki']['detect_cols'].issubset(cols):
            return 'islandsbanki', df
    except Exception:
        pass  # Sheet doesn't exist, try others

    # Try reading default sheet for other banks
    try:
        df = pd.read_excel(file_path)
        cols = set(df.columns)

        # Check if first row contains metadata headers instead of data columns
        # Arion has "Heiti" / "IBAN númer" at top, Landsbanki has "Netbanki fyrirtækja"
        first_col = str(df.columns[0]).lower() if len(df.columns) > 0 else ''
        has_metadata_header = any(keyword in first_col for keyword in ['heiti', 'netbanki', 'iban'])

        # Also check first cell value for metadata text
        if not has_metadata_header and len(df) > 0:
            first_val = str(df.iloc[0, 0]).lower() if pd.notna(df.iloc[0, 0]) else ''
            has_metadata_header = any(keyword in first_val for keyword in ['heiti', 'netbanki', 'iban', 'færslur'])

        if has_metadata_header:
            # Try reading with header on row 3 (0-indexed, Excel row 4)
            df_skip3 = pd.read_excel(file_path, header=3)
            cols_skip3 = set(df_skip3.columns)

            # Check Arion (usually row 4)
            if BANK_CONFIGS['arion']['detect_cols'].issubset(cols_skip3):
                return 'arion', df_skip3

            # Check Arion card transactions
            if BANK_CONFIGS['arion_kort']['detect_cols'].issubset(cols_skip3):
                return 'arion_kort', df_skip3

            # Landsbankinn card transactions also usually use row 4 as header.
            if BANK_CONFIGS['landsbankinn_kort']['detect_cols'].issubset(cols_skip3):
                return 'landsbankinn_kort', df_skip3

            # Try reading with header on row 4 (0-indexed, Excel row 5)
            df_skip4 = pd.read_excel(file_path, header=4)
            cols_skip4 = set(df_skip4.columns)

            # Check Landsbanki (usually row 5)
            if BANK_CONFIGS['landsbanki']['detect_cols'].issubset(cols_skip4):
                return 'landsbanki', df_skip4

            # Check Arion on row 5 as fallback
            if BANK_CONFIGS['arion']['detect_cols'].issubset(cols_skip4):
                return 'arion', df_skip4

            # Check Arion card transactions on row 5 as fallback
            if BANK_CONFIGS['arion_kort']['detect_cols'].issubset(cols_skip4):
                return 'arion_kort', df_skip4

            # Check Landsbankinn card transactions on row 5 as fallback
            if BANK_CONFIGS['landsbankinn_kort']['detect_cols'].issubset(cols_skip4):
                return 'landsbankinn_kort', df_skip4

        arion_kort_header = _find_header_row(file_path, BANK_CONFIGS['arion_kort']['detect_cols'])
        if arion_kort_header is not None:
            df_arion_kort = pd.read_excel(file_path, header=arion_kort_header)
            return 'arion_kort', df_arion_kort

        landsbankinn_kort_header = _find_header_row(file_path, BANK_CONFIGS['landsbankinn_kort']['detect_cols'])
        if landsbankinn_kort_header is not None:
            df_landsbankinn_kort = pd.read_excel(file_path, header=landsbankinn_kort_header)
            return 'landsbankinn_kort', df_landsbankinn_kort

        islandsbanki_kort_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_kort']['detect_cols'])
        if islandsbanki_kort_header is not None:
            df_islandsbanki_kort = pd.read_excel(file_path, header=islandsbanki_kort_header)
            return 'islandsbanki_kort', df_islandsbanki_kort

        islandsbanki_reikningsyfirlit_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_reikningsyfirlit']['detect_cols'])
        if islandsbanki_reikningsyfirlit_header is not None:
            df_islandsbanki_reikningsyfirlit = pd.read_excel(file_path, header=islandsbanki_reikningsyfirlit_header)
            return 'islandsbanki_reikningsyfirlit', df_islandsbanki_reikningsyfirlit

        islandsbanki_special_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_special']['detect_cols'])
        if islandsbanki_special_header is not None:
            df_islandsbanki_special = pd.read_excel(file_path, header=islandsbanki_special_header)
            return 'islandsbanki_special', df_islandsbanki_special

        # Islandsbanki Innheimta sometimes has metadata rows above the header.
        innheimta_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_innheimta']['detect_cols'])
        if innheimta_header is not None:
            df_innheimta = pd.read_excel(file_path, header=innheimta_header)
            return 'islandsbanki_innheimta', df_innheimta

        # Standard detection (no metadata rows)
        # Check Islandsbanki Innheimta first (more specific columns)
        if BANK_CONFIGS['islandsbanki_innheimta']['detect_cols'].issubset(cols):
            return 'islandsbanki_innheimta', df
        # Islandsbanki Kreditkort
        if BANK_CONFIGS['islandsbanki_kort']['detect_cols'].issubset(cols):
            return 'islandsbanki_kort', df
        # Islandsbanki Reikningsyfirlit
        if BANK_CONFIGS['islandsbanki_reikningsyfirlit']['detect_cols'].issubset(cols):
            return 'islandsbanki_reikningsyfirlit', df
        # Islandsbanki Special
        if BANK_CONFIGS['islandsbanki_special']['detect_cols'].issubset(cols):
            return 'islandsbanki_special', df
        # Sala Yfirlit
        if BANK_CONFIGS['sala_yfirlit']['detect_cols'].issubset(cols):
            return 'sala_yfirlit', df
        # Landsbankinn Kort
        if BANK_CONFIGS['landsbankinn_kort']['detect_cols'].issubset(cols):
            return 'landsbankinn_kort', df
        # Arion Kort
        if BANK_CONFIGS['arion_kort']['detect_cols'].issubset(cols):
            return 'arion_kort', df

        # Check Landsbanki
        if BANK_CONFIGS['landsbanki']['detect_cols'].issubset(cols):
            return 'landsbanki', df

        # Check Arion
        if BANK_CONFIGS['arion']['detect_cols'].issubset(cols):
            return 'arion', df
    except Exception:
        pass

    return None, None
