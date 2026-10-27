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
        'columns': ['Dagsetning', 'Mótaðili', 'Kennitala mótaðila', 'Upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Mótaðili': 'TEXT', 'Kennitala mótaðila': 'ID', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'metadata_labels': {'Eigandi', 'Reikningsnúmer', 'Yfirlit sótt'},
        'detect_cols': {
            'Dagsetning', 'Mótaðili', 'Kennitala mótaðila',
            'Tilvísun', 'Upphæð', 'Staða'
        },
    },
    'islandsbanki_reikningsyfirlit_metadata': {
        'name': 'Íslandsbanki Reikningsyfirlit',
        'columns': ['Dagsetning', 'Skýring', 'Kennitala', 'Upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Skýring': 'TEXT', 'Kennitala': 'ID', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {
            'Dagsetning', 'Skýring', 'Tilvísun', 'Upphæð', 'Staða', 'Mynt',
            'Nafn viðtakanda eða greiðanda', 'Kennitala'
        },
    },
    'islandsbanki_special': {
        'name': 'Íslandsbanki – sérsnið',
        'columns': ['Dagsetning', 'Mótaðili', 'Upphæð'],
        'rename': {'Dagsetning': 'DATE', 'Mótaðili': 'TEXT', 'Upphæð': 'AMOUNT'},
        'sheet': None,
        'detect_cols': {'Dagsetning', 'Mótaðili', 'Upphæð'},
        'custom_processor': True,  # no ID column in source file
        'skip_auto_coding': True,
        'auto_coding_note': 'Sjálfvirkri flokkun söluaðila og tilgangs er sleppt; þetta sérsnið byggir á útreikningi og frádrætti.'
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
    'arion_reikningsyfirlit': {
        'name': 'Arion Banki Reikningsyfirlit',
        'columns': ['Dagsetning', 'Skýring', 'Kennitala viðtakanda eða greiðanda', 'Upphæð'],
        'rename': {
            'Dagsetning': 'DATE',
            'Skýring': 'TEXT',
            'Kennitala viðtakanda eða greiðanda': 'ID',
            'Upphæð': 'AMOUNT',
        },
        'sheet': None,
        'detect_cols': {
            'Dagsetning', 'Upphæð', 'Staða', 'Mynt', 'Skýring',
            'Færslulykill', 'Kennitala viðtakanda eða greiðanda',
        },
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
        'auto_coding_note': 'Sjálfvirkri flokkun söluaðila og tilgangs er sleppt; þetta snið notar þegar fasta útreikningslykla fyrir innheimtu.'
    },
    'sala_yfirlit': {
        'name': 'Söluyfirlit',
        'columns': ['Nafn', 'Kennitala', 'Upphæð með vsk', 'Upphæð án vsk', 'Reikningur nr', 'Dagsetning'],
        'sheet': None,
        'detect_cols': {'Nafn', 'Kennitala', 'Upphæð með vsk', 'Reikningur nr', 'Dagsetning'},
        'custom_processor': True,  # uses special processing logic
        'skip_auto_coding': True,
        'auto_coding_note': 'Sjálfvirkri flokkun söluaðila og tilgangs er sleppt; „Reikningur nr“ er hlaupandi innra númer en ekki innkaupaflokkur.'
    }
}

def detect_bank_type(file_path):
    """
    Reads the Excel file and detects which bank format it is based on column names.
    Returns (bank_key, df) or (None, None) if unknown.
    """
    def _find_header_row(path, required_cols, max_rows=30, metadata_labels=None):
        try:
            preview = pd.read_excel(path, header=None, nrows=max_rows, dtype=object)
        except Exception:
            return None
        if metadata_labels:
            first_column_values = {
                str(value).strip()
                for value in preview.iloc[:, 0].tolist()
                if pd.notna(value)
            }
            if not set(metadata_labels).issubset(first_column_values):
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
        df = pd.read_excel(file_path, sheet_name='Yfirlit', dtype=object)
        cols = set(df.columns)
        if BANK_CONFIGS['islandsbanki_kort']['detect_cols'].issubset(cols):
            return 'islandsbanki_kort', df
        if BANK_CONFIGS['islandsbanki']['detect_cols'].issubset(cols):
            return 'islandsbanki', df
    except Exception:
        pass  # Sheet doesn't exist, try others
    
    # Try reading default sheet for other banks
    try:
        df = pd.read_excel(file_path, dtype=object)
        cols = set(df.columns)

        # Check if first row contains metadata headers instead of data columns.
        # Arion has "Heiti" / "IBAN númer" at top, Landsbanki has "Netbanki fyrirtækja",
        # and some Íslandsbanki exports have account/statement text above the real header.
        first_col = str(df.columns[0]).lower() if len(df.columns) > 0 else ''
        metadata_keywords = [
            'heiti', 'netbanki', 'iban', 'færslur', 'reikningseigandi',
            'nánari upplýsingar', 'stofnupplýsingar', 'tímabil', 'reikningsyfirlit'
        ]
        has_metadata_header = any(keyword in first_col for keyword in metadata_keywords)
        
        # Also check first cell value for metadata text
        if not has_metadata_header and len(df) > 0:
            first_val = str(df.iloc[0, 0]).lower() if pd.notna(df.iloc[0, 0]) else ''
            has_metadata_header = any(keyword in first_val for keyword in metadata_keywords)
        
        if has_metadata_header:
            # Try reading with header on row 3 (0-indexed, Excel row 4)
            df_skip3 = pd.read_excel(file_path, header=3, dtype=object)
            cols_skip3 = set(df_skip3.columns)

            # Check metadata-style Íslandsbanki Reikningsyfirlit export before generic Arion.
            # It also contains "Dagsetning" + "Skýring", so Arion's broad rule would match first.
            if BANK_CONFIGS['islandsbanki_reikningsyfirlit_metadata']['detect_cols'].issubset(cols_skip3):
                return 'islandsbanki_reikningsyfirlit_metadata', df_skip3

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
            df_skip4 = pd.read_excel(file_path, header=4, dtype=object)
            cols_skip4 = set(df_skip4.columns)

            # Check metadata-style Íslandsbanki Reikningsyfirlit export before generic Arion.
            if BANK_CONFIGS['islandsbanki_reikningsyfirlit_metadata']['detect_cols'].issubset(cols_skip4):
                return 'islandsbanki_reikningsyfirlit_metadata', df_skip4

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

        # New Íslandsbanki account statements have an account metadata block
        # and place the transaction header on row 12. Detect this distinctive
        # layout before broader bank/card rules.
        islandsbanki_statement_config = BANK_CONFIGS['islandsbanki_reikningsyfirlit']
        islandsbanki_statement_header = _find_header_row(
            file_path,
            islandsbanki_statement_config['detect_cols'],
            max_rows=60,
            metadata_labels=islandsbanki_statement_config.get('metadata_labels'),
        )
        if islandsbanki_statement_header is not None:
            df_islandsbanki_statement = pd.read_excel(
                file_path,
                header=islandsbanki_statement_header,
                dtype=object,
            )
            return 'islandsbanki_reikningsyfirlit', df_islandsbanki_statement

        arion_statement_header = _find_header_row(
            file_path,
            BANK_CONFIGS['arion_reikningsyfirlit']['detect_cols'],
            max_rows=60,
        )
        if arion_statement_header is not None:
            df_arion_statement = pd.read_excel(
                file_path,
                header=arion_statement_header,
                dtype=object,
            )
            return 'arion_reikningsyfirlit', df_arion_statement

        arion_kort_header = _find_header_row(file_path, BANK_CONFIGS['arion_kort']['detect_cols'])
        if arion_kort_header is not None:
            df_arion_kort = pd.read_excel(file_path, header=arion_kort_header, dtype=object)
            return 'arion_kort', df_arion_kort

        landsbankinn_kort_header = _find_header_row(file_path, BANK_CONFIGS['landsbankinn_kort']['detect_cols'])
        if landsbankinn_kort_header is not None:
            df_landsbankinn_kort = pd.read_excel(file_path, header=landsbankinn_kort_header, dtype=object)
            return 'landsbankinn_kort', df_landsbankinn_kort

        islandsbanki_kort_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_kort']['detect_cols'])
        if islandsbanki_kort_header is not None:
            df_islandsbanki_kort = pd.read_excel(file_path, header=islandsbanki_kort_header, dtype=object)
            return 'islandsbanki_kort', df_islandsbanki_kort

        islandsbanki_reikningsyfirlit_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_reikningsyfirlit']['detect_cols'])
        if islandsbanki_reikningsyfirlit_header is not None:
            df_islandsbanki_reikningsyfirlit = pd.read_excel(
                file_path,
                header=islandsbanki_reikningsyfirlit_header,
                dtype=object,
            )
            return 'islandsbanki_reikningsyfirlit', df_islandsbanki_reikningsyfirlit

        islandsbanki_reikningsyfirlit_metadata_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_reikningsyfirlit_metadata']['detect_cols'])
        if islandsbanki_reikningsyfirlit_metadata_header is not None:
            df_islandsbanki_reikningsyfirlit_metadata = pd.read_excel(
                file_path,
                header=islandsbanki_reikningsyfirlit_metadata_header,
                dtype=object,
            )
            return 'islandsbanki_reikningsyfirlit_metadata', df_islandsbanki_reikningsyfirlit_metadata

        islandsbanki_special_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_special']['detect_cols'])
        if islandsbanki_special_header is not None:
            df_islandsbanki_special = pd.read_excel(
                file_path,
                header=islandsbanki_special_header,
                dtype=object,
            )
            return 'islandsbanki_special', df_islandsbanki_special

        # Islandsbanki Innheimta sometimes has metadata rows above the header.
        innheimta_header = _find_header_row(file_path, BANK_CONFIGS['islandsbanki_innheimta']['detect_cols'])
        if innheimta_header is not None:
            df_innheimta = pd.read_excel(file_path, header=innheimta_header, dtype=object)
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
        if BANK_CONFIGS['arion_reikningsyfirlit']['detect_cols'].issubset(cols):
            return 'arion_reikningsyfirlit', df
        
        # Check Landsbanki
        if BANK_CONFIGS['landsbanki']['detect_cols'].issubset(cols):
            return 'landsbanki', df
        
        # Check Arion
        if BANK_CONFIGS['arion']['detect_cols'].issubset(cols):
            return 'arion', df
    except Exception:
        pass
    
    return None, None

