import openpyxl


def update_excel(sheet, excel_df, csv_df, csv_type):
    difficulty_columns = {
        'standard': 'I',
        'expert': 'J',
        'ultimate': 'K',
        'maniac': 'L',
        'connect': 'M',
    }
    difficulty_borders = {
        'standard': 500000,
        'expert': 600000,
        'ultimate': 700000,
        'maniac': 800000,
        'connect': 700000,
    }
    status_rank = {'FL': 0, 'CL': 1, 'FC': 2, 'AP': 3}
    normalized_titles = excel_df['Title'].fillna('').astype(str).str.rstrip()
    warnings = {}

    for _, row in csv_df.iterrows():
        if csv_type == 0:
            title = str(row['title']).rstrip()
            difficulty = str(row['difficulty']).lower()
            ap_count = row['APCount']
            fc_count = row['FCCount']
            high_score = row['highScore']
        else:
            title = str(row['楽曲名']).rstrip()
            difficulty = str(row['難易度']).lower()
            ap_count = row['パーフェクト回数']
            fc_count = row['フルコンボ回数']
            high_score = row['ハイスコア']

        excel_column = difficulty_columns.get(difficulty)
        if not excel_column:
            continue

        matching_rows = excel_df[normalized_titles == title].index
        if matching_rows.empty:
            warnings.setdefault(title, set()).add(difficulty)
            continue

        row_number = matching_rows[0] + 3
        column_number = openpyxl.utils.column_index_from_string(excel_column)
        cell = sheet.cell(row=row_number, column=column_number)

        if ap_count >= 1:
            new_value = 'AP'
        elif fc_count >= 1:
            new_value = 'FC'
        elif high_score >= difficulty_borders[difficulty]:
            new_value = 'CL'
        else:
            new_value = 'FL'

        if status_rank.get(new_value, -1) > status_rank.get(cell.value, -1):
            cell.value = new_value

    return warnings
