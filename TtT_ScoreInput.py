import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import shutil
import os
import glob
import webbrowser
import sys

# スクリプトのパスを取得する関数
def get_script_dir():
    if getattr(sys, 'frozen', False):
        # .exeファイルの場合
        return os.path.dirname(sys.executable)
    else:
        # .pyファイルの場合
        return os.path.dirname(os.path.abspath(__file__))

# スクリプトのパス
SCRIPT_DIR = get_script_dir()
SCRIPT_PATH = os.path.join(SCRIPT_DIR, 'TtT_ScoreInput.py') if not getattr(sys, 'frozen', False) else sys.executable

# ファイル検証
def get_valid_file_path(file_path, prompt_message, file_type):
    if os.path.exists(file_path) and file_path.lower().endswith(file_type):
        return file_path
    while True:
        file_path = filedialog.askopenfilename(title=prompt_message, filetypes=[(file_type.upper(), f"*{file_type}")])
        if not file_path:
            return ''  # キャンセルされた場合は空文字を返す
        if os.path.exists(file_path) and file_path.lower().endswith(file_type):
            return file_path.replace('/', '\\')  # パス区切り文字を \ に変換
        else:
            messagebox.showerror("エラー", f"'{file_path}' は存在しないか、{file_type}ファイルではありません。もう一度試してください。")

# CSVファイルを見つける
def find_csv_file(directory):
    csv_files = glob.glob(os.path.join(directory, '*.csv'))
    if csv_files:
        return csv_files[0]
    else:
        return None

# ファイルのコピー
def copy_file(original_file, temp_file):
    shutil.copy(original_file, temp_file)

# Excelファイルを読み込む
def read_excel_file(temp_file):
    try:
        excel_df = pd.read_excel(temp_file, sheet_name=0, engine='openpyxl', header=1)
        book = openpyxl.load_workbook(temp_file)
        sheet = book.active
        return excel_df, book, sheet
    except Exception as e:
        messagebox.showerror("エラー", f"Excelファイルの読み込みに失敗しました: {e}")
        return None, None, None

# CSVファイルを読み込む
def read_csv_file(csv_file):
    try:
        csv_df = pd.read_csv(csv_file)
        if 'title' in csv_df.columns and 'difficulty' in csv_df.columns:
            return csv_df
        else:
            messagebox.showerror("エラー", "CSVファイルに必要なカラムが含まれていません。")
            return None
    except Exception as e:
        messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")
        return None

# Excelファイルの更新
def update_excel(sheet, excel_df, csv_df):
    difficulty_columns = {
        'Standard': 'I',
        'Expert': 'J',
        'Ultimate': 'K',
        'Maniac': 'L',
        'Connect': 'M'
    }

    # 難易度ごとのスコアボーダー
    difficulty_borders = {
        'Standard': 500000,
        'Expert': 600000,
        'Ultimate': 700000,
        'Maniac': 800000,
        'Connect': 700000
    }

    warnings = {}

    for index, row in csv_df.iterrows():
        title = row['title'].rstrip()  # 最後の空白を削除
        difficulty = row['difficulty']
        ap_count = row['APCount']
        fc_count = row['FCCount']
        high_score = row['highScore']

        excel_column = difficulty_columns.get(difficulty)
        if excel_column:
            excel_row_index = excel_df[excel_df['Title'].str.rstrip() == title].index  # 最後の空白を削除
            if not excel_row_index.empty:
                excel_row_index = excel_row_index[0] + 3
                column_index = openpyxl.utils.column_index_from_string(excel_column)
                cell = sheet.cell(row=excel_row_index, column=column_index)

                if cell.value is None or cell.value == '':
                    if ap_count >= 1:
                        cell.value = 'AP'
                    elif fc_count >= 1:
                        cell.value = 'FC'
                    elif high_score >= difficulty_borders[difficulty]:
                        cell.value = 'CL'
                elif cell.value == 'FC' and ap_count >= 1:
                    cell.value = 'AP'
                elif cell.value == 'CL' and (fc_count >= 1 or ap_count >= 1):
                    cell.value = 'FC' if fc_count >= 1 else 'AP'
                elif cell.value == '' and high_score >= difficulty_borders[difficulty]:
                    cell.value = 'CL'
            else:
                if title not in warnings:
                    warnings[title] = set()
                warnings[title].add(difficulty)

    return warnings



# 警告メッセージを表示
def print_warnings(warnings, root):
    if warnings:
        warning_window = tk.Toplevel(root)
        warning_window.title("警告")
        warning_window.grab_set()  # モーダルにする

        warning_message = "以下の楽曲がExcelファイルに見つかりませんでした。\n\n"
        tk.Label(warning_window, text=warning_message).pack(pady=10)

        warning_text = tk.Text(warning_window, height=10, width=50)
        warning_text.pack(pady=10, padx=10)

        for title in warnings:
            warning_text.insert(tk.END, f"{title}\n")

        warning_text.config(state=tk.DISABLED)
        tk.Button(warning_window, text="閉じる", command=warning_window.destroy).pack(pady=10)

        root.wait_window(warning_window)  # ウィンドウが閉じられるまで待つ

# Excelファイルを保存
def save_excel_file(book, output_file):
    try:
        book.save(output_file)
        messagebox.showinfo("成功", f"データの追記が完了しました。出力ファイル: {output_file}")
        return True
    except PermissionError:
        messagebox.showerror("エラー", f"出力ファイル '{output_file}' にアクセスする権限がありません。")
    except Exception as e:
        messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")
    return False

# 処理開始
def process_files(excel_path, csv_path, root):
    missing_paths = []
    
    # ExcelファイルとCSVファイルの存在確認
    if not os.path.exists(excel_path):
        missing_paths.append("Excelファイル")
    if not os.path.exists(csv_path):
        missing_paths.append("CSVファイル")

    # ファイルの存在確認結果に応じてメッセージを表示
    if missing_paths:
        missing_files_message = "以下のファイルが見つかりません:\n" + "\n".join(missing_paths)
        messagebox.showerror("エラー", missing_files_message)
        return

    # 一時ファイルの作成と読み込み
    temp_file = 'temp_' + os.path.basename(excel_path)
    copy_file(excel_path, temp_file)
    excel_df, book, sheet = read_excel_file(temp_file)
    if excel_df is None or book is None or sheet is None:
        return

    # CSVファイルの読み込み
    csv_df = read_csv_file(csv_path)
    if csv_df is None:
        return

    # Excelファイルの更新
    warnings = update_excel(sheet, excel_df, csv_df)
    print_warnings(warnings, root)

    # Excelファイルの保存
    if save_excel_file(book, excel_path):
        os.remove(temp_file)
    else:
        os.remove(temp_file)  # アクセス権限エラーの場合もtempファイルを削除


# GUIの設定
def create_gui():
    def open_script_folder():
        webbrowser.open(SCRIPT_DIR)

    def set_default_paths():
        excel_default_path = os.path.join(SCRIPT_DIR, 'TtT_ClearSheet.xlsx')
        csv_default_path = find_csv_file(SCRIPT_DIR)

        if os.path.exists(excel_default_path):
            excel_path_var.set(excel_default_path)
        else:
            excel_path_var.set('')
            if not initial_setup_done:
                messagebox.showerror("エラー", "TtT_ClearSheet.xlsxが見つかりません。")

        if csv_default_path:
            csv_path_var.set(csv_default_path)
        else:
            csv_path_var.set('')
            '''if not initial_setup_done:
                messagebox.showerror("エラー", "CSVファイルが見つかりません。")'''

    def reset_paths():
        excel_path_var.set('')
        csv_path_var.set('')

    def open_twitter_profile():
        webbrowser.open("https://twitter.com/_ryuya_0124")

    global initial_setup_done
    initial_setup_done = False

    root = tk.Tk()
    root.title("TtT_ScoreInput")

    excel_path_var = tk.StringVar()
    csv_path_var = tk.StringVar()

    tk.Label(root, text="スクリプトのパス:").grid(row=0, column=0, padx=10, pady=5, sticky="e")
    tk.Entry(root, textvariable=tk.StringVar(value=SCRIPT_PATH), state='readonly', width=60).grid(row=0, column=1, padx=10, pady=5)
    tk.Button(root, text="フォルダを開く", command=open_script_folder).grid(row=0, column=2, padx=10, pady=5)

    tk.Label(root, text="Excelファイルのパス:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
    tk.Entry(root, textvariable=excel_path_var, width=60).grid(row=1, column=1, padx=10, pady=5)
    tk.Button(root, text="参照", command=lambda: excel_path_var.set(get_valid_file_path(excel_path_var.get(), "Excelファイルを選択してください", ".xlsx"))).grid(row=1, column=2, padx=10, pady=5)

    tk.Label(root, text="CSVファイルのパス:").grid(row=2, column=0, padx=10, pady=5, sticky="e")
    tk.Entry(root, textvariable=csv_path_var, width=60).grid(row=2, column=1, padx=10, pady=5)
    tk.Button(root, text="参照", command=lambda: csv_path_var.set(get_valid_file_path(csv_path_var.get(), "CSVファイルを選択してください", ".csv"))).grid(row=2, column=2, padx=10, pady=5)

    tk.Button(root, text="デフォルトに設定", command=set_default_paths).grid(row=4, column=0, padx=10, pady=5)
    tk.Button(root, text="パスをリセット", command=reset_paths).grid(row=4, column=1, padx=10, pady=5)

    tk.Button(root, text="処理を開始", command=lambda: process_files(excel_path_var.get(), csv_path_var.get(), root)).grid(row=5, column=0, columnspan=3, padx=10, pady=20)

    tk.Button(root, text="@_ryuya_0124", command=open_twitter_profile).grid(row=6, column=0, padx=10, pady=5, columnspan=3)

    # 初回起動時にデフォルトパスを設定
    if not initial_setup_done:
        initial_setup_done = True
        set_default_paths()

    root.mainloop()

# メイン関数の実行
if __name__ == "__main__":
    create_gui()
