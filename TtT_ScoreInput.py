import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import openpyxl
import shutil
import os
import glob
import webbrowser
import sys
import ctypes

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass

# スクリプトのパスを取得する関数
def get_script_dir():
    if getattr(sys, 'frozen', False):
        # .exeファイルの場合
        script_dir = os.path.dirname(sys.executable)
        
        # macOSの.appバンドルの場合
        if sys.platform.startswith('darwin'):
            script_dir = os.path.abspath(os.path.join(script_dir, '../../../'))
        
        return script_dir
    else:
        # .pyファイルの場合
        return os.path.dirname(os.path.abspath(__file__))

# スクリプトのパス
SCRIPT_DIR = get_script_dir()
if sys.platform.startswith('darwin'):
    SCRIPT_PATH = SCRIPT_DIR
else:
    SCRIPT_PATH = os.path.join(SCRIPT_DIR, 'TtT_ScoreInput.py') if not getattr(sys, 'frozen', False) else sys.executable

# ファイル検証
def get_valid_file_path(file_path, prompt_message, file_type):
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
    csv_type = None
    try:
        csv_df = pd.read_csv(csv_file)
        if 'title' in csv_df.columns and 'difficulty' in csv_df.columns and 'highScore' in csv_df.columns and 'FCCount' in csv_df.columns and 'APCount' in csv_df.columns:
            csv_type = 0
            return csv_df, csv_type
        elif '楽曲名' in csv_df.columns and '難易度' in csv_df.columns and 'ハイスコア' in csv_df.columns and 'フルコンボ回数' in csv_df.columns and 'パーフェクト回数' in csv_df.columns:
            csv_type = 1
            return csv_df, csv_type
        else:
            messagebox.showerror("エラー", "CSVファイルに必要なカラムが含まれていません。")
            return None, csv_type
    except Exception as e:
        messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました: {e}")
        return None, csv_type

# Excelファイルの更新
def update_excel(sheet, excel_df, csv_df, csv_type):
    difficulty_columns = {
        'standard': 'I',
        'expert': 'J',
        'ultimate': 'K',
        'maniac': 'L',
        'connect': 'M'
    }

    # 難易度ごとのスコアボーダー
    difficulty_borders = {
        'standard': 500000,
        'expert': 600000,
        'ultimate': 700000,
        'maniac': 800000,
        'connect': 700000
    }

    warnings = {}

    for index, row in csv_df.iterrows():

        if csv_type == 0:
            title = row['title'].rstrip()  # 最後の空白を削除
            difficulty = row['difficulty']
            ap_count = row['APCount']
            fc_count = row['FCCount']
            high_score = row['highScore']
        elif csv_type == 1:
            title = row["楽曲名"].rstrip()  # 最後の空白を削除
            difficulty = row["難易度"]
            ap_count = row['パーフェクト回数']
            fc_count = row['フルコンボ回数']
            high_score = row['ハイスコア']

        #csvタイプに依存しないようにすべて小文字へ
        difficulty = difficulty.lower()  

        excel_column = difficulty_columns.get(difficulty)
        if excel_column:
            excel_row_index = excel_df[excel_df['Title'].str.rstrip() == title].index  # 最後の空白を削除
            if not excel_row_index.empty:
                excel_row_index = excel_row_index[0] + 3
                column_index = openpyxl.utils.column_index_from_string(excel_column)
                cell = sheet.cell(row=excel_row_index, column=column_index)

                new_value = cell.value

                if ap_count >= 1:
                    new_value = 'AP'
                elif fc_count >= 1:
                    new_value = 'FC'
                elif high_score >= difficulty_borders[difficulty]:
                    new_value = 'CL'
                elif high_score >= 0:
                    new_value = 'FL'
                
                if not cell.value or (cell.value in ['FC', 'CL', 'FL'] and new_value == 'AP') or (cell.value == 'CL' and new_value == 'FC') or (cell.value == 'FL' and new_value == 'CL'):
                    cell.value = new_value

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
    csv_df, csv_type = read_csv_file(csv_path)
    if csv_df is None:
        return
    if csv_type is None:
        return

    # Excelファイルの更新
    warnings = update_excel(sheet, excel_df, csv_df, csv_type)
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

    def update_file_path(var, dialog_title, file_type):
        new_path = get_valid_file_path(var.get(), dialog_title, file_type)
        if new_path:  # new_path が空でない場合にのみ更新
            var.set(new_path)

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

    tk.Label(root, text="スクリプトのパス:").grid(
        row=0, column=0, padx=10, pady=5, sticky="e"
    )
    tk.Entry(root, textvariable=tk.StringVar(value=SCRIPT_PATH), state='readonly', width=60).grid(
        row=0, column=1, padx=10, pady=5
    )
    tk.Button(root, text="フォルダを開く", command=open_script_folder).grid(
        row=0, column=2, padx=10, pady=5
    )

    tk.Label(root, text="Excelファイルのパス:").grid(
        row=1, column=0, padx=10, pady=5, sticky="e"
    )
    tk.Entry(root, textvariable=excel_path_var, width=60).grid(
        row=1, column=1, padx=10, pady=5
    )
    
    # Excelファイルの参照ボタン
    tk.Button(
        root, 
        text="参照", 
        command=lambda: update_file_path(
            excel_path_var, 
            "Excelファイルを選択してください", 
            ".xlsx"
        )
    ).grid(row=1, column=2, padx=10, pady=5)

    tk.Label(root, text="CSVファイルのパス:").grid(
        row=2, column=0, padx=10, pady=5, sticky="e"
    )
    tk.Entry(root, textvariable=csv_path_var, width=60).grid(
        row=2, column=1, padx=10, pady=5
    )

    # CSVファイルの参照ボタン
    tk.Button(
        root, 
        text="参照", 
        command=lambda: update_file_path(
            csv_path_var, 
            "CSVファイルを選択してください", 
            ".csv"
        )
    ).grid(row=2, column=2, padx=10, pady=5)

    tk.Button(root, text="デフォルトに設定", command=set_default_paths).grid(
        row=4, column=0, padx=10, pady=5
    )
    tk.Button(root, text="パスをリセット", command=reset_paths).grid(
        row=4, column=1, padx=10, pady=5
    )

    tk.Button(
        root, 
        text="処理を開始", 
        command=lambda: process_files(excel_path_var.get(), csv_path_var.get(), root)
    ).grid(row=5, column=0, columnspan=3, padx=10, pady=20)

    tk.Button(
        root, 
        text="@_ryuya_0124", 
        command=open_twitter_profile
    ).grid(row=6, column=0, padx=10, pady=5, columnspan=3)

    # 初回起動時にデフォルトパスを設定
    if not initial_setup_done:
        initial_setup_done = True
        set_default_paths()

    root.mainloop()

# メイン関数の実行
if __name__ == "__main__":
    create_gui()
