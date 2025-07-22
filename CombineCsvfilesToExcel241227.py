import os
import pandas as pd
from tkinter import filedialog
import tkinter as tk

def select_folder():
    """
    フォルダを選択するダイアログを表示し、選択されたフォルダのパスを返す。

    Returns:
    - folder_path (str): 選択されたフォルダのパス
    """
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="CSVファイルが保存されているフォルダを選択してください")
    return folder_path

def extract_sheet_name(file_name):
    """
    ファイル名からシート名を抽出する。

    Args:
    - file_name (str): ファイル名（例: "prefix_part1_part2_ROI123.csv"）

    Returns:
    - sheet_name (str): シート名（例: "part1_part2ROI123"）
    """
    try:
        # ファイル名をアンダースコアで分割
        parts = file_name.split("_")
        if len(parts) < 4:
            raise ValueError(f"ファイル名 '{file_name}' が期待する形式ではありません。")

        # 1つ目と3つ目のアンダースコアの間の部分を抽出
        middle_part = "_".join(parts[1:3])
        
        # "ROI"以降の文字列を抽出
        roi_part = file_name.split("ROI")[-1].split(".")[0]

        # シート名を組み立てる
        sheet_name = f"{middle_part}ROI{roi_part}"
        
        # シート名の長さを調整（Excelの31文字制限）
        return sheet_name[:31]
    except Exception as e:
        print(f"シート名生成エラー: {str(e)}")
        return "InvalidSheetName"

def merge_all_csv_to_excel(folder_path):
    """
    指定されたフォルダ内のすべてのCSVファイルを1つのExcelファイルにまとめる。

    Args:
    - folder_path (str): CSVファイルが格納されているルートフォルダのパス

    Returns:
    - None
    """
    # 保存先のExcelファイル
    save_path = os.path.join(folder_path, "CombinedWorkbook.xlsx")
    
    # ExcelWriterオブジェクトを作成
    writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    
    for root, _, files in os.walk(folder_path):
        csv_files = [f for f in files if f.endswith('.csv')]
        
        for csv_file in csv_files:
            csv_path = os.path.join(root, csv_file)
            
            try:
                # シート名を抽出
                sheet_name = extract_sheet_name(csv_file)
                
                # CSVファイルをDataFrameとして読み込む
                df = pd.read_csv(csv_path, encoding='utf-8')
                
                # DataFrameをExcelの新しいシートに書き込む
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            except Exception as e:
                print(f"CSVファイル '{csv_file}' の処理中にエラーが発生しました: {str(e)}")
    
    # Excelファイルを保存
    writer._save()
    writer.close()
    print(f"CSVファイルがExcelファイル({save_path})にまとめられました！")

if __name__ == "__main__":
    # フォルダを選択
    root_folder = select_folder()
    if not root_folder:
        print("フォルダが選択されませんでした。処理を中止します。")
    else:
        merge_all_csv_to_excel(root_folder)
