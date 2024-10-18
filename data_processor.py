# 必要なモジュールのインポート
import os
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font
from bs4 import BeautifulSoup
import streamlit as st
from datetime import datetime
from github import Github

# シークレットからGitHubトークンを取得
GITHUB_TOKEN = st.secrets["github"]["token"]

# GitHubへのファイルアップロード関数
def upload_file_to_github(file_path, repo_name, file_name_in_repo, commit_message):
    # GitHubに認証
    g = Github(GITHUB_TOKEN)
    user = g.get_user()
    repo = user.get_repo(repo_name)

    # ファイルの読み込み
    with open(file_path, 'rb') as file:
        content = file.read()

    # リポジトリ内のファイルパス
    path = file_name_in_repo

    try:
        # 既存のファイルがあるか確認
        contents = repo.get_contents(path)
        # ファイルを更新
        repo.update_file(path, commit_message, content, contents.sha)
        st.info(f"{file_name_in_repo} を更新しました。")
    except Exception as e:
        # ファイルが存在しない場合は新規作成
        repo.create_file(path, commit_message, content)
        st.info(f"{file_name_in_repo} を作成しました。")

def extract_data_and_save_to_csv(html_path, output_csv_path, date):
    # HTMLファイルの内容を読み込み
    with open(html_path, "r", encoding="utf-8") as file:
        html_content = file.read()

    # BeautifulSoupを使ってHTMLを解析
    soup = BeautifulSoup(html_content, "lxml")

    # テーブル行（データを含む部分）を取得（ヘッダー行はスキップ）
    rows = soup.find_all("tr")[1:]

    # データを保存するリストを初期化
    data = {
        "台番号": [], "累計スタート": [], "BB回数": [], "RB回数": [], 
        "ART回数": [], "最大持玉": [], "BB確率": [], "RB確率": [], 
        "ART確率": [], "合成確率": []
    }

    # 各行をループし、セルの値を抽出
    for row in rows:
        cells = row.find_all("td")
        if len(cells) > 1:
            data["台番号"].append(cells[1].get_text())
            data["累計スタート"].append(cells[2].get_text())
            data["BB回数"].append(cells[3].get_text())
            data["RB回数"].append(cells[4].get_text())
            data["ART回数"].append(cells[5].get_text())
            data["最大持玉"].append(cells[6].get_text())
            data["BB確率"].append(cells[7].get_text())
            data["RB確率"].append(cells[8].get_text())
            data["ART確率"].append(cells[9].get_text())
            data["合成確率"].append(cells[10].get_text())

    # データをPandasのDataFrameに変換
    df = pd.DataFrame(data)

    # CSVファイルとして保存
    df.to_csv(output_csv_path, index=False, encoding="shift-jis")
    return df

def create_new_excel_with_all_data(output_csv_dir, excel_path):
    # フォルダー内のすべてのCSVファイルを取得
    csv_files = [os.path.join(output_csv_dir, f) for f in os.listdir(output_csv_dir) if f.endswith('.csv')]
    
    # 新しいExcelファイルの作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "合成確率"

    # ヘッダー行として日付を追加
    ws.cell(row=1, column=1, value="台番号")
    
    all_data = {}
    date_columns = []

    for csv_file in csv_files:
        # CSVファイルを読み込み
        df = pd.read_csv(csv_file, encoding="shift-jis")
        
        # ファイル名から日付を抽出（例: "slot_machine_data_2024-10-17.csv"）
        date = os.path.basename(csv_file).split('_')[-1].replace('.csv', '')
        formatted_date = pd.to_datetime(date).strftime('%Y/%m/%d')  # 日付を yyyy/mm/dd 形式に変換
        date_columns.append(formatted_date)
        
        # 各台番号ごとにデータをまとめる
        for index, row in df.iterrows():
            if row['台番号'] not in all_data:
                all_data[row['台番号']] = {}
            all_data[row['台番号']][formatted_date] = row['合成確率']
    
    # 日付列をヘッダーに追加
    for col_index, date in enumerate(sorted(date_columns), start=2):
        ws.cell(row=1, column=col_index, value=date)

    # 台番号を行に追加し、合成確率を各セルに追加
    for row_index, (machine_number, dates_data) in enumerate(all_data.items(), start=2):
        ws.cell(row=row_index, column=1, value=machine_number)
        for col_index, date in enumerate(sorted(date_columns), start=2):
            ws.cell(row=row_index, column=col_index, value=dates_data.get(date, None))

    # 列幅と行の高さを自動調整
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter  # 列の文字（例: "A", "B"）
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = max(max_length + 2, 10)  # 列幅を自動調整し、少なくとも10に設定
        ws.column_dimensions[column].width = adjusted_width

    # 行の高さも自動調整
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 20  # 20の固定高さを設定

    # すべてのセルにメイリオフォントを適用
    mei_font = Font(name="メイリオ")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.font = mei_font

    # 新しいExcelファイルを保存
    wb.save(excel_path)

def apply_color_fill_to_excel(excel_path):
    # Excelファイルを読み込み
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    # 色を定義
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # 黄色
    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # 水色

    # 各行をループし、確率に基づいて色を適用
    for row in ws.iter_rows(min_row=2, min_col=2, max_row=ws.max_row, max_col=ws.max_column):
        for cell in row:
            try:
                cell_value = float(cell.value)
                if cell_value < 125:
                    cell.fill = yellow_fill
                elif 125 <= cell_value < 140:
                    cell.fill = light_blue_fill
            except (TypeError, ValueError):
                pass

    # 色を塗りつぶしたExcelファイルを保存
    wb.save(excel_path)

def process_juggler_data(html_path, output_csv_dir, excel_path, date):
    # Step 1: データを抽出してCSVに保存
    output_csv_path = os.path.join(output_csv_dir, f"slot_machine_data_{date}.csv")
    df_new = extract_data_and_save_to_csv(html_path, output_csv_path, date)

    # Step 2: フォルダー内のすべてのCSVに基づいて新しいExcelファイルを作成
    create_new_excel_with_all_data(output_csv_dir, excel_path)

    # Step 3: 色を塗り分け
    apply_color_fill_to_excel(excel_path)

    print(f"データ処理が完了し、{excel_path} に保存されました")


# ヘッダー
st.title("データ処理アプリケーション")
st.write("HTMLファイルからデータを抽出し、Excelファイルを生成します。")

# サイドバーでユーザー入力を受け取る
st.sidebar.header("入力パラメータ")
uploaded_html = st.sidebar.file_uploader("HTMLファイルをアップロード", type=["html", "htm", "txt"])
output_csv_dir = st.sidebar.text_input("CSVファイルの保存フォルダ名", "マイジャグラーV")
excel_file_name = st.sidebar.text_input("Excelファイル名", "マイジャグラーV_塗りつぶし済み.xlsx")
date_input = st.sidebar.date_input("日付を選択", datetime.today())

# 処理開始ボタンがクリックされたときの動作
if st.sidebar.button("処理開始"):
    if uploaded_html is not None:
        # アップロードされたファイルを保存
        html_path = os.path.join(".", uploaded_html.name)
        with open(html_path, "wb") as f:
            f.write(uploaded_html.getbuffer())

        # 出力ディレクトリを作成
        if not os.path.exists(output_csv_dir):
            os.makedirs(output_csv_dir)

        # 日付を文字列に変換
        date_str = date_input.strftime("%Y-%m-%d")

        # 一連の処理を実行
        try:
            process_juggler_data(html_path, output_csv_dir, excel_file_name, date_str)
            st.success(f"データ処理が完了し、{excel_file_name} に保存されました。")

            # GitHubにファイルをアップロード
            repo_name = "your-username/your-repo-name"  # リポジトリ名を指定
            commit_message = f"Add data for {date_str}"

            # CSVファイルのパス
            output_csv_path = os.path.join(output_csv_dir, f"slot_machine_data_{date_str}.csv")

            # CSVファイルのアップロード
            upload_file_to_github(output_csv_path, repo_name, f"data/csv/slot_machine_data_{date_str}.csv", commit_message)

            # Excelファイルのアップロード
            upload_file_to_github(excel_file_name, repo_name, f"data/excel/{excel_file_name}", commit_message)

            # ダウンロードボタンの表示
            with open(excel_file_name, "rb") as f:
                st.download_button(
                    label="生成されたExcelファイルをダウンロード",
                    data=f,
                    file_name=excel_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            if os.path.exists(output_csv_path):
                with open(output_csv_path, "rb") as f:
                    st.download_button(
                        label="生成されたCSVファイルをダウンロード",
                        data=f,
                        file_name=os.path.basename(output_csv_path),
                        mime="text/csv"
                    )
            else:
                st.warning("CSVファイルが見つかりませんでした。")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
    else:
        st.warning("HTMLファイルをアップロードしてください。")
