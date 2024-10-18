# 必要なモジュールのインポート
import os
import pandas as pd
import openpyxl
import pytz
from openpyxl.styles import PatternFill, Font
from bs4 import BeautifulSoup
import streamlit as st
from datetime import datetime, timedelta, timezone
from github import Github


# GitHubへのファイルアップロード関数
def upload_file_to_github(file_path, repo_name, file_name_in_repo, commit_message):
    try:
        # GitHubに認証
        g = Github(GITHUB_TOKEN)
        # リポジトリを取得
        repo = g.get_repo(repo_name)

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
        except Exception as e_inner:
            # ファイルが存在しない場合は新規作成
            repo.create_file(path, commit_message, content)
            st.info(f"{file_name_in_repo} を作成しました。")
    except Exception as e_outer:
        st.error(f"GitHubへのファイルアップロード中にエラーが発生しました: {e_outer}")
        st.error(f"詳細: {e_outer.args}")

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

# シークレットからGitHubトークンを取得
GITHUB_TOKEN = st.secrets["github"]["token"]

# ヘッダー
st.markdown(
    """
    <style>
    .main-title {
        font-size: 40px;
        font-weight: bold;
        color: #2C3E50;
        text-align: center;
    }
    .subtitle {
        font-size: 18px;
        color: #34495E;
        text-align: center;
    }
    .section-title {
        font-size: 24px;
        font-weight: bold;
        color: #2980B9;
        margin-top: 20px;
    }
    .instruction {
        font-size: 16px;
        color: #2C3E50;
    }
    </style>
    <div class="main-title">🎰 データ処理アプリケーション 🎰</div>
    <div class="subtitle">HTMLファイルからデータを抽出し、Excelファイルを生成します。</div>
    """, unsafe_allow_html=True
)

# 台データオンラインのリンク
st.markdown(
    """
    <div class="section-title">📊 メッセ武蔵境店 台データオンライン</div>
    <div class="instruction">
        台データはこちらのリンクからご確認ください： 
        <a href="https://daidata.goraggio.com/100686" target="_blank">メッセ武蔵境店 - 台データオンライン</a>
    </div>
    """, unsafe_allow_html=True
)

# HTML取得方法の説明
st.markdown(
    """
    <div class="section-title">🔍 HTMLの取得方法</div>
    <div class="instruction">
        1. リンク先のページにアクセス<br>
        2. 右上にあるメニューを押し、「その他」を選択<br>
        3. 「デベロッパーツール」を選択<br>
        4. 画面左上の1行目に表示される「<html>」を右クリック<br>
        5. 「copy」を選択し、「copy element」をクリック<br>
    </div>
    """, unsafe_allow_html=True
)

# HTML貼り付け時の注意点の説明
st.markdown(
    """
    <div class="section-title">📋 HTMLを貼り付ける方法</div>
    <div class="instruction">
        HTMLを以下のフォームに貼り付けたら、<strong>Ctrl + Enter</strong>を押して処理を開始してください。
    </div>
    """, unsafe_allow_html=True
)

# 日本時間の今日の日付を取得
japan_time_zone = pytz.timezone('Asia/Tokyo')
current_date_japan = datetime.now(japan_time_zone)

# サイドバーでユーザー入力を受け取る
st.sidebar.markdown(
    """
    <style>
    .sidebar-title {
        font-size: 20px;
        font-weight: bold;
        color: #2980B9;
    }
    .sidebar-section {
        font-size: 16px;
        color: #34495E;
    }
    .sidebar-input {
        font-size: 14px;
        margin-bottom: 10px;
    }
    .processing-button {
        background-color: #3498DB;
        color: white;
        padding: 10px;
        font-size: 16px;
        border-radius: 5px;
        text-align: center;
        cursor: pointer;
    }
    </style>
    """, unsafe_allow_html=True
)

st.sidebar.markdown('<div class="sidebar-title">📋 入力パラメータ</div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sidebar-section">HTMLファイルの入力方法を選択してください</div>', unsafe_allow_html=True)

# デフォルトで"HTMLを貼り付け"が選択されるように設定
input_option = st.sidebar.radio("HTMLの入力方法を選択", ('ファイルをアップロード', 'HTMLを貼り付け'), index=1)

if input_option == 'ファイルをアップロード':
    st.sidebar.markdown('<div class="sidebar-section">HTMLファイルをアップロードしてください。</div>', unsafe_allow_html=True)
    uploaded_html = st.sidebar.file_uploader("HTMLファイルをアップロード", type=["html", "htm", "txt"])
    html_content = None
else:
    st.sidebar.markdown('<div class="sidebar-section">HTMLを貼り付けてください。<br>貼り付け後に Ctrl + Enter <br>を押してください。</div>', unsafe_allow_html=True)
    html_content = st.sidebar.text_area("HTMLを貼り付け", height=300)
    uploaded_html = None

# "CSVファイルの保存フォルダ名"を固定（ユーザーが変更できないようにする）
st.sidebar.text_input("CSVファイルの保存フォルダ名", "マイジャグラーV", disabled=True)

# Excelファイル名の入力欄
excel_file_name = st.sidebar.text_input("Excelファイル名", "マイジャグラーV_塗りつぶし済み.xlsx")

# 日本時間の今日の日付をデフォルトに設定
date_input = st.sidebar.date_input("日付を選択", current_date_japan)

# 処理開始ボタンがクリックされたときの動作
if st.sidebar.button("処理開始"):
    # 日付確認のポップアップを表示
    confirm_date = st.sidebar.checkbox(f"選択した日付は {date_input} です。確認しましたか？")
    
    if confirm_date:
        if uploaded_html is not None or html_content:
            if uploaded_html is not None:
                html_path = os.path.join(".", uploaded_html.name)
                with open(html_path, "wb") as f:
                    f.write(uploaded_html.getbuffer())
            else:
                html_path = os.path.join(".", "uploaded_html.html")
                with open(html_path, "w", encoding="utf-8") as f:
                    f.write(html_content)

            if not os.path.exists("マイジャグラーV"):
                os.makedirs("マイジャグラーV")

            date_str = date_input.strftime("%Y-%m-%d")

            try:
                process_juggler_data(html_path, "マイジャグラーV", excel_file_name, date_str)
                st.success(f"データ処理が完了し、{excel_file_name} に保存されました。")

                repo_name = "yudai4452/data-processor-app"
                commit_message = f"Add data for {date_str}"

                output_csv_path = os.path.join("マイジャグラーV", f"slot_machine_data_{date_str}.csv")

                upload_file_to_github(output_csv_path, repo_name, f"マイジャグラーV/slot_machine_data_{date_str}.csv", commit_message)
                upload_file_to_github(excel_file_name, repo_name, f"{excel_file_name}", commit_message)

                st.markdown("---")  # 区切り線を追加

                # ダウンロードボタンをおしゃれに表示
                st.markdown(
                    f"""
                    <style>
                    .download-button {{
                        background-color: #2ECC71;
                        color: white;
                        padding: 10px;
                        font-size: 16px;
                        border-radius: 5px;
                        text-align: center;
                        cursor: pointer;
                    }}
                    </style>
                    """, unsafe_allow_html=True
                )

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
            st.warning("HTMLファイルをアップロードするか、HTMLを貼り付けてください。")
    else:
        st.warning("日付の確認を行ってください。")
