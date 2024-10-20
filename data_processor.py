# 必要なモジュールのインポート
import os
import pandas as pd
import openpyxl
import pytz
from openpyxl.styles import PatternFill, Font
from bs4 import BeautifulSoup
import streamlit as st
from datetime import datetime
from github import Github
import plotly.graph_objects as go 

# ダウンロード用にファイルを読み込む関数
def download_excel_file(excel_path):
    with open(excel_path, "rb") as f:
        return f.read()

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

# Excelファイルの読み込み関数（追加）
def load_excel_data(excel_path):
    # Excelファイルを読み込み
    df = pd.read_excel(excel_path, sheet_name="合成確率", engine="openpyxl", index_col=0)
    return df

# 台番号ごとの合成確率の推移をプロットする関数（追加）
def plot_synthetic_probabilities(df, selected_machine_number):
    # 選択された台番号の合成確率データを取得
    machine_data = df.loc[selected_machine_number].dropna()

    # プロットするためのデータ
    dates = machine_data.index
    probabilities = machine_data.values

    # Plotlyを使用してグラフを作成
    fig = go.Figure()

    # 合成確率の推移をプロット
    fig.add_trace(go.Scatter(x=dates, y=probabilities, mode='lines+markers', name=f'合成確率: {selected_machine_number}'))

    # グラフのタイトルとラベルを設定
    fig.update_layout(
        title=f"台番号 {selected_machine_number} の合成確率の推移",
        xaxis_title="日付",
        yaxis_title="合成確率",
        xaxis=dict(tickformat="%Y-%m-%d"),
        hovermode="x"
    )

    # Streamlitでグラフを表示
    st.plotly_chart(fig, use_container_width=True)

# Streamlit UIの定義
# 日本時間の今日の日付を取得
japan_time_zone = pytz.timezone('Asia/Tokyo')
current_date_japan = datetime.now(japan_time_zone)

# ヘッダー
st.markdown(
    """
    <style>
    .main-title {
        font-size: 40px;
        font-weight: bold;
        color: #34495E;  /* 濃いグレー */
        text-align: center;
        margin-bottom: 20px;  /* タイトルの下に余白を追加 */
    }
    </style>
    <div class="main-title">🎯 Juggler Data Manager 🎯</div>
    <div class="subtitle">HTMLファイルからデータを抽出し、Excelファイルを生成します。</div>
    """, unsafe_allow_html=True
)

# サイドバーでユーザー入力を受け取る
st.sidebar.markdown('<div class="sidebar-title">📋 入力パラメータ</div>', unsafe_allow_html=True)

# Excelファイル名の入力欄
excel_file_name = st.sidebar.text_input("Excelファイル名", "マイジャグラーV_塗りつぶし済み.xlsx", key="excel_file_name")

# 日本時間の今日の日付をデフォルトに設定
date_input = st.sidebar.date_input("日付を選択", current_date_japan, key="date_input")

# 日付確認のポップアップを表示
confirm_date = st.sidebar.checkbox(f"選択した日付は {date_input} です。確認しましたか？", key="confirm_date")

# 既存のExcelファイルをプロットするオプション
if os.path.exists(excel_file_name):
    st.sidebar.markdown('<div class="sidebar-section">既存のExcelファイルからプロット</div>', unsafe_allow_html=True)
    df_synthetic = load_excel_data(excel_file_name)
    machine_numbers = df_synthetic.index.tolist()
    selected_machine_number = st.sidebar.selectbox("台番号を選択してプロットする", machine_numbers, key="existing_excel_plot")
    if selected_machine_number:
        plot_synthetic_probabilities(df_synthetic, selected_machine_number)

# 処理開始ボタンがクリックされたときの動作
if st.sidebar.button("処理開始"):
    if confirm_date:
        # スクロールの位置を保持するためのJavaScriptコードを追加
        st.markdown(
            """
            <script>
            document.querySelector('button[aria-label="処理開始"]').addEventListener('click', function() {
                window.scrollTo(0, 0);
            });
            </script>
            """, unsafe_allow_html=True
        )

        st.success(f"データ処理が完了しました。")

        # ダウンロードボタンを常に表示する
        if os.path.exists(excel_file_name):
            with open(excel_file_name, "rb") as f:
                st.download_button(
                    label="生成されたExcelファイルをダウンロード",
                    data=f,
                    file_name=excel_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="excel_download_button"
                )

    else:
        st.warning("日付の確認を行ってください。")
